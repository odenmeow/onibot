"""Shared, non-destructive image preparation used by previews and Ollama."""
import io
import math
import os


def normalize_crop_roi(first, second=None):
    values = first if second is None else (first[0], first[1], second[0], second[1])
    if len(values) != 4: raise ValueError("ROI 必須包含四個座標")
    x1, y1, x2, y2 = (max(0.0, min(1.0, float(x))) for x in values)
    left, right = sorted((x1, x2)); top, bottom = sorted((y1, y2))
    if right <= left or bottom <= top: raise ValueError("裁切範圍不可為空")
    return left, top, right, bottom


def map_canvas_point_to_image(x, y, image_width, image_height, scale=1.0,
                              offset=(0, 0), relative=False):
    """Undo a Canvas offset/zoom and clamp the point to the source image."""
    if scale <= 0: raise ValueError("scale 必須大於 0")
    ix = max(0.0, min(float(image_width), (float(x) - offset[0]) / scale))
    iy = max(0.0, min(float(image_height), (float(y) - offset[1]) / scale))
    return (ix / image_width, iy / image_height) if relative else (ix, iy)


def _pil(image):
    from PIL import Image
    if isinstance(image, Image.Image): return image.copy()
    if isinstance(image, (str, os.PathLike)):
        with Image.open(os.fspath(image)) as opened: return opened.copy()
    if isinstance(image, (bytes, bytearray, memoryview)):
        with Image.open(io.BytesIO(bytes(image))) as opened: return opened.copy()
    # OpenCV arrays are BGR/BGRA.
    if getattr(image, "ndim", 0) in (2, 3):
        import numpy as np
        data = np.asarray(image).copy()
        if data.ndim == 3 and data.shape[2] == 3: data = data[:, :, ::-1]
        elif data.ndim == 3 and data.shape[2] == 4: data = data[:, :, [2, 1, 0, 3]]
        return Image.fromarray(data)
    raise TypeError("不支援的圖片類型")


def prepare_ai_image(image, settings=None, source_type="camera"):
    """Crop, resize and pad a copy; return image plus dimensions metadata."""
    from PIL import Image, ImageOps
    settings = dict(settings or {}); result = _pil(image)
    # ``enabled`` is the master switch shown in the UI.  Settings dictionaries
    # created before that switch existed remain active for API compatibility.
    active = settings.get("enabled", True)
    original = result.size
    if active and settings.get("crop_enabled"):
        roi = normalize_crop_roi(settings.get("crop", (0, 0, 1, 1)))
        box = (round(roi[0] * result.width), round(roi[1] * result.height),
               round(roi[2] * result.width), round(roi[3] * result.height))
        result = result.crop(box)
    cropped = result.size
    if active and settings.get("resize_enabled"):
        target = (max(1, int(settings.get("target_width", 1280))),
                  max(1, int(settings.get("target_height", 720))))
        mode = settings.get("resize_mode", "contain")
        if not settings.get("allow_upscale", False):
            target = (min(target[0], result.width), min(target[1], result.height))
        if mode == "stretch": result = result.resize(target, Image.Resampling.LANCZOS)
        elif mode == "cover": result = ImageOps.fit(result, target, Image.Resampling.LANCZOS)
        else: result = ImageOps.contain(result, target, Image.Resampling.LANCZOS)
    if active and settings.get("align_qwen_grid"):
        size = (int(math.ceil(result.width / 32.0) * 32), int(math.ceil(result.height / 32.0) * 32))
        result = ImageOps.pad(result, size, method=Image.Resampling.LANCZOS, color=(0, 0, 0), centering=(.5, .5))
    metadata = {"source_type": source_type, "original_size": original,
                "cropped_size": cropped, "output_size": result.size,
                "pixel_count": result.width * result.height,
                "pixel_ratio": result.width * result.height / float(original[0] * original[1])}
    return result, metadata


def encode_ai_image(image, format_name="jpeg", jpeg_quality=95):
    result = _pil(image); name = str(format_name).upper()
    name = "JPEG" if name in ("JPG", "JPEG") else "PNG"
    if name == "JPEG" and result.mode not in ("RGB", "L"): result = result.convert("RGB")
    output = io.BytesIO(); kwargs = {"quality": int(jpeg_quality)} if name == "JPEG" else {}
    result.save(output, format=name, **kwargs)
    return output.getvalue()


def prepare_and_encode_ai_image(image, settings=None, source_type="camera"):
    settings = dict(settings or {})
    prepared, metadata = prepare_ai_image(image, settings, source_type)
    data = encode_ai_image(prepared, settings.get("output_format", "jpeg"), settings.get("jpeg_quality", 95))
    metadata["format"] = "JPEG" if settings.get("output_format", "jpeg").lower() in ("jpeg", "jpg") else "PNG"
    metadata["bytes"] = len(data)
    return data, metadata
