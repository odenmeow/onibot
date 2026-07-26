"""Persistent bounded AI image library with atomic manifest writes."""
import json
import os
import time
import uuid

try:
    import cv2
except ImportError:
    cv2 = None


class ImageLibraryFullError(RuntimeError):
    pass


class ImageLibrary:
    def __init__(self, directory="saved_ai_images", maximum=1000, image_writer=None):
        self.directory = os.path.abspath(directory)
        self.maximum = int(maximum)
        self.manifest_path = os.path.join(self.directory, "manifest.json")
        self.image_writer = image_writer or (cv2.imwrite if cv2 else None)
        os.makedirs(self.directory, exist_ok=True)
        self.warnings = []
        self.items = self._load()

    def _load(self):
        try:
            with open(self.manifest_path, "r", encoding="utf-8") as stream:
                raw = json.load(stream)
        except FileNotFoundError:
            return []
        except (OSError, ValueError) as exc:
            self.warnings.append("圖片清單損壞：{}".format(exc))
            return []
        valid = []
        for item in raw if isinstance(raw, list) else []:
            path = item.get("path", "") if isinstance(item, dict) else ""
            if path and os.path.isfile(path):
                valid.append(item)
            else:
                self.warnings.append("略過不存在的圖片：{}".format(path or "未知"))
        return valid

    def _write(self):
        tmp = self.manifest_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as stream:
            json.dump(self.items, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(tmp, self.manifest_path)

    def save(self, frame, source="camera", note=""):
        while len(self.items) >= self.maximum:
            candidates = [x for x in self.items if not x.get("favorite", False)]
            if not candidates:
                raise ImageLibraryFullError("圖片庫已滿，1000 張圖片全部是我的最愛")
            oldest = min(candidates, key=lambda x: float(x.get("saved_at", 0)))
            self.delete(oldest["id"])
        image_id = uuid.uuid4().hex
        path = os.path.join(self.directory, image_id + ".jpg")
        if not self.image_writer or not self.image_writer(path, frame):
            raise RuntimeError("圖片寫入失敗")
        item = {"id": image_id, "path": path, "saved_at": time.time(),
                "favorite": False, "source": source, "note": note}
        self.items.append(item)
        self._write()
        return dict(item)

    def set_favorite(self, image_id, favorite=True):
        item = next((x for x in self.items if x.get("id") == image_id), None)
        if not item:
            raise KeyError(image_id)
        item["favorite"] = bool(favorite)
        self._write()

    def delete(self, image_id):
        item = next((x for x in self.items if x.get("id") == image_id), None)
        if not item:
            return False
        try:
            os.remove(item["path"])
        except FileNotFoundError:
            pass
        self.items.remove(item)
        self._write()
        return True

    def list(self, favorites_only=False):
        return [dict(x) for x in self.items if not favorites_only or x.get("favorite")]

