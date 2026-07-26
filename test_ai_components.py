import json
import os
import tempfile
import time
import unittest
from unittest import mock

import camera_capture
from camera_capture import CameraCapture, classify_device
from image_library import ImageLibrary, ImageLibraryFullError
from qwen_client import QwenClient, QwenError, choose_model


class FakeCapture:
    def __init__(self, opened=True, frame=None): self.opened, self.frame, self.released = opened, frame, False
    def isOpened(self): return self.opened
    def read(self): return (self.frame is not None, self.frame)
    def release(self): self.released = True


class CameraTests(unittest.TestCase):
    def test_camera_names_distinguish_physical_integrated_and_virtual(self):
        self.assertEqual(classify_device("Logitech USB Camera"), "usb")
        self.assertEqual(classify_device("Integrated Camera"), "integrated")
        self.assertEqual(classify_device("OBS Virtual Camera"), "virtual")

    def test_stable_selection_prefers_id_then_name_then_index(self):
        devices = [
            {"device_id": "usb-a", "name": "Webcam", "index": 4, "device_type": "usb"},
            {"device_id": "obs", "name": "OBS Virtual Camera", "index": 0, "device_type": "virtual"},
        ]
        self.assertEqual(CameraCapture.select_device(devices, "usb-a", "", 0)["index"], 4)
        self.assertEqual(CameraCapture.select_device(devices, "missing", "Webcam", 0)["index"], 4)
        self.assertIsNone(CameraCapture.select_device(devices, "", "", 0))
    def test_missing_opencv_reports_dependency_error_without_thread(self):
        with mock.patch.object(camera_capture, "cv2", None):
            reader = CameraCapture(); reader.start()
        self.assertEqual(reader.state, "dependency_error")
        self.assertIn("OpenCV", reader.error); self.assertFalse(reader.running)

    def test_probe_releases_every_capture(self):
        captures = []
        def factory(index):
            cap = FakeCapture(index == 2); captures.append(cap); return cap
        self.assertEqual(CameraCapture.probe(4, factory), [2])
        self.assertTrue(all(cap.released for cap in captures))

    def test_open_failure_has_index_and_backend(self):
        reader = CameraCapture(3, lambda _i: FakeCapture(False), backend="dshow"); reader.start()
        while reader.running: time.sleep(.01)
        self.assertIn("3", reader.error); self.assertIn("dshow", reader.error)

    def test_configure_explicit_none_clears_requested_mode(self):
        reader = CameraCapture(width=1920, height=1080, fps=30, fourcc="MJPG")
        with mock.patch.object(reader, "stop", return_value=True), mock.patch.object(reader, "start", return_value=True):
            self.assertTrue(reader.configure(width=None, height=None, fps=None, fourcc=None))
        self.assertIsNone(reader.width); self.assertIsNone(reader.height)
        self.assertIsNone(reader.fps); self.assertIsNone(reader.fourcc)

    def test_configure_does_not_start_when_stop_times_out(self):
        reader = CameraCapture(width=640)
        with mock.patch.object(reader, "stop", return_value=False), mock.patch.object(reader, "start") as start:
            self.assertFalse(reader.configure(width=1920))
        start.assert_not_called(); self.assertEqual(reader.width, 640)


class QwenTests(unittest.TestCase):
    def test_choose_model_prefers_saved_then_vision(self):
        names = ["text:7b", "qwen3-vl:8b"]
        self.assertEqual(choose_model(names, "text:7b"), "text:7b")
        self.assertEqual(choose_model(names), "qwen3-vl:8b")

    def test_chat_rejects_blank_model_without_request(self):
        client = QwenClient(model="", opener=lambda *_a, **_k: self.fail("called"))
        with self.assertRaisesRegex(QwenError, "尚未選擇模型"): client.chat("hi")


class ImageLibraryTests(unittest.TestCase):
    def test_all_favorites_refuses_new_image(self):
        with tempfile.TemporaryDirectory() as directory:
            def writer(path, _frame):
                with open(path, "wb") as stream:
                    stream.write(b"x")
                return True
            library = ImageLibrary(directory, maximum=1, image_writer=writer)
            item = library.save(object()); library.set_favorite(item["id"])
            with self.assertRaises(ImageLibraryFullError): library.save(object())
            with open(library.manifest_path, encoding="utf-8") as stream: self.assertEqual(len(json.load(stream)), 1)


if __name__ == "__main__": unittest.main()
