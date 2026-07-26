import json
import os
import tempfile
import time
import unittest
from unittest import mock

import camera_capture
from camera_capture import CameraCapture
from image_library import ImageLibrary, ImageLibraryFullError
from qwen_client import QwenClient, QwenError, choose_model


class FakeCapture:
    def __init__(self, opened=True, frame=None): self.opened, self.frame, self.released = opened, frame, False
    def isOpened(self): return self.opened
    def read(self): return (self.frame is not None, self.frame)
    def release(self): self.released = True


class CameraTests(unittest.TestCase):
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
