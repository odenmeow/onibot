import json
import importlib.util
import os
import sys
import tempfile
import time
import types
import unittest
from unittest import mock

import camera_capture

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

import front
from ai_monitor import AIMonitor, AlarmPlayer, is_timeout_error
from ai_image_processing import prepare_and_encode_ai_image
from camera_capture import CameraCapture, classify_device
from image_library import ImageLibrary, ImageLibraryFullError
from qwen_client import QwenClient, QwenError, choose_model


class FakeCapture:
    def __init__(self, opened=True, frame=None): self.opened, self.frame, self.released = opened, frame, False
    def isOpened(self): return self.opened
    def read(self): return (self.frame is not None, self.frame)
    def release(self): self.released = True


@unittest.skipUnless(importlib.util.find_spec("PIL"), "Pillow is an optional AI dependency")
class AIImageProcessingTests(unittest.TestCase):
    def test_enabled_settings_crop_and_resize_the_transmitted_copy(self):
        from PIL import Image
        source = Image.new("RGB", (400, 200), "white")

        encoded, metadata = prepare_and_encode_ai_image(source, {
            "enabled": True, "crop_enabled": True, "crop": [.25, 0, .75, 1],
            "resize_enabled": True, "target_width": 80, "target_height": 80,
            "resize_mode": "contain", "output_format": "png",
        }, "attachment")

        self.assertTrue(encoded.startswith(b"\x89PNG"))
        self.assertEqual(metadata["original_size"], (400, 200))
        self.assertEqual(metadata["cropped_size"], (200, 200))
        self.assertEqual(metadata["output_size"], (80, 80))

    def test_attachment_file_path_is_opened_and_encoded(self):
        from PIL import Image
        with tempfile.TemporaryDirectory() as directory:
            source = os.path.join(directory, "attachment.png")
            Image.new("RGB", (40, 20), "white").save(source)

            encoded, metadata = prepare_and_encode_ai_image(source, {
                "enabled": True, "output_format": "png",
            }, "attachment")

        self.assertTrue(encoded.startswith(b"\x89PNG"))
        self.assertEqual(metadata["original_size"], (40, 20))
        self.assertEqual(metadata["output_size"], (40, 20))

    def test_master_switch_prevents_crop_and_resize(self):
        from PIL import Image
        source = Image.new("RGB", (400, 200), "white")

        _encoded, metadata = prepare_and_encode_ai_image(source, {
            "enabled": False, "crop_enabled": True, "crop": [.25, 0, .75, 1],
            "resize_enabled": True, "target_width": 80, "target_height": 80,
        }, "attachment")

        self.assertEqual(metadata["cropped_size"], (400, 200))
        self.assertEqual(metadata["output_size"], (400, 200))


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

    def test_chat_keeps_model_loaded(self):
        client = QwenClient(model="vision:8b", keep_alive="1h")
        self.assertEqual(client.build_payload("hi")["keep_alive"], "1h")

    def test_consecutive_chats_keep_model_loaded_but_never_resend_history(self):
        responses = []
        for answer in ("first-answer", "second-answer"):
            response = mock.Mock()
            response.readline.side_effect = [json.dumps({
                "message": {"content": answer}, "done": True,
            }).encode("utf-8") + b"\n"]
            responses.append(response)
        opener = mock.Mock(side_effect=responses)
        client = QwenClient(model="vision:8b", keep_alive="1h", opener=opener)

        self.assertEqual(client.chat("first-question", image=b"first-image"), "first-answer")
        self.assertEqual(client.chat("second-question", image=b"second-image"), "second-answer")

        first = json.loads(opener.call_args_list[0].args[0].data)
        second = json.loads(opener.call_args_list[1].args[0].data)
        self.assertEqual(first["keep_alive"], "1h")
        self.assertEqual(second["keep_alive"], "1h")
        self.assertEqual([item["content"] for item in first["messages"]], ["first-question"])
        self.assertEqual([item["content"] for item in second["messages"]], ["second-question"])
        self.assertNotIn("first-question", json.dumps(second))
        self.assertNotIn("first-answer", json.dumps(second))

    def test_chat_payload_defaults_to_fast_bounded_streaming(self):
        payload = QwenClient(model="vision:8b").build_payload("hi")
        self.assertTrue(payload["stream"])
        self.assertFalse(payload["think"])
        self.assertEqual(payload["options"]["num_predict"], 1024)

    def test_chat_raises_when_stream_contains_only_blank_content(self):
        response = mock.Mock()
        response.readline.side_effect = [
            b'{"message":{"content":""},"done":false}\n',
            b'{"message":{"content":"  "},"done":true}\n',
        ]
        client = QwenClient(model="vision:8b", opener=mock.Mock(return_value=response))
        with self.assertRaisesRegex(QwenError, "message.content"):
            client.chat("hi")
        response.close.assert_called_once_with()

    def test_chat_upgrades_too_small_num_predict(self):
        payload = QwenClient(model="vision:8b", num_predict=8).build_payload("hi")
        self.assertEqual(payload["options"]["num_predict"], 1024)

    def test_load_config_upgrades_legacy_num_predict(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "front_config.json")
            with open(path, "w", encoding="utf-8") as stream:
                json.dump({"ai": {"num_predict": 8}}, stream)
            with mock.patch.object(front, "CONFIG_FILE", path):
                self.assertEqual(front.load_config()["ai"]["num_predict"], 1024)

    def test_chat_explains_thinking_exhausted_token_budget(self):
        response = mock.Mock()
        response.readline.side_effect = [
            b'{"message":{"content":"","thinking":"checking"},"done":false}\n',
            b'{"message":{"content":""},"done":true,"done_reason":"length"}\n',
        ]
        client = QwenClient(model="vision:8b", opener=mock.Mock(return_value=response))
        with self.assertRaisesRegex(QwenError, "token.*用完"):
            client.chat("hi")

    def test_chat_does_not_mislabel_thinking_only_stop_as_token_exhaustion(self):
        response = mock.Mock()
        response.readline.side_effect = [
            b'{"message":{"content":"","thinking":"checking"},"done":true,"done_reason":"stop"}\n',
        ]
        client = QwenClient(model="vision:8b", opener=mock.Mock(return_value=response))
        with self.assertRaisesRegex(QwenError, "只回傳思考內容"):
            client.chat("hi")

    def test_chat_collects_stream_and_closes_response(self):
        response = mock.Mock()
        response.readline.side_effect = [
            b'{"message":{"content":"O"},"done":false}\n',
            b'{"message":{"content":"K"},"done":true}\n',
        ]
        client = QwenClient(model="vision:8b", opener=mock.Mock(return_value=response))
        self.assertEqual(client.chat("hi"), "OK")
        response.close.assert_called_once_with()

    def test_chat_cancel_closes_response(self):
        response = mock.Mock()
        cancelled = mock.Mock(); cancelled.is_set.return_value = True
        client = QwenClient(model="vision:8b", opener=mock.Mock(return_value=response))
        with self.assertRaisesRegex(QwenError, "cancelled"):
            client.chat("hi", cancel_event=cancelled)
        response.close.assert_called_once_with()

    def test_preload_uses_empty_generate_request(self):
        response = mock.Mock(status=200)
        response.read.return_value = b"{}"
        opener = mock.Mock(return_value=response)
        client = QwenClient(model="vision:8b", keep_alive="30m", opener=opener)
        self.assertTrue(client.preload())
        request = opener.call_args.args[0]
        self.assertEqual(request.full_url, "http://127.0.0.1:11434/api/generate")
        self.assertEqual(json.loads(request.data), {
            "model": "vision:8b", "prompt": "", "stream": False,
            "keep_alive": "30m",
        })


class AIMonitorTests(unittest.TestCase):
    def test_monitor_waits_for_a_new_camera_frame_each_attempt(self):
        camera = mock.Mock(); camera.error = ""
        camera.latest_frame.return_value = object()
        packets = iter(((object(), 10, 100.0), (object(), 11, 101.0), (object(), 12, 102.0)))
        camera.latest_frame_after.side_effect = lambda *_args, **_kwargs: next(packets)
        client = mock.Mock(timeout=30)
        monitor = AIMonitor(camera, client, [{"enabled": True, "prompt": "p"}])
        answers = iter(("X", "O", None))
        def chat(*_args, **_kwargs):
            answer = next(answers)
            if answer is None:
                monitor._stop.set()
                raise RuntimeError("done")
            return answer
        client.chat.side_effect = chat
        encoded = mock.Mock(); encoded.tobytes.return_value = b"jpg"
        cv2 = mock.Mock(); cv2.imencode.return_value = (True, encoded)
        with mock.patch.dict("sys.modules", {"cv2": cv2}):
            monitor.start()
            deadline = time.time() + 1
            while client.chat.call_count < 2 and time.time() < deadline: time.sleep(.01)
            monitor.stop()
        self.assertEqual(camera.latest_frame_after.call_args_list[:2], [
            mock.call(-1, timeout=1.0), mock.call(10, timeout=1.0)])
        _, _, first = monitor.results.get_nowait()
        _, _, second = monitor.results.get_nowait()
        self.assertEqual((first["frame_sequence"], second["frame_sequence"]), (10, 11))

    def test_monitor_uses_updated_crop_settings_on_next_attempt(self):
        monitor = AIMonitor(mock.Mock(), mock.Mock(), [])
        monitor.update_image_settings({"crop_enabled": True, "crop": [.1, .2, .8, .9]})
        self.assertEqual(monitor.image_settings["crop"], [.1, .2, .8, .9])

    def test_qwen_timeout_is_classified(self):
        self.assertTrue(is_timeout_error(QwenError("API timeout")))
        self.assertFalse(is_timeout_error(QwenError("無法連線")))

    def test_delay_occurs_after_completed_attempt(self):
        camera = mock.Mock(); camera.latest_frame.return_value = object(); camera.error = ""
        client = mock.Mock(timeout=30); client.chat.side_effect = ["X", RuntimeError("done")]
        monitor = AIMonitor(camera, client, [{"enabled": True, "prompt": "p"}], after_answer_delay=.15)
        encoded = mock.Mock(); encoded.tobytes.return_value = b"jpg"
        cv2 = mock.Mock(); cv2.imencode.return_value = (True, encoded)
        started = time.monotonic()
        with mock.patch.dict("sys.modules", {"cv2": cv2}):
            monitor.start()
            while client.chat.call_count < 2 and time.monotonic() - started < 1: time.sleep(.01)
            monitor.stop()
        self.assertGreaterEqual(time.monotonic() - started, .14)

    def test_monitor_sends_shared_system_prompt(self):
        camera = mock.Mock(); camera.latest_frame.return_value = object(); camera.error = ""
        client = mock.Mock(timeout=30); client.chat.side_effect = ["X", RuntimeError("done")]
        monitor = AIMonitor(camera, client, [{"enabled": True, "prompt": "判斷畫面"}],
                            system_prompt="只回答 O 或 X")
        encoded = mock.Mock(); encoded.tobytes.return_value = b"jpg"
        cv2 = mock.Mock(); cv2.imencode.return_value = (True, encoded)
        with mock.patch.dict("sys.modules", {"cv2": cv2}):
            monitor.start()
            deadline = time.time() + 1
            while client.chat.call_count < 1 and time.time() < deadline: time.sleep(.01)
            monitor.stop()

        client.chat.assert_any_call("判斷畫面", image=b"jpg",
                                    system_prompt="只回答 O 或 X",
                                    cancel_event=monitor._stop)

    def test_builtin_alarm_does_not_require_sound_path(self):
        alarm = AlarmPlayer(sound_path="", sound_mode="system_alarm")
        with mock.patch.object(alarm, "_beep") as beep:
            self.assertTrue(alarm.play("error"))
            deadline = time.time() + 1
            while alarm._playing and time.time() < deadline: time.sleep(.01)
        beep.assert_called_once_with("error")


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
