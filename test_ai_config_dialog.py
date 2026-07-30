import queue
import os
import tempfile
import time
import unittest
from types import SimpleNamespace
from unittest import mock

from ai_config_dialog import AIConfigDialog, QUESTION_HISTORY_LIMIT


class FakeText:
    def __init__(self, value=""):
        self.value = value
        self.state = None

    def config(self, **kwargs):
        self.state = kwargs.get("state", self.state)

    def delete(self, _start, _end):
        self.value = ""

    def get(self, _start, _end):
        return self.value

    def insert(self, _where, value):
        self.value += value

    def see(self, _where):
        pass


class FakeTree:
    def __init__(self):
        self.rows = []

    def get_children(self):
        return [row[0] for row in self.rows]

    def delete(self, row):
        self.rows = [item for item in self.rows if item[0] != row]

    def insert(self, _parent, _where, iid, values):
        self.rows.append((iid, values))

    def identify_region(self, _x, _y):
        return "cell"

    def identify_column(self, x):
        if x < 100: return "#1"
        if x < 200: return "#2"
        if x < 300: return "#3"
        return "#4"

    def identify_row(self, _y):
        return self.rows[0][0] if self.rows else ""


class FakeColumnTree:
    def __init__(self):
        self.columns = {}

    def column(self, name, **options):
        self.columns[name] = options


class FakeLabel:
    def __init__(self): self.options = {}
    def config(self, **options): self.options.update(options)


class AIConfigDialogTests(unittest.TestCase):
    def test_image_library_columns_can_shrink_without_hiding_source(self):
        tree = FakeColumnTree()

        AIConfigDialog._configure_library_columns(tree)

        self.assertEqual(tree.columns["favorite"], {"width": 45, "minwidth": 40, "stretch": False})
        self.assertEqual(tree.columns["time"]["minwidth"], 75)
        self.assertEqual(tree.columns["source"]["minwidth"], 45)
        self.assertTrue(tree.columns["source"]["stretch"])

    def test_sash_positions_capture_all_vertical_dividers(self):
        paned = SimpleNamespace(
            panes=lambda: ("one", "two", "three"),
            sash_coord=lambda index: ((100, 140), (100, 420))[index],
        )

        self.assertEqual(AIConfigDialog._sash_positions(paned), [140, 420])

    def test_save_ui_layout_does_not_validate_or_replace_ai_settings(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"timeout": "draft-invalid-value"}}
        dialog._save_ai_layout = mock.Mock()
        dialog.on_save = mock.Mock()
        dialog._append = mock.Mock()

        dialog.save_ui_layout()

        dialog._save_ai_layout.assert_called_once_with()
        dialog.on_save.assert_called_once_with(dialog.config)
        self.assertEqual(dialog.config["ai"]["timeout"], "draft-invalid-value")
        dialog._append.assert_called_once_with("系統", "UI 配置已保存")

    def test_clear_draft_only_clears_question_and_keeps_history(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.user_text = FakeText("提問")
        dialog.response = FakeText("console")
        dialog.config = {"ai": {"question_history": [{"answer": "old"}]}}
        dialog._save = mock.Mock()

        dialog.clear_draft()

        self.assertEqual(dialog.user_text.value, "")
        self.assertEqual(dialog.response.value, "console")
        self.assertEqual(dialog.config["ai"]["question_history"], [{"answer": "old"}])

    def test_clear_console_only_clears_console_and_keeps_history(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.user_text = FakeText("提問")
        dialog.response = FakeText("console")
        dialog.config = {"ai": {"question_history": [{"answer": "old"}]}}

        dialog.clear_console()

        self.assertEqual(dialog.response.value, "")
        self.assertEqual(dialog.response.state, "disabled")
        self.assertEqual(dialog.user_text.value, "提問")
        self.assertEqual(dialog.config["ai"]["question_history"], [{"answer": "old"}])

    def test_append_adds_timestamp_and_uses_explicit_time(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog); dialog.response = FakeText()
        dialog._append("系統", "完成", 0)
        expected = time.strftime("%H:%M:%S", time.localtime(0))
        self.assertEqual(dialog.response.value, "[{}] 系統：完成\n\n".format(expected))
        self.assertRegex(dialog.response.value, r"^\[\d{2}:\d{2}:\d{2}\]")

    def test_monitor_result_uses_ended_at(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.camera = SimpleNamespace(actual={}, width=None, height=None, fps=None, fallback={}, backend="dshow", state="connected", error="")
        dialog.fourcc = SimpleNamespace(get=lambda: "自動")
        dialog.camera_status = SimpleNamespace(config=lambda **_kwargs: None)
        dialog.mode = SimpleNamespace(get=lambda: "manual")
        dialog.camera_view_state = "docked"; dialog.show_latest = mock.Mock()
        results = queue.Queue(); results.put((None, "answer", {"text": "O", "ended_at": 1234}))
        dialog.monitor = SimpleNamespace(results=results)
        dialog._current_device = lambda: None
        dialog._record_history = mock.Mock()
        dialog._append = mock.Mock()
        dialog._schedule = lambda *_args: None

        dialog._poll_preview()

        dialog._append.assert_called_once_with("AI", "O", 1234)

    def test_monitor_o_opens_its_saved_capture_automatically(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {}}
        dialog.camera = SimpleNamespace(actual={}, width=None, height=None, fps=None, fallback={}, backend="dshow", state="connected", error="")
        dialog.fourcc = SimpleNamespace(get=lambda: "自動")
        dialog.camera_status = SimpleNamespace(config=lambda **_kwargs: None)
        dialog.mode = SimpleNamespace(get=lambda: "manual")
        dialog.camera_view_state = "hidden"
        result = {"text": "O", "ended_at": 1234, "path": "/history/detected.jpg"}
        results = queue.Queue(); results.put((None, "answer", result))
        dialog.monitor = SimpleNamespace(results=results)
        dialog._current_device = lambda: None
        dialog._record_history = mock.Mock(); dialog._append = mock.Mock()
        dialog._show_detected_history_image = mock.Mock()
        dialog._schedule = lambda *_args: None

        dialog._poll_preview()

        dialog._show_detected_history_image.assert_called_once_with(result)

    def test_new_detected_capture_reuses_image_viewer(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog._open_image_viewer = mock.Mock()
        with mock.patch("ai_config_dialog.os.path.isfile", return_value=True):
            item = {"path": "/history/new.jpg", "filename": "new.jpg"}
            dialog._show_detected_history_image(item)

        dialog._open_image_viewer.assert_called_once_with(
            "history", path="/history/new.jpg", title="偵測到 O", metadata=item)

    def test_space_in_viewer_stops_alarm(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.alarm = SimpleNamespace(stop=mock.Mock())
        dialog.viewer_status = FakeLabel()

        result = dialog._stop_alarm_from_viewer()

        dialog.alarm.stop.assert_called_once_with()
        self.assertEqual(result, "break")
        self.assertIn("鬧鐘已停止", dialog.viewer_status.options["text"])

    def test_visible_preview_refreshes_even_in_manual_mode(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.camera = SimpleNamespace(actual={}, width=None, height=None, fps=None, fallback={}, backend="dshow", state="connected", error="")
        dialog.fourcc = SimpleNamespace(get=lambda: "自動")
        dialog.camera_status = SimpleNamespace(config=lambda **_kwargs: None)
        dialog.mode = SimpleNamespace(get=lambda: "manual")
        dialog.camera_view_state = "detached"; dialog.monitor = None
        dialog._current_device = lambda: None
        dialog.show_latest = mock.Mock(); dialog._schedule = lambda *_args: None

        dialog._poll_preview()

        dialog.show_latest.assert_called_once_with()

    def test_history_keeps_latest_thousand_and_displays_newest_first(self):
        old_history = [{"history_id": str(i), "ended_at": i, "answer": str(i)} for i in range(QUESTION_HISTORY_LIMIT + 5)]
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"question_history": old_history}}
        dialog.on_save = mock.Mock(); dialog.history_tree = FakeTree(); dialog._history_by_id = {}

        dialog._record_history({"history_id": "new", "ended_at": 9999, "answer": "new"})

        history = dialog.config["ai"]["question_history"]
        self.assertEqual(len(history), QUESTION_HISTORY_LIMIT)
        self.assertEqual(history[-1]["history_id"], "new")
        self.assertEqual(dialog.history_tree.rows[0][0], "new")
        self.assertEqual(dialog.history_tree.rows[0][1][0], "☐")
        self.assertEqual(len(dialog.history_tree.rows), QUESTION_HISTORY_LIMIT)

    def test_history_error_checkbox_persists_human_feedback(self):
        item = {"history_id": "answer-1", "ended_at": 1, "answer": "O"}
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"question_history": [item]}}
        dialog.on_save = mock.Mock(); dialog.history_tree = FakeTree(); dialog._history_by_id = {}
        dialog._refresh_history()

        result = dialog._on_history_click(SimpleNamespace(x=10, y=5))

        self.assertEqual(result, "break")
        self.assertTrue(item["judgment_error"])
        self.assertEqual(dialog.history_tree.rows[0][1][0], "☑")
        dialog.on_save.assert_called_once_with(dialog.config)

    def test_history_click_outside_checkbox_does_not_change_feedback(self):
        item = {"history_id": "answer-1", "ended_at": 1, "answer": "O"}
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"question_history": [item]}}
        dialog.on_save = mock.Mock(); dialog.history_tree = FakeTree(); dialog._history_by_id = {}
        dialog._refresh_history()

        dialog._on_history_click(SimpleNamespace(x=350, y=5))

        self.assertNotIn("judgment_error", item)
        dialog.on_save.assert_not_called()

    def test_legacy_malformed_history_loads_without_error(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"question_history": {"legacy": True}}}
        dialog.history_tree = FakeTree(); dialog._history_by_id = {"stale": {}}
        dialog._refresh_history()
        self.assertEqual(dialog.history_tree.rows, [])
        self.assertEqual(dialog._history_by_id, {})

    def test_selected_history_exports_only_image_by_tag_and_result(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        with tempfile.TemporaryDirectory() as workspace:
            source = os.path.join(workspace, "capture.JPG")
            with open(source, "wb") as stream: stream.write(b"image")
            dialog._history_by_id = {
                "row-1": {"history_id": "row-1", "tag": "測試/範本", "answer": "O",
                          "filename": "capture.JPG", "path": source, "prompt": "請判斷這張圖"}
            }
            review_root = os.path.join(workspace, "ReviewFolder")

            exported, skipped, returned_root = dialog._export_history_rows(["row-1"], review_root)

            target = os.path.join(review_root, "測試_範本", "o")
            self.assertEqual(exported, 1)
            self.assertEqual(skipped, [])
            self.assertEqual(returned_root, review_root)
            with open(os.path.join(target, "capture.jpg"), "rb") as stream:
                self.assertEqual(stream.read(), b"image")
            self.assertFalse(os.path.exists(os.path.join(target, "capture.txt")))

    def test_started_and_endedat_include_every_history_row_between_them(self):
        history = [
            {"history_id": "before"},
            {"history_id": "start", "export_started": True},
            {"history_id": "middle-o", "answer": "O"},
            {"history_id": "middle-x", "answer": "X"},
            {"history_id": "end", "export_endedat": True},
            {"history_id": "after"},
        ]
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.config = {"ai": {"question_history": history}}

        self.assertEqual(dialog._marked_history_range(),
                         ["start", "middle-o", "middle-x", "end"])

    @mock.patch("ai_config_dialog.subprocess.Popen")
    def test_open_review_folder_creates_and_opens_directory(self, popen):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.window = mock.Mock()
        with tempfile.TemporaryDirectory() as workspace, \
                mock.patch("ai_config_dialog.__file__", os.path.join(workspace, "ai_config_dialog.py")), \
                mock.patch("ai_config_dialog.sys.platform", "linux"):
            dialog.open_review_folder()

            review_root = os.path.join(workspace, "ReviewFolder")
            self.assertTrue(os.path.isdir(review_root))
            popen.assert_called_once_with(["xdg-open", review_root])

    def test_judgment_error_is_exported_only_to_directional_wrong_folder(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        with tempfile.TemporaryDirectory() as workspace:
            source = os.path.join(workspace, "mistake.jpg")
            with open(source, "wb") as stream: stream.write(b"image")
            dialog._history_by_id = {
                "wrong-o": {"history_id": "wrong-o", "tag": "測試", "answer": "O",
                            "judgment_error": True, "filename": "mistake.jpg",
                            "path": source, "prompt": "prompt"}
            }

            exported, skipped, _ = dialog._export_history_rows(["wrong-o"], workspace)

            self.assertEqual((exported, skipped), (1, []))
            self.assertFalse(os.path.isfile(os.path.join(workspace, "測試", "o", "mistake.jpg")))
            wrong = os.path.join(workspace, "測試", "wrong", "seemsXasO")
            self.assertTrue(os.path.isfile(os.path.join(wrong, "mistake.jpg")))
            self.assertFalse(os.path.exists(os.path.join(wrong, "mistake.txt")))

    def test_history_export_skips_missing_images_and_non_ox_results(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog._history_by_id = {
            "timeout": {"answer": "", "error": "AI 回答逾時", "filename": "timeout.jpg"},
            "missing": {"answer": "X", "path": "/missing/image.jpg", "filename": "missing.jpg"},
        }
        with tempfile.TemporaryDirectory() as review_root:
            exported, skipped, _root = dialog._export_history_rows(["timeout", "missing"], review_root)

        self.assertEqual(exported, 0)
        self.assertEqual(len(skipped), 2)
        self.assertIn("結果不是 O/X", skipped[0])
        self.assertIn("圖片不存在", skipped[1])

    def test_manual_history_filename_has_manual_timestamp_prefix(self):
        filename = AIConfigDialog._manual_filename(0)
        expected_date = time.strftime("%Y%m%d_%H%M%S", time.localtime(0))
        self.assertEqual(filename, "manual_{}_000.jpg".format(expected_date))

    def test_manual_question_is_added_to_shared_history(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.profiles = [{"id": "profile-1", "name": "測試配置", "prompt": "配置內容", "enabled": True}]
        dialog.user_text = FakeText("手動問題")
        dialog.system = FakeText("系統提示")
        dialog.selected_image = "/source/image.png"
        dialog.config = {"ai": {"question_history": []}}
        dialog._save = mock.Mock()
        dialog._append = mock.Mock()
        dialog._archive_manual_image = mock.Mock(return_value=("manual_20260727_120000_000.jpg", "/history/manual.jpg"))
        dialog._client = mock.Mock(return_value=SimpleNamespace(chat=mock.Mock(return_value="O")))
        dialog._worker = lambda _operation, work, success, _failure: success(work())
        dialog._record_history = mock.Mock()

        dialog.send_test()

        history = dialog._record_history.call_args.args[0]
        self.assertEqual(history["mode"], "manual")
        self.assertEqual(history["filename"], "manual_20260727_120000_000.jpg")
        self.assertEqual(history["answer"], "O")
        self.assertEqual(history["tag"], "測試配置")
        self.assertIn("手動問題", history["prompt"])

    @mock.patch("ai_config_dialog.prepare_and_encode_ai_image")
    def test_manual_attachment_does_not_reuse_persisted_crop(self, prepare):
        prepare.return_value = (b"processed", {
            "original_size": (1516, 837), "cropped_size": (1516, 837),
            "output_size": (1280, 707), "format": "PNG", "bytes": 9,
        })
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.profiles = [{"id": "profile-1", "name": "測試", "prompt": "判斷", "enabled": True}]
        dialog.user_text = FakeText("問題"); dialog.system = FakeText()
        dialog.selected_image = "/source/new-image.png"
        dialog.attachment_crop_roi = None
        dialog.config = {"ai": {"attachment_ai_image": {
            "enabled": True, "crop_enabled": True, "crop": [.1, .1, .9, .9],
            "resize_enabled": True, "target_width": 1280, "target_height": 720,
        }}}
        dialog._save = mock.Mock(); dialog._append = mock.Mock()
        dialog._archive_manual_image = mock.Mock(return_value=("manual.png", "/history/manual.png"))
        dialog._client = mock.Mock(return_value=SimpleNamespace(chat=mock.Mock(return_value="O")))
        dialog._worker = lambda _operation, work, success, _failure: success(work())
        dialog._record_history = mock.Mock()

        dialog.send_test()

        sent_settings = prepare.call_args.args[1]
        self.assertFalse(sent_settings["crop_enabled"])
        self.assertEqual(sent_settings["crop"], [.1, .1, .9, .9])
        self.assertTrue(any("未裁切" in call.args[1] for call in dialog._append.call_args_list))

    def test_clear_attachment_resets_current_crop_and_status(self):
        dialog = AIConfigDialog.__new__(AIConfigDialog)
        dialog.selected_image = "/source/image.png"
        dialog.attachment_crop_roi = (.1, .1, .9, .9)
        dialog.image_info = FakeLabel(); dialog.attachment_preview = FakeLabel()
        dialog.attachment_crop_status = FakeLabel()

        dialog.clear_image()

        self.assertIsNone(dialog.attachment_crop_roi)
        self.assertIn("未套用", dialog.attachment_crop_status.options["text"])

    def test_viewer_zoom_keeps_pixel_below_pointer_fixed(self):
        scale, offset = AIConfigDialog._zoom_at(
            0.5, (100, 50), (400, 250), 2, minimum=0.25)

        self.assertEqual(scale, 1.0)
        self.assertEqual(offset, (-200, -150))
        before = ((400 - 100) / 0.5, (250 - 50) / 0.5)
        after = ((400 - offset[0]) / scale, (250 - offset[1]) / scale)
        self.assertEqual(after, before)

    def test_viewer_zoom_respects_fit_and_maximum_scale(self):
        zoomed_out = AIConfigDialog._zoom_at(1, (0, 0), (20, 20), 0.1, minimum=0.5)
        zoomed_in = AIConfigDialog._zoom_at(4, (0, 0), (20, 20), 4, minimum=0.5)

        self.assertEqual(zoomed_out[0], 0.5)
        self.assertEqual(zoomed_in[0], 8.0)


if __name__ == "__main__":
    unittest.main()
