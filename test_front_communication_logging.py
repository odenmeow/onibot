import json
import sys
import threading
import types
import unittest
from unittest import mock

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import App, PiRequestError, PI_PORT


class _FakeConn:
    def __init__(self):
        self.sent = []

    def sendall(self, data):
        self.sent.append(data)


class _FakeConnFile:
    def __init__(self, line):
        self.line = line

    def readline(self):
        return self.line


class CommunicationLoggingTests(unittest.TestCase):
    def _new_app(self, response_line='{"status":"ok"}\n'):
        app = App.__new__(App)
        app.config = {"pi_host": "pi.test", "communication_log_enabled": True}
        app.offline_mode = False
        app.connected = True
        app.control_request_lock = threading.Lock()
        app.status_request_lock = threading.Lock()
        app.control_conn = _FakeConn()
        app.control_conn_file = _FakeConnFile(response_line)
        app.status_conn = _FakeConn()
        app.status_conn_file = _FakeConnFile(response_line)
        app.communication_log_enabled = True
        app.communication_log_queue = types.SimpleNamespace(put_nowait=lambda entry: app.logged.append(entry))
        app.logged = []
        app.current_name = ""
        app.ensure_connection = lambda channel="control", timeout=None: None
        app.close_connection = lambda channel=None, silent=False: None
        app._clear_channel_error = lambda channel: None
        app._set_channel_error = lambda channel, err: None
        app.set_connected = lambda connected, message="": setattr(app, "connected", connected)
        app._format_backend_error = lambda res: str(res)
        app._log_message = lambda *_args, **_kwargs: None
        app.set_backend_error = lambda *_args, **_kwargs: None
        app.update_gpio_polarity_from_response = lambda res: None
        app.write_text = lambda data: None
        app.set_status = lambda msg: None
        app._classify_request_error = App._classify_request_error.__get__(app, App)
        app._build_request_diag_message = App._build_request_diag_message.__get__(app, App)
        app._get_handshake_contract_version = lambda: "v1"
        app._enqueue_communication_log = App._enqueue_communication_log.__get__(app, App)
        app._communication_log_base = App._communication_log_base.__get__(app, App)
        app.request_pi = App.request_pi.__get__(app, App)
        return app

    def test_outgoing_request_log_is_enqueued_before_send(self):
        app = self._new_app()
        order = []
        app._enqueue_communication_log = lambda entry: (order.append(entry["direction"]), app.logged.append(entry))
        app.control_conn.sendall = lambda data: (order.append("sendall"), app.control_conn.sent.append(data))

        app.request_pi({"action": "ping", "value": "完整"}, write_response=False)

        self.assertEqual(order[:2], ["give", "sendall"])
        give = app.logged[0]
        self.assertEqual(give["direction"], "give")
        self.assertEqual(give["channel"], "control")
        self.assertEqual(give["action"], "ping")
        self.assertEqual(give["pi_host"], "pi.test")
        self.assertEqual(give["port"], PI_PORT)
        self.assertEqual(give["payload_json_text"], app.control_conn.sent[0].decode("utf-8"))
        self.assertEqual(give["payload_byte_length"], len(app.control_conn.sent[0]))

    def test_incoming_response_log_contains_raw_response_and_elapsed_time(self):
        raw = '{"status":"ok","timeline_runtime":{"state":"running","processed_count":2,"events_total":3,"runtime_version":4,"execution_round":5}}\n'
        app = self._new_app(response_line=raw)

        app.request_pi({"action": "status"}, channel="status", write_response=False)

        take = next(entry for entry in app.logged if entry["direction"] == "take")
        parsed = next(entry for entry in app.logged if entry["direction"] == "parsed")
        self.assertEqual(take["raw_response_text"], raw)
        self.assertEqual(take["raw_response_byte_length"], len(raw.encode("utf-8")))
        self.assertIn("elapsed_ms", take)
        self.assertEqual(parsed["parsed_status"], "ok")
        self.assertEqual(parsed["timeline_runtime.state"], "running")
        self.assertEqual(parsed["timeline_runtime.processed_count"], 2)
        self.assertEqual(parsed["timeline_runtime.events_total"], 3)
        self.assertEqual(parsed["timeline_runtime.runtime_version"], 4)
        self.assertEqual(parsed["timeline_runtime.execution_round"], 5)

    def test_timeout_error_log_includes_channel_action_timeout_ms(self):
        app = self._new_app()
        app.control_conn.sendall = lambda data: (_ for _ in ()).throw(TimeoutError("boom"))

        with self.assertRaises(PiRequestError):
            app.request_pi({"action": "start_task"}, timeout=0.25, write_response=False)

        error = next(entry for entry in app.logged if entry["direction"] == "error")
        self.assertEqual(error["channel"], "control")
        self.assertEqual(error["action"], "start_task")
        self.assertEqual(error["timeout_ms"], 250)
        self.assertEqual(error["exception_class"], "TimeoutError")
        self.assertEqual(error["retry_attempt_index"], 0)
        self.assertEqual(error["max_attempts"], 1)

    def test_logging_failure_does_not_break_request_pi(self):
        app = self._new_app()
        app.communication_log_queue = types.SimpleNamespace(put_nowait=mock.Mock(side_effect=RuntimeError("queue failed")))

        response = app.request_pi({"action": "ping"}, write_response=False)

        self.assertEqual(response["status"], "ok")

    def test_communication_log_toggle_skips_enqueue(self):
        app = self._new_app()
        app.communication_log_enabled = False

        app._enqueue_communication_log({"direction": "give", "action": "ping"})

        self.assertEqual(app.logged, [])

    def test_daily_log_toggle_skips_queue(self):
        app = App.__new__(App)
        app.daily_log_enabled = False
        app.log_enabled = False
        app.daily_log_queue = types.SimpleNamespace(put_nowait=mock.Mock())
        app._log_message = App._log_message.__get__(app, App)

        app._log_message("前端", "不應寫入")

        app.daily_log_queue.put_nowait.assert_not_called()


if __name__ == "__main__":
    unittest.main()
