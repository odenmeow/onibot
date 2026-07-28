"""Stateless Ollama/Qwen client. Every call creates a brand-new conversation."""
import base64
import json
import socket
import time
import urllib.error
import urllib.request

# Qwen3-VL can consume a few hundred internal reasoning tokens even when Ollama
# is asked not to expose thinking.  A 256-token cap can therefore end the
# generation before the model emits its (very short) answer.
MIN_NUM_PREDICT = 1024


class QwenError(RuntimeError):
    pass


class QwenClient:
    def __init__(self, base_url="http://127.0.0.1:11434", model="", timeout=30,
                 opener=None, keep_alive="30m", think=False, num_predict=MIN_NUM_PREDICT,
                 options=None, option_mode="legacy"):
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = float(timeout)
        self.opener = opener or urllib.request.urlopen
        self.keep_alive = str(keep_alive).strip() or "30m"
        self.think = bool(think)
        self.num_predict = max(MIN_NUM_PREDICT, int(num_predict))
        self.options = None if options is None else dict(options)
        self.option_mode = option_mode

    def build_payload(self, text, image=None, system_prompt=""):
        """Build one isolated turn; model keep-alive never carries chat history."""
        # Deliberately create this list locally on every call.  Ollama keeps the
        # model weights resident via keep_alive, but only these messages form
        # the conversation; no prior user/image/assistant turn is resent.
        messages = []
        if system_prompt.strip():
            messages.append({"role": "system", "content": system_prompt.strip()})
        user = {"role": "user", "content": str(text)}
        if image is not None:
            if isinstance(image, bytes):
                raw = image
            else:
                with open(image, "rb") as stream:
                    raw = stream.read()
            user["images"] = [base64.b64encode(raw).decode("ascii")]
        messages.append(user)
        payload = {"model": self.model, "stream": True, "messages": messages,
                   "keep_alive": self.keep_alive, "think": self.think}
        # Legacy callers retain the prior safe cap.  New configuration modes
        # send only explicitly composed options, leaving Modelfile sampling
        # defaults untouched in model_default mode.
        payload["options"] = ({"num_predict": self.num_predict} if self.options is None
                              else dict(self.options))
        if not payload["options"]: payload.pop("options")
        return payload

    def _request(self, path, payload=None):
        data = None if payload is None else json.dumps(payload).encode("utf-8")
        request = urllib.request.Request(self.base_url + path, data=data,
                                         headers={"Content-Type": "application/json"})
        try:
            response = self.opener(request, timeout=self.timeout)
            status = getattr(response, "status", 200)
            raw = response.read()
        except (socket.timeout, TimeoutError) as exc:
            raise QwenError("API timeout") from exc
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")
            raise QwenError("HTTP {}：{}".format(exc.code, detail)) from exc
        except (urllib.error.URLError, OSError) as exc:
            raise QwenError("無法連線：{}".format(exc)) from exc
        if not 200 <= status < 300:
            raise QwenError("HTTP {}".format(status))
        try:
            return json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeError) as exc:
            raise QwenError("API 回傳無效 JSON") from exc

    def test_connection(self):
        data = self._request("/api/tags")
        names = [x.get("name", "") for x in data.get("models", []) if isinstance(x, dict)]
        if self.model and not any(x == self.model or x.split(":")[0] == self.model for x in names):
            raise QwenError("模型不存在：{}".format(self.model))
        return names

    def chat(self, text, image=None, system_prompt="", cancel_event=None):
        if not self.model.strip():
            raise QwenError("尚未選擇模型")
        payload = self.build_payload(text, image, system_prompt)
        request = urllib.request.Request(self.base_url + "/api/chat",
            data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"})
        response, pieces, thinking_pieces = None, [], []
        done_reason, deadline = "", time.monotonic() + self.timeout
        try:
            response = self.opener(request, timeout=self.timeout)
            while True:
                if cancel_event is not None and cancel_event.is_set(): raise QwenError("API cancelled")
                if time.monotonic() >= deadline: raise QwenError("API timeout")
                raw = response.readline()
                if not raw: break
                data = json.loads(raw.decode("utf-8"))
                if data.get("error"):
                    detail = str(data["error"])
                    if "image" in detail.lower(): raise QwenError("模型不支援圖片：{}".format(detail))
                    raise QwenError(detail)
                message = data.get("message", {})
                content = message.get("content", "") if isinstance(message, dict) else ""
                if isinstance(content, str): pieces.append(content)
                thinking = message.get("thinking", "") if isinstance(message, dict) else ""
                if isinstance(thinking, str): thinking_pieces.append(thinking)
                if data.get("done_reason"): done_reason = str(data["done_reason"])
                if data.get("done"): break
        except (socket.timeout, TimeoutError) as exc:
            raise QwenError("API timeout") from exc
        except (ValueError, UnicodeError) as exc:
            raise QwenError("API 回傳無效 JSON") from exc
        except urllib.error.HTTPError as exc:
            raise QwenError("HTTP {}：{}".format(exc.code, exc.read().decode("utf-8", "replace"))) from exc
        except (urllib.error.URLError, OSError) as exc:
            raise QwenError("無法連線：{}".format(exc)) from exc
        finally:
            if response is not None: response.close()
        answer = "".join(pieces)
        if not answer.strip():
            if done_reason == "length":
                raise QwenError("模型輸出 token 已用完，尚未產生答案；請提高「最多輸出 token」")
            if "".join(thinking_pieces).strip():
                raise QwenError("模型只回傳思考內容，沒有產生答案；此次提問仍未包含先前對話")
            raise QwenError("Ollama 回應缺少答案（message.content）")
        return answer

    def preload(self):
        """Load the selected model without generating an answer."""
        if not self.model.strip():
            raise QwenError("尚未選擇模型")
        self._request("/api/generate", {
            "model": self.model, "prompt": "", "stream": False,
            "keep_alive": self.keep_alive,
        })
        return True


def choose_model(names, preferred=""):
    """Choose deterministically: saved model, vision-looking model, then first."""
    names = [str(name).strip() for name in names if str(name).strip()]
    if preferred:
        for name in names:
            if name == preferred or name.split(":")[0] == preferred:
                return name
    for name in names:
        lowered = name.lower()
        if any(token in lowered for token in ("-vl", "vision", "llava")):
            return name
    return names[0] if names else ""
