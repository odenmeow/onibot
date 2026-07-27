"""Stateless Ollama/Qwen client. Every call creates a brand-new conversation."""
import base64
import json
import socket
import time
import urllib.error
import urllib.request


class QwenError(RuntimeError):
    pass


class QwenClient:
    def __init__(self, base_url="http://127.0.0.1:11434", model="", timeout=30,
                 opener=None, keep_alive="30m", think=False, num_predict=8):
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = float(timeout)
        self.opener = opener or urllib.request.urlopen
        self.keep_alive = str(keep_alive).strip() or "30m"
        self.think = bool(think)
        self.num_predict = max(1, int(num_predict))

    def build_payload(self, text, image=None, system_prompt=""):
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
        return {"model": self.model, "stream": True, "messages": messages,
                "keep_alive": self.keep_alive, "think": self.think,
                "options": {"num_predict": self.num_predict}}

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
        response, pieces, deadline = None, [], time.monotonic() + self.timeout
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
        if not pieces: raise QwenError("回應缺少 message.content")
        return "".join(pieces)

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
