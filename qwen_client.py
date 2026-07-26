"""Stateless Ollama/Qwen client. Every call creates a brand-new conversation."""
import base64
import json
import socket
import urllib.error
import urllib.request


class QwenError(RuntimeError):
    pass


class QwenClient:
    def __init__(self, base_url="http://127.0.0.1:11434", model="", timeout=30, opener=None):
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = float(timeout)
        self.opener = opener or urllib.request.urlopen

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
        return {"model": self.model, "stream": False, "messages": messages}

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

    def chat(self, text, image=None, system_prompt=""):
        if not self.model.strip():
            raise QwenError("尚未選擇模型")
        data = self._request("/api/chat", self.build_payload(text, image, system_prompt))
        message = data.get("message")
        if not isinstance(message, dict) or not isinstance(message.get("content"), str):
            detail = str(data.get("error", ""))
            if "image" in detail.lower():
                raise QwenError("模型不支援圖片：{}".format(detail))
            raise QwenError("回應缺少 message.content")
        return message["content"]


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
