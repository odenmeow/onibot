"""Validation and composition of optional Ollama runtime parameters."""
import json

QWEN_OX_PRESET = {
    "num_predict": 256, "num_ctx": 4096, "temperature": 1.0,
    "top_k": 20, "top_p": 0.95, "seed": 0,
}
KNOWN_OPTIONS = frozenset({
    "num_predict", "num_ctx", "temperature", "top_k", "top_p", "min_p",
    "seed", "repeat_last_n", "repeat_penalty", "stop",
})
RANGES = {
    "num_predict": (32, 4096, int), "num_ctx": (512, 262144, int),
    "temperature": (0, 2, float), "top_k": (1, 200, int),
    "top_p": (0, 1, float), "min_p": (0, 1, float),
    "repeat_penalty": (0, 3, float),
}


def parse_extra_options(value):
    """Parse an advanced JSON object without silently discarding any keys."""
    if value is None or value == "": return {}
    data = json.loads(value) if isinstance(value, str) else value
    if not isinstance(data, dict): raise ValueError("進階 JSON 必須是 JSON object")
    return dict(data)


def validate_options(options):
    result = {}
    for key, value in dict(options or {}).items():
        if key in RANGES:
            low, high, converter = RANGES[key]
            if converter is int and (isinstance(value, bool) or int(value) != float(value)):
                raise ValueError("{} 必須是整數".format(key))
            value = converter(value)
            if not low <= value <= high: raise ValueError("{} 必須介於 {}～{}".format(key, low, high))
        elif key == "seed":
            if isinstance(value, bool): raise ValueError("seed 必須是整數")
            value = int(value)
        elif key == "repeat_last_n":
            value = int(value)
            if value < 0 and value != -1: raise ValueError("repeat_last_n 必須為 -1 或 0 以上整數")
        elif key == "stop":
            if isinstance(value, str): value = [value]
            if not isinstance(value, list) or not all(isinstance(x, str) for x in value):
                raise ValueError("stop sequences 必須是字串陣列")
        result[key] = value
    return result


def build_ollama_options(mode="model_default", enabled_options=None, extra_options=None,
                         num_predict=None):
    """Return (payload options, unknown keys); advanced JSON takes precedence."""
    if mode == "qwen_ox": options = dict(QWEN_OX_PRESET)
    elif mode == "custom": options = dict(enabled_options or {})
    else: options = {}
    if num_predict is not None: options["num_predict"] = int(num_predict)
    options.update(parse_extra_options(extra_options))
    options = validate_options(options)
    return options, sorted(set(options) - KNOWN_OPTIONS)
