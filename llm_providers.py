"""llm_providers.py - the free AI services the agent can think with.

    gemini      Google Gemini (the default; uses Google's own SDK, see agent.py)
    groq        Groq - fast open models (gpt-oss), no credit card
    openrouter  OpenRouter - a changing set of free models from many companies
    mistral     Mistral - "Free mode", phone number needed, no credit card
    ollama      Ollama - runs a model on your own computer: no key, no limits

Groq, OpenRouter and Mistral speak the OpenAI "chat completions" format;
Ollama is used through its own /api/chat endpoint, because only that one lets
us ask for a context window big enough for the agent (Ollama's default on
most laptops is 4,096 tokens, and it silently cuts off the start of longer
requests - which would drop the agent's rules and tools).

Standard library only: nothing to install beyond requirements.txt.

To add another OpenAI-compatible service, add a ``Provider`` to
``PROVIDERS`` (base URL, default model, key environment variable, key page).
"""
import json
import os
import re
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass


@dataclass(frozen=True)
class Provider:
    id: str                    # used on the command line: --provider groq
    label: str                 # "Groq"
    name: str                  # short name used in messages: "Groq didn't accept..."
    company: str               # who checks the key: "Checking it with Groq..."
    kind: str                  # "gemini", "openai" (chat completions) or "ollama"
    default_model: str
    model_env: str             # environment variable that overrides the model
    base_url: str = ""
    base_url_env: str = ""     # environment variable that overrides base_url
    key_envs: tuple = ()       # environment variables that may hold the key
    key_url: str = ""          # where to get a key (or, for Ollama, the app)
    key_steps: tuple = ()      # printed before the key prompt; {url} = key_url
    site: str = ""             # where the user can see the key again
    check_path: str = ""       # GET this to test a key (must need the key)
    free: str = ""             # one line on the free allowance
    privacy: str = ""          # one line on what happens to the data
    max_chars: int = 0         # keep each request below this size (0 = no limit)

    @property
    def needs_key(self):
        return bool(self.key_envs)


PROVIDERS = {p.id: p for p in [
    Provider(
        id="gemini", label="Google Gemini", name="Gemini", company="Google",
        kind="gemini", default_model="gemini-flash-latest", model_env="GEMINI_MODEL",
        key_envs=("GEMINI_API_KEY", "GOOGLE_API_KEY"),
        key_url="https://aistudio.google.com/apikey",
        key_steps=("Open {url} and sign in with a Google account",
                   "Click 'Create API key', then the copy button next to the new key"),
        site="AI Studio",
        free="Daily limits (shown in AI Studio). No credit card.",
        privacy=("On the free tier Google may use what the agent sends to improve "
                 "its products."),
    ),
    Provider(
        id="groq", label="Groq", name="Groq", company="Groq",
        kind="openai", default_model="openai/gpt-oss-120b", model_env="GROQ_MODEL",
        base_url="https://api.groq.com/openai/v1", base_url_env="GROQ_BASE_URL",
        key_envs=("GROQ_API_KEY",), key_url="https://console.groq.com/keys",
        key_steps=("Open {url} and sign in (a Google or GitHub account works)",
                   "Click 'Create API Key', type any name, then copy the key"),
        site="the Groq console", check_path="/models",
        free="1,000 requests a day, 30 a minute, 8,000 tokens a minute. No credit card.",
        privacy=("Groq doesn't train on what the agent sends and doesn't keep it "
                 "(except up to 30 days if it investigates abuse)."),
        # 8,000 tokens a minute, counting the prompt AND the requested answer
        # length (GROQ_MAX_OUTPUT): keep each request well below that
        max_chars=16_000,
    ),
    Provider(
        id="openrouter", label="OpenRouter (free models)", name="OpenRouter",
        company="OpenRouter", kind="openai",
        # a router that picks one of the current free models that can use tools
        default_model="openrouter/free", model_env="OPENROUTER_MODEL",
        base_url="https://openrouter.ai/api/v1", base_url_env="OPENROUTER_BASE_URL",
        key_envs=("OPENROUTER_API_KEY",), key_url="https://openrouter.ai/settings/keys",
        key_steps=("Open {url} and sign in",
                   "Click 'Create API key' (leave the credit limit empty), then copy the key",
                   "Free models only work after you allow them at "
                   "https://openrouter.ai/settings/privacy (see the note above)"),
        site="OpenRouter", check_path="/key",
        free=("50 requests a day, 20 a minute (1,000 a day after a one-time $10 "
              "top-up). No credit card."),
        privacy=("Free models are run by other companies that may keep and train "
                 "on what the agent sends; you have to allow that in OpenRouter's "
                 "privacy settings."),
        max_chars=60_000,
    ),
    Provider(
        id="mistral", label="Mistral", name="Mistral", company="Mistral",
        kind="openai", default_model="mistral-small-latest", model_env="MISTRAL_MODEL",
        base_url="https://api.mistral.ai/v1", base_url_env="MISTRAL_BASE_URL",
        key_envs=("MISTRAL_API_KEY",), key_url="https://console.mistral.ai/api-keys",
        key_steps=("Open {url} and sign up (a phone number is needed, no credit card)",
                   "Click 'Create new key', then copy it (it is shown only once)"),
        site="the Mistral console", check_path="/models",
        free=("A monthly allowance in \"Free mode\" (your limits are shown in the "
              "Mistral console). No credit card; a phone number is needed."),
        privacy=("In Free mode Mistral may train on what the agent sends unless "
                 "you turn that off in its privacy settings."),
        max_chars=60_000,
    ),
    Provider(
        id="ollama", label="Ollama (on your computer)", name="Ollama", company="Ollama",
        kind="ollama", default_model="qwen3:8b", model_env="OLLAMA_MODEL",
        base_url="http://127.0.0.1:11434", base_url_env="OLLAMA_HOST",
        key_url="https://ollama.com/download",
        free=("No key and no limits: it runs on your own computer. The default "
              "model (qwen3:8b) needs about 8 GB of free memory and is slow "
              "without a graphics card."),
        privacy="Nothing leaves your computer.",
        max_chars=40_000,
    ),
]}

OLLAMA_NUM_CTX = 16_384        # tokens; room for max_chars plus the answer
GROQ_MAX_OUTPUT = 1_500        # tokens per answer (Groq counts it against the limit)
BACKOFF = (2, 5, 10)           # seconds between retries after busy errors
MAX_WAIT = 60                  # wait for a rate limit only if it resets this soon


class ProviderError(Exception):
    """An error from a provider's API (``status`` = HTTP status) or from the
    network (``status`` None)."""

    def __init__(self, provider, status, message, code="", retry_after=None,
                 network=False, refused=False, timeout=False):
        super().__init__(f"{provider} {status or 'network'}: {message}")
        self.provider = provider          # Provider.id
        self.status = status
        self.message = message or ""
        self.code = str(code or "")
        self.retry_after = retry_after    # seconds, if the provider said
        self.network = network
        self.refused = refused            # nothing is listening (Ollama not running)
        self.timeout = timeout

    @property
    def text(self):
        return f"{self.code} {self.message}".lower()


def is_key_error(exc):
    """The provider rejected the API key itself."""
    if not isinstance(exc, ProviderError) or exc.network:
        return False
    if exc.status == 401:
        return True
    return exc.status in (400, 403) and any(w in exc.text for w in (
        "api key", "api_key", "invalid key", "authentication", "unauthorized"))


def is_daily_limit(exc):
    return exc.status == 429 and (
        any(w in exc.text for w in ("per day", "per-day", "daily", "(rpd)", "(tpd)"))
        or (exc.retry_after or 0) > MAX_WAIT)


def retry_delay(exc, attempt):
    """Seconds to wait before retrying ``exc``, or None to give up."""
    backoff = BACKOFF[min(attempt, len(BACKOFF) - 1)]
    if exc.network:
        return None if (exc.refused or exc.timeout) else backoff
    if exc.status == 429:
        if is_daily_limit(exc):
            return None
        if exc.retry_after is not None:
            return exc.retry_after + 0.5
        return backoff
    if exc.status in (498, 500, 502, 503, 504):
        return backoff
    if exc.status == 400 and exc.code == "tool_use_failed" and attempt == 0:
        return 0                          # Groq: the model garbled a tool call
    return None


def _seconds(text):
    """'7.66s', '1m26.4s', '2h3m', '500ms' -> seconds (None if not a duration)."""
    parts = re.findall(r"(\d+(?:\.\d+)?)(ms|h|m|s)", text or "")
    if not parts:
        return None
    scale = {"ms": 0.001, "s": 1, "m": 60, "h": 3600}
    return sum(float(n) * scale[u] for n, u in parts)


def error_from_http(provider, status, raw, headers=None):
    """Build a ProviderError from an HTTP error response (any provider's
    format: OpenAI-style, Mistral, OpenRouter or Ollama)."""
    try:
        payload = json.loads(raw.decode("utf-8", "replace")) if raw else {}
    except ValueError:
        payload = {}
    message, code = "", ""
    if isinstance(payload, dict):
        err = payload.get("error", payload)
        if isinstance(err, str):
            message = err
        elif isinstance(err, dict):
            message = str(err.get("message") or err.get("detail") or "")
            code = err.get("code") or err.get("type") or ""
            meta = err.get("metadata") or {}
            if isinstance(meta, dict) and meta.get("raw"):   # OpenRouter: upstream error
                message = f"{message} ({str(meta['raw'])[:300]})"
        message = message or str(payload.get("message") or payload.get("detail") or "")
    if not message and raw:
        message = raw.decode("utf-8", "replace")[:300]
    retry_after = None
    value = (headers or {}).get("retry-after") or (headers or {}).get("Retry-After")
    if value:
        try:
            retry_after = float(value)
        except ValueError:
            retry_after = None
    if retry_after is None and status == 429:
        m = re.search(r"try again in\s+([\d.hms]+)", message, re.I)
        retry_after = _seconds(m.group(1)) if m else None
    return ProviderError(provider, status, message, code=code, retry_after=retry_after)


def base_url_for(provider):
    """The provider's address, or the one in its environment variable."""
    value = os.environ.get(provider.base_url_env, "").strip() if provider.base_url_env else ""
    if not value:
        return provider.base_url
    if provider.kind == "ollama":         # OLLAMA_HOST may be "0.0.0.0" or "host:port"
        if "://" not in value:
            value = "http://" + value
        parts = urllib.parse.urlsplit(value)
        host = parts.hostname or "127.0.0.1"
        if host in ("0.0.0.0", "::"):
            host = "127.0.0.1"
        if ":" in host:
            host = f"[{host}]"
        return f"{parts.scheme}://{host}:{parts.port or 11434}"
    return value.rstrip("/")


def _new_id():
    return uuid.uuid4().hex[:9]           # Mistral accepts only 9 letters/digits


def _object(arguments):
    """Tool-call arguments as a dict (they arrive as a JSON string)."""
    if isinstance(arguments, dict):
        return arguments
    try:
        value = json.loads(arguments or "{}")
    except ValueError:
        return {}
    return value if isinstance(value, dict) else {}


def _text(content):
    """Message content as plain text (Mistral may send a list of chunks)."""
    if isinstance(content, list):
        return "".join(c.get("text", "") for c in content
                       if isinstance(c, dict) and c.get("type") in (None, "text"))
    return content or ""


class ChatClient:
    """Talks to one OpenAI-compatible provider, or to Ollama.

    ``complete(model, messages, tools)`` always takes and returns messages in
    the OpenAI format: {"role": "assistant", "content": str,
    "tool_calls": [{"id", "type": "function", "function": {"name",
    "arguments": json-string}}]}. Busy and per-minute-limit errors are retried
    (waiting as long as the provider asks, up to a minute); everything else
    raises ProviderError."""

    def __init__(self, provider, api_key="", base_url=None, timeout=180,
                 retries=3, on_wait=None, sleep=time.sleep):
        self.provider = provider
        self.api_key = api_key or ""
        self.base_url = (base_url or base_url_for(provider)).rstrip("/")
        self.timeout = timeout
        self.retries = retries
        self.on_wait = on_wait            # on_wait(provider, seconds, exc)
        self.sleep = sleep

    # ---- HTTP -------------------------------------------------------------
    def request(self, method, path, body=None):
        headers = {"Accept": "application/json", "User-Agent": "citation-agent"}
        data = None
        if body is not None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        req = urllib.request.Request(self.base_url + path, data=data,
                                     headers=headers, method=method)
        pid = self.provider.id
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read()
        except urllib.error.HTTPError as e:
            raise error_from_http(pid, e.code, e.read(), e.headers) from None
        except urllib.error.URLError as e:
            reason = e.reason
            raise ProviderError(
                pid, None, str(reason), network=True,
                refused=isinstance(reason, ConnectionRefusedError),
                timeout=isinstance(reason, (socket.timeout, TimeoutError))) from None
        except (socket.timeout, TimeoutError) as e:
            raise ProviderError(pid, None, str(e) or "timed out", network=True,
                                timeout=True) from None
        except (ConnectionError, OSError) as e:           # reset, remote closed...
            raise ProviderError(pid, None, str(e), network=True,
                                refused=isinstance(e, ConnectionRefusedError)) from None
        try:
            payload = json.loads(raw.decode("utf-8") or "{}")
        except ValueError:
            raise ProviderError(pid, 502, "the answer wasn't JSON: "
                                + raw[:200].decode("utf-8", "replace")) from None
        if isinstance(payload, dict) and payload.get("error"):
            # OpenRouter can report an upstream failure with HTTP 200
            err = payload["error"]
            status = err.get("code") if isinstance(err, dict) else None
            raise error_from_http(pid, status if isinstance(status, int) else 502,
                                  json.dumps(payload).encode("utf-8"))
        return payload

    # ---- chat -------------------------------------------------------------
    def complete(self, model, messages, tools=None):
        """-> (assistant message in OpenAI format, finish reason)."""
        for attempt in range(self.retries + 1):
            try:
                if self.provider.kind == "ollama":
                    return self._ollama_chat(model, messages, tools)
                return self._openai_chat(model, messages, tools)
            except ProviderError as e:
                delay = None if attempt >= self.retries else retry_delay(e, attempt)
                if delay is None:
                    raise
                if self.on_wait and delay >= 1:
                    self.on_wait(self.provider, delay, e)
                self.sleep(delay)
        raise AssertionError("unreachable")      # pragma: no cover

    def _openai_chat(self, model, messages, tools):
        body = {"model": model, "messages": messages}
        if tools:
            body["tools"] = tools
        if self.provider.id == "groq":
            # Groq charges the *requested* answer length against its 8,000
            # tokens a minute, so ask only for what an answer needs.
            body["max_completion_tokens"] = GROQ_MAX_OUTPUT
            if "gpt-oss" in model:
                body["reasoning_effort"] = "low"  # fewer thinking tokens
        data = self.request("POST", "/chat/completions", body)
        choices = data.get("choices") or []
        if not choices:
            return {"role": "assistant", "content": ""}, "no answer"
        msg = choices[0].get("message") or {}
        calls = []
        for tc in msg.get("tool_calls") or []:
            fn = tc.get("function") or {}
            args = fn.get("arguments")
            if not isinstance(args, str):
                args = json.dumps(args or {})
            calls.append({"id": tc.get("id") or _new_id(), "type": "function",
                          "function": {"name": fn.get("name", ""),
                                       "arguments": args or "{}"}})
        out = {"role": "assistant", "content": _text(msg.get("content"))}
        if calls:
            out["tool_calls"] = calls
            out["content"] = out["content"] or None   # "no text", as in OpenAI's format
        if msg.get("reasoning_details"):         # OpenRouter wants these back
            out["reasoning_details"] = msg["reasoning_details"]
        return out, choices[0].get("finish_reason")

    def _ollama_chat(self, model, messages, tools):
        converted = []
        for m in messages:
            if m["role"] == "assistant" and m.get("tool_calls"):
                converted.append({"role": "assistant", "content": m.get("content") or "",
                                  "tool_calls": [{"type": "function", "function": {
                                      "name": c["function"]["name"],
                                      "arguments": _object(c["function"]["arguments"])}}
                                      for c in m["tool_calls"]]})
            elif m["role"] == "tool":
                converted.append({"role": "tool", "content": m["content"],
                                  "tool_name": m.get("name", "")})
            else:
                converted.append({"role": m["role"], "content": m.get("content") or ""})
        body = {"model": model, "messages": converted, "stream": False,
                "options": {"num_ctx": OLLAMA_NUM_CTX}}
        if tools:
            body["tools"] = tools
        data = self.request("POST", "/api/chat", body)
        msg = data.get("message") or {}
        calls = [{"id": c.get("id") or _new_id(), "type": "function",
                  "function": {"name": (c.get("function") or {}).get("name", ""),
                               "arguments": json.dumps(
                                   _object((c.get("function") or {}).get("arguments")))}}
                 for c in msg.get("tool_calls") or []]
        out = {"role": "assistant", "content": _text(msg.get("content"))}
        if calls:
            out["tool_calls"] = calls
        return out, data.get("done_reason")

    # ---- setup checks -----------------------------------------------------
    def check_key(self):
        """Raises ProviderError unless the provider accepts the key."""
        self.request("GET", self.provider.check_path)

    def ollama_models(self):
        """Names of the models downloaded into Ollama."""
        data = self.request("GET", "/api/tags")
        return [m.get("name") or m.get("model") for m in data.get("models") or []]
