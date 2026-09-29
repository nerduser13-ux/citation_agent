"""agent.py - chat with an AI agent that runs the citation tools for you.

    python agent.py                        # chat in the terminal (Google Gemini)
    python agent.py --provider groq        # think with another free AI
    python agent.py --providers            # list the free AIs and how to get a key
    python agent.py --once "check the links in my RALF document"
    python agent.py --model gemini-3.8-flash
    python agent.py --forget-key           # delete the saved Gemini key
    python agent.py --provider groq --forget-key

The "brain" is a free AI: Google Gemini by default, or Groq, OpenRouter,
Mistral - or Ollama, which runs on your own computer (see llm_providers.py).
In the chat, ``/use groq`` (etc.) switches AI and keeps the conversation.
The AI understands what you ask and decides which tool to use, but it can
ONLY use the four tools in ``agent_tools.Toolbox``:

    list_documents    - which Word files are in the input folder
    analyze_document  - preview citations / problems, saves nothing
    add_footnotes     - the same tested pipeline as main.py: a NEW copy in
                        output/, original untouched, links only from your
                        reference list
    check_links       - does each reference link work and lead to that paper?

It cannot edit your documents itself, open other files or run commands.

Privacy: the AI sees your messages, file names, section headings, the
reference list and the tools' results - not the body text of your document.
What a service may do with that differs (``python agent.py --providers``);
with Ollama nothing leaves your computer.

Needs ``pip install -r requirements-agent.txt``. Gemini needs Python 3.10+.
"""
import argparse
import getpass
import json
import math
import os
import re
import sys
import textwrap
from pathlib import Path

import llm_providers
from agent_tools import Toolbox
from llm_providers import PROVIDERS, ChatClient, ProviderError

HERE = Path(__file__).resolve().parent
GEMINI = PROVIDERS["gemini"]
DEFAULT_MODEL = GEMINI.default_model     # Google's alias for the newest Flash model
KEY_URL = GEMINI.key_url
# Saved OUTSIDE the project folder, so they can never end up on GitHub.
# (Other providers' keys go next to it: groq_api_key.txt, ...)
KEY_FILE = Path.home() / ".citation_agent" / "gemini_api_key.txt"

SYSTEM_PROMPT = """\
You are Citation Agent, a friendly assistant for students who write in Microsoft Word.
You help with two jobs:
1. Turning numbered citations such as [1], [2] or [3]-[5] into real Word footnotes that contain the source's link, taken from the document's own reference list.
2. Checking that each link in the reference list works and leads to the paper it names.

How you work:
- You only know what your tools tell you. Use them; never guess what a document contains.
- If the user doesn't name a document exactly, call list_documents and pick the obvious match. If more than one could fit, ask which one.
- Use analyze_document to preview or to answer questions ("how many citations?", "why would [7] fail?").
- Use add_footnotes to actually create the footnoted copy. It is safe (the original is never changed), so don't ask for permission when the user wants footnotes.
- Use check_links when the user wants the links or sources checked.

Rules you must never break:
- Never invent, guess, complete or "correct" a URL, DOI or reference, even if you think you know it. Say which reference needs fixing and why, and suggest the user fixes it in the reference list in Word and then asks you to add the footnotes again.
- You cannot edit documents yourself. The only change you can make is add_footnotes, which saves a new copy in the output folder. Never claim to have changed anything else.
- You cannot read the body text of documents; you only get the reference list, section headings, citation markers and results.

How to report:
- After add_footnotes: give the new file's name and folder, how many footnotes were added, and whether validation passed. If validation FAILED, say clearly that the new file should not be used. Then explain every problem in plain words: where it is (section and paragraph), what went wrong and how to fix it.
- After check_links: give the totals first. Then list the PROBLEM and CHECK links (reference number, short title, what is wrong), then the UNVERIFIED ones, which the user should open in a browser because the website blocks automatic checks. Don't list the OK links unless asked.
- Problem types: ERROR = the reference is missing from the list or has no link (fix the reference list). AMBIGUOUS = left unchanged for safety, e.g. the [n] is a hyperlink, comes from a reference manager such as EndNote or Zotero, or is split by formatting; the user can add that footnote by hand in Word (References > Insert Footnote). SKIPPED = already has a footnote from an earlier run.
- Word numbers footnotes 1, 2, 3... in reading order, so a citation [2] can become footnote 3, and a source cited twice gets two footnotes.

Style: this is a plain-text terminal. Write short, friendly sentences and simple lists that start with "- ". No Markdown: no **bold**, no # headings, no tables. Answer in the language the user writes in.
If asked about something unrelated, say briefly what you can help with.
"""

TOOL_ACTIVITY = {
    "list_documents": "looking at your documents",
    "analyze_document": "analysing {name}",
    "add_footnotes": "adding footnotes to {name}",
    "check_links": "checking the links in {name} (this can take up to a minute)",
}

TOO_MANY_STEPS = "Sorry - that needed too many steps. Please ask again, one thing at a time."


class CitationAgent:
    """The agent loop for Google Gemini: send the conversation; when Gemini
    asks for a tool, run it through the Toolbox and send the result back;
    repeat until Gemini answers in words.

    ``client`` is a ``google.genai.Client``. The model's own turns are kept
    exactly as returned (Gemini 3 needs its "thought signatures" back)."""

    provider = "gemini"

    def __init__(self, client, toolbox, model=DEFAULT_MODEL, on_tool=None,
                 max_steps=12):
        from google.genai import types
        self._types = types
        self.client = client
        self.toolbox = toolbox
        self.model = model
        self.on_tool = on_tool
        self.max_steps = max_steps
        self.history = []
        declarations = []
        for spec in toolbox.TOOL_SPECS:
            kw = {"name": spec["name"], "description": spec["description"]}
            if spec["parameters"]:           # Gemini rejects empty OBJECT schemas
                kw["parameters_json_schema"] = spec["parameters"]
            declarations.append(types.FunctionDeclaration(**kw))
        self.config = types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            tools=[types.Tool(function_declarations=declarations)],
            automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
        )

    def seed(self, transcript):
        """Start from an earlier conversation [(user text, answer), ...]
        (after switching AI). Plain text only, so no signatures are needed."""
        types = self._types
        for user, answer in transcript:
            self.history.append(types.Content(role="user", parts=[types.Part.from_text(text=user)]))
            self.history.append(types.Content(role="model", parts=[types.Part.from_text(text=answer)]))

    def ask(self, text):
        """Send one user message; returns the agent's final answer (str).
        On any error the conversation is rolled back to before this message
        and the exception re-raised."""
        types = self._types
        mark = len(self.history)
        self.history.append(types.Content(role="user", parts=[types.Part.from_text(text=text)]))
        try:
            for _ in range(self.max_steps):
                resp = self.client.models.generate_content(
                    model=self.model, contents=self.history, config=self.config)
                cand = resp.candidates[0] if resp.candidates else None
                content = cand.content if cand is not None else None
                if content is None or not content.parts:
                    reason = getattr(cand, "finish_reason", None) if cand is not None else None
                    del self.history[mark:]
                    return ("(Gemini sent an empty answer"
                            + (f" - reason: {reason}" if reason else "")
                            + ". Please try asking in a different way.)")
                self.history.append(content)      # unchanged: keeps thought signatures
                calls = [p.function_call for p in content.parts if p.function_call]
                if not calls:
                    answer = "".join(p.text for p in content.parts
                                     if p.text and not p.thought).strip()
                    return answer or "(no answer)"
                replies = []
                for call in calls:
                    args = dict(call.args or {})
                    if self.on_tool:
                        self.on_tool(call.name, args)
                    result = self.toolbox.call(call.name, args)
                    part = types.Part.from_function_response(name=call.name, response=result)
                    if call.id:
                        part.function_response.id = call.id
                    replies.append(part)
                self.history.append(types.Content(role="user", parts=replies))
            del self.history[mark:]
            return TOO_MANY_STEPS
        except BaseException:
            del self.history[mark:]
            raise


# Replaces an old tool result when a request has to be made smaller.
LEFT_OUT = json.dumps({"note": "Older result left out to save space. Call the "
                               "tool again if you need it."})


class ChatAgent:
    """The same agent loop for Groq, OpenRouter, Mistral and Ollama, which
    use OpenAI-style messages. ``client`` is an ``llm_providers.ChatClient``.

    Free tiers with small limits (Groq: 8,000 tokens a minute; Ollama: the
    context window) get a trimmed copy of the conversation: old tool results
    are replaced by a note first, then the oldest exchanges are left out."""

    def __init__(self, client, toolbox, model, on_tool=None, max_steps=12,
                 max_chars=None):
        self.client = client
        self.provider = client.provider.id
        self.toolbox = toolbox
        self.model = model
        self.on_tool = on_tool
        self.max_steps = max_steps
        self.max_chars = client.provider.max_chars if max_chars is None else max_chars
        self.history = []
        self.tools = [{"type": "function", "function": {
            "name": spec["name"], "description": spec["description"],
            "parameters": spec["parameters"] or {"type": "object", "properties": {}}}}
            for spec in toolbox.TOOL_SPECS]

    def seed(self, transcript):
        for user, answer in transcript:
            self.history += [{"role": "user", "content": user},
                             {"role": "assistant", "content": answer}]

    def window(self):
        """The part of the history to send (see the class docstring)."""
        msgs = list(self.history)
        if not self.max_chars:
            return msgs
        budget = self.max_chars - len(SYSTEM_PROMPT) - len(json.dumps(self.tools))

        def size():
            return sum(len(json.dumps(m, ensure_ascii=False)) for m in msgs)

        if size() <= budget:
            return msgs
        current = max(i for i, m in enumerate(msgs) if m["role"] == "user")
        for i in range(current):                  # never trim the current request
            if msgs[i]["role"] == "tool" and len(msgs[i]["content"]) > len(LEFT_OUT):
                msgs[i] = dict(msgs[i], content=LEFT_OUT)
                if size() <= budget:
                    return msgs
        while size() > budget:
            starts = [i for i, m in enumerate(msgs) if m["role"] == "user"]
            if len(starts) < 2:
                break
            del msgs[:starts[1]]                  # a whole exchange at a time
        return msgs

    def ask(self, text):
        """Send one user message; returns the final answer (str). On any error
        the conversation is rolled back and the exception re-raised."""
        label = self.client.provider.name
        mark = len(self.history)
        self.history.append({"role": "user", "content": text})
        try:
            for _ in range(self.max_steps):
                messages = [{"role": "system", "content": SYSTEM_PROMPT}] + self.window()
                msg, finish = self.client.complete(self.model, messages, self.tools)
                calls = msg.get("tool_calls") or []
                if not calls:
                    answer = re.sub(r"(?s)<think>.*?</think>", "", msg.get("content") or "").strip()
                    if not answer:
                        del self.history[mark:]
                        return (f"({label} sent an empty answer"
                                + (f" - reason: {finish}" if finish and finish != "stop" else "")
                                + ". Please try asking in a different way.)")
                    self.history.append({"role": "assistant", "content": answer})
                    if finish == "length":
                        answer += "\n(The answer was cut short - ask for the rest if you need it.)"
                    return answer
                self.history.append(msg)
                for call in calls:
                    name = call["function"]["name"]
                    try:
                        args = json.loads(call["function"]["arguments"] or "{}")
                        if args is None:              # some models send "null"
                            args = {}
                        if not isinstance(args, dict):
                            raise ValueError
                    except ValueError:
                        result = {"error": "The arguments were not a valid JSON object. "
                                           "Call the tool again."}
                    else:
                        if self.on_tool:
                            self.on_tool(name, args)
                        result = self.toolbox.call(name, args)
                    self.history.append({"role": "tool", "tool_call_id": call["id"],
                                         "name": name,
                                         "content": json.dumps(result, ensure_ascii=False,
                                                               separators=(",", ":"))})
            del self.history[mark:]
            return TOO_MANY_STEPS
        except BaseException:
            del self.history[mark:]
            raise


# --------------------------------------------------------------------------
# Setup helpers
# --------------------------------------------------------------------------
def make_client(api_key, base_url=None, retries=True):
    """A google.genai Client for Gemini."""
    from google import genai
    from google.genai import types
    retry = types.HttpRetryOptions(
        attempts=4, initial_delay=2.0, max_delay=30.0,
        http_status_codes=[429, 500, 502, 503, 504]) if retries else None
    return genai.Client(api_key=api_key, http_options=types.HttpOptions(
        base_url=base_url,
        timeout=180_000,                   # milliseconds
        retry_options=retry,
    ))


def new_client(provider, key):
    """The right client for ``provider`` (a PROVIDERS key)."""
    if PROVIDERS[provider].kind == "gemini":
        return make_client(key)
    return ChatClient(PROVIDERS[provider], key, on_wait=_show_wait)


def make_agent(provider, key, model, toolbox):
    if PROVIDERS[provider].kind == "gemini":
        return CitationAgent(new_client(provider, key), toolbox, model=model,
                             on_tool=_show_activity)
    return ChatAgent(new_client(provider, key), toolbox, model=model,
                     on_tool=_show_activity)


def is_key_error(exc):
    """True if the AI service rejected the API key itself (mis-pasted,
    deleted, blocked) - not for limits, outages or an unknown model name."""
    if isinstance(exc, ProviderError):
        return llm_providers.is_key_error(exc)
    try:
        from google.genai import errors
    except ImportError:          # pragma: no cover - checked in open_session()
        return False
    if not isinstance(exc, errors.APIError):
        return False
    code = exc.code or 0
    blob = f"{exc.status} {exc.message} {exc.details}".lower()
    if code == 401:              # "Request had invalid authentication credentials"
        return True
    return code in (400, 403) and any(
        w in blob for w in ("api key", "api_key", "credential", "unregistered callers"))


def is_limit_error(exc):
    """A free-tier limit or an outage: worth offering another AI."""
    if isinstance(exc, ProviderError):
        return exc.status in (402, 429) or (exc.status or 0) >= 500
    code = getattr(exc, "code", None)
    return isinstance(code, int) and (code == 429 or code >= 500)


def key_problem(key):
    """Why ``key`` can't be an API key ('' if it could be). Catches failed
    pastes into the hidden prompt without assuming a key format (Google's
    new "auth keys" exist since 2026)."""
    if not key:
        return "nothing was pasted"
    if any(ord(c) < 32 or c.isspace() for c in key):
        return "it contains spaces or invisible characters, so the paste didn't work"
    if len(key) < 20:
        return f"it is only {len(key)} characters long"
    return ""


def mask(key):
    """'AIza...x9Qk (39 characters)': enough to compare with the website,
    without printing the key."""
    if len(key) >= 12:
        return f"{key[:4]}...{key[-4:]} ({len(key)} characters)"
    return f"({len(key)} characters)"


def key_path(provider="gemini"):
    if provider == "gemini":
        return KEY_FILE
    return KEY_FILE.parent / f"{provider}_api_key.txt"


def forget_command(provider="gemini"):
    if provider == "gemini":
        return "python agent.py --forget-key"
    return f"python agent.py --provider {provider} --forget-key"


def check_key(key, model=None, provider="gemini"):
    """Ask the service whether it accepts ``key`` (a free request).
    (True, "") accepted - (False, reason) rejected - (None, reason) could not
    tell right now (offline, busy)."""
    p = PROVIDERS[provider]
    model = model or p.default_model
    if p.kind != "gemini":
        try:
            ChatClient(p, key, retries=0, timeout=30).check_key()
            return True, ""
        except ProviderError as e:
            if is_key_error(e):
                return False, f"{p.company} didn't accept this key."
            return None, friendly_error(e, model, provider)
    # Keep a reference: google-genai closes its connection when the Client
    # object is garbage-collected, even in the middle of a chained call.
    client = make_client(key, retries=False)
    try:
        client.models.get(model=model)
        return True, ""
    except Exception as e:
        if is_key_error(e):
            return False, "Google didn't accept this key."
        if getattr(e, "code", None) == 404:   # key accepted; model name is handled later
            return True, ""
        return None, friendly_error(e, model)


def save_key(key, provider="gemini"):
    path = key_path(provider)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(key, encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    print(f"Saved in {path} (outside the project, so it can't end up on GitHub).")


def forget_key(provider="gemini"):
    """Delete the saved key. True if there was one."""
    path = key_path(provider)
    if path.is_file():
        path.unlink()
        return True
    return False


def has_key(provider):
    p = PROVIDERS[provider]
    return (any(os.environ.get(v, "").strip() for v in p.key_envs)
            or key_path(provider).is_file())


def ask_for_key(model=None, intro=True, attempts=3, provider="gemini"):
    """Hidden prompt for a key. Each key is checked with the service BEFORE
    it is used or saved, so a wrong paste is never stored. Returns key or None."""
    p = PROVIDERS[provider]
    if intro and p.kind == "gemini":
        print("To think, the agent uses Google Gemini, which needs a free API key.")
        print("(On the free tier Google may use what the agent sends - your messages")
        print(" and reference list, not your essay text - to improve its products.)")
        print("(Other free AIs work too:  python agent.py --providers)")
    elif intro:
        print(f"To think, the agent uses {p.label}, which needs a free API key.")
        print(f"  Free: {p.free}")
        print(f"  Privacy: {p.privacy} It gets your messages and reference list, "
              "not your essay text.")
    for i, step in enumerate(p.key_steps, 1):
        print(f"  {i}. {step.format(url=p.key_url)}")
    n = len(p.key_steps) + 1
    for attempt in range(1, attempts + 1):
        key = getpass.getpass(f"  {n}. Paste it here (Ctrl+V or right-click) and press "
                              "Enter. It stays hidden: ").strip()
        problem = key_problem(key)
        if problem:
            print(f"     That isn't a key: {problem}.")
        else:
            print(f"     Got {mask(key)}. Checking it with {p.company}...", flush=True)
            ok, reason = check_key(key, model, provider)
            if ok is not False:
                print("     The key works." if ok else
                      f"     Couldn't check it right now. {reason}")
                answer = input("Save the key on this computer so you don't have to "
                               "paste it again? [Y/n] ").strip().lower()
                if answer in ("", "y", "yes"):
                    save_key(key, provider)
                return key
            print(f"     {reason} Copy the key itself (the long code), not its name "
                  f"or project, and check that its last 4 characters match {p.site}.")
        if attempt < attempts:
            print("     Please try again.")
    print("No working key yet - start the agent again when you have one.")
    return None


def get_api_key(interactive=True, model=None, provider="gemini"):
    """(key, source) with source "env", "file" or "typed"; (None, None) if
    there is no key. A typed key has already been checked with the service."""
    p = PROVIDERS[provider]
    for var in p.key_envs:
        if os.environ.get(var, "").strip():
            return os.environ[var].strip(), "env"
    path = key_path(provider)
    if path.is_file():
        key = path.read_text(encoding="utf-8").strip()
        if key:
            return key, "file"
    if not interactive:
        return None, None
    key = ask_for_key(model, provider=provider)
    return (key, "typed") if key else (None, None)


def replace_rejected_key(agent, source, model, provider="gemini"):
    """The service rejected the key mid-chat: offer to paste a new one right
    away. True if the agent now uses a new (checked) key."""
    p = PROVIDERS[provider]
    if source == "env":
        names = p.key_envs[0] + "".join(f" (or {v})" for v in p.key_envs[1:])
        print(f"\n{p.name} didn't accept the API key from the {names} environment "
              "variable. Change or remove it, then start the agent again.")
        return False
    where = " saved on this computer" if source == "file" else ""
    print(f"\n{p.name} didn't accept the API key{where}. It may have been pasted "
          "wrongly, deleted or blocked.")
    answer = input("Paste a new key now? [Y/n] ").strip().lower()
    if answer not in ("", "y", "yes"):
        print(f"OK. Later you can run:  {forget_command(provider)}")
        return False
    if forget_key(provider):
        print("Removed the old saved key.")
    key = ask_for_key(model, intro=False, provider=provider)
    if not key:
        return False
    agent.client = new_client(provider, key)
    return True


def _model_missing(e):
    t = e.text
    return e.status == 404 or (e.status == 400 and (
        "invalid_model" in t or "invalid model" in t or "not a valid model" in t
        or ("model" in t and ("not found" in t or "does not exist" in t))))


def _provider_advice(e, model):
    """Advice for a ProviderError (Groq, OpenRouter, Mistral, Ollama)."""
    p = PROVIDERS.get(e.provider, GEMINI)
    name, t, status = p.name, e.text, e.status or 0
    local = p.kind == "ollama"
    if e.network:
        if local and not e.timeout:
            return (f"Ollama isn't running on this computer. Start the Ollama app "
                    f"(get it from {p.key_url}), then try again.")
        if e.timeout:
            return (f"{name} took too long to answer. Please try again"
                    + (" - a model on your own computer can be slow, especially "
                       "the first time it loads." if local else "."))
        return f"Couldn't reach {name}. Check your internet connection and try again."
    if status == 429:
        if llm_providers.is_daily_limit(e):
            return (f"You've used up {name}'s free requests for today. Try again "
                    "later (the limit resets within a day).")
        return f"You've reached {name}'s free-tier limit. Wait a minute and try again."
    if is_key_error(e):
        return (f"{name} didn't accept your API key (pasted wrongly, deleted or "
                f"blocked). To enter a new one, run:  {forget_command(p.id)}  and "
                "then start the agent again.")
    if status == 402:
        return (f"{name} says your account has no credit left (free models stop "
                "working too when the balance is below zero). Check your account "
                "on the website.")
    if "data policy" in t:
        return ("OpenRouter's free models only work after you allow them: turn on "
                "the free-model options at https://openrouter.ai/settings/privacy "
                "(those companies may keep what the agent sends).")
    if re.search(r"supports? (tool|function)|does not support tools|tools? (are|is) "
                 r"not supported|tool calling is not supported", t):
        return (f"The model '{model}' can't use tools, which the agent needs. "
                "Choose another one with --model"
                + (f" (for example: ollama pull {p.default_model})." if local else "."))
    if _model_missing(e):
        if local:
            return (f"The model '{model}' isn't on this computer yet. Download it "
                    f"once with:  ollama pull {model}")
        return (f"{name} doesn't offer the model '{model}' to you. Try:  "
                f"python agent.py --provider {p.id} --model {p.default_model}")
    if status == 413 or "too large" in t or "context length" in t or "context window" in t:
        return (f"That was too much text for {name}'s free tier at once. Start the "
                "agent again for a fresh chat, or use another free AI.")
    if status >= 500:
        return f"{name} is busy or having problems right now. Please try again shortly."
    if status == 403:
        return f"{name} refused the request: {e.message}"
    return f"{name} error {status}: {e.message}"


def friendly_error(exc, model=DEFAULT_MODEL, provider="gemini"):
    """Turn an exception from the AI service / the network into advice."""
    if isinstance(exc, ProviderError):
        return _provider_advice(exc, model)
    try:
        from google.genai import errors
    except ImportError:          # pragma: no cover - checked in open_session()
        errors = None
    if errors is not None and isinstance(exc, errors.APIError):
        code = exc.code or 0
        msg = (exc.message or "").strip()
        blob = f"{exc.status} {msg} {exc.details}".lower()
        if code == 429:
            return ("You've reached Gemini's free-tier limit. Wait a minute and "
                    "try again (there is also a daily limit).")
        if is_key_error(exc):
            return ("Gemini didn't accept your API key (pasted wrongly, deleted "
                    "or blocked). To enter a new one, run:  python agent.py "
                    "--forget-key  and then start the agent again.")
        if "location is not supported" in blob:
            return "Gemini isn't available where you are (Google blocks some regions)."
        if code == 404:
            return (f"Gemini doesn't offer the model '{model}' to you. Try:  "
                    "python agent.py --model gemini-3.8-flash")
        if code == 403:
            return f"Gemini refused the request: {msg}"
        if code >= 500:
            return "Gemini is busy or having problems right now. Please try again shortly."
        return f"Gemini error {code}: {msg}"
    name = type(exc).__name__
    if isinstance(exc, (ConnectionError, TimeoutError)) or name in (
            "ConnectError", "ConnectTimeout", "ReadTimeout", "WriteTimeout",
            "PoolTimeout", "TimeoutException", "RemoteProtocolError"):
        return (f"Couldn't reach {PROVIDERS[provider].name}. Check your internet "
                "connection and try again.")
    return f"Something went wrong: {name}: {exc}"


def plain(text):
    """Remove Markdown that models like to add anyway (it shows up as
    literal ** and # in a terminal)."""
    text = re.sub(r"\*\*(.+?)\*\*|__(.+?)__", lambda m: m.group(1) or m.group(2), text)
    text = re.sub(r"(?m)^#{1,6}\s+", "", text)
    text = re.sub(r"(?m)^(\s*)[*\u2022]\s+", r"\1- ", text)
    return text


def _show_activity(name, args):
    label = TOOL_ACTIVITY.get(name, name).format(name=f"\"{args.get('name', '')}\"")
    print(f"  ... {label}", flush=True)


def _show_wait(provider, seconds, exc):
    if exc.network or (exc.status or 0) >= 500:
        why = f"{provider.name} is busy"
    else:
        why = f"{provider.name}'s per-minute limit is reached"
    print(f"  ... {why} - waiting {math.ceil(seconds)} seconds", flush=True)


# --------------------------------------------------------------------------
# Choosing and switching the AI
# --------------------------------------------------------------------------
class Session:
    """The AI currently in use."""

    def __init__(self, provider, model, source, agent):
        self.provider = provider          # PROVIDERS key
        self.model = model
        self.source = source              # where the key came from
        self.agent = agent

    @property
    def p(self):
        return PROVIDERS[self.provider]


def _ollama_ready(model):
    """Is Ollama running and is ``model`` downloaded? Prints advice if not."""
    client = ChatClient(PROVIDERS["ollama"], retries=0, timeout=15)
    try:
        names = client.ollama_models()
    except ProviderError as e:
        print(friendly_error(e, model, "ollama"))
        return False
    if model in names or f"{model}:latest" in names:
        return True
    print(f"The model '{model}' isn't in Ollama yet. Download it once with:  "
          f"ollama pull {model}")
    if names:
        print("Models you already have: " + ", ".join(names))
        print("(use one:  python agent.py --provider ollama --model NAME  - or in "
              "the chat:  /use ollama NAME)")
    return False


def open_session(provider, model, toolbox, interactive):
    """Set up ``provider`` (asking for its key if needed). None if that's not
    possible; the reason has been printed."""
    p = PROVIDERS[provider]
    model = model or os.environ.get(p.model_env, "").strip() or p.default_model
    if p.kind == "gemini":
        if sys.version_info < (3, 10):
            print("Gemini needs Python 3.10 or newer (main.py still works on 3.9, and "
                  "so do the other AIs:  python agent.py --providers).")
            return None
        try:
            import google.genai  # noqa: F401
        except ImportError:
            print("The AI agent needs Google's Gemini package. Install it with:\n"
                  "    pip install -r requirements-agent.txt\n"
                  "(The other free AIs don't need it:  python agent.py --providers)")
            return None
    key, source = "", "none"
    if p.needs_key:
        key, source = get_api_key(interactive=interactive, model=model, provider=provider)
        if not key:
            print(f"No API key. Get a free one at {p.key_url} and start the agent again "
                  f"(or set the {p.key_envs[0]} environment variable).")
            return None
    elif not _ollama_ready(model):
        return None
    return Session(provider, model, source, make_agent(provider, key, model, toolbox))


def _switch_to(provider, model, toolbox, transcript):
    new = open_session(provider, model, toolbox, interactive=True)
    if new is None:
        return None
    new.agent.seed(transcript)
    print(f"Now using {new.p.label}, model {new.model}"
          + (" - the conversation so far comes along." if transcript else "."))
    return new


def _use_command(text, session, toolbox, transcript):
    """/use groq [model]"""
    parts = text.split()
    if len(parts) < 2 or parts[1].lower() not in PROVIDERS:
        print("Type /use and one of: " + ", ".join(PROVIDERS)
              + "   (for example: /use groq)")
        return session
    provider = parts[1].lower()
    model = parts[2] if len(parts) > 2 else None
    if provider == session.provider and model in (None, session.model):
        print(f"Already using {session.p.label}, model {session.model}.")
        return session
    return _switch_to(provider, model, toolbox, transcript) or session


def _offer_switch(session, toolbox, transcript):
    """After a limit/outage: offer an AI the user already has a key for."""
    others = [pid for pid in PROVIDERS if pid != session.provider]
    ready = [pid for pid in others if PROVIDERS[pid].needs_key and has_key(pid)]
    if not ready:
        print("Tip: you can carry on with another free AI - type "
              + ", ".join(f"/use {pid}" for pid in others)
              + "  (python agent.py --providers shows how to get a key).")
        return None
    alt = PROVIDERS[ready[0]]
    answer = input(f"Switch this chat to {alt.label} (you have a key for it) and "
                   "try again? [Y/n] ").strip().lower()
    if answer not in ("", "y", "yes"):
        return None
    return _switch_to(alt.id, None, toolbox, transcript)


def _ask(session, text, toolbox, transcript):
    """Send one message, recovering from a rejected key or a used-up limit.
    Returns (session, answer or None)."""
    for attempt in (1, 2):
        try:
            return session, session.agent.ask(text)
        except KeyboardInterrupt:
            print("\n(stopped)")
            return session, None
        except Exception as e:
            if attempt == 1 and is_key_error(e) and replace_rejected_key(
                    session.agent, session.source, session.model, session.provider):
                session.source = "file" if key_path(session.provider).is_file() else "typed"
                print("Trying your message again...")
                continue
            print("\n" + friendly_error(e, session.model, session.provider))
            if attempt == 1 and is_limit_error(e):
                new = _offer_switch(session, toolbox, transcript)
                if new is not None:
                    session = new
                    print("Trying your message again...")
                    continue
            return session, None
    return session, None                   # pragma: no cover


def print_providers():
    wrap = textwrap.TextWrapper(width=78, initial_indent=" " * 6,
                                subsequent_indent=" " * 6)
    print("Free AIs the agent can think with (free limits as of September 2026):")
    for p in PROVIDERS.values():
        if not p.needs_key:
            status = "no key needed"
        elif key_path(p.id).is_file():
            status = "key saved"
        elif any(os.environ.get(v, "").strip() for v in p.key_envs):
            status = "key in an environment variable"
        else:
            status = "no key yet"
        default = ", the default" if p.id == "gemini" else ""
        print(f"\n  {p.id} - {p.label}{default} ({status})")
        print(wrap.fill(f"Free: {p.free}"))
        print(wrap.fill(f"Privacy: {p.privacy}"))
        print(wrap.fill(f"Model: {p.default_model}   "
                        + (f"Key: {p.key_url}" if p.needs_key else f"Get it: {p.key_url}")))
    print("\nStart with one:     python agent.py --provider groq")
    print("Switch in the chat: /use groq   (the conversation comes along)")


def _banner(toolbox, session):
    print("Citation Agent - AI helper for footnotes and reference links")
    print(f"Brain: {session.p.label}, model {session.model}")
    print(f"Documents folder: {toolbox.input_dir.resolve()}")
    docs = toolbox.documents()
    if docs:
        print("Your documents:")
        for p in docs:
            print(f"  - {p.name}")
    else:
        print("No .docx files there yet - copy your document into that folder.")
    print('Try: "add footnotes to my RALF document" or "check the links in the AI essay".')
    print("Another free AI any time: "
          + ", ".join(f"/use {pid}" for pid in PROVIDERS if pid != session.provider))
    print("Type exit to quit.")


def _parse_args(argv):
    p = argparse.ArgumentParser(description=(
        "Chat with an AI agent that adds Word footnotes for [n] citations and "
        "checks reference links. The AI can be Google Gemini (default), Groq, "
        "OpenRouter, Mistral or Ollama."))
    p.add_argument("--provider", choices=list(PROVIDERS), default="gemini",
                   help="which free AI to use (default: gemini). See --providers")
    p.add_argument("--providers", action="store_true",
                   help="list the free AIs, their limits and where to get a key")
    p.add_argument("--model", default=None,
                   help=("model name (default: the provider's, e.g. "
                         f"{DEFAULT_MODEL}; or $GEMINI_MODEL, $GROQ_MODEL, ...)"))
    p.add_argument("--once", metavar="MESSAGE",
                   help="Send one message, print the answer and exit")
    p.add_argument("--forget-key", action="store_true",
                   help="Delete the saved API key of --provider and exit")
    p.add_argument("--input-dir", default=str(HERE / "input"))
    p.add_argument("--output-dir", default=str(HERE / "output"))
    return p.parse_args(argv)


def main(argv=None):
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    try:                                   # never crash on an unprintable character
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    if args.providers:
        print_providers()
        return 0
    provider = PROVIDERS[args.provider]
    if args.forget_key:
        if not provider.needs_key:
            print(f"{provider.name} doesn't use an API key.")
        elif forget_key(provider.id):
            print(f"Deleted the saved key ({key_path(provider.id)}). Start the agent "
                  "again to paste a new one.")
        else:
            print("No saved key found.")
        return 0

    toolbox = Toolbox(args.input_dir, args.output_dir)
    session = open_session(provider.id, args.model, toolbox,
                           interactive=sys.stdin.isatty() and not args.once)
    if session is None:
        return 2
    if args.once:
        try:
            print(plain(session.agent.ask(args.once)))
            return 0
        except Exception as e:
            print(friendly_error(e, session.model, session.provider))
            return 1

    _banner(toolbox, session)
    transcript = []                        # (message, answer) - comes along on /use
    while True:
        try:
            text = input("\nYou> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not text:
            continue
        if text.lower() in ("exit", "quit", "bye", "q"):
            break
        if text.split()[0].lower() in ("/use", "/switch"):
            session = _use_command(text, session, toolbox, transcript)
            continue
        session, answer = _ask(session, text, toolbox, transcript)
        if answer is not None:
            transcript.append((text, answer))
            print("\nAgent> " + plain(answer))
    print("Goodbye!")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
