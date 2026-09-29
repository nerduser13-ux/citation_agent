"""agent.py - chat with an AI agent that runs the citation tools for you.

    python agent.py                        # chat in the terminal
    python agent.py --once "check the links in my RALF document"
    python agent.py --model gemini-3.8-flash
    python agent.py --forget-key           # delete the saved API key

The "brain" is Google Gemini (free API key: https://aistudio.google.com/apikey).
It understands what you ask and decides which tool to use, but it can ONLY
use the four tools in ``agent_tools.Toolbox``:

    list_documents    - which Word files are in the input folder
    analyze_document  - preview citations / problems, saves nothing
    add_footnotes     - the same tested pipeline as main.py: a NEW copy in
                        output/, original untouched, links only from your
                        reference list
    check_links       - does each reference link work and lead to that paper?

It cannot edit your documents itself, open other files or run commands.

Privacy: Gemini sees your messages, file names, section headings, the
reference list and the tools' results - not the body text of your document.
On Gemini's free tier Google may use this to improve its products.

Needs Python 3.10+ and ``pip install -r requirements-agent.txt``.
"""
import argparse
import getpass
import os
import re
import sys
from pathlib import Path

from agent_tools import Toolbox

HERE = Path(__file__).resolve().parent
DEFAULT_MODEL = "gemini-flash-latest"   # Google's alias for the newest Flash model
KEY_URL = "https://aistudio.google.com/apikey"
# Saved OUTSIDE the project folder, so it can never end up on GitHub.
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


class CitationAgent:
    """The agent loop: send the conversation to Gemini; when it asks for a
    tool, run it through the Toolbox and send the result back; repeat until
    Gemini answers in words.

    ``client`` is a ``google.genai.Client``. The model's own turns are kept
    exactly as returned (Gemini 3 needs its "thought signatures" back)."""

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
            return ("Sorry - that needed too many steps. Please ask again, one "
                    "thing at a time.")
        except BaseException:
            del self.history[mark:]
            raise


# --------------------------------------------------------------------------
# Setup helpers
# --------------------------------------------------------------------------
def make_client(api_key, base_url=None, retries=True):
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


def is_key_error(exc):
    """True if Gemini rejected the API key itself (mis-pasted, deleted,
    blocked) - not for limits, outages or an unknown model name."""
    try:
        from google.genai import errors
    except ImportError:          # pragma: no cover - checked in main()
        return False
    if not isinstance(exc, errors.APIError):
        return False
    code = exc.code or 0
    blob = f"{exc.status} {exc.message} {exc.details}".lower()
    if code == 401:              # "Request had invalid authentication credentials"
        return True
    return code in (400, 403) and any(
        w in blob for w in ("api key", "api_key", "credential", "unregistered callers"))


def key_problem(key):
    """Why ``key`` can't be an API key ('' if it could be). Catches failed
    pastes into the hidden prompt without assuming Google's key format
    (new "auth keys" exist since 2026)."""
    if not key:
        return "nothing was pasted"
    if any(ord(c) < 32 or c.isspace() for c in key):
        return "it contains spaces or invisible characters, so the paste didn't work"
    if len(key) < 20:
        return f"it is only {len(key)} characters long"
    return ""


def mask(key):
    """'AIza...x9Qk (39 characters)': enough to compare with AI Studio,
    without printing the key."""
    if len(key) >= 12:
        return f"{key[:4]}...{key[-4:]} ({len(key)} characters)"
    return f"({len(key)} characters)"


def check_key(key, model=DEFAULT_MODEL):
    """Ask Google whether it accepts ``key`` (a free metadata request).
    (True, "") accepted - (False, reason) rejected - (None, reason) could not
    tell right now (offline, busy)."""
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


def save_key(key):
    KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
    KEY_FILE.write_text(key, encoding="utf-8")
    try:
        os.chmod(KEY_FILE, 0o600)
    except OSError:
        pass
    print(f"Saved in {KEY_FILE} (outside the project, so it can't end up on GitHub).")


def forget_key():
    """Delete the saved key. True if there was one."""
    if KEY_FILE.is_file():
        KEY_FILE.unlink()
        return True
    return False


def ask_for_key(model=DEFAULT_MODEL, intro=True, attempts=3):
    """Hidden prompt for a key. Each key is checked with Google BEFORE it is
    used or saved, so a wrong paste is never stored. Returns key or None."""
    if intro:
        print("To think, the agent uses Google Gemini, which needs a free API key.")
        print("(On the free tier Google may use what the agent sends - your messages")
        print(" and reference list, not your essay text - to improve its products.)")
    print(f"  1. Open {KEY_URL} and sign in with a Google account")
    print("  2. Click 'Create API key', then the copy button next to the new key")
    for attempt in range(1, attempts + 1):
        key = getpass.getpass("  3. Paste it here (Ctrl+V or right-click) and press "
                              "Enter. It stays hidden: ").strip()
        problem = key_problem(key)
        if problem:
            print(f"     That isn't a key: {problem}.")
        else:
            print(f"     Got {mask(key)}. Checking it with Google...", flush=True)
            ok, reason = check_key(key, model)
            if ok is not False:
                print("     The key works." if ok else
                      f"     Couldn't check it right now. {reason}")
                answer = input("Save the key on this computer so you don't have to "
                               "paste it again? [Y/n] ").strip().lower()
                if answer in ("", "y", "yes"):
                    save_key(key)
                return key
            print(f"     {reason} Copy the key itself (the long code), not its name "
                  "or project, and check that its last 4 characters match AI Studio.")
        if attempt < attempts:
            print("     Please try again.")
    print("No working key yet - start the agent again when you have one.")
    return None


def get_api_key(interactive=True, model=DEFAULT_MODEL):
    """(key, source) with source "env", "file" or "typed"; (None, None) if
    there is no key. A typed key has already been checked with Google."""
    for var in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        if os.environ.get(var, "").strip():
            return os.environ[var].strip(), "env"
    if KEY_FILE.is_file():
        key = KEY_FILE.read_text(encoding="utf-8").strip()
        if key:
            return key, "file"
    if not interactive:
        return None, None
    key = ask_for_key(model)
    return (key, "typed") if key else (None, None)


def replace_rejected_key(agent, source, model):
    """Gemini rejected the key mid-chat: offer to paste a new one right away.
    True if the agent now uses a new (checked) key."""
    if source == "env":
        print("\nGemini didn't accept the API key from the GEMINI_API_KEY (or "
              "GOOGLE_API_KEY) environment variable. Change or remove it, then "
              "start the agent again.")
        return False
    where = " saved on this computer" if source == "file" else ""
    print(f"\nGemini didn't accept the API key{where}. It may have been pasted "
          "wrongly, deleted or blocked.")
    answer = input("Paste a new key now? [Y/n] ").strip().lower()
    if answer not in ("", "y", "yes"):
        print("OK. Later you can run:  python agent.py --forget-key")
        return False
    if forget_key():
        print("Removed the old saved key.")
    key = ask_for_key(model, intro=False)
    if not key:
        return False
    agent.client = make_client(key)
    return True


def friendly_error(exc, model=DEFAULT_MODEL):
    """Turn an exception from Gemini / the network into advice."""
    try:
        from google.genai import errors
    except ImportError:          # pragma: no cover - checked in main()
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
        return "Couldn't reach Gemini. Check your internet connection and try again."
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


def _banner(toolbox, model):
    print("Citation Agent - AI helper for footnotes and reference links")
    print(f"Brain: Google Gemini ({model})")
    print(f"Documents folder: {toolbox.input_dir.resolve()}")
    docs = toolbox.documents()
    if docs:
        print("Your documents:")
        for p in docs:
            print(f"  - {p.name}")
    else:
        print("No .docx files there yet - copy your document into that folder.")
    print('Try: "add footnotes to my RALF document" or "check the links in the AI essay".')
    print("Type exit to quit.")


def _parse_args(argv):
    p = argparse.ArgumentParser(description=(
        "Chat with an AI agent (Google Gemini) that adds Word footnotes for "
        "[n] citations and checks reference links."))
    p.add_argument("--model", default=os.environ.get("GEMINI_MODEL", DEFAULT_MODEL),
                   help=f"Gemini model (default: {DEFAULT_MODEL}, or $GEMINI_MODEL)")
    p.add_argument("--once", metavar="MESSAGE",
                   help="Send one message, print the answer and exit")
    p.add_argument("--forget-key", action="store_true",
                   help="Delete the saved API key and exit")
    p.add_argument("--input-dir", default=str(HERE / "input"))
    p.add_argument("--output-dir", default=str(HERE / "output"))
    return p.parse_args(argv)


def main(argv=None):
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    try:                                   # never crash on an unprintable character
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    if args.forget_key:
        if forget_key():
            print(f"Deleted the saved key ({KEY_FILE}). Start the agent again to "
                  "paste a new one.")
        else:
            print("No saved key found.")
        return 0
    if sys.version_info < (3, 10):
        print("The AI agent needs Python 3.10 or newer (main.py still works on 3.9).")
        return 2
    try:
        import google.genai  # noqa: F401
    except ImportError:
        print("The AI agent needs Google's Gemini package. Install it with:\n"
              "    pip install -r requirements-agent.txt")
        return 2

    key, source = get_api_key(interactive=sys.stdin.isatty() and not args.once,
                              model=args.model)
    if not key:
        print(f"No API key. Get a free one at {KEY_URL} and start the agent again "
              "(or set the GEMINI_API_KEY environment variable).")
        return 2

    toolbox = Toolbox(args.input_dir, args.output_dir)
    agent = CitationAgent(make_client(key), toolbox, model=args.model,
                          on_tool=_show_activity)
    if args.once:
        try:
            print(plain(agent.ask(args.once)))
            return 0
        except Exception as e:
            print(friendly_error(e, args.model))
            return 1

    _banner(toolbox, args.model)
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
        for attempt in (1, 2):
            try:
                answer = agent.ask(text)
            except KeyboardInterrupt:
                print("\n(stopped)")
                break
            except Exception as e:
                if attempt == 1 and is_key_error(e) and replace_rejected_key(
                        agent, source, args.model):
                    source = "file" if KEY_FILE.is_file() else "typed"
                    print("Trying your message again...")
                    continue
                print("\n" + friendly_error(e, args.model))
                break
            print("\nAgent> " + plain(answer))
            break
    print("Goodbye!")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
