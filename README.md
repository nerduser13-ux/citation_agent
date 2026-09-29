# Word Citation Footnote Agent

A document-preserving tool that inserts **real Microsoft Word footnotes**
containing the explicit source URL for every numbered citation marker
(`[1]`, `[1], [2]`, `[1]–[3]`, `[4]-[6]`, …) found in an existing `.docx`
file. The References/Bibliography section of the document is used as the only
authoritative number → source-URL mapping.

* The **input file is never modified** — its SHA-256 hash is verified before
  and after processing, and the result is written to a new output path.
* The reference section is detected **generically** (no hard-coded
  "5. References"); if it cannot be identified with confidence the tool
  **aborts without producing output**.
* Only URLs **explicitly present** in a reference entry are used. Nothing is
  invented — no URLs, DOIs, authors, titles, publications, or numbers.
  Bare "DOI: 10.1000/abc" text is **not** turned into a URL.
* Editing is **surgical**: only the exact `<w:r>` run containing a safe
  citation marker is replaced. Paragraph/character formatting, bold, italic,
  underline, fonts, sizes, colors, styles, headings, lists, tables,
  hyperlinks, bookmarks, fields, drawings, images, headers, footers, page
  layout, section breaks, content controls, existing footnotes and all other
  unrelated OOXML are preserved.
* Markers that cannot be safely modified (split across runs, inside a
  hyperlink, inside a field/drawing/complex run, reversed ranges) are
  reported as `AMBIGUOUS` and left **completely untouched**.
* Re-running on an already-processed file does **not** create duplicate
  footnotes (idempotency guard), while unrelated pre-existing footnotes never
  block new citations and are preserved verbatim.

**AI agent (optional).** `python agent.py` starts a chat in which a free AI
(Google Gemini by default; Groq, OpenRouter, Mistral or Ollama on your own
computer also work) runs these tools for you ("add footnotes to my RALF
essay", "check the links in my AI essay") and explains any problems in plain
English. It can also check that every reference link works and leads to the
paper it names. See
[AI agent (chat with Gemini or another free AI)](#ai-agent-chat-with-gemini-or-another-free-ai).

---

## Folder structure

```
citation_agent/
├── main.py                # CLI + workflow orchestration
├── config.py              # all paths / behaviour switches / detection thresholds
├── document_reader.py     # read-only docx access; single shared paragraph ordering
│                          #   (body + tables incl. nested + content controls),
│                          #   heading detection, raw-text helpers, footnote snapshot
├── reference_parser.py    # reference-section location + number -> URL extraction
├── citation_detector.py   # [n] / [a]-[b] detection in the body only, range expansion
├── word_footnotes.py      # surgical OOXML footnote insertion + experimental COM backend
├── validator.py           # input-unchanged + full post-save output validation
├── report.py              # citation_review_report.csv writer
├── agent_tools.py         # the AI agent's tools (Toolbox) + narrow pipeline wrappers
├── agent.py               # AI agent: terminal chat, the AI decides which tool to run
├── llm_providers.py       # the other free AIs: Groq, OpenRouter, Mistral, Ollama
├── link_checker.py        # does each reference link work / lead to that paper? (Crossref, ...)
├── requirements.txt
├── requirements-agent.txt # requirements.txt + Google's Gemini SDK (for agent.py)
├── input/                 # drop your .docx here (or use --input)
├── output/                # generated files land here
└── tests/
    ├── create_sample.py   # builds a realistic sample .docx (incl. a manual footnote)
    ├── validate_sample.py # deep checks of the sample output (OOXML level)
    ├── test_pipeline.py   # pytest end-to-end suite (48 tests)
    └── test_agent.py      # agent, AIs, tools, link checker, fully offline (105 tests)
```

## Requirements

Python 3.9+ and:

```
python-docx>=1.1.0
lxml>=4.9.0
# optional, Windows-only COM backend:
# pywin32>=306
# optional, for the test suite:
# pytest
```

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/pip install pytest
```

For the AI agent (`agent.py`) install `requirements-agent.txt` instead (it
includes `requirements.txt`). Gemini needs Python **3.10+**; the other free
AIs need nothing beyond `requirements.txt`.

## How to run

```bash
# 1) Use the first .docx found in ./input
python main.py

# 2) Explicit input
python main.py --input path/to/thesis.docx

# 3) Explicit input + output
python main.py --input thesis.docx --output thesis_with_footnotes.docx
```

Options:

| Flag | Meaning |
| --- | --- |
| `--backend xml\|com` | `xml` = OOXML manipulation (default, cross-platform, validated). `com` = Windows Microsoft Word COM automation (experimental, Windows-only). |
| `--keep-marker` | Keep the `[n]` text and insert the footnote reference **after** it. Default: replace `[n]` with the footnote reference. |
| `--overwrite` | Allow overwriting an existing output file (default: create `<name>_with_footnotes_1.docx`, `_2`, …). |
| `--input-dir` | Directory scanned for the default input (default `input`). |
| `--output-dir` | Output directory (default `output`). |

Exit codes: `0` success (output produced **and validated**); `1` aborted
(reference section not found) or validation failed; `2` usage error /
unsupported backend; `3` input file changed during processing (fatal).

Example session (sample document):

```
Word Citation Footnote Agent
Input : input/sample.docx
Output: output/sample_with_footnotes.docx
Analyzing document...
Reference section : FOUND (heading, high) - Reference heading detected: "References"
References detected: 11
URLs detected     : 10
Citation markers  : 9
Processing...

Footnotes inserted : 11
Unresolved citations: 0
Missing URLs       : 1
Ambiguous citations: 0
Validation         : PASS
```

## AI agent (chat with Gemini or another free AI)

`agent.py` lets you use the tool by chatting instead of typing commands.
A free AI is the "brain" (Google Gemini by default, or
[Groq, OpenRouter, Mistral or Ollama](#other-free-ais-groq-openrouter-mistral-ollama)):
it understands what you ask, decides which tool to run and explains the
results in plain English. The tools are the tested code of this project, so
every guarantee above still holds whichever AI you pick: the AI cannot edit
your documents itself, invent links or open other files.

```
You> add footnotes to my RALF essay
  ... looking at your documents
  ... adding footnotes to "RALF Reflective Professional Competency Account.docx"

Agent> Done! I added 7 footnotes to your RALF document.
- New file: RALF Reflective Professional Competency Account_with_footnotes.docx
- Folder: output
- Validation: passed - your original file was not changed.
```

Things you can ask, for example:

* "add footnotes to my RALF essay" (or "... but keep the [n] numbers")
* "check the links in my AI essay": does each link work and lead to that paper?
* "what problems would there be in the AI document?" (a preview that saves nothing)
* "why wasn't citation [7] turned into a footnote?"

### Setup (once)

1. **Python 3.10 or newer** for Gemini (`main.py` and the other AIs also
   work on 3.9).
2. In the project folder, with the virtual environment active:
   ```
   pip install -r requirements-agent.txt
   ```
3. Get a **free Gemini API key** at <https://aistudio.google.com/apikey>:
   sign in with a Google account and click *Create API key*. No credit card
   is needed. (Prefer another AI? See
   [below](#other-free-ais-groq-openrouter-mistral-ollama).)

### Start it

```
python agent.py
```

The first time, it asks for the key. Paste it with Ctrl+V or a right-click.
It stays hidden, but you then see a masked preview such as
`AIza...x9Qk (39 characters)`, so you know the paste worked. The agent
**checks the key with Google before using or saving it**, so a wrong paste is
never stored. It then offers to save the key in `%USERPROFILE%\.citation_agent\`
(`~/.citation_agent/` on macOS/Linux). That is **outside** the project folder,
so the key can never be pushed to GitHub. Type `exit` to quit.

**Wrong or old key?** If Gemini rejects the saved key during a chat, the agent
offers to paste a new one on the spot and then retries your message. To
replace a key yourself, run `python agent.py --forget-key` (or delete
`%USERPROFILE%\.citation_agent\gemini_api_key.txt`) and start the agent again.

| Option | Effect |
| --- | --- |
| `--once "check the links in my RALF essay"` | one question, one answer, then exit |
| `--provider groq` | think with another free AI: `gemini` (default), `groq`, `openrouter`, `mistral`, `ollama` (see below) |
| `--providers` | list the free AIs, their limits, privacy terms and where to get a key |
| `--model gemini-3.8-flash` | use a specific model. Default: the provider's; for Gemini `gemini-flash-latest`, Google's alias for its newest Flash model (or set `GEMINI_MODEL`, `GROQ_MODEL`, ...) |
| `--forget-key` | delete the saved key (of `--provider`, default Gemini) |
| `--input-dir`, `--output-dir` | folders to use (default: `input/` and `output/` next to `agent.py`) |
| `GEMINI_API_KEY` environment variable | used instead of the saved key (likewise `GROQ_API_KEY`, `OPENROUTER_API_KEY`, `MISTRAL_API_KEY`) |

In the chat, `/use groq` (or any other provider name) switches AI; see below.

### Other free AIs: Groq, OpenRouter, Mistral, Ollama

Gemini stays the default. For a second AI when Gemini's daily limit is used
up, or for different privacy terms, the agent can also think with these
(free limits as of September 2026; `python agent.py --providers` shows the
same list):

| `--provider` | Free allowance | Card needed? | What happens to what the agent sends | Default model |
| --- | --- | --- | --- | --- |
| `gemini` | daily limits, shown in AI Studio | no | free tier: Google may use it to improve its products | `gemini-flash-latest` |
| `groq` | 1,000 requests a day, 30 a minute, 8,000 tokens a minute | no | not used for training, not kept (except up to 30 days for abuse checks) | `openai/gpt-oss-120b` |
| `openrouter` | 50 requests a day, 20 a minute (1,000 a day after a one-time $10 top-up) | no | free models are run by other companies that may keep and train on it; you have to allow this in OpenRouter's privacy settings | `openrouter/free` (picks a current free model that can use tools) |
| `mistral` | a monthly allowance in "Free mode", limits shown in the console | no (phone number needed) | may be used for training unless you turn that off | `mistral-small-latest` |
| `ollama` | unlimited: runs on your own computer | no key at all | never leaves your computer | `qwen3:8b` (5 GB download, about 8 GB free memory) |

Every provider needs a model that can call tools; all the defaults can.

```
python agent.py --provider groq         # asks for a Groq key the first time
python agent.py --provider openrouter
python agent.py --provider mistral
python agent.py --provider ollama       # after installing Ollama and: ollama pull qwen3:8b
```

Keys: Groq <https://console.groq.com/keys>, OpenRouter
<https://openrouter.ai/settings/keys>, Mistral
<https://console.mistral.ai/api-keys>. Like the Gemini key, each key is
checked before it is saved, is kept in its own file next to the Gemini key
(`groq_api_key.txt`, ...) and can be removed with
`python agent.py --provider groq --forget-key`.

**Switching during a chat.** Type `/use groq` (or `/use gemini`,
`/use ollama`, ..., optionally followed by a model name). The conversation so
far comes along as text; the new AI calls the tools again when it needs the
details. When an AI's free limit is used up, the agent offers to switch to
one you already have a key for and sends your message again.

Good to know:

* **Groq** is very fast, but its free 8,000 tokens a minute cover only one
  or two steps of a question. When the minute is used up, the agent waits as
  long as Groq asks (seconds, at most a minute) and carries on by itself:
  `... Groq's per-minute limit is reached - waiting 20 seconds`. To use as
  little as possible, it asks Groq for short answers (at most 1,500 tokens)
  and, in long chats, leaves out old tool results.
* **OpenRouter**'s free models change often; `openrouter/free` always picks
  one that is currently free and can use tools. If the agent says free models
  "only work after you allow them", turn on the free-model options at
  <https://openrouter.ai/settings/privacy>. That is the setting that lets
  those companies keep your prompts.
* **Ollama**: install it from <https://ollama.com/download>, then run
  `ollama pull qwen3:8b` once. On a computer with little memory try
  `ollama pull granite4.1:3b` and `--model granite4.1:3b` (faster, less
  accurate). The agent asks Ollama for a 16K-token context window, because
  Ollama's default on most laptops (4K) silently cuts off longer requests.
* Smaller free models follow instructions less reliably than Gemini. They
  can't do any damage, though: footnotes still come only from the tested
  pipeline and your reference list.
* **Not free (September 2026):** GitHub Models was retired on 30 July 2026;
  Cerebras replaced its free tier with a $5 trial that needs a card;
  OpenAI, Anthropic, DeepSeek and xAI have no free API. Other
  OpenAI-compatible services (for example Cloudflare Workers AI, Cohere's
  trial key or NVIDIA's free endpoints) can be added with one entry in
  `llm_providers.PROVIDERS`.

### The four tools the AI can use

| Tool | What it does |
| --- | --- |
| `list_documents` | lists the `.docx` files in `input/` (and what is in `output/`) |
| `analyze_document` | dry run: which citations would get footnotes, which have problems and why; saves nothing |
| `add_footnotes` | the normal pipeline: new file in `output/`, original untouched, validated |
| `check_links` | checks each reference link (below) |

That is everything it can do. A document name from the AI is matched against
the files that actually exist in `input/`, so nothing else can be reached, and
the only file it can create is the pipeline's output.

### Link checking

`check_links` is also available without the AI:
`python link_checker.py "input\My Essay.docx"`. For every reference link:

* **DOI links** (`https://doi.org/10...`, or a DOI inside a publisher URL) are
  checked against **Crossref**, the official DOI registry: the DOI must exist
  and its registered title must match the reference. Publisher websites often
  block automated checks; Crossref does not.
* **ScienceDirect** (`/pii/...`) links are looked up at Crossref too, and
  **PubMed Central** (`PMC...`) links at Europe PMC.
* Any other link: the page or PDF is opened and its title compared.

| Verdict | Meaning |
| --- | --- |
| `OK` | works, and the title matches the reference |
| `CHECK` | works, but the title only partly matches or can't be read; compare it yourself |
| `PROBLEM` | broken (not found, DOI doesn't exist, website gone) or leads to a **different paper** |
| `UNVERIFIED` | couldn't be checked automatically (site blocks robots, timeout, offline); open it in a browser |
| `NO_LINK` | the reference has no URL |

Only registry data can say "different paper"; a website title that doesn't
match is only `CHECK`. The checker never changes anything and never suggests
replacement links.

### Privacy and cost

* The AI receives your messages, file names, section headings, citation
  markers, the **reference list** and the tools' results. It does **not**
  receive the body text of your document.
* On Gemini's **free tier, Google may use this content to improve its
  products** (see the [pricing page](https://ai.google.dev/gemini-api/docs/pricing)).
  Don't use the free tier for confidential documents. The other AIs' terms
  are in the table above; with Ollama nothing leaves your computer.
* Free tiers have per-minute and per-day limits. If you hit one, the agent
  tells you (and offers another AI you have a key for); wait a minute and try
  again.

### Troubleshooting

| Message | Fix |
| --- | --- |
| "didn't accept your API key" | the agent offers to paste a new key. Copy the key itself with the copy button in AI Studio, not its name or project. Or run `python agent.py --forget-key` and start again |
| "That isn't a key: ... the paste didn't work" | paste again: Ctrl+V or a right-click in the Command Prompt window |
| an old key stopped working | since May 28, 2026, AI Studio creates "auth keys" and the Gemini API rejects old *unrestricted* keys, so create a new key |
| "free-tier limit" | wait a minute (or until tomorrow for the daily limit) |
| "doesn't offer the model" | `python agent.py --model gemini-3.8-flash` |
| "needs Google's Gemini package" | `pip install -r requirements-agent.txt` |
| "needs Python 3.10 or newer" | install a newer Python from python.org, or use another AI (`--provider groq`) |
| "per-minute limit is reached - waiting ..." | normal on Groq's free plan; the agent retries by itself |
| "used up ...'s free requests for today" | try again later, or `/use` another AI |
| "free models only work after you allow them" | OpenRouter: allow the free-model options at <https://openrouter.ai/settings/privacy> (or use another AI) |
| "Ollama isn't running" | start the Ollama app (install it from <https://ollama.com/download>) |
| "isn't on this computer yet" / "isn't in Ollama yet" | `ollama pull qwen3:8b` (or the model named in the message) |
| "can't use tools" | choose a model that supports tools (`--model ...`) |

## Citation syntax supported

* `[1]` — single citation
* `[1] and [2]`, `[1], [2], and [3]` — multiple distinct citations
* `[2] … [2]` — repeated citation numbers (each occurrence gets a footnote)
* `[1]-[3]`, `[1]–[3]` (en dash), `[1]—[3]` (em dash) — ranges, expanded
  (`[4]-[6]` → references 4, 5, 6 → three footnotes)

Detection operates on the actual OOXML structure:

* Only the **direct runs** of a paragraph are used as the coordinate space,
  so a marker next to a hyperlink is never mis-attributed to the link.
* Detection **stops at the start of the reference section** — the reference
  list itself is never scanned for body citations.
* Arbitrary numbers in prose are never citations (brackets are required).

## Reference-section detection

The heading is matched generically, most-conservative-first:

1. **Heading (high confidence)** — a styled heading (Heading 1–9 / Title /
   outline level, or an all-bold short line) whose text, after stripping
   leading numbering (`5.`, `1.2.`, `A.`) and trailing punctuation, is
   *exactly* a keyword: *references, reference, bibliography, works cited,
   reference list, literature cited, sources*. The **last** such heading wins
   (reference sections sit at the end of the document).
2. **Heading + block (medium)** — a heading that merely *contains* a keyword
   is accepted only if it is directly followed by a qualifying numbered block
   (≥ 3 numbered entries, ≥ 1 URL in entry lines).
3. **Block detection (medium)** — with no heading at all, a numbered block
   near the end of the document (starting at/after 40% of the paragraph list)
   with ≥ 3 numbered entries and ≥ 1 URL.
4. **Nothing** → the tool **aborts**: no output file, input unchanged, and
   the CSV report says `ABORTED - reference section could not be
   confidently identified`. It never guesses.

A short body sentence that merely contains a keyword (e.g. "Our data sources
were collected …") is **not** treated as a heading, and neither is a table
cell (e.g. a bold "Reference" / "Sources" column header in an appendix table).

### Reference parsing rules

* Entry lines look like `1. …`, `1) …`, `1 …` or `[1] …`.
* A line starting with a year (`2020. …`) is **not** a numbered entry; it ends
  the current entry (its URLs must not be attributed to a reference number).
* Wrapped continuation lines are appended to the current entry.
* The list ends at the next styled heading **or** at a bold-only heading line
  (short, entirely bold, not itself a numbered entry — e.g. "Appendix A" in a
  document that formats headings with direct bold). Otherwise a following
  section would be glued onto the last entry and its URLs attributed to that
  reference. Consequence: bold category sub-headings *inside* a reference
  list ("Books", "Web sources") also end it; later citations are then
  reported as unresolved, never mis-linked.
* Decimal Word auto-numbering (`w:numPr` on the paragraph) is read as the
  entry number when the number is not typed in the text.
* Duplicate numbers: the first occurrence wins (warning recorded).
* Only explicit `http(s)://` URLs are extracted. Trailing sentence
  punctuation (`. , ; : ! ?`) is stripped from the end of a URL.
* **Multiple URLs in one entry** — deterministic selection rule: the **last**
  URL that is not a DOI resolver (`doi.org`); if every URL is a DOI resolver,
  the **last** URL. URLs are used exactly as written.

## Safety rules (what is never modified)

A citation marker is modified only when **all** of the following hold:

1. it lies entirely inside a **single direct run** of the paragraph,
2. that run contains **only text** (its children are exclusively
   `w:rPr`/`w:t` — no fields, drawings, OLE objects, tabs, breaks, or inline
   content controls). The one exception is `w:lastRenderedPageBreak`, Word's
   content-free layout-cache hint present in almost every multi-page
   document: it is allowed and re-emitted before exactly the same character
   (the validator checks its count is unchanged),
3. the run is **not part of a complex field** (between `fldChar` begin and
   end — how EndNote, Zotero and Mendeley Desktop store `[1]`, and how
   cross-references work; a footnote placed there would be lost or break the
   field on the next refresh),
4. its number exists in the reference section **and** that reference has an
   explicit URL.

Otherwise the marker is left byte-for-byte untouched and reported:

| Status | Meaning |
| --- | --- |
| `SUCCESS` | footnote(s) inserted |
| `ERROR` | reference number missing, or reference has no URL (unresolved) |
| `AMBIGUOUS` | split across runs / inside hyperlink / inside a field result / complex run / reversed range |
| `SKIPPED` | already processed by a previous run of this tool |

Specifically, markers inside `<w:hyperlink>`, fields, or runs containing
drawings/objects/breaks are **not** modified; unrelated hyperlinks, fields,
bookmarks and drawings are always left exactly where they are.

## Real Word footnotes (XML backend)

The output contains genuine OOXML footnotes:

* `word/footnotes.xml` with real `<w:footnote>` definitions (including the
  required `separator` / `continuationSeparator` entries) and a
  `footnotes.xml` relationship on the main document part,
* corresponding `<w:footnoteReference>` elements in the body,
* the FootnoteText paragraph style and FootnoteReference character style
  (reused if the document already defines them, created otherwise),
* the displayed numbering is Word's own footnote mechanism (automatic
  `<w:footnoteRef>`), never typed text.

Each generated footnote contains **exactly** the URL from the corresponding
reference entry, unaltered. For a range like `[4]-[6]` three footnotes are
created (references 4, 5, 6).

The insertion is surgical at the XML level: only the affected run is replaced
by `[before-text] [footnote-ref…] [after-text]` (or
`[before-text] [marker-text] [footnote-ref…] [after-text]` with
`--keep-marker`). The reference runs deep-copy the original run's character
formatting (so e.g. bold is kept) and gain the FootnoteReference character
style, inserted as the schema-correct first child of `w:rPr`.

## Idempotency

* **Replace mode** (default): markers disappear after the first run, so a
  second run finds nothing to do and preserves everything.
* **Keep-marker mode**: on the second run the marker is still present; it is
  recognized as previously processed when a footnote reference immediately
  adjacent to it carries **one of the citation's own URLs** — and is skipped
  without creating a duplicate. A pre-existing footnote with different
  content never triggers a skip.

## Input protection & output naming

* SHA-256 of the input is computed before processing and re-verified after;
  any change aborts with a loud failure (exit 3).
* Processing happens on a separate in-memory `Document` instance; the only
  file written is the output.
* Default output: `output/<input-name>_with_footnotes.docx`. If that file
  exists, `<input-name>_with_footnotes_1.docx`, `_2`, … is used unless
  `--overwrite` is given. An output path equal to the input path is rejected.

## Validation

After saving, the tool automatically verifies (and reports `PASS`/`FAIL`):

* output exists, is a valid ZIP, reopens with python-docx;
* `word/footnotes.xml` exists and the **count** of regular footnotes and body
  footnote references is exactly right (existing + newly inserted);
* every **new** footnote contains exactly the URL of its reference (no
  fabricated URLs);
* **existing footnotes are preserved** — every pre-existing definition still
  exists with identical text, and every pre-existing body reference is intact;
* **per-paragraph raw-text fidelity** — the output text equals the input text
  minus exactly the removed markers (this also proves no text was lost from
  tables, hyperlinks, or any other structure);
* structural preservation — paragraph, table, hyperlink, bookmark, field,
  drawing, content-control, page-size and section-break counts unchanged, and
  every input hyperlink target is still present;
* the input file's hash is unchanged.

If any check fails the run exits non-zero and the failures are printed.

## CSV review report

`citation_review_report.csv` (next to the output) contains one row per
citation occurrence:

`Citation Number, Location, Citation Text, Reference Text, URL, Action,
Status, Notes`

followed by a `SUMMARY` block: Input, Output, Reference section
(+ method + confidence), References detected, URLs detected, Citation markers
detected, Footnotes inserted, Existing footnotes detected / preserved,
Unresolved citations, Missing URLs, Ambiguous citations, Uncited references,
Previously processed, Output valid, Reopened with python-docx, All footnote
URLs from refs, Input unchanged, Status.

"Uncited references" are references never cited in the body. A reference
whose markers were found but could not be processed (`AMBIGUOUS` / `ERROR`)
counts as cited — it shows up in those rows instead.

## Windows / Microsoft Word COM backend

```bash
pip install pywin32
python main.py --input thesis.docx --backend com
```

Opens the document in Microsoft Word, finds each marker (one `Find` per
citation, in document order, with all sticky Find options reset), validates
every reference number *before* touching the text, deletes the marker when
replacing, inserts automatically numbered footnotes with
`Footnotes.Add(Range=…, Text=url)` at a collapsed range (each one after the
previous reference mark, so `[4]-[6]` stays in order), and `SaveAs`-es to the
output path (input untouched). Markers classified unsafe by the detector are
found and stepped over, never modified.

The **experimental** backend is Windows-only and was **not executed** in the
Linux build/test environment (it fails loudly on other platforms); it follows
the documented Word object model but has not been run against Word. The XML
backend is the validated default and is what the test suite exercises.

## Tests

```bash
python -m pytest tests/ -v                 # 153 tests: 48 pipeline (OOXML-level)
                                           #   + 105 agent / link checker (offline)
python tests/create_sample.py              # (re)generate input/sample.docx
                                           #   (or: create_sample.py OUT.docx)
python main.py                             # process the sample
python tests/validate_sample.py            # deep checks of the sample output
```

The pytest suite covers: simple `[1]`; multiple citations; repeated numbers;
`[1]-[3]`, `[1]–[3]` and `[1]—[3]` ranges; reversed range (ambiguous);
citation in bold text; citation in a table (incl. nested tables); hyperlink
preservation; citation **inside** a hyperlink (untouched); citation **split
across runs** (untouched); fields/bookmarks preservation; missing reference
URL; missing reference number; uncited reference; reference section not
scanned for citations; year-starting line not an entry; DOI-preference URL
selection; bare DOI text never invented; pre-existing footnotes preserved;
unrelated existing footnote not blocking; idempotent second run (both
modes); input hash unchanged; non-standard headings ("Bibliography",
"5. References", bold unstyled "Works Cited"); block fallback without
heading; keyword-sentence false-positive avoidance; abort when no reference
section; keep-marker mode; output-collision safe naming; `--overwrite`;
output==input rejection; COM backend rejection on Linux; sample end-to-end
(generated into a temp dir — running the suite never modifies tracked files).

`tests/test_agent.py` runs **without internet or an API key**. A local fake
server plays Crossref, Europe PMC, doi.org and ordinary websites (OK / partly
matching / different paper / not found / robot check / 403 / PDF / redirect /
offline), and another plays Gemini, driven through Google's real
`google-genai` SDK. It covers: DOI/PII/PMCID extraction; title matching;
every link verdict; the toolbox's file restrictions (path tricks, Word lock
files, ambiguous names); previews that write nothing; footnotes via the
agent with the original unchanged; problems explained; multi-step and
parallel tool calls; Gemini 3 thought signatures sent back unchanged; tool
errors returned to the model; API errors turned into advice with the
conversation rolled back; the step limit; API-key handling. Fake Groq,
OpenRouter, Mistral and Ollama servers cover the other AIs: OpenAI-format
tool calls, Ollama's own format and context size, waiting for per-minute
limits, retries when a service is busy, error advice for each service,
trimming long chats, one key file per service, `/use` switching with the
conversation, and the offer to switch when a limit is used up. The Gemini
tests are skipped when `google-genai` is not installed.

Regression tests for issues found on real Word documents: runs carrying
`w:lastRenderedPageBreak` (8 layouts × both modes, checked against an oracle
for exact hint placement); citation in a block-level content control; bold
"Appendix" heading after the list must not leak its URL into the last
reference; bold "Reference" table header not taken as the reference heading;
EndNote-style field result left untouched; "uncited" excludes references
cited by unprocessed markers; `agent_tools.run_pipeline` returns the real
(collision-safe) output path; `w:vertAlign` inserted in schema order.

The sample document deliberately includes: a heading-based reference section,
inline/multiple/repeated/range citations, a citation inside a bold run, a
real hyperlink, a citation inside a table cell, a **pre-existing manual
footnote**, one reference with **no URL** (reported unresolved), and one
**uncited** reference (reported).

## Known limitations / design choices

* **Citations inside inline content controls, text boxes, headers, footers
  or footnotes themselves** are out of scope: they are never modified (their
  surrounding structure is preserved) but not reported either, because they
  are not part of the body-paragraph coordinate space.
* **Keep-marker idempotency** is position-sensitive: it matches a footnote
  reference *immediately* adjacent to the marker with matching content. If a
  user inserts extra text between marker and footnote after processing, a
  re-run may add another footnote (conservative, documented behavior).
* **`[n]` inside a reference entry that has no URL** and other
  unresolvable markers are left in place by design (never invented).
* **Style-linked list numbering** (entries numbered only via a paragraph
  style such as "List Number", with no `w:numPr` on the paragraph itself) is
  not read: such a list yields no references and every citation is reported
  unresolved. Direct list numbering (the usual Numbering button) works.
* Markers inside **tracked insertions** (`w:ins`), simple fields
  (`w:fldSimple`) or smart tags are, like inline content controls, outside
  the direct-run coordinate space: never modified, not reported.
* **COM backend** is experimental, Windows-only, and untested here; prefer
  the XML backend.
* The tool operates on the **first document part only** (`word/document.xml`);
  sub-documents (e.g. attached OLE Word objects) are preserved but not
  processed.
* **Future Mendeley stage** is intentionally out of scope; the pipeline stops
  at `citation marker → reference URL → real Word footnote`. `agent_tools.py`
  exposes narrow, explicit functions (`inspect_document`,
  `find_reference_section`, `extract_references`, `find_citations`,
  `validate_citation`, `verify_output`, `generate_report`, `run_pipeline`)
  so an AI agent can drive the same deterministic pipeline without
  unrestricted filesystem or shell access.

## AI-agent design

All deterministic work (reading the file, extracting URLs, matching numbers,
insertion, validation, link checking) is plain Python and fully
reproducible. The AI only decides *which* tool to call and explains the
results:

```
you -> agent.py -> the AI (decides) -> agent_tools.Toolbox -> main.run / link_checker
                          ^------------- plain JSON results --------------'
```

* Gemini is driven through Google's `google-genai` SDK (`CitationAgent`).
  Groq, OpenRouter and Mistral use the OpenAI chat-completions format and
  Ollama its own `/api/chat`, through `llm_providers.ChatClient`
  (standard library only) and `ChatAgent`. Every AI gets the same system
  prompt, the same four tools and the same rules.

* `agent_tools.Toolbox` is the entire interface: four tools with JSON inputs
  and outputs. Errors come back as `{"error": ...}` so the model can explain
  them instead of the chat crashing.
* Documents are chosen by name from the listing of `input/`; a path from the
  model is never opened.
* Footnotes are only inserted by the deterministic pipeline, with its
  validation. The model never touches document XML and cannot supply URLs.
* The model's turns are sent back exactly as received, including Gemini 3
  "thought signatures", which multi-step tool use requires.
* `agent_tools.py` still offers the lower-level functions
  (`inspect_document`, `find_reference_section`, ..., `run_pipeline`) for
  other integrations. The planned Mendeley stage (identify publication ->
  Mendeley reference -> Mendeley Cite citation) can be added as more tools in
  the same way.
