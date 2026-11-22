# Doc Sync Agent

An AI agent that keeps your documentation up to date. Clone it, open it in **any AI coding tool**
(GitHub Copilot, Google Antigravity, Claude Code, Cursor, ...) and say **`run docsync`**.

The agent:

1. Asks **once** for the root folder where your documents live (and optionally your code repo).
2. Reads all your documents and learns their structure (folders, naming, headings, style).
3. Finds what changed in the code since the last run (git commits and diffs, or notes you give it).
4. Shows a **report**: which document needs what change, what should be added, and which new documents are needed.
5. After you approve, it either **creates a new version** (`API_Guide_v2.docx`) or **updates the existing file**
   (backup kept) - your choice, remembered for next time.
6. If a completely new feature appears, it creates a **new document in the right folder, with the same name pattern,
   template and heading structure** as your existing documents, and adds it to your index document.

No API key is needed: the AI tool you already use does the thinking, and a small Python script does the file work.

```
        You: "run docsync"
              |
   AI tool reads AGENTS.md
              |
   docsync.py scan  ->  .docsync/index.md      (all docs as text)
                        .docsync/structure.md  (learned conventions)
                        .docsync/changes.md    (what changed in code)
              |
   AI writes the report  ->  you approve
              |
   AI writes edits.json  ->  docsync.py apply  ->  _v2 files / updated files / new docs
```

## Install

Requirements: Python 3.9+ and Git.

```bash
git clone <your-repo-url> doc-sync-agent
cd doc-sync-agent
pip install -r requirements.txt
```

## Use

Open the `doc-sync-agent` folder in your AI tool and type:

> **run docsync**

On the first run the agent asks for the documents root folder, e.g.
`C:\Users\you\OneDrive - Company\Project Docs`. After that it never asks again.

| Tool | How it picks up the instructions |
|---|---|
| GitHub Copilot (agent mode) | `AGENTS.md`, `.github/copilot-instructions.md`, prompt `/docsync` |
| Google Antigravity | `AGENTS.md` / `GEMINI.md` |
| Claude Code | `CLAUDE.md`, slash command `/docsync` |
| Cursor | `AGENTS.md`, `.cursor/rules/docsync.mdc` |
| Anything else | Tell it: *"Read AGENTS.md and follow it. Run docsync."* |

All of these files point to `AGENTS.md`, which holds the single source of truth. If a tool does not load
it automatically, the last row always works. The tool must be able to run terminal commands (agent mode).

You can also run the helper yourself:

```bash
python docsync.py init              # asks for the docs root (one time)
python docsync.py scan              # index docs + collect changes
python docsync.py note "Added export to CSV on the Reports page"   # when work is not in git
python docsync.py set-mode new-version      # or: in-place
python docsync.py apply .docsync/edits.json --dry-run
python docsync.py apply .docsync/edits.json
python docsync.py status
```

## SharePoint / OneDrive

Sync the SharePoint document library with OneDrive and set the docs root to the **synced local folder**.
The script edits local files and OneDrive uploads them, so no SharePoint API permissions are needed. New
versions, new folders and the changelog all sync automatically, and SharePoint's version history is an
extra safety net. Close a document in Word before applying changes to it.

## Output modes

| Mode | What happens |
|---|---|
| `new-version` | The original stays untouched. Changes go to `Name_v2.docx` (next free number; `_v3` after that). |
| `in-place` | The file is updated. The previous copy is saved to `<docs root>/_docsync/backups/<timestamp>/`. |

New documents are always created as new files. Every run is appended to `<docs root>/_docsync/changelog.md`.
Tip: when several versions exist, only the newest one is edited (older ones are marked *superseded* in the index).

## What the agent can do to a document

| Edit | Use |
|---|---|
| `update_section` | rewrite the text under one heading |
| `add_section` | insert a new heading + text at the right place |
| `append_to_section` | add a bullet/paragraph to the end of a section (e.g. link in an index) |
| `new_document` | create a document using a sibling `.docx` as the template |

The full JSON format is in `AGENTS.md`. Supported files: `.docx` and `.md` (editable), `.txt` (read-only).

## Safety design

- **Report first, apply after approval.** Nothing is written before you say yes.
- **All-or-nothing validation.** If any edit in a batch is invalid (heading not found, file exists, path outside the
  docs root), nothing at all is written.
- **Originals protected** by versioning or backups. Paths outside the docs root are rejected.
- **No invented facts.** The agent is instructed to write only what the diff/notes support and to mark gaps as
  `TODO: confirm with developer`.
- **Images are protected.** A section containing an image or embedded object is not overwritten unless explicitly allowed.

## Known limitations

- **Word formatting:** new text copies the look of existing paragraphs, headings, bullets and table styles. Complex
  layouts (text boxes, tracked changes, content controls, cross-references) are not interpreted; check the first
  results on a copy of your docs.
- **Numbered lists** are written as plain text (`1. ...`), not Word auto-numbering.
- **Quality depends on your AI tool's model.** Always read the report before approving.
- **Code sent to the AI:** the diffs in `changes.md` are read by whatever AI tool you use. Follow your company's
  policy on which tools and repositories are allowed.
- Large document sets are truncated per section in `index.md` (`--max-section-chars`) to fit the AI's context.

## Test it

```bash
python tests/smoke_test.py
```

Builds sample documents and a git repo in a temp folder and runs the whole flow (scan, update, add, new document,
both modes, safety checks). It does not touch your real configuration.

## Files

```
AGENTS.md                         agent instructions (single source of truth)
CLAUDE.md, GEMINI.md              pointers for Claude Code / Antigravity
.github/copilot-instructions.md   pointer for GitHub Copilot
.github/prompts/docsync.prompt.md Copilot prompt  (/docsync)
.claude/commands/docsync.md       Claude Code command (/docsync)
.cursor/rules/docsync.mdc         Cursor rule
docsync.py                        helper CLI
tests/smoke_test.py               end-to-end test
.docsync/                         created on first run (config + generated files, git-ignored)
```
