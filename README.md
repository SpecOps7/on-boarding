# Box Q&A

A local web app that hooks into your **Box Drive** through the mounted Windows folder (no Box API key), lets you pick folders, indexes their documents, and answers questions about them with citations — powered by the **Claude Code CLI** (`claude -p`, no Anthropic API key either).

## How it works

- Runs inside **WSL**. Box Drive's virtual filesystem can't be read from WSL directly, so all Box access goes through small PowerShell helper scripts (`app/ps/`) that list folders and copy the ones you pick into `.cache/box/` (incrementally — unchanged files are skipped).
- Copied files are extracted (PDF / Word / Excel / text) and chunked into a per-folder SQLite index (`.cache/index/`). Retrieval is local BM25 — nothing needs an API key.
- Questions go to `claude -p` with the most relevant excerpts and strict citation instructions. Claude is restricted to the Read tool, scoped to the synced folder, so it can also open scanned PDFs and images directly when they're relevant.

## Setup (new machine)

Requirements: Windows with **Box Drive** installed and signed in, **WSL** (Ubuntu or similar), and a Claude Code login.

```bash
bash setup.sh    # installs Python 3 + venv if missing, deps, and the claude CLI if missing
bash run.sh      # starts the app and opens http://localhost:8712
```

Notes:
- The Box root is auto-detected as `%USERPROFILE%\Box`. If yours lives elsewhere: `export BOX_ROOT='D:\Path\To\Box'` before `run.sh`.
- If `claude` was just installed by setup, run `claude` once first to log in.
- Keep the project on a Windows drive (e.g. `/mnt/c/...`) — the PowerShell helpers and the sync cache must be visible from Windows.

## Using it

1. Browse Box folders in the left pane (click a name to drill down).
2. Click **Index** on a folder — progress shows the file currently being copied/indexed. First run downloads the folder's files from Box; re-indexing is fast.
3. Select the folder under **Indexed folders** and ask questions in the chat. Answers cite `[filename, p.X]` and show source chips underneath.

File handling: PDFs are indexed per page; fully scanned PDFs and images aren't OCR'd but can be opened by Claude when relevant. Legacy `.doc/.xls` and design files (`.indd/.psd/.ai`) are listed but skipped. Files over 50 MB are excluded from sync.

## Deal pipeline dashboard

`http://localhost:8712/dashboard` (or the 📊 button in the app) tracks every property folder through the deal workflow from `.Finalize Checklist/Deal Timeline 1.docx`: **Listing → Prep & Marketing → LOI → PSA/Contract → Due Diligence/Escrow → Closing**.

- **Scan now** lists each property folder's file names (nothing is downloaded) and matches them against per-stage evidence (e.g. `*LOI*`, `*estoppel*`, `*settlement*`). A property's stage is the furthest one with evidence; hover the evidence chips to see which files matched.
- Stat tiles, a stage funnel chart, and a category chart aggregate the filtered set; filter by category, stage, or search.
- Template noise is ignored: every property folder is a copy of the brokerage "Continuum" tree (`Agent Information`, MREIS blank forms, `*TEMPLATE*`, checklists, design assets under `Images & Resources` / `Links`), so only real file names count as evidence and folder names like `Phase 5 - Closing` are never used. Tune the rules in `app/status.py` (`ITEMS`, `NOISE_*`).
- **What's missing** shows the most commonly missing checklist items (chart) and each property's gaps (list, click to jump to its checklist); tick "only properties with missing docs" to filter the whole dashboard to them.
- Click a property name to open its **Deal Timeline checklist** — every item expected by its stage marked ✓ (with the matching file) or ✗ missing — alongside its critical dates. The **Missing** column counts the ✗ items.
- **📅 Extract dates** has Claude read each LOI-or-later property's deal documents (PSA, LOI, settlement statement, estoppel, or a filled-in Critical Dates Timeline; `.docx` is converted to text first) and pull the Critical Dates from your Deal Timeline: effective date, deposit, DD/inspection end, deposit non-refundable, financing contingency, closing. Results feed the **Upcoming deadlines** panel (overdue → critical, ≤7 days → serious, ≤14 → warning) and each property's detail view. Cached in `.cache/status/dates.json`; unchanged properties are skipped on re-runs.
- The stage dropdown on each row sets a **manual override** (marked "manual"; choose "auto" to go back to detection). Overrides persist in `.cache/status/`.
- Categories (Urgent Care / Dental Care / Retail / Other) resolve from the `0-*` folder a property sits in, falling back to the name mapping in `app/categories.py`.

## Engagement advisor — action items

The **Action items** card lists what each engagement needs next, derived from its stage, missing Deal Timeline documents, critical dates, and recent email: overdue/imminent dates are critical ("DD ends in 5d — deposit goes non-refundable"), stage gaps are high/medium ("under contract, no estoppel on file"), quiet deals get a check-in nudge, and pending email proposals surface as reviews. Rules live in `app/advisor.py`. The **🧭 advise** button asks Claude for a deeper prioritized read of one engagement (cached until its state changes). Closed deals only get low-priority housekeeping.

## Outreach — AI-drafted, human-approved email

The **✍ draft email** buttons in Action items ask Claude to write stage-appropriate outreach for an engagement — it picks the right recipient from the deal's actual email contacts (derived from matched threads), replies inside an existing thread when that's the natural place, and cites real documents and dates. Drafts land in the **Outreach drafts** card where every field (To/Cc/Subject/Body) is editable; **nothing sends until you click Approve & send** (confirm dialog shows the recipient), or choose **Save to Outlook Drafts** to send from Outlook yourself, or Discard. Sending goes through classic Outlook COM — no API — from your own account, threaded into the conversation. History is kept in `.cache/outlook/outreach.json`.

## Outlook hook (no API) — email → workflow with go/no-go approval

The dashboard's **✉ Check mail** pulls recent Inbox + Sent mail from **classic Outlook** through its local COM automation (no Graph API, no app registration), then:

1. **Categorizes each email like the Box folders**: matches it to a property by name/address (brand + street number, exact folder name, distinctive location words; known senders learned from your approvals; replies inherit their thread's property) and tags the document type from attachment/subject names using the Deal Timeline items. Ambiguous or non-deal mail is left alone.
2. **Files matched Inbox mail into Outlook subfolders** `Inbox\Deals\<Category>\<Property>` (created on demand). Sent mail and unmatched mail stay where they are. Turn off with `file_into_folders` in `.cache/outlook/state.json` settings.
3. **Proposes workflow changes** for review — a stage advance (only for executed, non-draft documents), critical-date updates (Claude reads the email + attachments), and filing the attachment into the property's Box folder under `Inbox Filing\<Doc type>`. **Nothing is applied until you click Approve (go)**; Reject (no-go) discards it. Approved dates are locked against later re-extraction.

Setup (one time): the user's mailbox must exist in **classic** Outlook — new Outlook has no automation interface. Open classic Outlook → File → Add Account → sign in, let it sync, and leave it running minimized (the bridge will start it if needed). Auto-check every 15/60 min is a header toggle. Everything pulled stays local under `.cache/outlook/` (gitignored); Claude only reads property-matched deal mail.

## Configuration

| Env var | Default | Meaning |
|---|---|---|
| `BOX_ROOT` | `%USERPROFILE%\Box` | Windows path of the Box Drive root |
| `VENV_DIR` | `~/.venvs/onboarding` | Python virtualenv location (WSL-native for speed) |
| `QA_MODEL` | `sonnet` | Model passed to `claude -p` |

## Terminal smoke test

```bash
~/.venvs/onboarding/bin/python -m app.qa <slug> "What are these documents about?"
```

(slugs are the `.db` filenames in `.cache/index/`)
