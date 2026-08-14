# weekly-report

Editor that turns a Markdown weekly report into Outlook-safe HTML email.
Write Markdown on the left, see the rendered email on the right, click
**Copy email**, paste into your mail compose window, send.

Two equivalent frontends:

- `app.py` — local Flask app.
- `docs/` — static web version for GitHub Pages. Runs fully client-side
  (marked + highlight.js, vendored); reads and writes the same
  `weekly-reports/` folder through the File System Access API
  (Chrome/Edge), so no report content ever leaves your machine.

## Deploy the web version

The repo must be public for free GitHub Pages; `weekly-reports/` is
git-ignored so report content is never committed.

```sh
gh repo create weekly-report --public --source . --push
```

Then on GitHub: Settings → Pages → Deploy from a branch → `main`, folder
`/docs`. The app appears at `https://<user>.github.io/weekly-report/`.
On first visit, pick your `weekly-reports/` folder; the browser remembers
the choice.

## Setup

```sh
python3 -m venv .venv
.venv/bin/pip install flask markdown pygments pillow
```

## Run

```sh
.venv/bin/python app.py            # http://127.0.0.1:8765
.venv/bin/python app.py --dir PATH --port N
```

## Report format

One `.md` file per report in `weekly-reports/`:

```markdown
---
week: 33
productivity: high
---

## Project name — 2 days
> optional subtitle

This week:
- what happened
    - a detail underneath
        - and one level deeper

Status:
- Ongoing

Upcoming:
- what comes next

## Others — 0.5 days
- Weekly meeting
```

- `— N days` in each `##` heading; the total is summed automatically.
  An optional `days:` in the frontmatter is checked against the sum.
- `productivity:` is free text; the row is omitted when empty.
- `This week:` / `Status:` / `Upcoming:` labels are bolded automatically.
- Bullets nest three levels — • then ◦ then ▪ (numbers go 1. then a. then i.).
  Press Tab to make a sub-bullet, Shift+Tab to come back out; Enter on an
  empty sub-bullet also steps out one level. Indent by four spaces if you
  type it by hand: two nests in the web version but not in the Flask one.
- Emphasis: select a word and press Ctrl+B, Ctrl+I or Ctrl+U — which writes
  `**bold**`, `*italic*` and `<u>underline</u>`. Markdown has no underline
  syntax, so that one is a literal HTML tag; both renderers pass it through
  and Outlook honors it. Pressing the same key again removes the marks.
- Paste or drag an image into the editor to attach it; fenced code blocks
  get inline syntax highlighting.
- Images are shown 600 px wide but embedded at up to 1200, so they stay sharp
  on high-DPI screens; anything already under 1200 px goes in untouched. That
  costs email size — the header warns past ~95 KB, where Gmail starts
  clipping. The file in `images/` is always the untouched original.

## History

The **History** button aggregates every report in the folder: days per
project (totals, share, meter) and a chronological timeline of each
project's *This week* entries — useful for quarterly summaries and
appraisals. Click a project row to see its timeline; Esc closes.

Project identity ignores case and extra whitespace; for real renames,
select the project and use **Same project as… → Merge** — mappings are
stored in `.project-aliases.json` next to the reports (so they sync
through git). Projects without This week/Status/Upcoming labels (like
Others) show their whole content in the timeline. The pencil button
next to the file selector renames the current report file.

## Sync reports between machines

`weekly-reports/` is git-ignored here; make it its own repo pushed to a
**private** remote:

```sh
cd weekly-reports
git init -b main && git add . && git commit -m "import reports"
git remote add origin git@github.com:<you>/weekly-reports.git
git push -u origin main
```

On the other machine: clone it anywhere and point the web app at that
folder. Routine: `git pull --rebase` before writing, commit and push
after.

## Weekly workflow

The outline sidebar lists projects in order with their day counts and a
live total: the arrows (or Alt+Up/Down inside the editor) move a project,
clicking a name jumps to it. The web editor continues bullet lists on
Enter, indents with Tab/Shift+Tab, and bolds/italicizes/underlines with
Ctrl+B/I/U — press `?` (or the ? button) for the full shortcut list.

**New week** scaffolds the next report from the latest one: projects and
subtitles carry over, each project's *Status* and *Upcoming* are kept,
*This week* is cleared, day counts reset to 0.

**Copy email** puts the rendered HTML on the clipboard with images embedded;
the compose window converts them to proper inline attachments on paste.
Send one to yourself first to sanity-check rendering in your client.

## Use it

The hosted web version — no install, nothing to sign up for:

**https://ybkahveci.github.io/weekly-report/**

Open it in Chrome or Edge (the File System Access API is required), pick a
folder for your reports on first visit, and start writing. Everything runs
in the browser and your reports stay on your own machine.
