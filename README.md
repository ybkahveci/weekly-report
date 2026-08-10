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
- Paste or drag an image into the editor to attach it; fenced code blocks
  get inline syntax highlighting.

## Weekly workflow

The outline sidebar lists projects in order: the arrows (or Alt+Up/Down
inside the editor) move a project, clicking a name jumps to it.

**New week** scaffolds the next report from the latest one: projects and
subtitles carry over, each project's *Status* and *Upcoming* are kept,
*This week* is cleared, day counts reset to 0.

**Copy email** puts the rendered HTML on the clipboard with images embedded;
the compose window converts them to proper inline attachments on paste.
Send one to yourself first to sanity-check rendering in your client.
