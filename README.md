# weekly-report

Editor that turns a Markdown weekly report into Outlook-safe HTML email.
Write Markdown on the left, see the rendered email on the right, click
**Copy email**, paste into your mail compose window, send.

Two frontends over the same report format:

- `docs/`: static web version for GitHub Pages, and the one that gets new
  features (settings, toolbar, history charts, `.eml` export, import/export).
  Runs fully client-side (marked + highlight.js, vendored), so no report
  content ever leaves your machine. Reports live either in a folder on disk
  (File System Access API, Chrome/Edge) or in browser storage (any current
  browser, including Firefox and Safari). Installable and works offline.
- `app.py`: the original local Flask app; it keeps the default section
  names and email style and ignores the web version's settings.

## Deploy the web version

The repo must be public for free GitHub Pages; `weekly-reports/` is
git-ignored so report content is never committed.

```sh
gh repo create weekly-report --public --source . --push
```

Then on GitHub: Settings → Pages → Deploy from a branch → `main`, folder
`/docs`. The app appears at `https://<user>.github.io/weekly-report/`.
On first visit, pick your `weekly-reports/` folder (the browser remembers
it; with Chrome's "allow on every visit" it opens straight away) or choose
browser storage. **Storage…** switches between the two, imports a folder
or `.zip`, and exports everything as a `.zip` (the backup for browser
storage, which lives in that one browser only).

Bump the `?v=` on the script tags in `docs/index.html` together with
`APP_VERSION` in `docs/app.js` on every change so Pages never serves a stale
script.

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

## Project name - 2 days
> optional subtitle

This week:
- what happened
    - a detail underneath
        - and one level deeper

Status:
- Ongoing

Upcoming:
- what comes next

## Others - 0.5 days
- Weekly meeting
```

- ` - N days` ends each `##` heading (an em or en dash works too); the
  total is summed automatically.
  An optional `days:` in the frontmatter is checked against the sum;
  without it the web version checks against the working days per week
  from Settings (default 5).
- `productivity:` is free text; the row is omitted when empty.
- `This week:` / `Status:` / `Upcoming:` labels are styled automatically.
  The web version can rename them in Settings; the English names keep
  working, so older reports stay valid.
- Bullets nest three levels: • then ◦ then ▪ (numbers go 1. then a. then i.).
  Press Tab to make a sub-bullet, Shift+Tab to come back out; Enter on an
  empty sub-bullet also steps out one level. Indent by four spaces if you
  type it by hand: two nests in the web version but not in the Flask one.
- Emphasis: select a word and press Ctrl+B, Ctrl+I or Ctrl+U, which writes
  `**bold**`, `*italic*` and `<u>underline</u>`. Markdown has no underline
  syntax, so that one is a literal HTML tag; both renderers pass it through
  and Outlook honors it. Pressing the same key again removes the marks.
- Paste or drag an image into the editor to attach it; fenced code blocks
  get inline syntax highlighting.
- Images are shown 600 px wide but embedded at up to 1200, so they stay sharp
  on high-DPI screens; anything already under 1200 px goes in untouched. That
  costs email size: the header warns past ~95 KB, where Gmail starts
  clipping. The file in `images/` is always the untouched original.

## Settings (web)

Name, greeting, signature, subject template (`{week}`, `{year}`, `{name}`),
To/Cc for the `.eml` download, section names, email color theme, density,
and working days per week. Stored as `.weekly-report.json` next to the
reports, so they sync with the folder.

## History

The **History** button aggregates every report in the folder: a weekly
stacked chart of days per project (click a column to open that report),
days per project (totals, share, meter) and a chronological timeline of
each project's *This week* entries, useful for quarterly summaries and
appraisals. Filter by year or quarter (from `YYYY-Www.md` file names),
search all reports, download the days as CSV, or **Copy summary** for a
per-project write-up of the period. Click a project row to see its
timeline; Esc closes.

Project identity ignores case and extra whitespace; for real renames,
select the project and use **Same project as… → Merge**; mappings are
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
live total; hovering a project shows − / + (half a day) and arrows to
move it (also Alt+Up/Down inside the editor), clicking a name jumps to
it. The web editor has a toolbar (bold, italic, underline, lists, link,
code, image, new project), continues bullet lists on Enter, indents with
Tab/Shift+Tab, bolds/italicizes/underlines with Ctrl+B/I/U, and keeps all
of it on the undo stack; press `?` (or the ? button) for the full
shortcut list.

**New week** scaffolds the next report from the latest one: projects and
subtitles carry over, last week's *Upcoming* items become the draft of
*This week* (delete what didn't happen), *Status* and *Upcoming* are
kept, day counts reset to 0.

**Copy email** (Ctrl+Shift+C) puts the rendered HTML on the clipboard with
images embedded; the compose window converts them to proper inline
attachments on paste. **Subject** copies the subject line. **.eml**
downloads the email with To/Cc/Subject filled in and images as inline
attachments; classic Outlook opens it as a ready draft. Send one to
yourself first to sanity-check rendering in your client.

## Use it

The hosted web version (no install, nothing to sign up for):

**https://ybkahveci.github.io/weekly-report/**

Open it in any current browser. In Chrome or Edge you can pick a folder for
your reports; elsewhere they stay in the browser's own storage (export a
`.zip` backup from **Storage…**). Everything runs in the browser and your
reports stay on your own machine.
