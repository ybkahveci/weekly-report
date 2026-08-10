#!/usr/bin/env python3
# =============================================================================
# weekly_report — local web editor that renders Markdown weekly reports as
# Outlook-safe HTML email.
#
# Classic Outlook renders mail with the Word engine: no <style> blocks, no
# flexbox, poor CSS inheritance. Everything is therefore emitted as tables
# with styles inlined on every element. Images are base64-embedded only in
# the copied email: compose windows convert pasted data URIs into real inline
# attachments, while received mail would strip them.
#
# Copyright (c) 2026 Yusuf Bunyamin Kahveci <yusuf.kahveci@tii.ae>
# SPDX-License-Identifier: Apache-2.0
# =============================================================================

import argparse
import base64
import datetime as dt
import io
import re
from dataclasses import dataclass, field
from html import escape
from pathlib import Path

import markdown
from flask import Flask, Response, abort, jsonify, request, send_from_directory

try:
    from PIL import Image
except ImportError:
    Image = None

REPORTS_DIR = Path(__file__).resolve().parent / "weekly-reports"

# Content column is 640 - 150 (name column) = 490 px wide.
MAX_IMG_WIDTH = 480
TAB_PX = 36

FONT = "Calibri,Arial,Helvetica,sans-serif"
MONO = "Consolas,'Courier New',monospace"
NAVY = "#1f3864"
BORDER = "#d9d9d9"
RULE = "#e3e6ea"
MUTED = "#7a8699"
TD_FONT = f"font-family:{FONT};font-size:15px;line-height:1.45;color:#1a1a1a;"
SECTION_LABEL = (
    f"font-family:{FONT};font-size:11px;font-weight:bold;letter-spacing:1px;"
    f"color:{MUTED};"
)

TAG_STYLES = {
    "p": "margin:0 0 8px 0;",
    "ul": "margin:2px 0 10px 0;padding:0 0 0 24px;",
    "ol": "margin:2px 0 10px 0;padding:0 0 0 24px;",
    "li": "margin:0 0 3px 0;",
    "blockquote": (
        "margin:4px 0 10px 0;padding:2px 0 2px 10px;"
        "border-left:3px solid #cccccc;color:#555555;"
    ),
    "h1": f"font-size:17px;margin:12px 0 6px 0;color:{NAVY};",
    "h2": f"font-size:16px;margin:12px 0 6px 0;color:{NAVY};",
    "h3": f"font-size:15px;margin:10px 0 4px 0;color:{NAVY};",
    "h4": "font-size:15px;margin:10px 0 4px 0;",
}
A_STYLE = "color:#1f6fb2;"
CODE_INLINE_STYLE = (
    f"font-family:{MONO};font-size:13px;background-color:#f5f5f5;padding:0 3px;"
)

MD_EXTENSIONS = ["fenced_code", "sane_lists", "nl2br", "codehilite"]
MD_CONFIGS = {"codehilite": {"noclasses": True, "pygments_style": "default"}}

LABEL_RE = re.compile(r"^(this week|status|upcoming)\s*:(.*)$", re.IGNORECASE)
LIST_RE = re.compile(r"^\s*(?:[-*+]\s|\d+\.\s)")
HEAD_DAYS_RE = re.compile(
    r"^(?P<name>.+?)\s+[—–-]+\s+(?P<days>\d+(?:[.,]\d+)?)\s*days?\s*$", re.IGNORECASE
)
CODEBLOCK_RE = re.compile(
    r'<div class="codehilite"[^>]*>\s*<pre[^>]*>'
    r"(?:<span></span>)?(?:<code[^>]*>)?(.*?)(?:</code>)?</pre>\s*</div>",
    re.DOTALL,
)
IMG_RE = re.compile(r"<img\b[^>]*>")
NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._&-]*\.md$")

IMAGE_MIMES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
}


@dataclass
class Project:
    name: str
    days: float | None
    subtitle: str | None
    body_md: str


@dataclass
class Report:
    meta: dict[str, str]
    projects: list[Project]
    warnings: list[str] = field(default_factory=list)


def parse_report(text: str) -> Report:
    meta: dict[str, str] = {}
    body = text
    if text.startswith("---"):
        end = text.find("\n---", 3)
        if end != -1:
            for line in text[3:end].strip().splitlines():
                if ":" in line:
                    key, value = line.split(":", 1)
                    meta[key.strip().lower()] = value.strip()
            body = text[end + 4 :]

    raw: list[dict] = []
    current: dict | None = None
    for line in body.splitlines():
        if line.startswith("## "):
            current = {"head": line[3:].strip(), "lines": []}
            raw.append(current)
        elif current is not None:
            current["lines"].append(line)

    report = Report(meta, [])
    for item in raw:
        match = HEAD_DAYS_RE.match(item["head"])
        if match:
            name = match.group("name").strip()
            days = float(match.group("days").replace(",", "."))
        else:
            name, days = item["head"], None
            report.warnings.append(f'No "— N days" found in heading "{item["head"]}"')
        lines = item["lines"]
        while lines and not lines[0].strip():
            lines.pop(0)
        subtitle = None
        if lines and lines[0].lstrip().startswith(">"):
            subtitle = lines.pop(0).lstrip()[1:].strip()
        report.projects.append(
            Project(name, days, subtitle, "\n".join(lines).strip("\n"))
        )
    return report


def preprocess_body(text: str) -> str:
    """python-markdown ignores a list glued to a paragraph line; insert the
    blank line the author almost certainly meant."""
    lines = text.split("\n")
    out: list[str] = []
    in_fence = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith(("```", "~~~")):
            in_fence = not in_fence
            out.append(line)
            continue
        if in_fence:
            out.append(line)
            continue
        out.append(line)
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        if stripped and not LIST_RE.match(line) and LIST_RE.match(nxt):
            out.append("")
    return "\n".join(out)


def split_body(text: str) -> list[tuple[str | None, str]]:
    """Split a project body into ordered (label, content) segments on the
    This week/Status/Upcoming labels; label is None for unlabeled content."""
    segments: list[tuple[str | None, str]] = []
    label: str | None = None
    current: list[str] = []
    in_fence = False

    def flush() -> None:
        content = "\n".join(current).strip("\n")
        if content or label is not None:
            segments.append((label, content))

    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.startswith(("```", "~~~")):
            in_fence = not in_fence
        match = None if in_fence or LIST_RE.match(line) else LABEL_RE.match(stripped)
        if match:
            flush()
            label = match.group(1).lower()
            current = [match.group(2).strip()] if match.group(2).strip() else []
        else:
            current.append(line)
    flush()
    return segments


SPAN_STYLE_RE = re.compile(r'<span style="([^"]*)">(.*?)</span>', re.DOTALL)


def _legacy_span(match: re.Match) -> str:
    """Pygments token span -> font/b/i, which survive an Outlook paste
    where inline CSS colors do not."""
    style, inner = match.group(1), match.group(2)
    prefix, suffix = "", ""
    color = re.search(r"color:\s*(#[0-9a-fA-F]{6})", style)
    if color:
        prefix, suffix = f'<font color="{color.group(1).lower()}">', "</font>"
    if "bold" in style:
        prefix, suffix = prefix + "<b>", "</b>" + suffix
    if "italic" in style:
        prefix, suffix = prefix + "<i>", "</i>" + suffix
    return f"{prefix}{inner}{suffix}"


def harden_code_blocks(html: str) -> str:
    """Word ignores white-space CSS, so newlines become <br> and space runs
    become &nbsp; to keep code formatting intact."""

    def repl(match: re.Match) -> str:
        code_html = SPAN_STYLE_RE.sub(_legacy_span, match.group(1))
        lines = []
        for line in code_html.strip("\n").split("\n"):
            line = re.sub(r"^ +", lambda m: "&nbsp;" * len(m.group()), line)
            line = re.sub(r"  +", lambda m: "&nbsp;" * len(m.group()), line)
            lines.append(line)
        code = "<br>".join(lines)
        return (
            '<table cellpadding="0" cellspacing="0" border="0" width="100%"'
            ' style="margin:4px 0 10px 0;"><tr>'
            f'<td bgcolor="#F5F5F5" style="background-color:#f5f5f5;'
            f"border:1px solid {BORDER};padding:8px 10px;"
            f'font-family:{MONO};font-size:13px;line-height:1.4;">'
            f"{code}</td></tr></table>"
        )

    return CODEBLOCK_RE.sub(repl, html)


def inject_styles(html: str) -> str:
    html = re.sub(
        r"<(p|ul|ol|li|blockquote|h1|h2|h3|h4)>",
        lambda m: f'<{m.group(1)} style="{TAG_STYLES[m.group(1)]}">',
        html,
    )
    html = html.replace("<a ", f'<a style="{A_STYLE}" ')
    html = html.replace("<code>", f'<code style="{CODE_INLINE_STYLE}">')
    return html


def _load_local_image(path: Path) -> tuple[str, int | None]:
    """Return (data URI, display width). Downscales when Pillow is present so
    the base64 payload stays under Gmail's ~100 KB body clip."""
    data = path.read_bytes()
    mime = IMAGE_MIMES[path.suffix.lower()]
    width = None
    if Image is not None and mime != "image/gif":
        img = Image.open(path)
        width = min(img.width, MAX_IMG_WIDTH)
        if img.width > MAX_IMG_WIDTH:
            ratio = MAX_IMG_WIDTH / img.width
            img = img.resize((MAX_IMG_WIDTH, round(img.height * ratio)), Image.LANCZOS)
            buf = io.BytesIO()
            if mime == "image/jpeg":
                img.save(buf, "JPEG", quality=85)
            else:
                img.convert("RGBA").save(buf, "PNG")
            data = buf.getvalue()
    return f"data:{mime};base64,{base64.b64encode(data).decode()}", width


def process_images(html: str, embed: bool, warnings: list[str]) -> str:
    def repl(match: re.Match) -> str:
        tag = match.group(0)
        src_m = re.search(r'src="([^"]*)"', tag)
        alt_m = re.search(r'alt="([^"]*)"', tag)
        src = src_m.group(1) if src_m else ""
        alt = alt_m.group(1) if alt_m else ""
        width: int | None = None

        if src.startswith(("http://", "https://")):
            if embed:
                warnings.append(f"Remote image kept as a link (may be blocked): {src}")
        else:
            path = (REPORTS_DIR / src).resolve()
            if REPORTS_DIR not in path.parents or not path.is_file():
                warnings.append(f"Image not found: {src}")
                return (
                    f'<span style="color:#b00;">[missing image: {escape(src)}]</span>'
                )
            if path.suffix.lower() not in IMAGE_MIMES:
                warnings.append(f"Unsupported image type: {src}")
                return (
                    '<span style="color:#b00;">'
                    f"[unsupported image: {escape(src)}]</span>"
                )
            if embed:
                src, width = _load_local_image(path)
            else:
                if Image is not None:
                    with Image.open(path) as img:
                        width = min(img.width, MAX_IMG_WIDTH)
                src = "/" + src.lstrip("/")

        width_attr = f' width="{width}"' if width else ""
        return (
            f'<img src="{src}" alt="{escape(alt)}"{width_attr} '
            'style="display:block;max-width:100%;height:auto;border:0;margin:6px 0;">'
        )

    return IMG_RE.sub(repl, html)


def _md_fragment(text: str, embed: bool, warnings: list[str]) -> str:
    html = markdown.markdown(
        preprocess_body(text),
        extensions=MD_EXTENSIONS,
        extension_configs=MD_CONFIGS,
    )
    html = harden_code_blocks(html)
    html = inject_styles(html)
    return process_images(html, embed, warnings)


def _italicize(frag: str) -> str:
    """Both CSS and legacy <i>: Outlook drops font-style on paste, and Word
    does not inherit font-style into list items either."""
    frag = frag.replace('<p style="', '<p style="font-style:italic;')
    frag = frag.replace('<li style="', '<li style="font-style:italic;')
    frag = re.sub(r'(<(?:p|li) style="[^"]*">)', r"\1<i>", frag)
    return re.sub(r"</(p|li)>", r"</i></\1>", frag)


def _indent(frag: str) -> str:
    """Spacer-cell table: the only indent the Word engine renders reliably."""
    if not frag:
        return frag
    return (
        '<table cellpadding="0" cellspacing="0" border="0" width="100%">'
        f'<tr><td width="{TAB_PX}" style="font-size:1px;line-height:1px;">'
        "&nbsp;</td>"
        f'<td valign="top" style="{TD_FONT}">{frag}</td></tr></table>'
    )


def render_body(body_md: str, embed: bool, warnings: list[str]) -> str:
    if not body_md.strip():
        return ""
    parts: list[str] = []
    first = True
    for label, content in split_body(body_md):
        if label is not None:
            top = "0" if first else "12px"
            parts.append(
                f'<p style="{SECTION_LABEL}margin:{top} 0 4px 0;">'
                f'<b><font color="{MUTED}">{label.upper()}</font></b></p>'
            )
        first = False
        frag = _md_fragment(content, embed, warnings) if content.strip() else ""
        if label == "this week":
            frag = _italicize(frag)
        elif label == "status" and not any(
            # bullet lists indent themselves; only plain text needs the tab
            LIST_RE.match(ln)
            for ln in content.splitlines()
        ):
            frag = _indent(frag)
        parts.append(frag)
    return "".join(parts)


def format_days(days: float) -> str:
    return f"{days:g} {'day' if days == 1 else 'days'}"


def render_email(report: Report, embed: bool) -> tuple[str, list[str]]:
    warnings = list(report.warnings)
    total = sum(p.days for p in report.projects if p.days is not None)

    declared = report.meta.get("days")
    if declared:
        try:
            if abs(float(declared.replace(",", ".")) - total) > 1e-9:
                warnings.append(
                    f"Frontmatter days: {declared} != sum of project days ({total:g})"
                )
        except ValueError:
            warnings.append(f"Frontmatter days is not a number: {declared}")

    summary = [
        ("Week", report.meta.get("week", "")),
        ("Total working days", f"{total:g}"),
    ]
    if report.meta.get("productivity"):
        summary.append(("Productivity", report.meta["productivity"]))
    cells = []
    for i, (label, value) in enumerate(summary):
        divider = f"border-right:1px solid {RULE};" if i < len(summary) - 1 else ""
        cells.append(
            f'<td bgcolor="#f2f4f8" style="{TD_FONT}background-color:#f2f4f8;'
            f'padding:10px 22px 10px 14px;{divider}">'
            f'<div style="{SECTION_LABEL}">'
            f'<b><font color="{MUTED}">{escape(label.upper())}</font></b></div>'
            f'<div style="font-size:16px;font-weight:bold;color:{NAVY};'
            f'margin:2px 0 0 0;">'
            f'<b><font color="{NAVY}">{escape(value)}</font></b></div></td>'
        )
    header = (
        '<table width="640" cellpadding="0" cellspacing="0" border="0"'
        f' style="margin:0 0 16px 0;border:1px solid {RULE};">'
        f"<tr>{''.join(cells)}</tr></table>"
    )

    cell_border = f"border-bottom:1px solid {RULE};"
    rows = [
        f'<tr><td colspan="2" bgcolor="{NAVY}" style="{TD_FONT}color:#ffffff;'
        "font-weight:bold;font-size:12px;letter-spacing:1px;"
        'padding:8px 12px;"><b><font color="#ffffff">PROJECTS</font></b></td></tr>'
    ]
    for p in report.projects:
        subtitle = (
            '<div style="font-size:12px;font-weight:normal;font-style:italic;'
            f'color:{MUTED};margin:4px 0 0 0;">'
            f'<i><font color="{MUTED}">{escape(p.subtitle)}</font></i></div>'
            if p.subtitle
            else ""
        )
        days_txt = format_days(p.days) if p.days is not None else "—"
        days = (
            f'<div style="font-size:12px;font-weight:normal;color:{MUTED};'
            f'margin:5px 0 0 0;"><font color="{MUTED}">{days_txt}</font></div>'
        )
        rows.append(
            f'<tr><td valign="top" width="150" style="{TD_FONT}font-weight:bold;'
            f"color:{NAVY};padding:14px 10px 10px 12px;{cell_border}"
            f'border-right:1px solid {RULE};">'
            f'<b><font color="{NAVY}">{escape(p.name)}</font></b>'
            f"{subtitle}{days}</td>"
            f'<td valign="top" style="{TD_FONT}padding:14px 12px 10px 14px;'
            f'{cell_border}">'
            f"{render_body(p.body_md, embed, warnings)}</td></tr>"
        )

    html = (
        '<meta charset="utf-8">'
        f'<div style="{TD_FONT}">'
        '<p style="margin:0 0 8px 0;">Dear all,</p>'
        '<p style="margin:0 0 14px 0;">This is my weekly progress report.</p>'
        f"{header}"
        '<table width="640" cellpadding="0" cellspacing="0" border="0"'
        ' style="border-collapse:collapse;">' + "".join(rows) + "</table></div>"
    )
    return html, warnings


def _sections(body_md: str) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    current: str | None = None
    for line in body_md.splitlines():
        match = LABEL_RE.match(line.strip())
        if match and not LIST_RE.match(line):
            current = match.group(1).lower()
            sections[current] = []
            if match.group(2).strip():
                sections[current].append(match.group(2).strip())
        elif current is not None:
            sections[current].append(line)
    for lines in sections.values():
        while lines and not lines[-1].strip():
            lines.pop()
    return sections


def scaffold(previous: str | None, week: int) -> str:
    out = ["---", f"week: {week}", "productivity:", "---", ""]
    if previous is None:
        out += [
            "## Project name — 0 days",
            "This week:",
            "- ",
            "",
            "Status:",
            "- Ongoing",
            "",
            "Upcoming:",
            "- ",
            "",
            "## Others — 0 days",
            "- Weekly meeting",
            "- Weekly report",
            "",
        ]
        return "\n".join(out)

    for p in parse_report(previous).projects:
        out.append(f"## {p.name} — 0 days")
        if p.subtitle:
            out.append(f"> {p.subtitle}")
        out.append("")
        sections = _sections(p.body_md)
        if sections:
            out += ["This week:", "- ", ""]
            out += ["Status:"] + (sections.get("status") or ["- Ongoing"]) + [""]
            out += ["Upcoming:"] + (sections.get("upcoming") or ["- "]) + [""]
        elif p.body_md:
            out += [p.body_md, ""]
    return "\n".join(out).rstrip() + "\n"


app = Flask(__name__)


def _report_path(name: str | None) -> Path:
    if not name or not NAME_RE.match(name):
        abort(400, "invalid report name")
    path = REPORTS_DIR / name
    if not path.is_file():
        abort(404, f"no such report: {name}")
    return path


@app.get("/")
def editor() -> Response:
    return Response(EDITOR_HTML, mimetype="text/html")


@app.get("/api/files")
def list_files() -> Response:
    return jsonify(sorted((p.name for p in REPORTS_DIR.glob("*.md")), reverse=True))


@app.get("/api/file")
def get_file() -> Response:
    text = _report_path(request.args.get("name")).read_text(encoding="utf-8")
    return Response(text, mimetype="text/plain; charset=utf-8")


@app.post("/api/save")
def save_file() -> Response:
    data = request.get_json(silent=True) or {}
    name, content = data.get("name"), data.get("content")
    if not name or not NAME_RE.match(name) or content is None:
        abort(400, "invalid save request")
    (REPORTS_DIR / name).write_text(content, encoding="utf-8")
    return jsonify(ok=True)


@app.post("/api/new")
def new_week() -> Response:
    iso = dt.date.today().isocalendar()
    name = f"{iso.year}-W{iso.week:02d}.md"
    path = REPORTS_DIR / name
    if path.exists():
        abort(409, f"{name} already exists")
    existing = sorted(REPORTS_DIR.glob("*.md"))
    previous = existing[-1].read_text(encoding="utf-8") if existing else None
    path.write_text(scaffold(previous, iso.week), encoding="utf-8")
    return jsonify(name=name)


@app.post("/api/upload")
def upload_image() -> Response:
    file = request.files.get("image")
    if file is None:
        abort(400, "no image in request")
    ext = {"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif"}.get(
        file.mimetype
    )
    if ext is None:
        abort(415, f"unsupported image type: {file.mimetype}")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    images_dir = REPORTS_DIR / "images"
    images_dir.mkdir(parents=True, exist_ok=True)
    path = images_dir / f"img-{stamp}{ext}"
    counter = 1
    while path.exists():
        path = images_dir / f"img-{stamp}-{counter}{ext}"
        counter += 1
    file.save(path)
    return jsonify(markdown=f"![](images/{path.name})")


@app.get("/images/<path:filename>")
def serve_image(filename: str) -> Response:
    return send_from_directory(REPORTS_DIR / "images", filename)


@app.get("/preview")
def preview() -> Response:
    report = parse_report(
        _report_path(request.args.get("name")).read_text(encoding="utf-8")
    )
    email_html, warnings = render_email(report, embed=True)
    preview_html, _ = render_email(report, embed=False)
    size_kb = len(email_html.encode()) / 1024
    size_note = f"Email size: {size_kb:.0f} KB"
    if size_kb > 95:
        warnings.append("Email exceeds ~100 KB — Gmail may clip the message")
    banner = "".join(
        '<div style="background:#fff8e1;border:1px solid #f0dc9a;color:#6b5d1f;'
        "border-radius:4px;padding:7px 12px;margin:0 0 8px 0;"
        f'font:12.5px system-ui;">{escape(w)}</div>'
        for w in warnings
    )
    doc = (
        '<!doctype html><html><head><meta charset="utf-8"></head>'
        '<body style="margin:0;padding:18px;background:#eef1f5;">'
        '<div style="max-width:700px;margin:0 auto;">'
        f"{banner}"
        f'<div style="color:#8a94a6;font:12px system-ui;text-align:center;'
        f'margin:0 0 10px 0;">{size_note} — preview approximates Outlook rendering'
        "</div>"
        '<div style="background:#ffffff;border:1px solid #dce0e6;border-radius:6px;'
        f'padding:26px 30px;">{preview_html}</div>'
        "</div></body></html>"
    )
    return Response(doc, mimetype="text/html")


@app.get("/email")
def email() -> Response:
    report = parse_report(
        _report_path(request.args.get("name")).read_text(encoding="utf-8")
    )
    email_html, _ = render_email(report, embed=True)
    return Response(email_html, mimetype="text/html; charset=utf-8")


EDITOR_HTML = r"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Weekly Report</title>
<style>
* { box-sizing: border-box; }
body { margin: 0; height: 100vh; display: flex; flex-direction: column;
       font: 14px system-ui, sans-serif; }
header { display: flex; gap: 8px; align-items: center; padding: 8px 12px;
         background: #1c2128; color: #dfe3ea; }
.brand { font-weight: 600; letter-spacing: .3px; margin-right: 10px; }
header select, header button { font: inherit; padding: 5px 12px;
  background: #2c3340; color: #dfe3ea; border: 1px solid #3a4250;
  border-radius: 4px; cursor: pointer; }
header button:hover { background: #384050; }
.grow { flex: 1; }
#status { color: #8fa3c0; font-size: 13px; }
#copy { background: #2d6cdf; border-color: #2d6cdf; color: #fff;
        padding: 6px 18px; font-weight: 600; }
#copy:hover { background: #245bc0; }
main { flex: 1; display: flex; min-height: 0; }
#outline { width: 190px; background: #f7f8fa; border-right: 1px solid #dfe3e8;
           overflow-y: auto; padding: 4px 0 8px; }
.ohead { font-size: 10.5px; font-weight: 700; letter-spacing: 1px;
         color: #7a8699; text-transform: uppercase; padding: 8px 12px 4px; }
.oitem { display: flex; align-items: center; gap: 3px; padding: 4px 8px 4px 12px;
         font-size: 12.5px; border-radius: 4px; margin: 0 4px; }
.oitem:hover { background: #eceef2; }
.oitem span { flex: 1; overflow: hidden; text-overflow: ellipsis;
              white-space: nowrap; cursor: pointer; color: #2a3342; }
.oitem span:hover { color: #2d6cdf; }
.oitem button { border: 1px solid #d3d8df; background: #fff; border-radius: 3px;
                cursor: pointer; padding: 0 5px; font-size: 11px; color: #4a5568; }
.oitem button:hover { border-color: #2d6cdf; color: #2d6cdf; }
.oitem button:disabled { opacity: .3; cursor: default; }
#editor { flex: 1; border: none; outline: none; resize: none; padding: 16px;
          font: 13px/1.55 Consolas, monospace; background: #fdfdfd;
          color: #24292f; }
#preview { flex: 1; border: none; border-left: 1px solid #dfe3e8;
           background: #eef1f5; }
</style>
</head>
<body>
<header>
  <span class="brand">Weekly Report</span>
  <select id="file"></select>
  <button id="new">New week</button>
  <span id="status"></span>
  <span class="grow"></span>
  <button id="copy">Copy email</button>
</header>
<main>
  <aside id="outline"
         title="Project order — arrows move, Alt+Up/Down in the editor">
    <div class="ohead">Projects</div>
    <div id="olist"></div>
  </aside>
  <textarea id="editor" spellcheck="false"></textarea>
  <iframe id="preview"></iframe>
</main>
<script>
const $ = id => document.getElementById(id);
let current = null, saveTimer = null, dirty = false, lastEmail = null;

async function api(url, opts) {
  const r = await fetch(url, opts);
  if (!r.ok) throw new Error(await r.text());
  return r;
}
function setStatus(t) { $('status').textContent = t; }

async function loadList(preferred) {
  const files = await (await api('/api/files')).json();
  const el = $('file');
  el.innerHTML = '';
  for (const f of files) {
    const o = document.createElement('option');
    o.value = o.textContent = f;
    el.appendChild(o);
  }
  if (files.length) {
    current = preferred && files.includes(preferred) ? preferred : files[0];
    el.value = current;
    await openFile(current);
  } else {
    setStatus('No reports yet — click "New week"');
  }
}
async function openFile(name) {
  current = name;
  const r = await api('/api/file?name=' + encodeURIComponent(name));
  $('editor').value = await r.text();
  dirty = false;
  renderOutline();
  refresh();
}
function refresh() {
  const fr = $('preview');
  const y = fr.contentWindow ? fr.contentWindow.scrollY : 0;
  fr.onload = () => fr.contentWindow.scrollTo(0, y);
  fr.src = '/preview?name=' + encodeURIComponent(current) + '&t=' + Date.now();
  fetch('/email?name=' + encodeURIComponent(current))
    .then(r => r.ok ? r.text() : null)
    .then(t => { if (t) lastEmail = t; });
}
async function save() {
  if (!current) return;
  await api('/api/save', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: current, content: $('editor').value }),
  });
  dirty = false;
  setStatus('Saved');
  refresh();
}
function scheduleSave() {
  dirty = true;
  setStatus('…');
  renderOutline();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(
    () => save().catch(e => setStatus('Save failed: ' + e.message)), 600);
}

function splitSections(text) {
  const lines = text.split('\n');
  const sections = [];
  let fence = false, cur = null;
  lines.forEach((ln, i) => {
    const t = ln.trim();
    if (t.startsWith('```') || t.startsWith('~~~')) fence = !fence;
    if (!fence && ln.startsWith('## ')) {
      if (cur) cur.end = i;
      cur = { start: i, end: lines.length, name: ln.slice(3).trim() };
      sections.push(cur);
    }
  });
  return { lines, sections };
}
function lineToChar(lines, lineIdx) {
  let n = 0;
  for (let i = 0; i < lineIdx; i++) n += lines[i].length + 1;
  return n;
}
function moveSection(idx, dir, focusEditor) {
  const ed = $('editor');
  const { lines, sections } = splitSections(ed.value);
  const j = idx + dir;
  if (j < 0 || j >= sections.length) return;
  const lo = sections[Math.min(idx, j)], hi = sections[Math.max(idx, j)];
  const norm = arr => {
    const c = [...arr];
    while (c.length && !c[c.length - 1].trim()) c.pop();
    c.push('');
    return c;
  };
  const blockLo = norm(lines.slice(lo.start, lo.end));
  const blockHi = norm(lines.slice(hi.start, hi.end));
  const assembled = [
    ...lines.slice(0, lo.start), ...blockHi, ...blockLo, ...lines.slice(hi.end),
  ];
  ed.value = assembled.join('\n');
  const movedLine = dir < 0 ? lo.start : lo.start + blockHi.length;
  if (focusEditor) {
    const pos = lineToChar(assembled, movedLine);
    ed.focus();
    ed.setSelectionRange(pos, pos);
  }
  scheduleSave();
}
function jumpTo(startLine) {
  const ed = $('editor');
  const { lines } = splitSections(ed.value);
  const pos = lineToChar(lines, startLine);
  ed.focus();
  ed.setSelectionRange(pos, pos);
  ed.scrollTop = Math.max(0, (startLine / lines.length) * ed.scrollHeight - 40);
}
function renderOutline() {
  const { sections } = splitSections($('editor').value);
  const el = $('olist');
  el.innerHTML = '';
  sections.forEach((s, i) => {
    const row = document.createElement('div');
    row.className = 'oitem';
    const name = document.createElement('span');
    name.textContent = s.name.replace(/\s+[—–-]+\s+\d.*$/, '');
    name.onclick = () => jumpTo(s.start);
    const up = document.createElement('button');
    up.textContent = '↑';
    up.disabled = i === 0;
    up.onclick = () => moveSection(i, -1, false);
    const dn = document.createElement('button');
    dn.textContent = '↓';
    dn.disabled = i === sections.length - 1;
    dn.onclick = () => moveSection(i, 1, false);
    row.append(name, up, dn);
    el.appendChild(row);
  });
}
async function upload(file) {
  const fd = new FormData();
  fd.append('image', file);
  const r = await api('/api/upload', { method: 'POST', body: fd });
  const { markdown } = await r.json();
  const ed = $('editor');
  ed.setRangeText(markdown + '\n', ed.selectionStart, ed.selectionEnd, 'end');
  scheduleSave();
}

$('editor').addEventListener('input', scheduleSave);
$('file').addEventListener('change', e => openFile(e.target.value));
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 's') {
    e.preventDefault();
    clearTimeout(saveTimer);
    save().catch(err => setStatus('Save failed: ' + err.message));
  }
  if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
    const ed = $('editor');
    if (document.activeElement !== ed) return;
    const { sections } = splitSections(ed.value);
    const line = ed.value.slice(0, ed.selectionStart).split('\n').length - 1;
    const idx = sections.findIndex(s => line >= s.start && line < s.end);
    if (idx >= 0) {
      e.preventDefault();
      moveSection(idx, e.key === 'ArrowUp' ? -1 : 1, true);
    }
  }
});
$('new').addEventListener('click', async () => {
  try {
    const { name } = await (await api('/api/new', { method: 'POST' })).json();
    await loadList(name);
  } catch (e) { alert(e.message); }
});
function copyRich(html, plain) {
  // copy-event path writes the payload verbatim; the async clipboard API
  // sanitizes text/html and strips inline styles that Outlook needs
  let ok = false;
  const listener = e => {
    e.clipboardData.setData('text/html', html);
    e.clipboardData.setData('text/plain', plain);
    e.preventDefault();
    ok = true;
  };
  const ta = document.createElement('textarea');
  ta.value = ' ';
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  document.addEventListener('copy', listener);
  try {
    document.execCommand('copy');
  } finally {
    document.removeEventListener('copy', listener);
    document.body.removeChild(ta);
    $('editor').focus();
  }
  return ok;
}
$('copy').addEventListener('click', () => {
  // execCommand only works synchronously inside the click stack, so copy
  // the cached render (kept fresh by refresh()) with no awaits before it
  const plain = $('editor').value;
  if (!lastEmail) {
    setStatus('Nothing rendered yet');
    return;
  }
  if (copyRich(lastEmail, plain)) {
    setStatus('Copied — paste into your mail compose window');
  } else {
    navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([lastEmail], { type: 'text/html' }),
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })]).then(
      () => setStatus('Copied (fallback path — styling may degrade)'),
      e => setStatus('Copy failed: ' + e.message));
  }
  if (dirty) {
    clearTimeout(saveTimer);
    save().catch(e => setStatus('Save failed: ' + e.message));
  }
});
$('editor').addEventListener('paste', e => {
  for (const item of e.clipboardData.items) {
    if (item.type.startsWith('image/')) {
      e.preventDefault();
      upload(item.getAsFile()).catch(err => setStatus('Upload failed: ' + err.message));
      return;
    }
  }
});
$('editor').addEventListener('drop', e => {
  const f = [...e.dataTransfer.files].find(f => f.type.startsWith('image/'));
  if (f) {
    e.preventDefault();
    upload(f).catch(err => setStatus('Upload failed: ' + err.message));
  }
});

loadList().catch(e => setStatus(e.message));
</script>
</body>
</html>
"""


def main() -> None:
    global REPORTS_DIR
    parser = argparse.ArgumentParser(
        description="Weekly report editor — Markdown in, Outlook-safe HTML out"
    )
    parser.add_argument(
        "--dir",
        type=Path,
        default=REPORTS_DIR,
        help="reports directory (default: ./weekly-reports)",
    )
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    REPORTS_DIR = args.dir.resolve()
    (REPORTS_DIR / "images").mkdir(parents=True, exist_ok=True)

    print(f"Reports dir: {REPORTS_DIR}")
    print(f"Open http://127.0.0.1:{args.port}")
    app.run(host="127.0.0.1", port=args.port)


if __name__ == "__main__":
    main()
