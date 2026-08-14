/*
 * weekly-report web — static browser port of the weekly report editor.
 *
 * Runs entirely client-side (GitHub Pages): Markdown is rendered to
 * Outlook-safe HTML with marked + highlight.js, and reports are read and
 * written straight to a local folder via the File System Access API, so no
 * report content ever leaves the machine. The email markup mirrors app.py:
 * tables with inline styles on every element, because classic Outlook
 * renders mail with the Word engine.
 *
 * Copyright (c) 2026 Yusuf Bunyamin Kahveci <yusuf.kahveci@tii.ae>
 * SPDX-License-Identifier: Apache-2.0
 */

'use strict';

const APP_VERSION = 'v11';

/* ---------------------------------------------------------------- email CSS */

const MAX_IMG_WIDTH = 600;  /* full-width layout: 640 minus content padding */

const FONT = 'Calibri,Arial,Helvetica,sans-serif';
const MONO = "Consolas,'Courier New',monospace";
const NAVY = '#1f3864';
const BORDER = '#d9d9d9';
const RULE = '#e3e6ea';
const MUTED = '#7a8699';
const TD_FONT = `font-family:${FONT};font-size:15px;line-height:1.45;color:#1a1a1a;`;
const SECTION_LABEL =
  `font-family:${FONT};font-size:11px;font-weight:bold;letter-spacing:1px;` +
  `color:${MUTED};`;

const TAG_STYLES = {
  P: 'margin:0 0 8px 0;',
  UL: 'margin:2px 0 10px 0;padding:0 0 0 24px;',
  OL: 'margin:2px 0 10px 0;padding:0 0 0 24px;',
  LI: 'margin:0 0 3px 0;',
  BLOCKQUOTE: 'margin:4px 0 10px 0;padding:2px 0 2px 10px;' +
              'border-left:3px solid #cccccc;color:#555555;',
  H1: `font-size:17px;margin:12px 0 6px 0;color:${NAVY};`,
  H2: `font-size:16px;margin:12px 0 6px 0;color:${NAVY};`,
  H3: `font-size:15px;margin:10px 0 4px 0;color:${NAVY};`,
  H4: 'font-size:15px;margin:10px 0 4px 0;',
};
/* A sub-list sits inside its parent <li>, so the 10px tail a top-level list
   needs would open a gap before the item that follows it. */
const NESTED_LIST_STYLE = 'margin:3px 0 2px 0;padding:0 0 0 22px;';
/* [legacy type attribute, CSS equivalent] per depth: browsers vary the marker
   by nesting level on their own, the Word engine only obeys type=. */
const LIST_MARKERS = {
  UL: [['disc', 'disc'], ['circle', 'circle'], ['square', 'square']],
  OL: [['1', 'decimal'], ['a', 'lower-alpha'], ['i', 'lower-roman']],
};
const A_STYLE = 'color:#1f6fb2;';
const CODE_INLINE_STYLE =
  `font-family:${MONO};font-size:13px;background-color:#f5f5f5;padding:0 3px;`;

/* Credit line closing every email; the spacer row is the only top gap the
   Word engine renders reliably above a bordered cell. */
const APP_URL = 'https://ybkahveci.github.io/weekly-report/';
const FOOTER =
  '<table width="640" cellpadding="0" cellspacing="0" border="0">' +
  '<tr><td height="16" style="font-size:1px;line-height:16px;">&nbsp;</td></tr>' +
  `<tr><td style="font-family:${FONT};font-size:12px;line-height:1.45;` +
  `color:${MUTED};border-top:1px solid ${RULE};padding:8px 0 0 0;">` +
  `<font color="${MUTED}">Made with weekly-report — write in Markdown, ` +
  `paste into Outlook: </font><a href="${APP_URL}" style="${A_STYLE}">` +
  `<font color="#1f6fb2">${APP_URL}</font></a></td></tr></table>`;

/* Pygments-default-like colors for highlight.js token classes. */
const HLJS_COLORS = {
  'hljs-keyword': 'color:#008000;font-weight:bold;',
  'hljs-built_in': 'color:#008000;',
  'hljs-type': 'color:#b00040;',
  'hljs-literal': 'color:#666666;',
  'hljs-number': 'color:#666666;',
  'hljs-string': 'color:#ba2121;',
  'hljs-comment': 'color:#3d7b7b;font-style:italic;',
  'hljs-doctag': 'color:#3d7b7b;font-style:italic;',
  'hljs-meta': 'color:#9c6500;',
  'hljs-title': 'color:#0000ff;',
  'hljs-attr': 'color:#008000;',
  'hljs-symbol': 'color:#19177c;',
  'hljs-variable': 'color:#19177c;',
  'hljs-operator': 'color:#666666;',
  'hljs-params': '',
};

const LABEL_RE = /^(this week|status|upcoming)\s*:(.*)$/i;
const LIST_RE = /^\s*(?:[-*+]\s|\d+\.\s)/;
/* One nesting level. Four spaces, not two: python-markdown (the Flask
   frontend) only nests a sub-list at four, while marked accepts either. */
const INDENT = '    ';
/* Ctrl+key -> what wraps the selection. Markdown has no underline, so that
   one emits a literal <u> tag; both renderers pass inline HTML through and
   the Word engine honors the tag. */
const MARKS = {
  b: ['**'],
  i: ['*'],
  u: ['<u>', '</u>'],
};
const HEAD_DAYS_RE = /^(.+?)\s+[—–-]+\s+(\d+(?:[.,]\d+)?)\s*days?\s*$/i;

const esc = s => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ------------------------------------------------------------------ parsing */

function parseReport(text) {
  const meta = {};
  let body = text;
  if (text.startsWith('---')) {
    const end = text.indexOf('\n---', 3);
    if (end !== -1) {
      for (const line of text.slice(3, end).trim().split('\n')) {
        const i = line.indexOf(':');
        if (i > 0) meta[line.slice(0, i).trim().toLowerCase()] =
          line.slice(i + 1).trim();
      }
      body = text.slice(end + 4);
    }
  }

  const raw = [];
  let current = null;
  for (const line of body.split('\n')) {
    if (line.startsWith('## ')) {
      current = { head: line.slice(3).trim(), lines: [] };
      raw.push(current);
    } else if (current) {
      current.lines.push(line);
    }
  }

  const report = { meta, projects: [], warnings: [] };
  for (const item of raw) {
    const m = item.head.match(HEAD_DAYS_RE);
    let name, days;
    if (m) {
      name = m[1].trim();
      days = parseFloat(m[2].replace(',', '.'));
    } else {
      name = item.head;
      days = null;
      report.warnings.push(`No "— N days" found in heading "${item.head}"`);
    }
    const lines = item.lines;
    while (lines.length && !lines[0].trim()) lines.shift();
    let subtitle = null;
    if (lines.length && lines[0].trimStart().startsWith('>')) {
      subtitle = lines.shift().trimStart().slice(1).trim();
    }
    report.projects.push({
      name, days, subtitle,
      bodyMd: lines.join('\n').replace(/^\n+|\n+$/g, ''),
    });
  }
  return report;
}

function splitBody(text) {
  const segments = [];
  let label = null;
  let current = [];
  let inFence = false;

  const flush = () => {
    const content = current.join('\n').replace(/^\n+|\n+$/g, '');
    if (content || label !== null) segments.push([label, content]);
  };

  for (const line of text.split('\n')) {
    const stripped = line.trim();
    if (stripped.startsWith('```') || stripped.startsWith('~~~')) {
      inFence = !inFence;
    }
    const m = (inFence || LIST_RE.test(line)) ? null : stripped.match(LABEL_RE);
    if (m) {
      flush();
      label = m[1].toLowerCase();
      current = m[2].trim() ? [m[2].trim()] : [];
    } else {
      current.push(line);
    }
  }
  flush();
  return segments;
}

/* ---------------------------------------------------------------- rendering */

function highlightAndHarden(container) {
  for (const pre of [...container.querySelectorAll('pre')]) {
    const code = pre.querySelector('code') || pre;
    const langCls = [...code.classList].find(c => c.startsWith('language-'));
    const lang = langCls ? langCls.slice(9) : null;
    if (lang && hljs.getLanguage(lang)) {
      code.innerHTML = hljs.highlight(code.textContent, { language: lang }).value;
      /* legacy font/b/i tags: Word drops CSS colors on paste but honors these;
         reversed order processes children before their parent is serialized */
      for (const span of [...code.querySelectorAll('span')].reverse()) {
        const cls = [...span.classList].find(c => c.startsWith('hljs-'));
        const spec = HLJS_COLORS[cls];
        if (!spec) {
          span.replaceWith(...span.childNodes);
          continue;
        }
        let inner = span.innerHTML;
        if (spec.includes('font-style:italic')) inner = `<i>${inner}</i>`;
        if (spec.includes('font-weight:bold')) inner = `<b>${inner}</b>`;
        const color = spec.match(/color:(#[0-9a-f]{6})/);
        span.outerHTML = color ? `<font color="${color[1]}">${inner}</font>` : inner;
      }
    }
    /* Word ignores white-space CSS: newlines -> <br>, space runs -> &nbsp; */
    const lines = code.innerHTML.replace(/\n+$/, '').split('\n').map(ln => ln
      .replace(/^ +/, m => '&nbsp;'.repeat(m.length))
      .replace(/ {2,}/g, m => '&nbsp;'.repeat(m.length)));
    pre.outerHTML =
      '<table cellpadding="0" cellspacing="0" border="0" width="100%"' +
      ' style="margin:4px 0 10px 0;"><tr>' +
      `<td bgcolor="#F5F5F5" style="background-color:#f5f5f5;` +
      `border:1px solid ${BORDER};padding:8px 10px;` +
      `font-family:${MONO};font-size:13px;line-height:1.4;">` +
      lines.join('<br>') + '</td></tr></table>';
  }
}

function styleLists(container) {
  /* Style ul/ol by nesting depth: markers stay distinct three levels deep and
     only the outermost list keeps the paragraph-sized bottom margin. */
  for (const list of container.querySelectorAll('ul,ol')) {
    let depth = 0;
    for (let p = list.parentElement; p && p !== container; p = p.parentElement) {
      if (p.tagName === 'UL' || p.tagName === 'OL') depth++;
    }
    const markers = LIST_MARKERS[list.tagName];
    const [attr, css] = markers[Math.min(depth, markers.length - 1)];
    list.setAttribute('type', attr);
    list.setAttribute('style',
      (depth ? NESTED_LIST_STYLE : TAG_STYLES[list.tagName]) +
      `list-style-type:${css};`);
  }
}

function injectStyles(container) {
  for (const el of container.querySelectorAll('p,li,blockquote,h1,h2,h3,h4')) {
    el.setAttribute('style', TAG_STYLES[el.tagName]);
  }
  styleLists(container);
  for (const a of container.querySelectorAll('a')) a.setAttribute('style', A_STYLE);
  for (const c of container.querySelectorAll('code')) {
    if (c.parentElement.tagName !== 'PRE') c.setAttribute('style', CODE_INLINE_STYLE);
  }
}

async function readLocalImage(src) {
  let dir = state.dir;
  const parts = src.split('/').filter(Boolean);
  for (const part of parts.slice(0, -1)) {
    dir = await dir.getDirectoryHandle(part);
  }
  return (await dir.getFileHandle(parts[parts.length - 1])).getFile();
}

function fileDataUri(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(r.error);
    r.readAsDataURL(file);
  });
}

async function embedImage(file) {
  if (file.type === 'image/gif') return { uri: await fileDataUri(file), width: null };
  const bmp = await createImageBitmap(file);
  if (bmp.width <= MAX_IMG_WIDTH) {
    return { uri: await fileDataUri(file), width: bmp.width };
  }
  const canvas = document.createElement('canvas');
  canvas.width = MAX_IMG_WIDTH;
  canvas.height = Math.round(bmp.height * MAX_IMG_WIDTH / bmp.width);
  canvas.getContext('2d').drawImage(bmp, 0, 0, canvas.width, canvas.height);
  const mime = file.type === 'image/jpeg' ? 'image/jpeg' : 'image/png';
  return { uri: canvas.toDataURL(mime, 0.85), width: MAX_IMG_WIDTH };
}

async function processImages(container, warnings) {
  for (const img of [...container.querySelectorAll('img')]) {
    const src = img.getAttribute('src') || '';
    const alt = img.getAttribute('alt') || '';
    if (/^https?:\/\//.test(src)) {
      warnings.push(`Remote image kept as a link (may be blocked): ${src}`);
    } else {
      let uri, width;
      try {
        ({ uri, width } = await embedImage(await readLocalImage(src)));
      } catch {
        warnings.push(`Image not found: ${src}`);
        const span = document.createElement('span');
        span.setAttribute('style', 'color:#b00;');
        span.textContent = `[missing image: ${src}]`;
        img.replaceWith(span);
        continue;
      }
      img.setAttribute('src', uri);
      if (width) img.setAttribute('width', width);
    }
    img.setAttribute('alt', alt);
    img.setAttribute('style',
      'display:block;max-width:100%;height:auto;border:0;margin:6px 0;');
  }
}

async function mdFragment(text, warnings) {
  const container = document.createElement('div');
  container.innerHTML = marked.parse(text, { breaks: true, gfm: true });
  injectStyles(container);
  highlightAndHarden(container);
  await processImages(container, warnings);
  return container.innerHTML;
}

function italicize(frag) {
  /* Both CSS and legacy <i>: Word drops font-style on paste. */
  return frag
    .replaceAll('<p style="', '<p style="font-style:italic;')
    .replaceAll('<li style="', '<li style="font-style:italic;')
    .replace(/(<(?:p|li) style="[^"]*">)/g, '$1<i>')
    .replace(/<\/(p|li)>/g, '</i></$1>');
}

async function renderBody(bodyMd, warnings) {
  if (!bodyMd.trim()) return '';
  const parts = [];
  let first = true;
  for (const [label, content] of splitBody(bodyMd)) {
    if (label !== null) {
      const top = first ? '0' : '12px';
      parts.push(`<p style="${SECTION_LABEL}margin:${top} 0 4px 0;">` +
        `<b><font color="${MUTED}">${label.toUpperCase()}</font></b></p>`);
    }
    first = false;
    let frag = content.trim() ? await mdFragment(content, warnings) : '';
    if (label === 'this week') {
      frag = italicize(frag);
    }
    parts.push(frag);
  }
  return parts.join('');
}

function formatDays(days) {
  return `${days} ${days === 1 ? 'day' : 'days'}`;
}

async function renderEmail(report) {
  const warnings = [...report.warnings];
  const total = report.projects.reduce((s, p) => s + (p.days ?? 0), 0);

  const declared = report.meta.days;
  if (declared) {
    const d = parseFloat(declared.replace(',', '.'));
    if (Number.isNaN(d)) {
      warnings.push(`Frontmatter days is not a number: ${declared}`);
    } else if (Math.abs(d - total) > 1e-9) {
      warnings.push(`Frontmatter days: ${declared} != sum of project days (${total})`);
    }
  }

  const summary = [
    ['Week', report.meta.week || ''],
    ['Total working days', String(total)],
  ];
  if (report.meta.productivity) summary.push(['Productivity', report.meta.productivity]);
  const cells = summary.map(([label, value], i) => {
    const divider = i < summary.length - 1 ? `border-right:1px solid ${RULE};` : '';
    return `<td bgcolor="#f2f4f8" style="${TD_FONT}background-color:#f2f4f8;` +
      `padding:10px 22px 10px 14px;${divider}">` +
      `<div style="${SECTION_LABEL}">` +
      `<b><font color="${MUTED}">${esc(label.toUpperCase())}</font></b></div>` +
      `<div style="font-size:16px;font-weight:bold;color:${NAVY};` +
      `margin:2px 0 0 0;"><b><font color="${NAVY}">${esc(value)}</font></b>` +
      '</div></td>';
  });
  const header =
    '<table width="640" cellpadding="0" cellspacing="0" border="0"' +
    ` style="margin:0 0 16px 0;border:1px solid ${RULE};">` +
    `<tr>${cells.join('')}</tr></table>`;

  const rows = [
    `<tr><td bgcolor="${NAVY}" style="${TD_FONT}color:#ffffff;` +
    'font-weight:bold;font-size:12px;letter-spacing:1px;' +
    'padding:8px 12px;"><b><font color="#ffffff">PROJECTS</font></b></td></tr>',
  ];
  for (const p of report.projects) {
    const subtitle = p.subtitle
      ? '<div style="font-size:12px;font-weight:normal;font-style:italic;' +
        `color:${MUTED};margin:2px 0 0 0;">` +
        `<i><font color="${MUTED}">${esc(p.subtitle)}</font></i></div>`
      : '';
    const daysTxt = p.days !== null ? formatDays(p.days) : '—';
    const chip =
      '<table cellpadding="0" cellspacing="0" border="0" align="right"><tr>' +
      '<td bgcolor="#ffffff" style="background-color:#ffffff;' +
      `border:1px solid ${RULE};border-radius:10px;padding:2px 10px;` +
      `font-family:${FONT};font-size:12px;color:${MUTED};white-space:nowrap;">` +
      `<font color="${MUTED}">${daysTxt}</font></td></tr></table>`;
    rows.push(
      '<tr><td bgcolor="#f2f4f8" style="background-color:#f2f4f8;' +
      `border-left:3px solid ${NAVY};padding:0;">` +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      `<td style="${TD_FONT}font-size:16px;font-weight:bold;color:${NAVY};` +
      'padding:9px 12px;">' +
      `<b><font color="${NAVY}">${esc(p.name)}</font></b>${subtitle}</td>` +
      `<td align="right" valign="middle" style="${TD_FONT}padding:9px 12px;">` +
      `${chip}</td></tr></table></td></tr>`);
    rows.push(
      `<tr><td style="${TD_FONT}padding:12px 12px 18px 15px;">` +
      `${await renderBody(p.bodyMd, warnings)}</td></tr>`);
  }

  const html =
    '<meta charset="utf-8">' +
    `<div style="${TD_FONT}">` +
    '<p style="margin:0 0 8px 0;">Dear all,</p>' +
    '<p style="margin:0 0 14px 0;">This is my weekly progress report.</p>' +
    header +
    '<table width="640" cellpadding="0" cellspacing="0" border="0"' +
    ' style="border-collapse:collapse;">' + rows.join('') + '</table>' +
    FOOTER + '</div>';
  return { html, warnings };
}

/* ----------------------------------------------------------------- scaffold */

function sectionsOf(bodyMd) {
  const sections = {};
  let current = null;
  for (const line of bodyMd.split('\n')) {
    const m = line.trim().match(LABEL_RE);
    if (m && !LIST_RE.test(line)) {
      current = m[1].toLowerCase();
      sections[current] = [];
      if (m[2].trim()) sections[current].push(m[2].trim());
    } else if (current) {
      sections[current].push(line);
    }
  }
  for (const lines of Object.values(sections)) {
    while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  }
  return sections;
}

function scaffold(previous, week) {
  const out = ['---', `week: ${week}`, 'productivity:', '---', ''];
  if (previous === null) {
    out.push(
      '## Project name — 0 days', 'This week:', '- ', '',
      'Status:', '- Ongoing', '', 'Upcoming:', '- ', '',
      '## Others — 0 days', '- Weekly meeting', '- Weekly report', '');
    return out.join('\n');
  }
  for (const p of parseReport(previous).projects) {
    out.push(`## ${p.name} — 0 days`);
    if (p.subtitle) out.push(`> ${p.subtitle}`);
    out.push('');
    const sections = sectionsOf(p.bodyMd);
    if (Object.keys(sections).length) {
      out.push('This week:', '- ', '');
      out.push('Status:', ...(sections.status?.length ? sections.status : ['- Ongoing']), '');
      out.push('Upcoming:', ...(sections.upcoming?.length ? sections.upcoming : ['- ']), '');
    } else if (p.bodyMd) {
      out.push(p.bodyMd, '');
    }
  }
  return out.join('\n').replace(/\n+$/, '') + '\n';
}

/* ------------------------------------------------------------------ history */

function normKey(name) {
  return name.toLowerCase().replace(/\s+/g, ' ').trim();
}

function resolveAlias(name, aliases) {
  let display = name;
  const seen = new Set();
  let key = normKey(display);
  while (aliases[key] && !seen.has(key)) {
    seen.add(key);
    display = aliases[key];
    key = normKey(display);
  }
  return display;
}

function aggregateHistory(files, aliases = {}) {
  /* files: [{name, text}] in chronological (ascending) order */
  const projects = new Map();
  for (const f of files) {
    const rep = parseReport(f.text);
    const week = rep.meta.week || f.name.replace(/\.md$/, '');
    for (const p of rep.projects) {
      const display = resolveAlias(p.name, aliases);
      const key = normKey(display);
      if (!projects.has(key)) {
        projects.set(key, { name: display, total: 0, weeks: [] });
      }
      const entry = projects.get(key);
      entry.name = display;  /* latest spelling wins */
      const days = p.days === null ? 0 : p.days;
      entry.total += days;
      const segs = splitBody(p.bodyMd);
      const tw = segs.find(s => s[0] === 'this week');
      const hasLabels = segs.some(s => s[0] !== null);
      entry.weeks.push({
        file: f.name, week, days: p.days,
        /* label-less projects (e.g. Others) show their whole body */
        thisWeek: tw ? tw[1] : (hasLabels ? '' : p.bodyMd),
      });
    }
  }
  const list = [...projects.values()].sort((a, b) => b.total - a.total);
  const grandTotal = list.reduce((s, p) => s + p.total, 0);
  return { projects: list, grandTotal, reports: files.length };
}

const ALIAS_FILE = '.project-aliases.json';

async function readAliases() {
  try {
    const text = await (
      await (await state.dir.getFileHandle(ALIAS_FILE)).getFile()
    ).text();
    const norm = {};
    for (const [from, to] of Object.entries(JSON.parse(text))) {
      norm[normKey(from)] = to;
    }
    return norm;
  } catch {
    return {};
  }
}

async function writeAlias(from, to) {
  let raw = {};
  try {
    raw = JSON.parse(await (
      await (await state.dir.getFileHandle(ALIAS_FILE)).getFile()
    ).text());
  } catch {
    /* first alias */
  }
  raw[from] = to;
  await writeFile(ALIAS_FILE, JSON.stringify(raw, null, 2) + '\n');
}

function renderTimeline(project, hist) {
  document.getElementById('tl-title').textContent = `Timeline — ${project.name}`;
  const bar = document.getElementById('mergebar');
  const sel = document.getElementById('mergesel');
  const others = hist.projects.filter(p => p !== project);
  bar.hidden = others.length === 0;
  sel.innerHTML = '';
  for (const o of others) {
    const opt = document.createElement('option');
    opt.value = opt.textContent = o.name;
    sel.appendChild(opt);
  }
  document.getElementById('mergebtn').onclick = async () => {
    await writeAlias(project.name, sel.value);
    await showHistory();
  };
  const tl = document.getElementById('histtl');
  tl.innerHTML = '';
  for (const w of project.weeks) {
    const entry = document.createElement('div');
    entry.className = 'tlentry';
    const head = document.createElement('div');
    const wk = document.createElement('span');
    wk.className = 'tlweek';
    wk.textContent = `Week ${w.week}`;
    const days = document.createElement('span');
    days.className = 'tldays';
    days.textContent =
      `${w.file.replace(/\.md$/, '')} · ` +
      (w.days !== null ? formatDays(w.days) : '—');
    head.append(wk, days);
    const md = document.createElement('div');
    md.className = 'tlmd';
    if (w.thisWeek.trim()) {
      md.innerHTML = marked.parse(w.thisWeek, { breaks: true, gfm: true });
    } else {
      md.innerHTML = '<div class="tlnone">no activity recorded</div>';
    }
    entry.append(head, md);
    tl.appendChild(entry);
  }
}

async function showHistory() {
  const names = (await listFiles()).sort();
  const files = [];
  for (const n of names) files.push({ name: n, text: await readFile(n) });
  const hist = aggregateHistory(files, await readAliases());

  document.getElementById('hv-meta').textContent =
    `${hist.reports} report${hist.reports === 1 ? '' : 's'} · ` +
    `${hist.grandTotal} days total`;

  const tbody = document.getElementById('hsum-body');
  tbody.innerHTML = '';
  const maxTotal = hist.projects.length ? hist.projects[0].total : 0;
  hist.projects.forEach((p, i) => {
    const tr = document.createElement('tr');
    const mk = (cls, text) => {
      const td = document.createElement('td');
      td.className = cls;
      td.textContent = text;
      return td;
    };
    tr.append(
      mk('hname', p.name),
      mk('num', String(p.weeks.length)),
      mk('num', String(p.total)),
      mk('num', hist.grandTotal
        ? `${Math.round(100 * p.total / hist.grandTotal)}%` : '—'),
    );
    const barTd = document.createElement('td');
    const track = document.createElement('div');
    track.className = 'track';
    const meter = document.createElement('div');
    meter.className = 'meter';
    meter.style.width = maxTotal ? `${100 * p.total / maxTotal}%` : '0';
    track.appendChild(meter);
    barTd.appendChild(track);
    tr.appendChild(barTd);
    tr.onclick = () => {
      for (const row of tbody.children) row.classList.remove('sel');
      tr.classList.add('sel');
      renderTimeline(p, hist);
    };
    tbody.appendChild(tr);
    if (i === 0) tr.onclick();
  });
  if (!hist.projects.length) {
    document.getElementById('tl-title').textContent = 'Timeline';
    document.getElementById('mergebar').hidden = true;
    document.getElementById('histtl').innerHTML =
      '<div class="tlnone">no reports found</div>';
  }
  document.getElementById('histview').classList.add('open');
}

function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const year = t.getUTCFullYear();
  const start = new Date(Date.UTC(year, 0, 1));
  return { year, week: Math.ceil(((t - start) / 86400000 + 1) / 7) };
}

/* --------------------------------------------------------- folder + storage */

const state = {
  dir: null, current: null, dirty: false, saveTimer: null, lastEmail: null,
};

function idbOpen() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('weekly-report', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idbGet(key) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const t = db.transaction('kv').objectStore('kv').get(key);
    t.onsuccess = () => res(t.result);
    t.onerror = () => rej(t.error);
  });
}
async function idbSet(key, value) {
  const db = await idbOpen();
  return new Promise((res, rej) => {
    const t = db.transaction('kv', 'readwrite').objectStore('kv').put(value, key);
    t.onsuccess = () => res();
    t.onerror = () => rej(t.error);
  });
}

async function listFiles() {
  const names = [];
  for await (const entry of state.dir.values()) {
    if (entry.kind === 'file' && entry.name.endsWith('.md')) names.push(entry.name);
  }
  return names.sort().reverse();
}
async function readFile(name) {
  return (await (await state.dir.getFileHandle(name)).getFile()).text();
}
async function writeFile(name, content) {
  const fh = await state.dir.getFileHandle(name, { create: true });
  const w = await fh.createWritable();
  await w.write(content);
  await w.close();
}
async function fileExists(name) {
  try {
    await state.dir.getFileHandle(name);
    return true;
  } catch {
    return false;
  }
}

/* ----------------------------------------------------------------------- UI */

const $ = id => document.getElementById(id);
const setStatus = t => { $('status').textContent = t; };

async function loadList(preferred) {
  const files = await listFiles();
  const el = $('file');
  el.innerHTML = '';
  for (const f of files) {
    const o = document.createElement('option');
    o.value = o.textContent = f;
    el.appendChild(o);
  }
  if (files.length) {
    state.current = preferred && files.includes(preferred) ? preferred : files[0];
    el.value = state.current;
    await openFile(state.current);
  } else {
    setStatus('No reports yet — click "New week"');
  }
}

async function openFile(name) {
  state.current = name;
  $('editor').value = await readFile(name);
  state.dirty = false;
  $('dirty').hidden = true;
  renderOutline();
  await refresh();
}

async function refresh() {
  const report = parseReport($('editor').value);
  const { html, warnings } = await renderEmail(report);
  state.lastEmail = html;
  const sizeKb = new Blob([html]).size / 1024;
  if (sizeKb > 95) {
    warnings.push('Email exceeds ~100 KB — Gmail may clip the message');
  }
  const banner = warnings.map(w =>
    '<div style="background:#fff8e1;border:1px solid #f0dc9a;color:#6b5d1f;' +
    'border-radius:4px;padding:7px 12px;margin:0 0 8px 0;' +
    `font:12.5px system-ui;">${esc(w)}</div>`).join('');
  const doc =
    '<!doctype html><html><head><meta charset="utf-8"></head>' +
    '<body style="margin:0;padding:18px;background:#eef1f5;' +
    `zoom:${previewZoom()};">` +
    '<div style="max-width:700px;margin:0 auto;">' + banner +
    '<div style="color:#8a94a6;font:12px system-ui;text-align:center;' +
    `margin:0 0 10px 0;">Email size: ${sizeKb.toFixed(0)} KB` +
    ' — preview approximates Outlook rendering</div>' +
    '<div style="background:#ffffff;border:1px solid #dce0e6;border-radius:6px;' +
    `padding:26px 30px;">${html}</div>` +
    '</div></body></html>';
  const fr = $('preview');
  const y = fr.contentWindow ? fr.contentWindow.scrollY : 0;
  fr.onload = () => {
    fr.contentWindow.scrollTo(0, y);
    fr.contentWindow.addEventListener('wheel', ev => {
      if (!ev.ctrlKey) return;
      ev.preventDefault();
      setPreviewZoom(previewZoom() + (ev.deltaY < 0 ? 0.1 : -0.1));
    }, { passive: false });
  };
  fr.srcdoc = doc;
}

function previewZoom() {
  return parseFloat(localStorage.getItem('previewzoom')) || 1;
}

function setPreviewZoom(z) {
  z = Math.min(3, Math.max(0.5, Math.round(z * 10) / 10));
  localStorage.setItem('previewzoom', String(z));
  const w = $('preview').contentWindow;
  if (w && w.document.body) w.document.body.style.zoom = z;
  $('zoomlvl').textContent = Math.round(z * 100) + '%';
}

async function save() {
  if (!state.current) return;
  await writeFile(state.current, $('editor').value);
  state.dirty = false;
  $('dirty').hidden = true;
  setStatus('Saved');
  await refresh();
}

function scheduleSave() {
  state.dirty = true;
  $('dirty').hidden = false;
  setStatus('…');
  renderOutline();
  clearTimeout(state.saveTimer);
  state.saveTimer = setTimeout(
    () => save().catch(e => setStatus('Save failed: ' + e.message)), 600);
}

/* outline + reorder */

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
  let total = 0;
  let allParsed = true;
  sections.forEach((s, i) => {
    const row = document.createElement('div');
    row.className = 'oitem';
    const name = document.createElement('span');
    name.textContent = s.name.replace(/\s+[—–-]+\s+\d.*$/, '');
    name.onclick = () => jumpTo(s.start);
    const days = document.createElement('span');
    const m = s.name.match(HEAD_DAYS_RE);
    if (m) {
      const d = parseFloat(m[2].replace(',', '.'));
      total += d;
      days.className = 'odays';
      days.textContent = d + 'd';
    } else {
      allParsed = false;
      days.className = 'obadge';
      days.textContent = '!';
      days.title = 'No "— N days" in this heading';
    }
    const up = document.createElement('button');
    up.textContent = '↑';
    up.disabled = i === 0;
    up.onclick = () => moveSection(i, -1, false);
    const dn = document.createElement('button');
    dn.textContent = '↓';
    dn.disabled = i === sections.length - 1;
    dn.onclick = () => moveSection(i, 1, false);
    row.append(name, days, up, dn);
    el.appendChild(row);
  });
  $('ototal').textContent = sections.length
    ? `Total: ${total} ${total === 1 ? 'day' : 'days'}` +
      (allParsed ? '' : ' (incomplete)')
    : '';
}

/* editor keybindings */

function indentSelection(ed, dir) {
  const start = ed.value.lastIndexOf('\n', ed.selectionStart - 1) + 1;
  let end = ed.value.indexOf('\n', ed.selectionEnd);
  if (end === -1) end = ed.value.length;
  const block = ed.value.slice(start, end);
  const out = block.split('\n').map(ln =>
    dir > 0 ? INDENT + ln : ln.replace(/^ {1,4}/, '')).join('\n');
  ed.setRangeText(out, start, end, 'select');
}

function wrapSelection(ed, open, close = open) {
  const s = ed.selectionStart;
  const e = ed.selectionEnd;
  const sel = ed.value.slice(s, e);
  if (sel.startsWith(open) && sel.endsWith(close) &&
      sel.length >= open.length + close.length) {
    ed.setRangeText(sel.slice(open.length, sel.length - close.length),
      s, e, 'select');
    return;
  }
  ed.setRangeText(open + sel + close, s, e, 'select');
  if (sel) {
    ed.setSelectionRange(s, e + open.length + close.length);
  } else {
    ed.setSelectionRange(s + open.length, s + open.length);
  }
}

function editorKeydown(e) {
  const ed = e.target;
  if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.altKey &&
      !e.shiftKey && ed.selectionStart === ed.selectionEnd) {
    const pos = ed.selectionStart;
    const lineStart = ed.value.lastIndexOf('\n', pos - 1) + 1;
    const m = ed.value.slice(lineStart, pos)
      .match(/^(\s*)([-*+]|\d+[.)])\s(.*)$/);
    if (m) {
      e.preventDefault();
      if (!m[3].trim()) {
        if (m[1].length >= INDENT.length) {   /* step out one level first */
          ed.setRangeText(`${m[1].slice(INDENT.length)}${m[2]} `,
            lineStart, pos, 'end');
        } else {
          ed.setRangeText('', lineStart, pos, 'start');   /* end the list */
        }
      } else {
        const num = m[2].match(/^(\d+)([.)])$/);
        const marker = num ? `${Number(num[1]) + 1}${num[2]}` : m[2];
        ed.setRangeText(`\n${m[1]}${marker} `, pos, pos, 'end');
      }
      scheduleSave();
      return;
    }
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    indentSelection(ed, e.shiftKey ? -1 : 1);
    scheduleSave();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && MARKS[e.key]) {
    e.preventDefault();
    wrapSelection(ed, ...MARKS[e.key]);
    scheduleSave();
  }
}

/* images */

async function upload(file) {
  const images = await state.dir.getDirectoryHandle('images', { create: true });
  const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
  const ext = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/gif': '.gif' }[file.type];
  if (!ext) throw new Error(`unsupported image type: ${file.type}`);
  let name = `img-${stamp}${ext}`;
  let counter = 1;
  while (await images.getFileHandle(name).then(() => true, () => false)) {
    name = `img-${stamp}-${counter++}${ext}`;
  }
  const fh = await images.getFileHandle(name, { create: true });
  const w = await fh.createWritable();
  await w.write(file);
  await w.close();
  insertImageRef(name);
}

function insertImageRef(name) {
  const ed = $('editor');
  ed.setRangeText(`![](images/${name})\n`, ed.selectionStart, ed.selectionEnd,
    'end');
  $('imgview').classList.remove('open');
  ed.focus();
  scheduleSave();
}

let imgUrls = [];

async function openImagePanel() {
  const grid = $('imggrid');
  grid.innerHTML = '';
  imgUrls.forEach(URL.revokeObjectURL);
  imgUrls = [];
  const names = [];
  let images = null;
  try {
    images = await state.dir.getDirectoryHandle('images');
    for await (const e of images.values()) {
      if (e.kind === 'file' && /\.(png|jpe?g|gif)$/i.test(e.name)) {
        names.push(e.name);
      }
    }
  } catch {
    /* no images yet */
  }
  names.sort().reverse();
  for (const n of names) {
    const file = await (await images.getFileHandle(n)).getFile();
    const url = URL.createObjectURL(file);
    imgUrls.push(url);
    const card = document.createElement('div');
    card.className = 'imgcard';
    card.title = `Insert ${n}`;
    const im = document.createElement('img');
    im.src = url;
    const cap = document.createElement('div');
    cap.className = 'imgname';
    cap.textContent = n;
    const del = document.createElement('button');
    del.className = 'imgdel';
    del.textContent = '✕';
    del.title = 'Delete image file';
    del.onclick = async ev => {
      ev.stopPropagation();
      if (!confirm(`Delete ${n}? Reports referencing it will show a ` +
                   'missing image.')) return;
      await images.removeEntry(n);
      await openImagePanel();
    };
    card.onclick = () => insertImageRef(n);
    card.append(im, cap, del);
    grid.appendChild(card);
  }
  if (!grid.children.length) {
    grid.innerHTML = '<div class="tlnone">no images yet</div>';
  }
  $('imgview').classList.add('open');
}

/* folder gate */

async function start(dir) {
  state.dir = dir;
  $('gate').style.display = 'none';
  await loadList();
}

async function init() {
  if (!window.showDirectoryPicker) {
    $('open-folder').disabled = true;
    $('gate-err').textContent =
      'This browser has no File System Access API — use Chrome or Edge.';
    return;
  }
  const stored = await idbGet('dir').catch(() => null);
  if (stored) {
    const btn = $('reopen-folder');
    btn.hidden = false;
    btn.textContent = `Reopen "${stored.name}"`;
    btn.onclick = async () => {
      try {
        const perm = await stored.requestPermission({ mode: 'readwrite' });
        if (perm === 'granted') await start(stored);
        else $('gate-err').textContent = 'Permission denied.';
      } catch (e) { $('gate-err').textContent = e.message; }
    };
  }
  $('open-folder').onclick = async () => {
    try {
      const dir = await window.showDirectoryPicker({ mode: 'readwrite' });
      await idbSet('dir', dir);
      await start(dir);
    } catch (e) {
      if (e.name !== 'AbortError') $('gate-err').textContent = e.message;
    }
  };
}

/* event wiring */

$('editor').addEventListener('input', scheduleSave);
$('editor').addEventListener('keydown', editorKeydown);
$('file').addEventListener('change', e => openFile(e.target.value));
$('rename').addEventListener('click', async () => {
  if (!state.current) return;
  let name = prompt('New file name:', state.current);
  if (!name || name === state.current) return;
  if (!name.endsWith('.md')) name += '.md';
  if (!/^[A-Za-z0-9][A-Za-z0-9 ._&-]*\.md$/.test(name)) {
    alert('Invalid file name');
    return;
  }
  try {
    if (await fileExists(name)) {
      alert(`${name} already exists`);
      return;
    }
    await writeFile(name, $('editor').value);
    await state.dir.removeEntry(state.current);
    await loadList(name);
    setStatus('Renamed');
  } catch (e) { setStatus('Rename failed: ' + e.message); }
});

/* pane layout: drag #split to resize, double-click to hide the preview */
function applyLayout() {
  const px = localStorage.getItem('splitpx');
  const hidden = localStorage.getItem('previewhidden') === '1';
  $('pwrap').style.display = hidden ? 'none' : '';
  $('editor').style.flex = !hidden && px ? `0 0 ${px}px` : '1 1 auto';
}
$('split').addEventListener('dblclick', () => {
  localStorage.setItem('previewhidden',
    localStorage.getItem('previewhidden') === '1' ? '' : '1');
  applyLayout();
});
$('split').addEventListener('pointerdown', e => {
  /* pointer capture: move/up/cancel reach the splitter even when the
     pointer is released over the iframe or outside the window, so the
     preview's pointer-events can never get stuck disabled */
  const split = $('split');
  const pv = $('preview');
  split.setPointerCapture(e.pointerId);
  pv.style.pointerEvents = 'none';
  const move = ev => {
    const left = $('editor').getBoundingClientRect().left;
    localStorage.setItem('splitpx', String(Math.max(240, ev.clientX - left)));
    localStorage.setItem('previewhidden', '');
    applyLayout();
  };
  const finish = () => {
    pv.style.pointerEvents = '';
    /* resizing the iframe leaves Chrome's compositor scroll state stale
       (wheel events arrive but nothing scrolls); a hidden reflow rebuilds it */
    pv.style.display = 'none';
    void pv.offsetHeight;
    pv.style.display = '';
    applyLayout();
    split.removeEventListener('pointermove', move);
    split.removeEventListener('pointerup', finish);
    split.removeEventListener('pointercancel', finish);
  };
  split.addEventListener('pointermove', move);
  split.addEventListener('pointerup', finish);
  split.addEventListener('pointercancel', finish);
});
window.addEventListener('blur', () => {
  $('preview').style.pointerEvents = '';
});
applyLayout();

$('zoomin').addEventListener('click', () => setPreviewZoom(previewZoom() + 0.1));
$('zoomout').addEventListener('click', () => setPreviewZoom(previewZoom() - 0.1));
$('zoomlvl').addEventListener('click', () => setPreviewZoom(1));
setPreviewZoom(previewZoom());

$('help').addEventListener('click', () => {
  $('helpview').classList.toggle('open');
});
$('image').addEventListener('click', () => {
  openImagePanel().catch(e => setStatus('Images failed: ' + e.message));
});
$('img-close').addEventListener('click', () => {
  $('imgview').classList.remove('open');
});
$('imgupload').addEventListener('click', () => $('imgfile').click());
$('imgfile').addEventListener('change', async e => {
  try {
    for (const f of e.target.files) await upload(f);
    e.target.value = '';
  } catch (err) { setStatus('Upload failed: ' + err.message); }
});
window.addEventListener('beforeunload', e => {
  if (state.dirty) {
    e.preventDefault();
    e.returnValue = '';
  }
});
$('folder').addEventListener('click', async () => {
  try {
    const dir = await window.showDirectoryPicker({ mode: 'readwrite' });
    await idbSet('dir', dir);
    await start(dir);
  } catch (e) {
    if (e.name !== 'AbortError') setStatus(e.message);
  }
});
$('history').addEventListener('click', () => {
  showHistory().catch(e => setStatus('History failed: ' + e.message));
});
$('hv-close').addEventListener('click', () => {
  $('histview').classList.remove('open');
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    $('histview').classList.remove('open');
    $('helpview').classList.remove('open');
    $('imgview').classList.remove('open');
  }
  if (e.key === '?' && document.activeElement !== $('editor')) {
    $('helpview').classList.toggle('open');
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 's') {
    e.preventDefault();
    clearTimeout(state.saveTimer);
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
    const { year, week } = isoWeek(new Date());
    const name = `${year}-W${String(week).padStart(2, '0')}.md`;
    if (await fileExists(name)) {
      alert(`${name} already exists`);
      return;
    }
    const files = await listFiles();
    const previous = files.length ? await readFile(files[0]) : null;
    await writeFile(name, scaffold(previous, week));
    await loadList(name);
    const ed = $('editor');
    const marker = 'This week:\n- ';
    const at = ed.value.indexOf(marker);
    if (at >= 0) {
      ed.focus();
      ed.setSelectionRange(at + marker.length, at + marker.length);
    }
  } catch (e) { alert(e.message); }
});
function flashCopied() {
  const b = $('copy');
  b.classList.add('ok');
  b.textContent = 'Copied ✓';
  setTimeout(() => {
    b.classList.remove('ok');
    b.textContent = 'Copy email';
  }, 1500);
}

function copyRich(html, plain) {
  /* copy-event path writes the payload verbatim; the async clipboard API
     sanitizes text/html and strips inline styles that Outlook needs */
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
  /* execCommand only works synchronously inside the click stack, so copy
     the cached render (kept fresh by refresh()) with no awaits before it */
  const plain = $('editor').value;
  const html = state.lastEmail;
  if (!html) {
    setStatus('Nothing rendered yet');
    return;
  }
  if (copyRich(html, plain)) {
    setStatus('Copied — paste into your mail compose window');
    flashCopied();
  } else {
    navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })]).then(
      () => { setStatus('Copied (fallback path — styling may degrade)'); flashCopied(); },
      e => setStatus('Copy failed: ' + e.message));
  }
  if (state.dirty) {
    clearTimeout(state.saveTimer);
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

function applySpell(on) {
  $('editor').spellcheck = on;
  $('spell').textContent = 'Spell: ' + (on ? 'on' : 'off');
  localStorage.setItem('spellcheck', on ? 'on' : 'off');
}
$('spell').addEventListener('click', () => {
  applySpell(!$('editor').spellcheck);
  /* blur/focus makes the browser re-scan existing text immediately */
  $('editor').blur();
  $('editor').focus();
});
applySpell(localStorage.getItem('spellcheck') !== 'off');

$('appver').textContent = APP_VERSION;
console.log('weekly-report ' + APP_VERSION);
init().catch(e => { $('gate-err').textContent = e.message; });
