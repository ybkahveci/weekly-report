/*
 * weekly-report web: static browser port of the weekly report editor.
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

const APP_VERSION = 'v22';

/* ---------------------------------------------------------------- email CSS */

const MAX_IMG_WIDTH = 600;  /* full-width layout: 640 minus content padding */
/* Images are embedded at twice the width they are displayed at, so they stay
   sharp on high-DPI screens instead of being upscaled from 600 real pixels. */
const EMBED_SCALE = 2;

const FONT = 'Calibri,Arial,Helvetica,sans-serif';
const MONO = "Consolas,'Courier New',monospace";
const BORDER = '#d9d9d9';
const RULE = '#e3e6ea';
const MUTED = '#7a8699';
const SECTION_LABEL =
  `font-family:${FONT};font-size:11px;font-weight:bold;letter-spacing:1px;` +
  `color:${MUTED};`;
const CODE_INLINE_STYLE =
  `font-family:${MONO};font-size:13px;background-color:#f5f5f5;padding:0 3px;`;
/* A sub-list sits inside its parent <li>, so the 10px tail a top-level list
   needs would open a gap before the item that follows it. */
const NESTED_LIST_STYLE = 'margin:3px 0 2px 0;padding:0 0 0 22px;';
/* [legacy type attribute, CSS equivalent] per depth: browsers vary the marker
   by nesting level on their own, the Word engine only obeys type=. */
const LIST_MARKERS = {
  UL: [['disc', 'disc'], ['circle', 'circle'], ['square', 'square']],
  OL: [['1', 'decimal'], ['a', 'lower-alpha'], ['i', 'lower-roman']],
};
const APP_URL = 'https://ybkahveci.github.io/weekly-report/';

const EMAIL_THEMES = {
  navy: { label: 'Navy', accent: '#1f3864', soft: '#f2f4f8', link: '#1f6fb2' },
  teal: { label: 'Teal', accent: '#0b5563', soft: '#edf5f6', link: '#0b7285' },
  graphite: { label: 'Graphite', accent: '#2f3a48', soft: '#f3f4f6', link: '#2d6cdf' },
};

/* rebuilt whenever settings change: Outlook needs every style inline, so
   theme and density cannot live in a stylesheet */
let ES = null;

function buildEmailStyle() {
  const compact = settings.density === 'compact';
  const t = EMAIL_THEMES[settings.theme] || EMAIL_THEMES.navy;
  ES = {
    ...t,
    td: `font-family:${FONT};font-size:${compact ? 14 : 15}px;` +
        `line-height:${compact ? 1.35 : 1.45};color:#1a1a1a;`,
    headPad: compact ? '6px 12px' : '9px 12px',
    bodyPad: compact ? '8px 12px 12px 15px' : '12px 12px 18px 15px',
    cellPad: compact ? '7px 18px 7px 12px' : '10px 22px 10px 14px',
    tags: {
      P: `margin:0 0 ${compact ? 6 : 8}px 0;`,
      UL: `margin:2px 0 ${compact ? 8 : 10}px 0;padding:0 0 0 24px;`,
      OL: `margin:2px 0 ${compact ? 8 : 10}px 0;padding:0 0 0 24px;`,
      LI: `margin:0 0 ${compact ? 2 : 3}px 0;`,
      BLOCKQUOTE: 'margin:4px 0 10px 0;padding:2px 0 2px 10px;' +
                  'border-left:3px solid #cccccc;color:#555555;',
      H1: `font-size:17px;margin:12px 0 6px 0;color:${t.accent};`,
      H2: `font-size:16px;margin:12px 0 6px 0;color:${t.accent};`,
      H3: `font-size:15px;margin:10px 0 4px 0;color:${t.accent};`,
      H4: 'font-size:15px;margin:10px 0 4px 0;',
    },
  };
}

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
const HEAD_DAYS_RE = /^(.+?)\s+[\u2014\u2013-]+\s+(\d+(?:[.,]\d+)?)\s*days?\s*$/i;
const FENCE_RE = /^\s*(?:```|~~~)/;

const esc = s => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const fillTemplate = (tpl, vars) =>
  tpl.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? vars[k] : m));

/* sums of decimal day counts pick up float noise (2.1 + 2.2) */
const fmtNum = n => String(Math.round(n * 100) / 100);

/* ----------------------------------------------------------------- settings */

const SETTINGS_FILE = '.weekly-report.json';
/* section roles: "this week" is italicized and fills the history timeline,
   status and upcoming carry over into the next week's scaffold */
const ROLES = ['this week', 'status', 'upcoming'];
const DEFAULT_SETTINGS = {
  name: '',
  greeting: 'Dear all,\n\nThis is my weekly progress report.',
  signature: '',
  subject: 'Progress: report of Week {week}, {year} - {name}',
  to: '',
  cc: '',
  labels: { 'this week': 'This week', status: 'Status', upcoming: 'Upcoming' },
  theme: 'navy',
  density: 'comfortable',
  weekDays: 5,
};
let settings = structuredClone(DEFAULT_SETTINGS);
let LABEL_RE = null;

function applySettings(raw = {}) {
  settings = { ...structuredClone(DEFAULT_SETTINGS), ...raw };
  settings.labels = { ...DEFAULT_SETTINGS.labels, ...raw.labels };
  for (const role of ROLES) {
    settings.labels[role] =
      String(settings.labels[role]).trim() || DEFAULT_SETTINGS.labels[role];
  }
  const wd = Number(settings.weekDays);
  settings.weekDays = Number.isFinite(wd) && wd >= 0 ? wd : DEFAULT_SETTINGS.weekDays;
  /* the English names keep matching so older reports keep their sections */
  const names = new Set(ROLES);
  for (const role of ROLES) names.add(settings.labels[role].toLowerCase());
  LABEL_RE = new RegExp(`^(${[...names].map(escapeRe).join('|')})\\s*:(.*)$`, 'i');
  buildEmailStyle();
}

function labelRole(text) {
  const t = text.toLowerCase();
  return ROLES.find(r => r === t) ||
    ROLES.find(r => settings.labels[r].toLowerCase() === t);
}

applySettings();

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
  let inFence = false;
  for (const line of body.split('\n')) {
    if (FENCE_RE.test(line)) inFence = !inFence;
    if (!inFence && line.startsWith('## ')) {
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
      report.warnings.push(`No day count (" - N days") in heading "${item.head}"`);
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
    if (FENCE_RE.test(line)) inFence = !inFence;
    const m = (inFence || LIST_RE.test(line)) ? null : stripped.match(LABEL_RE);
    if (m) {
      flush();
      label = labelRole(m[1]);
      current = m[2].trim() ? [m[2].trim()] : [];
    } else {
      current.push(line);
    }
  }
  flush();
  return segments;
}

/* ---------------------------------------------------------------- rendering */

/* 4-column stops like python-markdown; each match starts at a stop, so the
   pad only depends on the match itself */
const expandTabs = s => s.replace(/[^\t\n]*\t/g,
  m => m.slice(0, -1) + ' '.repeat(4 - (m.length - 1) % 4));

function highlightAndHarden(container) {
  for (const pre of [...container.querySelectorAll('pre')]) {
    const code = pre.querySelector('code') || pre;
    /* Word collapses tabs to a single space */
    code.textContent = expandTabs(code.textContent);
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
      (depth ? NESTED_LIST_STYLE : ES.tags[list.tagName]) +
      `list-style-type:${css};`);
  }
}

function injectStyles(container) {
  for (const el of container.querySelectorAll('p,li,blockquote,h1,h2,h3,h4')) {
    el.setAttribute('style', ES.tags[el.tagName]);
  }
  styleLists(container);
  for (const a of container.querySelectorAll('a')) {
    a.setAttribute('style', `color:${ES.link};`);
  }
  for (const c of container.querySelectorAll('code')) {
    if (c.parentElement?.tagName !== 'PRE') c.setAttribute('style', CODE_INLINE_STYLE);
  }
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
  try {
    const width = Math.min(bmp.width, MAX_IMG_WIDTH);
    const target = MAX_IMG_WIDTH * EMBED_SCALE;
    /* Anything that already fits the pixel budget is embedded byte for byte;
       re-encoding it would only throw detail away. */
    if (bmp.width <= target) return { uri: await fileDataUri(file), width };
    const canvas = document.createElement('canvas');
    canvas.width = target;
    canvas.height = Math.round(bmp.height * target / bmp.width);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingQuality = 'high';   /* the default 'low' aliases badly */
    ctx.drawImage(bmp, 0, 0, canvas.width, canvas.height);
    const mime = file.type === 'image/jpeg' ? 'image/jpeg' : 'image/png';
    return { uri: canvas.toDataURL(mime, 0.92), width };
  } finally {
    bmp.close();
  }
}

/* every autosave re-renders; re-encoding unchanged images made typing lag */
const imageCache = new Map();

function embedImageCached(src, file) {
  const key = `${src}|${file.size}|${file.lastModified}`;
  if (!imageCache.has(key)) {
    const p = embedImage(file);
    p.catch(() => imageCache.delete(key));
    imageCache.set(key, p);
  }
  return imageCache.get(key);
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
        ({ uri, width } = await embedImageCached(src, await fileAt(src)));
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
  /* template content is inert: images are not fetched from the web server
     before processImages swaps in data URIs */
  const tpl = document.createElement('template');
  tpl.innerHTML = marked.parse(text, { breaks: true, gfm: true });
  injectStyles(tpl.content);
  highlightAndHarden(tpl.content);
  await processImages(tpl.content, warnings);
  return tpl.innerHTML;
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
        `<b><font color="${MUTED}">${esc(settings.labels[label].toUpperCase())}` +
        '</font></b></p>');
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
  const n = fmtNum(days);
  return `${n} ${n === '1' ? 'day' : 'days'}`;
}

async function renderEmail(report) {
  const warnings = [...report.warnings, ...lintReport(report)];
  const total = report.projects.reduce((s, p) => s + (p.days ?? 0), 0);

  const declared = report.meta.days;
  if (declared) {
    const d = parseFloat(declared.replace(',', '.'));
    if (Number.isNaN(d)) {
      warnings.push(`Frontmatter days is not a number: ${declared}`);
    } else if (Math.abs(d - total) > 1e-9) {
      warnings.push(
        `Frontmatter days: ${declared} != sum of project days (${fmtNum(total)})`);
    }
  } else if (settings.weekDays && report.projects.length &&
             report.projects.every(p => p.days !== null) &&
             Math.abs(settings.weekDays - total) > 1e-9) {
    warnings.push(`Total is ${formatDays(total)}, expected ` +
      `${formatDays(settings.weekDays)} (change in Settings, or add days: to ` +
      'the frontmatter for a short week)');
  }

  const summary = [
    ['Week', report.meta.week || ''],
    ['Total working days', fmtNum(total)],
  ];
  if (report.meta.productivity) summary.push(['Productivity', report.meta.productivity]);
  const cells = summary.map(([label, value], i) => {
    const divider = i < summary.length - 1 ? `border-right:1px solid ${RULE};` : '';
    return `<td bgcolor="${ES.soft}" style="${ES.td}background-color:${ES.soft};` +
      `padding:${ES.cellPad};${divider}">` +
      `<div style="${SECTION_LABEL}">` +
      `<b><font color="${MUTED}">${esc(label.toUpperCase())}</font></b></div>` +
      `<div style="font-size:16px;font-weight:bold;color:${ES.accent};` +
      `margin:2px 0 0 0;"><b><font color="${ES.accent}">${esc(value)}</font></b>` +
      '</div></td>';
  });
  const header =
    '<table width="640" cellpadding="0" cellspacing="0" border="0"' +
    ` style="margin:0 0 16px 0;border:1px solid ${RULE};">` +
    `<tr>${cells.join('')}</tr></table>`;

  const rows = [
    `<tr><td bgcolor="${ES.accent}" style="${ES.td}color:#ffffff;` +
    'font-weight:bold;font-size:12px;letter-spacing:1px;' +
    'padding:8px 12px;"><b><font color="#ffffff">PROJECTS</font></b></td></tr>',
  ];
  for (const p of report.projects) {
    const subtitle = p.subtitle
      ? '<div style="font-size:12px;font-weight:normal;font-style:italic;' +
        `color:${MUTED};margin:2px 0 0 0;">` +
        `<i><font color="${MUTED}">${esc(p.subtitle)}</font></i></div>`
      : '';
    const daysTxt = p.days !== null ? formatDays(p.days) : 'n/a';
    const chip =
      '<table cellpadding="0" cellspacing="0" border="0" align="right"><tr>' +
      '<td bgcolor="#ffffff" style="background-color:#ffffff;' +
      `border:1px solid ${RULE};border-radius:10px;padding:2px 10px;` +
      `font-family:${FONT};font-size:12px;color:${MUTED};white-space:nowrap;">` +
      `<font color="${MUTED}">${daysTxt}</font></td></tr></table>`;
    rows.push(
      `<tr><td bgcolor="${ES.soft}" style="background-color:${ES.soft};` +
      `border-left:3px solid ${ES.accent};padding:0;">` +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      `<td style="${ES.td}font-size:16px;font-weight:bold;color:${ES.accent};` +
      `padding:${ES.headPad};">` +
      `<b><font color="${ES.accent}">${esc(p.name)}</font></b>${subtitle}</td>` +
      `<td align="right" valign="middle" style="${ES.td}padding:${ES.headPad};">` +
      `${chip}</td></tr></table></td></tr>`);
    rows.push(
      `<tr><td style="${ES.td}padding:${ES.bodyPad};">` +
      `${await renderBody(p.bodyMd, warnings)}</td></tr>`);
  }

  const html =
    '<meta charset="utf-8">' +
    `<div style="${ES.td}">` +
    textParagraphs(settings.greeting, 0, 14) +
    header +
    '<table width="640" cellpadding="0" cellspacing="0" border="0"' +
    ' style="border-collapse:collapse;">' + rows.join('') + '</table>' +
    textParagraphs(settings.signature, 16, 0) + footer() + '</div>';
  return { html, warnings };
}

/* Credit line closing every email; the spacer row is the only top gap the
   Word engine renders reliably above a bordered cell. */
function footer() {
  return '<table width="640" cellpadding="0" cellspacing="0" border="0">' +
    '<tr><td height="16" style="font-size:1px;line-height:16px;">&nbsp;</td></tr>' +
    `<tr><td style="font-family:${FONT};font-size:12px;line-height:1.45;` +
    `color:${MUTED};border-top:1px solid ${RULE};padding:8px 0 0 0;">` +
    `<font color="${MUTED}">Made with weekly-report (write in Markdown, ` +
    `paste into Outlook): </font><a href="${APP_URL}" style="color:${ES.link};">` +
    `<font color="${ES.link}">${APP_URL}</font></a></td></tr></table>`;
}

function textParagraphs(text, topPx, bottomPx) {
  const paras = fillTemplate(text, { name: settings.name })
    .split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);
  return paras.map((p, i) =>
    `<p style="margin:${i === 0 ? topPx : 0}px 0 ` +
    `${i === paras.length - 1 ? bottomPx : 8}px 0;">` +
    `${esc(p).replace(/\n/g, '<br>')}</p>`).join('');
}

/* ----------------------------------------------------------------- scaffold */

function sectionsOf(bodyMd) {
  const sections = {};
  for (const [role, content] of splitBody(bodyMd)) {
    if (role) sections[role] = content ? content.split('\n') : [];
  }
  return sections;
}

const isEmptyItem = ln => /^\s*(?:[-*+]|\d+[.)])?\s*$/.test(ln);
/* New week's default for every project; replaced by hand where work happened */
const NO_ACTIVITY = '- No activities this week';

function newProjectLines(name) {
  const L = settings.labels;
  return [`## ${name} - 0 days`, '', `${L['this week']}:`, '- ', '',
    `${L.status}:`, '- Ongoing', '', `${L.upcoming}:`, '- ', ''];
}

function scaffold(previous, week) {
  const out = ['---', `week: ${week}`, 'productivity:', '---', ''];
  if (previous === null) {
    const lines = newProjectLines('Project name');
    lines[lines.indexOf('- ')] = NO_ACTIVITY;
    out.push(...lines,
      '## Others - 0 days', '- Weekly meeting', '- Weekly report', '');
    return out.join('\n');
  }
  const L = settings.labels;
  for (const p of parseReport(previous).projects) {
    out.push(`## ${p.name} - 0 days`);
    if (p.subtitle) out.push(`> ${p.subtitle}`);
    out.push('');
    const sections = sectionsOf(p.bodyMd);
    if (Object.keys(sections).length) {
      out.push(`${L['this week']}:`, NO_ACTIVITY, '');
      out.push(`${L.status}:`,
        ...(sections.status?.length ? sections.status : ['- Ongoing']), '');
      out.push(`${L.upcoming}:`,
        ...(sections.upcoming?.length ? sections.upcoming : ['- ']), '');
    } else if (p.bodyMd) {
      out.push(p.bodyMd, '');
    }
  }
  return out.join('\n').replace(/\n+$/, '') + '\n';
}

/* ----------------------------------------------------------- markdown tools */

const hl = (cls, html) => `<span class="${cls}">${html}</span>`;

function highlightInline(s) {
  return esc(s)
    .replace(/`[^`]+`/g, m => hl('hl-code', m))
    .replace(/\*\*[^*]+\*\*/g, m => hl('hl-em', m))
    .replace(/(^|[^*])(\*[^*\s][^*]*\*)(?!\*)/g, (m, pre, em) => pre + hl('hl-em', em))
    .replace(/!?\[[^\]]*\]\([^)]*\)/g, m => hl('hl-link', m));
}

/* Colors only, never weight or style: the overlay must keep the exact glyph
   advances of the textarea above it or the caret drifts off the text. */
function highlightMd(text) {
  const lines = text.split('\n');
  let front = lines[0] === '---';
  let fence = false;
  return lines.map((ln, i) => {
    if (front) {
      if (i > 0 && ln === '---') front = false;
      return hl('hl-front', esc(ln));
    }
    if (FENCE_RE.test(ln)) {
      fence = !fence;
      return hl('hl-code', esc(ln));
    }
    if (fence) return hl('hl-code', esc(ln));
    if (ln.startsWith('## ')) {
      const m = ln.slice(3).match(/^(.*?)(\s+[\u2014\u2013-]+\s+\d+(?:[.,]\d+)?\s*days?\s*)$/i);
      return hl('hl-mark', '## ') + (m
        ? hl('hl-head', esc(m[1])) + hl('hl-days', esc(m[2]))
        : hl('hl-head', esc(ln.slice(3))));
    }
    if (/^\s*>/.test(ln)) return hl('hl-sub', esc(ln));
    if (!LIST_RE.test(ln) && LABEL_RE.test(ln.trim())) {
      const colon = ln.indexOf(':') + 1;
      return hl('hl-label', esc(ln.slice(0, colon))) + highlightInline(ln.slice(colon));
    }
    const li = ln.match(/^(\s*)([-*+]|\d+[.)])(\s.*|)$/);
    if (li) return esc(li[1]) + hl('hl-mark', esc(li[2])) + highlightInline(li[3]);
    return highlightInline(ln);
  }).join('\n');
}

/* One layout for every report: a blank line around headings and before
   section labels, labels spelled as in Settings, "-" bullets indented by
   INDENT per level, no runs of blank lines. Fenced code is left as is. */
function tidyReport(text) {
  const src = text.replace(/\r\n?/g, '\n').split('\n');
  const front = [];
  let start = 0;
  if (src[0] === '---') {
    const end = src.indexOf('---', 1);
    if (end > 0) {
      front.push(...src.slice(0, end + 1).map(l => l.trimEnd()), '');
      start = end + 1;
    }
  }
  const lines = src.slice(start);

  const items = [];
  /* indents of the open list levels: nesting is relative to the parent item,
     since a top-level bullet may itself be indented up to three spaces */
  let open = [];
  let fence = false;
  for (const raw of lines) {
    if (FENCE_RE.test(raw)) {
      fence = !fence;
      items.push({ kind: 'code', text: raw.trimEnd() });
      continue;
    }
    if (fence) {
      items.push({ kind: 'code', text: raw });
      continue;
    }
    const ln = raw.replace(/\t/g, INDENT).trimEnd();
    if (!ln.trim()) {
      items.push({ kind: 'blank' });
      continue;
    }
    if (!/^\s/.test(ln) && !LIST_RE.test(ln) && !/^(?:[-*+]|\d+[.)])$/.test(ln)) open = [];
    if (ln.startsWith('## ')) {
      const head = ln.slice(3).trim();
      const m = head.match(HEAD_DAYS_RE);
      items.push({ kind: 'head', text: m
        ? `## ${m[1].trim()} - ${formatDays(parseFloat(m[2].replace(',', '.')))}`
        : `## ${head}` });
      continue;
    }
    const lm = !LIST_RE.test(ln) && ln.trim().match(LABEL_RE);
    if (lm) {
      const rest = lm[2].trim();
      items.push({ kind: 'label',
        text: `${settings.labels[labelRole(lm[1])]}:${rest ? ` ${rest}` : ''}` });
      continue;
    }
    const li = ln.match(/^( *)([-*+]|\d+[.)])(?:\s+(.*))?$/);
    if (li) {
      const n = li[1].length;
      while (open.length && n < open[open.length - 1]) open.pop();
      if (!open.length || n > open[open.length - 1]) open.push(n);
      const level = open.length - 1;
      const marker = /\d/.test(li[2]) ? li[2] : '-';
      items.push({ kind: li[3] ? 'item' : 'empty',
        text: `${INDENT.repeat(level)}${marker} ${li[3] || ''}` });
      continue;
    }
    items.push({ kind: /^\s*>/.test(ln) ? 'quote' : 'text', text: ln });
  }

  /* placeholder bullets go once their section has real content; a lone one
     stays, since "-" on its own after a line would turn it into a heading */
  const keep = items.map(() => true);
  let segStart = 0;
  const closeSegment = end => {
    const seg = items.slice(segStart, end);
    if (seg.some(t => ['item', 'text', 'quote', 'code'].includes(t.kind))) {
      for (let k = segStart; k < end; k++) if (items[k].kind === 'empty') keep[k] = false;
    }
  };
  items.forEach((t, k) => {
    if (t.kind === 'head' || t.kind === 'label') {
      closeSegment(k);
      segStart = k + 1;
    }
  });
  closeSegment(items.length);

  const body = [];
  let prev = null;
  let pendingBlank = false;
  items.forEach((t, k) => {
    if (!keep[k]) return;
    if (t.kind === 'blank') {
      pendingBlank = true;
      return;
    }
    const kind = t.kind === 'quote' && prev === 'head' ? 'sub' : t.kind;
    let blank = pendingBlank;
    if (!body.length || kind === 'sub' || prev === 'label') blank = false;
    else if (kind === 'head' || kind === 'label' || prev === 'head' || prev === 'sub') blank = true;
    if (blank) body.push('');
    body.push(t.text);
    prev = kind;
    pendingBlank = false;
  });
  return [...front, ...body].join('\n').replace(/\n*$/, '\n');
}

/* Preview-only hints for slips a spell checker misses. */
function lintReport(report) {
  const out = [];
  const tw = settings.labels['this week'];
  const noCode = md => {
    let fence = false;
    return md.split('\n').filter(ln => {
      if (FENCE_RE.test(ln)) fence = !fence;
      return !fence && !FENCE_RE.test(ln);
    }).join('\n').replace(/`[^`]*`/g, '');
  };
  for (const p of report.projects) {
    const segs = splitBody(p.bodyMd);
    const done = segs.find(s => s[0] === 'this week');
    if (done) {
      const work = done[1].split('\n')
        .filter(ln => !isEmptyItem(ln) && !/no activit/i.test(ln));
      if (p.days === 0 && work.length) {
        out.push(`${p.name}: 0 days, but work is listed under "${tw}"`);
      } else if (p.days > 0 && !work.length) {
        out.push(`${p.name}: ${formatDays(p.days)}, but nothing is listed under "${tw}"`);
      }
    }
    if (p.bodyMd.split('\n').some(ln => /^\s*(?:[-*+]|\d+[.)])\s*$/.test(ln))) {
      out.push(`${p.name}: empty bullet`);
    }
    const repeats = new Set();
    for (const m of noCode(p.bodyMd).matchAll(/(?<![\p{L}\p{N}])(\p{L}+)[ \t]+\1(?![\p{L}\p{N}])/giu)) {
      repeats.add(m[0].toLowerCase());
    }
    for (const r of repeats) out.push(`${p.name}: repeated word "${r}"`);
  }
  return out;
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
  const columns = [];
  for (const f of files) {
    const rep = parseReport(f.text);
    const week = rep.meta.week || f.name.replace(/\.md$/, '');
    const col = { file: f.name, week, days: new Map() };
    columns.push(col);
    for (const p of rep.projects) {
      const display = resolveAlias(p.name, aliases);
      const key = normKey(display);
      if (!projects.has(key)) {
        projects.set(key, { key, name: display, total: 0, weeks: [] });
      }
      const entry = projects.get(key);
      entry.name = display;  /* latest spelling wins */
      const days = p.days === null ? 0 : p.days;
      entry.total += days;
      col.days.set(key, (col.days.get(key) || 0) + days);
      const segs = splitBody(p.bodyMd);
      const tw = segs.find(s => s[0] === 'this week');
      const hasLabels = segs.some(s => s[0] !== null);
      entry.weeks.push({
        file: f.name, week, days: p.days, body: p.bodyMd,
        /* label-less projects (e.g. Others) show their whole body */
        thisWeek: tw ? tw[1] : (hasLabels ? '' : p.bodyMd),
      });
    }
  }
  const list = [...projects.values()].sort((a, b) => b.total - a.total);
  const grandTotal = list.reduce((s, p) => s + p.total, 0);
  return { projects: list, grandTotal, reports: files.length, columns };
}

/* "2026-W43.md" -> ISO year, week and quarter; other names have no period */
function reportPeriod(file) {
  const m = file.match(/^(\d{4})-W(\d{1,2})\b/);
  if (!m) return null;
  const year = Number(m[1]);
  const week = Number(m[2]);
  /* the week's Thursday decides its month, as ISO 8601 does for the year */
  const jan4 = new Date(Date.UTC(year, 0, 4));
  const thu = new Date(jan4);
  thu.setUTCDate(4 - ((jan4.getUTCDay() || 7) - 1) + (week - 1) * 7 + 3);
  return { year, week, quarter: Math.floor(thu.getUTCMonth() / 3) + 1 };
}

function inRange(file, range) {
  if (range === 'all') return true;
  const p = reportPeriod(file);
  if (!p) return false;
  const [kind, year, quarter] = range.split(':');
  return p.year === Number(year) && (kind === 'y' || p.quarter === Number(quarter));
}

function rangeOptions(files) {
  const years = new Map();
  for (const f of files) {
    const p = reportPeriod(f.name);
    if (!p) continue;
    if (!years.has(p.year)) years.set(p.year, new Set());
    years.get(p.year).add(p.quarter);
  }
  const opts = [['all', 'All time']];
  for (const year of [...years.keys()].sort((a, b) => b - a)) {
    opts.push([`y:${year}`, String(year)]);
    for (const q of [...years.get(year)].sort((a, b) => b - a)) {
      opts.push([`q:${year}:${q}`, `${year} Q${q}`]);
    }
  }
  return opts;
}

function summaryMarkdown(hist, rangeLabel) {
  const out = [`# Summary: ${rangeLabel}`, '',
    `${hist.reports} report${hist.reports === 1 ? '' : 's'} · ` +
    `${formatDays(hist.grandTotal)}`, ''];
  for (const p of hist.projects) {
    const share = hist.grandTotal
      ? ` (${Math.round(100 * p.total / hist.grandTotal)}%)` : '';
    out.push(`## ${p.name}: ${formatDays(p.total)}${share}`, '');
    for (const w of p.weeks) {
      if (w.thisWeek.trim()) out.push(`**Week ${w.week}**`, '', w.thisWeek.trim(), '');
    }
  }
  return out.join('\n');
}

function historyCsv(hist) {
  const cell = s => (/[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const rows = [];
  for (const p of hist.projects) {
    for (const w of p.weeks) {
      rows.push([w.file, String(w.week), p.name, w.days === null ? '' : fmtNum(w.days)]);
    }
  }
  rows.sort((a, b) => a[0].localeCompare(b[0]) || a[2].localeCompare(b[2]));
  return [['file', 'week', 'project', 'days'], ...rows]
    .map(r => r.map(cell).join(',')).join('\r\n') + '\r\n';
}

const ALIAS_FILE = '.project-aliases.json';

async function readAliases() {
  const norm = {};
  for (const [from, to] of Object.entries(await readJson(ALIAS_FILE, {}))) {
    norm[normKey(from)] = to;
  }
  return norm;
}

async function writeAlias(from, to) {
  const raw = await readJson(ALIAS_FILE, {});
  raw[from] = to;
  await writeFile(ALIAS_FILE, JSON.stringify(raw, null, 2) + '\n');
}

/* validated categorical slots (adjacent-pair CVD safe); projects past the
   last slot fold into a neutral "Other" instead of a generated hue */
const CHART_COLORS = {
  light: ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6'],
  dark: ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9'],
};
const CHART_OTHER = { light: '#c3c2b7', dark: '#5f5e5a' };

const hv = {
  files: [], aliases: {}, slots: new Map(), hist: null,
  range: 'all', query: '', selected: null,
};

function chartColor(key) {
  const mode = isDarkUi() ? 'dark' : 'light';
  const slot = hv.slots.get(key);
  return slot === undefined ? CHART_OTHER[mode] : CHART_COLORS[mode][slot];
}

function swatch(key) {
  const sw = document.createElement('span');
  sw.className = 'swatch';
  sw.style.background = chartColor(key);
  return sw;
}

function markMatches(root, needle) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    const text = node.nodeValue;
    const lower = text.toLowerCase();
    let i = lower.indexOf(needle);
    if (i < 0) continue;
    const frag = document.createDocumentFragment();
    let last = 0;
    while (i >= 0) {
      const mark = document.createElement('mark');
      mark.textContent = text.slice(i, i + needle.length);
      frag.append(text.slice(last, i), mark);
      last = i + needle.length;
      i = lower.indexOf(needle, last);
    }
    frag.append(text.slice(last));
    node.replaceWith(frag);
  }
}

function timelineEntry(title, sub, md, needle) {
  const entry = document.createElement('div');
  entry.className = 'tlentry';
  const head = document.createElement('div');
  const wk = document.createElement('span');
  wk.className = 'tlweek';
  wk.textContent = title;
  const meta = document.createElement('span');
  meta.className = 'tldays';
  meta.textContent = sub;
  head.append(wk, meta);
  const body = document.createElement('div');
  body.className = 'tlmd';
  if (md.trim()) {
    const tpl = document.createElement('template');
    tpl.innerHTML = marked.parse(md, { breaks: true, gfm: true });
    if (needle) markMatches(tpl.content, needle);
    /* relative srcs would be fetched from the web server, not the reports */
    for (const img of tpl.content.querySelectorAll('img')) {
      const src = img.getAttribute('src') || '';
      if (/^(https?|data|blob):/.test(src)) continue;
      img.removeAttribute('src');
      fileAt(src).then(f => {
        img.onload = () => URL.revokeObjectURL(img.src);
        img.src = URL.createObjectURL(f);
      }, () => { img.alt = `[missing image: ${src}]`; });
    }
    body.append(tpl.content);
  } else {
    body.innerHTML = '<div class="tlnone">no activity recorded</div>';
  }
  entry.append(head, body);
  return entry;
}

function renderTimeline(project, hist) {
  $('tl-title').textContent = `Timeline: ${project.name}`;
  const sel = $('mergesel');
  const others = hist.projects.filter(p => p !== project);
  $('mergebar').hidden = others.length === 0;
  sel.innerHTML = '';
  for (const o of others) {
    const opt = document.createElement('option');
    opt.value = opt.textContent = o.name;
    sel.appendChild(opt);
  }
  $('mergebtn').onclick = async () => {
    await writeAlias(project.name, sel.value);
    await showHistory();
  };
  const tl = $('histtl');
  tl.innerHTML = '';
  for (const w of project.weeks) {
    tl.appendChild(timelineEntry(`Week ${w.week}`,
      `${w.file.replace(/\.md$/, '')} · ` +
      (w.days !== null ? formatDays(w.days) : 'n/a'), w.thisWeek));
  }
}

function renderSearch(hist, query) {
  const needle = query.toLowerCase();
  $('tl-title').textContent = `Search: "${query}"`;
  $('mergebar').hidden = true;
  const hits = [];
  for (const p of hist.projects) {
    for (const w of p.weeks) {
      if (p.name.toLowerCase().includes(needle) ||
          w.body.toLowerCase().includes(needle)) hits.push({ p, w });
    }
  }
  hits.sort((a, b) => a.w.file.localeCompare(b.w.file));
  const tl = $('histtl');
  tl.innerHTML = hits.length ? '' : '<div class="tlnone">no matches</div>';
  for (const { p, w } of hits) {
    tl.appendChild(timelineEntry(`Week ${w.week} · ${p.name}`,
      w.file.replace(/\.md$/, ''), w.body, needle));
  }
}

function renderSummaryTable(hist) {
  const tbody = $('hsum-body');
  tbody.innerHTML = '';
  const maxTotal = hist.projects.length ? hist.projects[0].total : 0;
  const mk = (cls, ...content) => {
    const td = document.createElement('td');
    td.className = cls;
    td.append(...content);
    return td;
  };
  for (const p of hist.projects) {
    const tr = document.createElement('tr');
    const track = document.createElement('div');
    track.className = 'track';
    const meter = document.createElement('div');
    meter.className = 'meter';
    meter.style.width = maxTotal ? `${100 * p.total / maxTotal}%` : '0';
    track.appendChild(meter);
    tr.append(
      mk('hname', swatch(p.key), p.name),
      mk('num', String(p.weeks.length)),
      mk('num', fmtNum(p.total)),
      mk('num', hist.grandTotal
        ? `${Math.round(100 * p.total / hist.grandTotal)}%` : '0%'),
      mk('', track),
    );
    tr.classList.toggle('sel', !hv.query && p.key === hv.selected);
    tr.onclick = () => {
      hv.selected = p.key;
      hv.query = '';
      $('hsearch').value = '';
      renderHistory();
    };
    tbody.appendChild(tr);
  }
}

const colLabel = c => (/^\d+$/.test(c.week) ? `W${c.week}` : c.file.replace(/\.md$/, ''));

function roundTop(x, y, w, h, r) {
  r = Math.min(r, w / 2, h);
  return `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}` +
    `H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
}

function renderChart(hist) {
  const box = $('hchart');
  const legend = $('hlegend');
  const cols = hist.columns;
  $('hchartcard').hidden = !cols.length;
  box.innerHTML = '';
  legend.innerHTML = '';
  if (!cols.length) return;

  const order = k => (hv.slots.has(k) ? hv.slots.get(k) : Infinity);
  const name = k => hist.projects.find(p => p.key === k)?.name ?? k;
  /* '' collects every project without a color slot */
  const stacks = cols.map(c => {
    const segs = new Map();
    for (const [key, d] of c.days) {
      if (d <= 0) continue;
      const k = hv.slots.has(key) ? key : '';
      segs.set(k, (segs.get(k) || 0) + d);
    }
    return [...segs].sort((a, b) => order(a[0]) - order(b[0]));
  });
  const totals = stacks.map(s => s.reduce((t, [, d]) => t + d, 0));

  const W = Math.max(320, box.clientWidth);
  const H = 190;
  const pad = { l: 30, r: 6, t: 8, b: 24 };
  const plotW = W - pad.l - pad.r;
  const plotH = H - pad.t - pad.b;
  const top = Math.max(1, Math.ceil(Math.max(settings.weekDays, ...totals)));
  const step = top <= 6 ? 1 : Math.ceil(top / 5);
  const y = v => pad.t + plotH * (1 - v / top);
  const band = plotW / cols.length;
  const bw = Math.max(3, Math.min(24, band * 0.6));
  const every = Math.max(1, Math.ceil(36 / band));

  const parts = [];
  for (let v = 0; v <= top; v += step) {
    parts.push(
      `<line class="${v ? 'grid' : 'base'}" x1="${pad.l}" x2="${W - pad.r}" ` +
      `y1="${y(v)}" y2="${y(v)}"/>`,
      `<text class="tick" x="${pad.l - 6}" y="${y(v) + 4}" text-anchor="end">${v}</text>`);
  }
  cols.forEach((c, i) => {
    const x = pad.l + band * i + (band - bw) / 2;
    let base = 0;
    stacks[i].forEach(([key, d], j) => {
      const isTop = j === stacks[i].length - 1;
      const y0 = y(base);
      const y1 = y(base + d) + (isTop ? 0 : 2);  /* 2px surface gap */
      base += d;
      if (y0 - y1 <= 0) return;
      const fill = chartColor(key);
      parts.push(isTop
        ? `<path d="${roundTop(x, y1, bw, y0 - y1, 4)}" fill="${fill}"/>`
        : `<rect x="${x}" y="${y1}" width="${bw}" height="${y0 - y1}" fill="${fill}"/>`);
    });
    if (i % every === 0) {
      parts.push(`<text class="tick" x="${pad.l + band * (i + 0.5)}" y="${H - 6}" ` +
        `text-anchor="middle">${esc(colLabel(c))}</text>`);
    }
    parts.push(`<rect class="hit" data-i="${i}" x="${pad.l + band * i}" y="${pad.t}" ` +
      `width="${band}" height="${plotH}"/>`);
  });
  box.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ` +
    `role="img" aria-label="Days per project per week">${parts.join('')}</svg>`;

  const keys = [...new Set(stacks.flat().map(([k]) => k))]
    .sort((a, b) => order(a) - order(b));
  for (const k of keys) {
    const item = document.createElement('span');
    item.className = 'litem';
    /* not "Other": reports commonly have a real "Others" project */
    item.append(swatch(k), k ? name(k) : 'Smaller projects');
    legend.appendChild(item);
  }

  const svg = box.firstChild;
  const tip = $('htip');
  svg.addEventListener('mousemove', e => {
    const hit = e.target.closest('.hit');
    if (!hit) {
      tip.hidden = true;
      return;
    }
    const i = Number(hit.dataset.i);
    const c = cols[i];
    const rows = [...c.days].filter(([, d]) => d > 0).sort((a, b) => b[1] - a[1]);
    tip.innerHTML =
      `<div class="tiph">Week ${esc(c.week)} · ${esc(c.file.replace(/\.md$/, ''))}</div>` +
      rows.map(([k, d]) => '<div class="tipr">' +
        `<span class="swatch" style="background:${chartColor(k)}"></span>` +
        `<span>${esc(name(k))}</span><b>${fmtNum(d)}</b></div>`).join('') +
      `<div class="tipr tipt"><span></span><span>Total</span><b>${fmtNum(totals[i])}</b></div>` +
      '<div class="tipf">click to open</div>';
    tip.hidden = false;
    const r = box.getBoundingClientRect();
    const left = e.clientX - r.left + 14;
    tip.style.left = `${Math.min(left, r.width - tip.offsetWidth - 4)}px`;
    tip.style.top = `${e.clientY - r.top + 14}px`;
  });
  svg.addEventListener('mouseleave', () => { tip.hidden = true; });
  svg.addEventListener('click', e => {
    const hit = e.target.closest('.hit');
    if (hit) openReport(cols[Number(hit.dataset.i)].file);
  });
}

function renderHistory() {
  const hist = aggregateHistory(
    hv.files.filter(f => inRange(f.name, hv.range)), hv.aliases);
  hv.hist = hist;
  $('hv-meta').textContent =
    `${hist.reports} report${hist.reports === 1 ? '' : 's'} · ` +
    `${formatDays(hist.grandTotal)} total`;
  if (!hist.projects.some(p => p.key === hv.selected)) {
    hv.selected = hist.projects.length ? hist.projects[0].key : null;
  }
  renderSummaryTable(hist);
  renderChart(hist);
  if (hv.query) {
    renderSearch(hist, hv.query);
  } else if (hv.selected) {
    renderTimeline(hist.projects.find(p => p.key === hv.selected), hist);
  } else {
    $('tl-title').textContent = 'Timeline';
    $('mergebar').hidden = true;
    $('histtl').innerHTML = '<div class="tlnone">no reports found</div>';
  }
}

async function showHistory() {
  const names = (await listFiles()).sort();
  hv.files = [];
  for (const n of names) hv.files.push({ name: n, text: await readFile(n) });
  hv.aliases = await readAliases();
  /* rank over all reports so a project keeps its color across period filters */
  hv.slots = new Map(aggregateHistory(hv.files, hv.aliases).projects
    .slice(0, CHART_COLORS.light.length).map((p, i) => [p.key, i]));
  const opts = rangeOptions(hv.files);
  if (!opts.some(([v]) => v === hv.range)) hv.range = 'all';
  const sel = $('hrange');
  sel.innerHTML = '';
  for (const [value, label] of opts) sel.add(new Option(label, value));
  sel.value = hv.range;
  $('histview').classList.add('open');
  renderHistory();
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
  dir: null, mode: null, current: null, dirty: false, saveTimer: null,
  lastEmail: null, renderedText: null, renderGen: 0,
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
async function readJson(name, fallback) {
  try {
    return JSON.parse(await readFile(name));
  } catch {
    return fallback;
  }
}
async function fileAt(path) {
  const parts = path.split('/').filter(Boolean);
  let dir = state.dir;
  for (const part of parts.slice(0, -1)) dir = await dir.getDirectoryHandle(part);
  return (await dir.getFileHandle(parts[parts.length - 1])).getFile();
}
async function writePath(path, data) {
  const parts = path.split('/').filter(Boolean);
  let dir = state.dir;
  for (const part of parts.slice(0, -1)) {
    dir = await dir.getDirectoryHandle(part, { create: true });
  }
  const w = await (await dir.getFileHandle(parts[parts.length - 1], { create: true }))
    .createWritable();
  await w.write(data);
  await w.close();
}

async function loadSettings() {
  applySettings(await readJson(SETTINGS_FILE, {}));
}

async function saveSettings(raw) {
  applySettings(raw);
  await writeFile(SETTINGS_FILE, JSON.stringify(settings, null, 2) + '\n');
}

async function browserStorageDir() {
  const root = await navigator.storage.getDirectory();
  /* without it the browser may evict the reports under storage pressure;
     not awaited: Firefox asks the user and the promise waits for the answer */
  navigator.storage.persist?.().catch(() => {});
  return root.getDirectoryHandle('weekly-reports', { create: true });
}

/* .git and other dot entries stay out of exports */
const KEEP_DOTFILES = new Set([ALIAS_FILE, SETTINGS_FILE]);
const IMPORTABLE =
  /^(?:[^/]+\.md|\.project-aliases\.json|\.weekly-report\.json|images\/[^/]+\.(?:png|jpe?g|gif))$/i;

async function collectFiles(dir, prefix = '') {
  const out = [];
  for await (const entry of dir.values()) {
    if (entry.name.startsWith('.') && !KEEP_DOTFILES.has(entry.name)) continue;
    const path = prefix + entry.name;
    if (entry.kind === 'directory') {
      out.push(...await collectFiles(entry, `${path}/`));
    } else {
      out.push({ path, file: await entry.getFile() });
    }
  }
  return out;
}

async function exportZip() {
  const entries = [];
  for (const { path, file } of await collectFiles(state.dir)) {
    entries.push({
      name: `weekly-reports/${path}`,
      data: new Uint8Array(await file.arrayBuffer()),
      date: new Date(file.lastModified),
    });
  }
  download(`weekly-reports-${new Date().toISOString().slice(0, 10)}.zip`,
    zipStore(entries));
  return entries.length;
}

/* entries: [{path, data}]; a folder pick or zip wraps everything in one
   top-level directory, which is dropped */
async function importEntries(entries) {
  /* macOS zips carry __MACOSX/ and ._ shadow files outside the top folder */
  entries = entries.filter(e => !/(^|\/)(__MACOSX\/|\._|\.DS_Store$)/.test(e.path));
  const top = entries[0]?.path.split('/')[0];
  if (top && entries.every(e => e.path.startsWith(`${top}/`))) {
    for (const e of entries) e.path = e.path.slice(top.length + 1);
  }
  const ok = entries.filter(e => IMPORTABLE.test(e.path));
  if (!ok.length) throw new Error('nothing to import: expected .md reports and images/');
  let clashes = 0;
  for (const e of ok) {
    if (await fileAt(e.path).then(() => true, () => false)) clashes++;
  }
  if (clashes && !confirm(`${clashes} of ${ok.length} files already exist and will ` +
                          'be overwritten. Continue?')) return 0;
  for (const e of ok) await writePath(e.path, e.data);
  imageCache.clear();
  return ok.length;
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
  /* typing with no file open would be silently lost */
  $('editor').readOnly = !files.length;
  $('empty').hidden = files.length > 0;
  if (files.length) {
    state.current = preferred && files.includes(preferred) ? preferred : files[0];
    el.value = state.current;
    await openFile(state.current);
  } else {
    state.current = null;
    state.lastEmail = state.renderedText = null;
    $('editor').value = '';
    renderHighlight();
    $('preview').srcdoc = '';
    renderOutline();
    setStatus('');
  }
}

async function openReport(name) {
  await flushSave();
  $('histview').classList.remove('open');
  $('file').value = name;
  await openFile(name);
}

async function openFile(name) {
  const text = await readFile(name);
  state.current = name;
  $('editor').value = text;
  renderHighlight();
  state.dirty = false;
  $('dirty').hidden = true;
  setStatus('');
  renderOutline();
  await refresh();
}

async function refresh() {
  /* image embedding awaits, so an older render can finish after a newer one */
  const gen = ++state.renderGen;
  const text = $('editor').value;
  const { html, warnings } = await renderEmail(parseReport(text));
  if (gen !== state.renderGen) return;
  state.lastEmail = html;
  state.renderedText = text;
  const sizeKb = new Blob([html]).size / 1024;
  if (sizeKb > 95) {
    warnings.push('Email exceeds ~100 KB; Gmail may clip the message');
  }
  const banner = warnings.map(w =>
    '<div style="background:#fff8e1;border:1px solid #f0dc9a;color:#6b5d1f;' +
    'border-radius:4px;padding:7px 12px;margin:0 0 8px 0;' +
    `font:12.5px system-ui;">${esc(w)}</div>`).join('');
  const bg = getComputedStyle(document.documentElement)
    .getPropertyValue('--preview-bg').trim();
  const doc =
    '<!doctype html><html><head><meta charset="utf-8"></head>' +
    `<body style="margin:0;padding:18px;background:${bg};` +
    `zoom:${previewZoom()};">` +
    '<div style="max-width:700px;margin:0 auto;">' + banner +
    '<div style="color:#8a94a6;font:12px system-ui;text-align:center;' +
    `margin:0 0 10px 0;">Email size: ${sizeKb.toFixed(0)} KB` +
    ' · preview approximates Outlook rendering</div>' +
    '<div style="background:#ffffff;border:1px solid #dce0e6;border-radius:6px;' +
    `padding:26px 30px;">${html}</div>` +
    '</div></body></html>';
  const fr = $('preview');
  const y = fr.contentWindow ? fr.contentWindow.scrollY : 0;
  fr.onload = () => {
    fr.contentWindow.scrollTo(0, y);
    blockFileDrops(fr.contentWindow);
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
  const text = $('editor').value;
  await writeFile(state.current, text);
  /* typing during the write leaves a newer autosave pending */
  if ($('editor').value === text) {
    state.dirty = false;
    $('dirty').hidden = true;
    setStatus('Saved');
  }
  await refresh();
}

/* before anything that changes state.current or state.dir */
async function flushSave() {
  clearTimeout(state.saveTimer);
  if (state.dirty) await save();
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
    if (FENCE_RE.test(ln)) fence = !fence;
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
function moveSection(idx, dir) {
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
  setText(ed, assembled.join('\n'));
  const pos = lineToChar(assembled, dir < 0 ? lo.start : lo.start + blockHi.length);
  ed.setSelectionRange(pos, pos);
}
function adjustDays(idx, delta) {
  const ed = $('editor');
  const { lines, sections } = splitSections(ed.value);
  const s = sections[idx];
  const m = s.name.match(HEAD_DAYS_RE);
  const name = m ? m[1].trim() : s.name;
  const days = Math.max(0, (m ? parseFloat(m[2].replace(',', '.')) : 0) + delta);
  const start = lineToChar(lines, s.start);
  replaceRange(ed, start, start + lines[s.start].length,
    `## ${name} - ${formatDays(days)}`);
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
  const button = (text, title, onclick, cls = '') => {
    const b = document.createElement('button');
    b.textContent = text;
    b.title = title;
    b.className = cls;
    b.onclick = onclick;
    return b;
  };
  sections.forEach((s, i) => {
    const row = document.createElement('div');
    row.className = 'oitem';
    const m = s.name.match(HEAD_DAYS_RE);
    const name = document.createElement('span');
    name.className = 'oname';
    name.textContent = m ? m[1].trim() : s.name;
    name.onclick = () => jumpTo(s.start);
    const days = document.createElement('span');
    if (m) {
      const d = parseFloat(m[2].replace(',', '.'));
      total += d;
      days.className = 'odays';
      days.textContent = fmtNum(d) + 'd';
    } else {
      allParsed = false;
      days.className = 'obadge';
      days.textContent = '!';
      days.title = 'No day count (" - N days") in this heading';
    }
    const up = button('↑', 'Move up (Alt+Up)', () => moveSection(i, -1));
    up.disabled = i === 0;
    const dn = button('↓', 'Move down (Alt+Down)', () => moveSection(i, 1));
    dn.disabled = i === sections.length - 1;
    row.append(name,
      button('−', 'Half a day less', () => adjustDays(i, -0.5), 'dbtn'),
      days,
      button('+', 'Half a day more', () => adjustDays(i, 0.5), 'dbtn'),
      up, dn);
    el.appendChild(row);
  });
  const expected = settings.weekDays;
  const t = $('ototal');
  t.textContent = sections.length
    ? `Total: ${formatDays(total)}` +
      (expected ? ` of ${fmtNum(expected)}` : '') +
      (allParsed ? '' : ' (incomplete)')
    : '';
  t.classList.toggle('warn',
    Boolean(expected && allParsed && Math.abs(total - expected) > 1e-9));
}

/* editor edits */

/* execCommand keeps the edit on the native undo stack (and fires input);
   setRangeText and assigning .value bypass it */
function replaceRange(ed, start, end, text) {
  ed.focus();
  ed.setSelectionRange(start, end);
  const ok = text
    ? document.execCommand('insertText', false, text)
    : start === end || document.execCommand('delete');
  if (!ok) {
    ed.setRangeText(text, start, end, 'end');
    ed.dispatchEvent(new Event('input'));
  }
}

/* replaces only the changed middle so undo and scroll stay local */
function setText(ed, text) {
  const old = ed.value;
  if (old === text) return;
  let a = 0;
  while (a < old.length && a < text.length && old[a] === text[a]) a++;
  let b = 0;
  while (b < old.length - a && b < text.length - a &&
         old[old.length - 1 - b] === text[text.length - 1 - b]) b++;
  replaceRange(ed, a, old.length - b, text.slice(a, text.length - b));
}

function lineBlock(ed) {
  const { selectionStart: s, selectionEnd: e } = ed;
  const start = ed.value.lastIndexOf('\n', s - 1) + 1;
  /* a selection ending at column 0 does not include that line */
  let end = ed.value.indexOf('\n', e > s && ed.value[e - 1] === '\n' ? e - 1 : e);
  if (end === -1) end = ed.value.length;
  return { s, e, start, end, block: ed.value.slice(start, end) };
}

function indentSelection(ed, dir) {
  const { s, e, start, end, block } = lineBlock(ed);
  const out = block.split('\n').map(ln =>
    dir > 0 ? INDENT + ln : ln.replace(/^ {1,4}/, '')).join('\n');
  if (out === block) return;
  replaceRange(ed, start, end, out);
  if (s === e) {
    const pos = Math.max(start, s + out.length - block.length);
    ed.setSelectionRange(pos, pos);
  } else {
    ed.setSelectionRange(start, start + out.length);
  }
}

function toggleList(ed, numbered) {
  const { s, e, start, end, block } = lineBlock(ed);
  const lines = block.split('\n');
  const re = numbered ? /^(\s*)\d+[.)]\s/ : /^(\s*)[-*+]\s/;
  const remove = lines.some(ln => ln.trim()) &&
    lines.every(ln => !ln.trim() || re.test(ln));
  let n = 0;
  const out = lines.map(ln => {
    if (remove) return ln.replace(re, '$1');
    if (!ln.trim() && lines.length > 1) return ln;
    const bare = ln.replace(/^(\s*)(?:[-*+]|\d+[.)])\s/, '$1');
    const lead = bare.match(/^\s*/)[0];
    return `${lead}${numbered ? `${++n}.` : '-'} ${bare.trimStart()}`;
  }).join('\n');
  replaceRange(ed, start, end, out);
  if (s !== e) ed.setSelectionRange(start, start + out.length);
}

function wrapSelection(ed, open, close = open) {
  const s = ed.selectionStart;
  const e = ed.selectionEnd;
  const sel = ed.value.slice(s, e);
  if (sel.startsWith(open) && sel.endsWith(close) &&
      sel.length >= open.length + close.length) {
    const inner = sel.slice(open.length, sel.length - close.length);
    replaceRange(ed, s, e, inner);
    ed.setSelectionRange(s, s + inner.length);
    return;
  }
  replaceRange(ed, s, e, open + sel + close);
  if (sel) {
    ed.setSelectionRange(s, e + open.length + close.length);
  } else {
    ed.setSelectionRange(s + open.length, s + open.length);
  }
}

function insertLink(ed) {
  const s = ed.selectionStart;
  const e = ed.selectionEnd;
  const sel = ed.value.slice(s, e).trim();
  const url = prompt('Link address:', /^https?:\/\//.test(sel) ? sel : 'https://');
  if (!url || url === 'https://') {
    ed.focus();
    return;
  }
  const text = sel && sel !== url ? sel : url;
  replaceRange(ed, s, e, `[${text}](${url})`);
}

function insertCode(ed) {
  const s = ed.selectionStart;
  const e = ed.selectionEnd;
  const sel = ed.value.slice(s, e);
  if (sel && !sel.includes('\n')) {
    wrapSelection(ed, '`');
    return;
  }
  const lead = s > 0 && ed.value[s - 1] !== '\n' ? '\n' : '';
  replaceRange(ed, s, e, `${lead}\`\`\`\n${sel.replace(/\n$/, '')}\n\`\`\`\n`);
  if (!sel) {
    const pos = s + lead.length + 4;
    ed.setSelectionRange(pos, pos);
  }
}

function addProject(ed) {
  const { lines, sections } = splitSections(ed.value);
  const line = ed.value.slice(0, ed.selectionStart).split('\n').length - 1;
  const cur = sections.find(s => line >= s.start && line < s.end);
  const at = cur ? cur.end : (sections.length ? sections[0].start : lines.length);
  const block = newProjectLines('New project').join('\n');
  let pos;
  let text;
  if (at < lines.length) {
    pos = lineToChar(lines, at);
    text = `${block}\n`;
  } else {
    pos = ed.value.length;
    const sep = ed.value.endsWith('\n\n') || !ed.value ? ''
      : ed.value.endsWith('\n') ? '\n' : '\n\n';
    text = sep + block;
  }
  replaceRange(ed, pos, pos, text);
  const nameAt = pos + text.indexOf('New project');
  ed.setSelectionRange(nameAt, nameAt + 'New project'.length);
}

function editorKeydown(e) {
  const ed = e.target;
  if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.altKey &&
      !e.shiftKey && !e.isComposing && ed.selectionStart === ed.selectionEnd) {
    const pos = ed.selectionStart;
    const lineStart = ed.value.lastIndexOf('\n', pos - 1) + 1;
    const m = ed.value.slice(lineStart, pos)
      .match(/^(\s*)([-*+]|\d+[.)])\s(.*)$/);
    if (m) {
      e.preventDefault();
      if (!m[3].trim()) {
        if (m[1].length >= INDENT.length) {   /* step out one level first */
          replaceRange(ed, lineStart, pos, `${m[1].slice(INDENT.length)}${m[2]} `);
        } else {
          replaceRange(ed, lineStart, pos, '');   /* end the list */
        }
      } else {
        const num = m[2].match(/^(\d+)([.)])$/);
        const marker = num ? `${Number(num[1]) + 1}${num[2]}` : m[2];
        replaceRange(ed, pos, pos, `\n${m[1]}${marker} `);
      }
    }
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    indentSelection(ed, e.shiftKey ? -1 : 1);
    return;
  }
  const key = e.key.toLowerCase();
  if ((e.ctrlKey || e.metaKey) && !e.shiftKey && MARKS[key]) {
    e.preventDefault();
    wrapSelection(ed, ...MARKS[key]);
  } else if ((e.ctrlKey || e.metaKey) && !e.shiftKey && key === 'k') {
    e.preventDefault();
    insertLink(ed);
  }
}

/* The textarea's own text is transparent; this layer behind it shows the
   same text colored. The extra newline gives a trailing empty line height,
   as the textarea does. */
function renderHighlight() {
  const layer = $('edhl');
  layer.innerHTML = highlightMd($('editor').value) + '\n';
  layer.scrollTop = $('editor').scrollTop;
}

function tidy(ed) {
  const out = tidyReport(ed.value);
  if (out === ed.value) {
    setStatus('Already tidy');
    return;
  }
  setText(ed, out);
  setStatus('Tidied (Ctrl+Z undoes it)');
}

const TOOLBAR = {
  bold: ed => wrapSelection(ed, ...MARKS.b),
  italic: ed => wrapSelection(ed, ...MARKS.i),
  underline: ed => wrapSelection(ed, ...MARKS.u),
  bullets: ed => toggleList(ed, false),
  numbers: ed => toggleList(ed, true),
  link: insertLink,
  code: insertCode,
  project: addProject,
  tidy,
};

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
  $('imgview').classList.remove('open');
  replaceRange(ed, ed.selectionStart, ed.selectionEnd, `![](images/${name})\n`);
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

/* storage gate */

async function start(dir, mode) {
  state.dir = dir;
  state.mode = mode;
  imageCache.clear();
  localStorage.setItem('storage', mode);
  $('gate').style.display = 'none';
  await loadSettings();
  await loadList();
}

async function pickFolder() {
  const dir = await window.showDirectoryPicker({ mode: 'readwrite' });
  await flushSave();
  await idbSet('dir', dir);
  await start(dir, 'folder');
}

async function useBrowserStorage() {
  await flushSave();
  await start(await browserStorageDir(), 'browser');
}

async function init() {
  const canPick = Boolean(window.showDirectoryPicker);
  const canStore = Boolean(navigator.storage && navigator.storage.getDirectory);
  $('open-folder').hidden = !canPick;
  $('use-browser').hidden = !canStore;
  $('gate-note').hidden = canPick;
  if (!canPick && !canStore) {
    $('gate-err').textContent =
      'This browser cannot store files; use a current Chrome, Edge, Firefox or Safari.';
    return;
  }
  if (canStore && localStorage.getItem('storage') === 'browser') {
    await start(await browserStorageDir(), 'browser');
    return;
  }
  if (!canPick) return;
  const stored = await idbGet('dir').catch(() => null);
  if (!stored) return;
  /* Chrome's "allow on every visit" keeps the grant across sessions */
  if (await stored.queryPermission({ mode: 'readwrite' }) === 'granted') {
    await start(stored, 'folder');
    return;
  }
  const btn = $('reopen-folder');
  btn.hidden = false;
  btn.textContent = `Reopen "${stored.name}"`;
  btn.onclick = async () => {
    try {
      const perm = await stored.requestPermission({ mode: 'readwrite' });
      if (perm === 'granted') await start(stored, 'folder');
      else $('gate-err').textContent = 'Permission denied.';
    } catch (e) { $('gate-err').textContent = e.message; }
  };
}

$('open-folder').addEventListener('click', () => {
  pickFolder().catch(e => {
    if (e.name !== 'AbortError') $('gate-err').textContent = e.message;
  });
});
$('use-browser').addEventListener('click', () => {
  useBrowserStorage().catch(e => { $('gate-err').textContent = e.message; });
});

/* storage dialog */

function openStorage() {
  const browser = state.mode === 'browser';
  $('st-where').textContent = browser
    ? 'Browser storage: the reports live inside this browser only. Export a ' +
      'zip now and then as a backup, or to move them to another machine.'
    : `Folder "${state.dir.name}" on your disk.`;
  $('st-folder').hidden = !window.showDirectoryPicker;
  $('st-folder').textContent = browser ? 'Open a folder…' : 'Open another folder…';
  $('st-browser').hidden = browser || !navigator.storage?.getDirectory;
  $('st-msg').textContent = '';
  $('storeview').classList.add('open');
}

async function runImport(entries) {
  try {
    await flushSave();
    const n = await importEntries(entries);
    if (!n) return;
    await loadSettings();
    await loadList(state.current);
    $('st-msg').textContent = `Imported ${n} file${n === 1 ? '' : 's'}.`;
  } catch (err) {
    $('st-msg').textContent = 'Import failed: ' + err.message;
  }
}

$('storage').addEventListener('click', () => { if (state.dir) openStorage(); });
$('st-close').addEventListener('click', () => $('storeview').classList.remove('open'));
$('st-folder').addEventListener('click', () => {
  pickFolder().then(() => $('storeview').classList.remove('open'), e => {
    if (e.name !== 'AbortError') $('st-msg').textContent = e.message;
  });
});
$('st-browser').addEventListener('click', () => {
  useBrowserStorage().then(() => $('storeview').classList.remove('open'),
    e => { $('st-msg').textContent = e.message; });
});
$('st-import-dir').addEventListener('click', () => $('importdir').click());
$('st-import-zip').addEventListener('click', () => $('importzip').click());
$('importdir').addEventListener('change', async e => {
  const files = [...e.target.files];
  e.target.value = '';
  await runImport(files.map(f => ({ path: f.webkitRelativePath || f.name, data: f })));
});
$('importzip').addEventListener('change', async e => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  try {
    await runImport(await unzip(f));
  } catch (err) {
    $('st-msg').textContent = 'Import failed: ' + err.message;
  }
});
$('st-export').addEventListener('click', async () => {
  try {
    await flushSave();
    const n = await exportZip();
    $('st-msg').textContent = `Exported ${n} file${n === 1 ? '' : 's'}.`;
  } catch (err) {
    $('st-msg').textContent = 'Export failed: ' + err.message;
  }
});

/* settings dialog */

const SET_FIELDS = {
  'set-name': 'name', 'set-greeting': 'greeting', 'set-signature': 'signature',
  'set-subject': 'subject', 'set-to': 'to', 'set-cc': 'cc',
  'set-density': 'density', 'set-weekdays': 'weekDays',
};
const SET_LABELS = {
  'set-l-this': 'this week', 'set-l-status': 'status', 'set-l-upcoming': 'upcoming',
};

function fillSettingsForm(s) {
  for (const [id, key] of Object.entries(SET_FIELDS)) $(id).value = s[key];
  for (const [id, role] of Object.entries(SET_LABELS)) $(id).value = s.labels[role];
  const radios = [...document.querySelectorAll('input[name="set-theme"]')];
  (radios.find(r => r.value === s.theme) || radios[0]).checked = true;
}

function openSettings() {
  const themes = $('set-themes');
  if (!themes.children.length) {
    for (const [key, t] of Object.entries(EMAIL_THEMES)) {
      const label = document.createElement('label');
      label.className = 'tswatch';
      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'set-theme';
      input.value = key;
      const chip = document.createElement('span');
      chip.style.background = t.accent;
      label.append(input, chip, t.label);
      themes.appendChild(label);
    }
  }
  fillSettingsForm(settings);
  $('set-msg').textContent = '';
  $('setview').classList.add('open');
  $('set-name').focus();
}

$('settings').addEventListener('click', () => { if (state.dir) openSettings(); });
$('set-cancel').addEventListener('click', () => $('setview').classList.remove('open'));
$('set-reset').addEventListener('click', () => fillSettingsForm(DEFAULT_SETTINGS));
$('set-form').addEventListener('submit', async e => {
  e.preventDefault();
  const raw = { labels: {} };
  for (const [id, key] of Object.entries(SET_FIELDS)) raw[key] = $(id).value;
  raw.weekDays = Number(raw.weekDays);
  for (const [id, role] of Object.entries(SET_LABELS)) raw.labels[role] = $(id).value;
  raw.theme = document.querySelector('input[name="set-theme"]:checked')?.value;
  try {
    await saveSettings(raw);
    $('setview').classList.remove('open');
    renderOutline();
    if (state.current) await refresh();
    setStatus('Settings saved');
  } catch (err) {
    $('set-msg').textContent = 'Save failed: ' + err.message;
  }
});

/* app theme: auto follows the OS; the email itself always renders light */

const UI_THEMES = ['auto', 'light', 'dark'];

function isDarkUi() {
  const t = document.documentElement.dataset.theme;
  return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
}

function themeChanged() {
  if (state.current) refresh().catch(e => setStatus('Render failed: ' + e.message));
  if ($('histview').classList.contains('open') && hv.hist) renderHistory();
}

function applyUiTheme(t) {
  if (t === 'auto') delete document.documentElement.dataset.theme;
  else document.documentElement.dataset.theme = t;
  localStorage.setItem('uitheme', t);
  $('theme').title = `Theme: ${t} (click to change)`;
  themeChanged();
}

$('theme').addEventListener('click', () => {
  const cur = UI_THEMES.indexOf(localStorage.getItem('uitheme') || 'auto');
  applyUiTheme(UI_THEMES[(cur + 1) % UI_THEMES.length]);
});
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
  if (!document.documentElement.dataset.theme) themeChanged();
});
applyUiTheme(localStorage.getItem('uitheme') || 'auto');

/* email export */

function emailSubject() {
  const rep = parseReport($('editor').value);
  const period = state.current ? reportPeriod(state.current) : null;
  return fillTemplate(settings.subject, {
    week: rep.meta.week || (period ? period.week : ''),
    year: period ? period.year : new Date().getFullYear(),
    name: settings.name,
  }).replace(/[\s,:;·-]+$/, '');  /* an empty {name} leaves a dangling " -" */
}

function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(s);
}
const wrap76 = s => s.replace(/.{76}(?=.)/g, '$&\r\n');
const mimeWord = v =>
  (/^[\x20-\x7e]*$/.test(v) ? v : `=?UTF-8?B?${b64(new TextEncoder().encode(v))}?=`);

/* X-Unsent makes classic Outlook open the file as an editable draft;
   images become cid: parts so they arrive as real inline attachments */
function buildEml(html, subject) {
  const boundary = `=_weekly_report_${Date.now().toString(36)}`;
  const images = [];
  const body = html.replace(/^<meta charset="utf-8">/, '')
    .replace(/src="data:([\w/+.-]+);base64,([^"]+)"/g, (m, mime, data) => {
      const n = images.length + 1;
      const file = `image${n}.${mime === 'image/jpeg' ? 'jpg' : mime.split('/')[1]}`;
      images.push([`Content-Type: ${mime}; name="${file}"`,
        'Content-Transfer-Encoding: base64', `Content-ID: <img${n}@weekly-report>`,
        `Content-Disposition: inline; filename="${file}"`, '', wrap76(data)]);
      return `src="cid:img${n}@weekly-report"`;
    });
  const doc = '<!doctype html><html><head><meta charset="utf-8"></head>' +
    `<body>${body}</body></html>`;
  const lines = ['X-Unsent: 1'];
  if (settings.to.trim()) lines.push(`To: ${settings.to.trim()}`);
  if (settings.cc.trim()) lines.push(`Cc: ${settings.cc.trim()}`);
  lines.push(`Subject: ${mimeWord(subject)}`, 'MIME-Version: 1.0',
    `Content-Type: multipart/related; type="text/html"; boundary="${boundary}"`,
    '', `--${boundary}`, 'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: base64', '',
    wrap76(b64(new TextEncoder().encode(doc))));
  for (const part of images) lines.push(`--${boundary}`, ...part);
  lines.push(`--${boundary}--`, '');
  return lines.join('\r\n');
}

function download(name, blob) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

function flashButton(btn, text) {
  if (!btn.dataset.label) btn.dataset.label = btn.textContent;
  btn.classList.add('ok');
  btn.textContent = text;
  clearTimeout(btn.flashTimer);
  btn.flashTimer = setTimeout(() => {
    btn.classList.remove('ok');
    btn.textContent = btn.dataset.label;
  }, 1500);
}

function copyRich(html, plain, refocus) {
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
    refocus.focus();
  }
  return ok;
}

function copyEmail() {
  /* execCommand only works synchronously inside the user gesture, so copy
     the cached render (kept fresh by refresh()) with no awaits before it */
  const plain = $('editor').value;
  const html = state.lastEmail;
  if (!html) {
    setStatus('Nothing rendered yet');
    return;
  }
  if (plain !== state.renderedText) {
    setStatus('Rendering…');
    clearTimeout(state.saveTimer);
    (state.dirty ? save() : refresh()).then(
      () => setStatus('Updated: click Copy email again'),
      e => setStatus('Render failed: ' + e.message));
    return;
  }
  if (copyRich(html, plain, $('editor'))) {
    setStatus('Copied: paste into your mail compose window');
    flashButton($('copy'), 'Copied ✓');
  } else {
    navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })]).then(() => {
      setStatus('Copied (fallback path, styling may degrade)');
      flashButton($('copy'), 'Copied ✓');
    }, e => setStatus('Copy failed: ' + e.message));
  }
  flushSave().catch(e => setStatus('Save failed: ' + e.message));
}

$('copy').addEventListener('click', copyEmail);
$('subject').addEventListener('click', () => {
  if (!state.current) return;
  const subject = emailSubject();
  navigator.clipboard.writeText(subject).then(() => {
    setStatus(`Subject copied: ${subject}`);
    flashButton($('subject'), 'Copied ✓');
  }, e => setStatus('Copy failed: ' + e.message));
});
$('eml').addEventListener('click', async () => {
  if (!state.current) return;
  try {
    await flushSave();
    if ($('editor').value !== state.renderedText) await refresh();
    const eml = buildEml(state.lastEmail, emailSubject());
    download(state.current.replace(/\.md$/, '.eml'),
      new Blob([eml], { type: 'message/rfc822' }));
    setStatus('Saved .eml: open it to get a ready draft in Outlook');
  } catch (e) {
    setStatus('Export failed: ' + e.message);
  }
});

/* event wiring */

$('editor').addEventListener('input', renderHighlight);
$('editor').addEventListener('input', scheduleSave);
$('editor').addEventListener('scroll', () => {
  $('edhl').scrollTop = $('editor').scrollTop;
});
$('editor').addEventListener('keydown', editorKeydown);
$('file').addEventListener('change', e => {
  flushSave().then(() => openFile(e.target.value))
    .catch(err => setStatus('Open failed: ' + err.message));
});
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
    /* a pending autosave would recreate the old file after removeEntry */
    await flushSave();
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

$('toolbar').addEventListener('mousedown', e => {
  /* buttons must not take focus, or the editor loses its selection */
  if (e.target.closest('button')) e.preventDefault();
});
$('toolbar').addEventListener('click', e => {
  const b = e.target.closest('button[data-act]');
  if (b && state.current) TOOLBAR[b.dataset.act]($('editor'));
});
$('empty-new').addEventListener('click', () => $('new').click());
$('empty-settings').addEventListener('click', openSettings);

/* pane layout: drag #split to resize, double-click to hide the preview */
function applyLayout() {
  const px = localStorage.getItem('splitpx');
  const hidden = localStorage.getItem('previewhidden') === '1';
  $('pwrap').style.display = hidden ? 'none' : '';
  $('ewrap').style.flex = !hidden && px ? `0 0 ${px}px` : '1 1 0';
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
    const left = $('ewrap').getBoundingClientRect().left;
    localStorage.setItem('splitpx', String(Math.max(320, ev.clientX - left)));
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
  if (!state.current) return;
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
/* a file dropped outside the editor would otherwise open in place of the app */
function blockFileDrops(win) {
  for (const type of ['dragover', 'drop']) {
    win.addEventListener(type, e => {
      if (e.dataTransfer.types.includes('Files')) e.preventDefault();
    });
  }
}
blockFileDrops(window);

$('history').addEventListener('click', () => {
  if (!state.dir) return;
  flushSave().then(showHistory).catch(e => setStatus('History failed: ' + e.message));
});
$('hv-close').addEventListener('click', () => {
  $('histview').classList.remove('open');
});
$('hrange').addEventListener('change', e => {
  hv.range = e.target.value;
  renderHistory();
});
let searchTimer = null;
$('hsearch').addEventListener('input', e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    hv.query = e.target.value.trim();
    renderHistory();
  }, 150);
});
$('hcopy').addEventListener('click', () => {
  if (!hv.hist) return;
  const md = summaryMarkdown(hv.hist, $('hrange').selectedOptions[0].textContent);
  if (copyRich(marked.parse(md, { gfm: true }), md, $('hcopy'))) {
    flashButton($('hcopy'), 'Copied ✓');
  }
});
$('hcsv').addEventListener('click', () => {
  if (!hv.hist) return;
  const slug = hv.range === 'all' ? 'all' : hv.range.split(':').slice(1).join('-q');
  download(`weekly-history-${slug}.csv`,
    new Blob([historyCsv(hv.hist)], { type: 'text/csv' }));
});
let resizeTimer = null;
window.addEventListener('resize', () => {
  if (!$('histview').classList.contains('open') || !hv.hist) return;
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => renderChart(hv.hist), 120);
});

const OVERLAYS = ['histview', 'helpview', 'imgview', 'setview', 'storeview'];
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    for (const id of OVERLAYS) $(id).classList.remove('open');
  }
  const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);
  if (e.key === '?' && !typing) {
    $('helpview').classList.toggle('open');
  }
  const mod = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  if (mod && !e.shiftKey && key === 's') {
    e.preventDefault();
    clearTimeout(state.saveTimer);
    save().catch(err => setStatus('Save failed: ' + err.message));
  }
  if (mod && e.shiftKey && key === 'c' && state.current) {
    e.preventDefault();
    copyEmail();
  }
  if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
    const ed = $('editor');
    if (document.activeElement !== ed) return;
    const { sections } = splitSections(ed.value);
    const line = ed.value.slice(0, ed.selectionStart).split('\n').length - 1;
    const idx = sections.findIndex(s => line >= s.start && line < s.end);
    if (idx >= 0) {
      e.preventDefault();
      moveSection(idx, e.key === 'ArrowUp' ? -1 : 1);
    }
  }
});
$('new').addEventListener('click', async () => {
  if (!state.dir) return;
  try {
    await flushSave();
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
    /* the first placeholder is selected, so typing replaces it */
    const ed = $('editor');
    const at = ed.value.indexOf(NO_ACTIVITY);
    if (at >= 0) {
      ed.focus();
      ed.setSelectionRange(at + 2, at + NO_ACTIVITY.length);
    }
  } catch (e) { alert(e.message); }
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
  $('spell').textContent = on ? 'Spell ✓' : 'Spell ✗';
  localStorage.setItem('spellcheck', on ? 'on' : 'off');
}
$('spell').addEventListener('click', () => {
  applySpell(!$('editor').spellcheck);
  /* blur/focus makes the browser re-scan existing text immediately */
  $('editor').blur();
  $('editor').focus();
});
applySpell(localStorage.getItem('spellcheck') !== 'off');

function applyEditorFont(sans) {
  $('edwrap').classList.toggle('sans', sans);
  $('edfont').textContent = sans ? 'Sans' : 'Mono';
  localStorage.setItem('edfont', sans ? 'sans' : 'mono');
}
$('edfont').addEventListener('click', () => {
  applyEditorFont(!$('edwrap').classList.contains('sans'));
});
applyEditorFont(localStorage.getItem('edfont') === 'sans');

/* offline + installable; service workers need https or localhost */
if ('serviceWorker' in navigator && window.isSecureContext) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

$('appver').textContent = APP_VERSION;
console.log('weekly-report ' + APP_VERSION);
init().catch(e => { $('gate-err').textContent = e.message; });
