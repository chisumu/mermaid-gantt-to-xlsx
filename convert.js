#!/usr/bin/env node
'use strict';

/*
 * Convert a Mermaid gantt chart into an Excel (.xlsx) spreadsheet.
 *
 *   node convert.js <input.md> [output.xlsx]
 *
 * The input may be a markdown file containing a ```mermaid gantt block,
 * or a bare .mmd file whose contents are the gantt definition.
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const TASK_TAGS = ['done', 'active', 'crit', 'milestone'];

function fail(msg) {
  console.error('Error: ' + msg);
  process.exit(1);
}

// --- Date handling ----------------------------------------------------------

// Build a parser from a Mermaid dateFormat string (e.g. "YYYY-MM-DD").
// Recognised tokens: YYYY MM DD HH mm ss. Everything else is a literal.
function makeDateParser(fmt) {
  const tokenRe = /YYYY|MM|DD|HH|mm|ss/g;
  const groups = [];
  let regexStr = '^';
  let last = 0;
  let m;
  while ((m = tokenRe.exec(fmt)) !== null) {
    regexStr += escapeRegex(fmt.slice(last, m.index));
    switch (m[0]) {
      case 'YYYY': regexStr += '(\\d{4})'; groups.push('year'); break;
      case 'MM':   regexStr += '(\\d{1,2})'; groups.push('month'); break;
      case 'DD':   regexStr += '(\\d{1,2})'; groups.push('day'); break;
      case 'HH':   regexStr += '(\\d{1,2})'; groups.push('hour'); break;
      case 'mm':   regexStr += '(\\d{1,2})'; groups.push('minute'); break;
      case 'ss':   regexStr += '(\\d{1,2})'; groups.push('second'); break;
    }
    last = m.index + m[0].length;
  }
  regexStr += escapeRegex(fmt.slice(last)) + '$';
  const re = new RegExp(regexStr);

  return function parse(str) {
    const match = re.exec(str.trim());
    if (!match) return null;
    const parts = { year: 1970, month: 1, day: 1, hour: 0, minute: 0, second: 0 };
    groups.forEach((g, i) => { parts[g] = parseInt(match[i + 1], 10); });
    // Use UTC so spreadsheet dates are not shifted by the local timezone.
    return new Date(Date.UTC(parts.year, parts.month - 1, parts.day,
                             parts.hour, parts.minute, parts.second));
  };
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const MS_PER_DAY = 24 * 60 * 60 * 1000;

// Parse a Mermaid duration token ("3d", "24h", "2w", "30m", "45s") into ms.
function parseDuration(str) {
  const m = /^(\d+(?:\.\d+)?)\s*(ms|s|m|h|d|w)$/.exec(str.trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2]) {
    case 'ms': return n;
    case 's':  return n * 1000;
    case 'm':  return n * 60 * 1000;
    case 'h':  return n * 60 * 60 * 1000;
    case 'd':  return n * MS_PER_DAY;
    case 'w':  return n * 7 * MS_PER_DAY;
  }
  return null;
}

function addDays(date, ms) {
  return new Date(date.getTime() + ms);
}

// --- Extraction -------------------------------------------------------------

// Pull the gantt definition out of a markdown / mmd file. Returns the lines
// of the gantt block (without fences). Throws if no gantt is found.
function extractGantt(text) {
  const lines = text.split(/\r?\n/);
  // Look for a fenced ```mermaid block first.
  let start = -1;
  for (let i = 0; i < lines.length; i++) {
    if (/^```+\s*mermaid\b/i.test(lines[i])) { start = i + 1; break; }
  }
  let block;
  if (start !== -1) {
    let end = lines.length;
    for (let i = start; i < lines.length; i++) {
      if (/^```+/.test(lines[i])) { end = i; break; }
    }
    block = lines.slice(start, end);
  } else {
    block = lines;
  }
  // Confirm it's a gantt and drop everything before the `gantt` keyword.
  const ganttIdx = block.findIndex(l => /^\s*gantt\b/.test(l));
  if (ganttIdx === -1) fail('no Mermaid gantt chart found in input');
  return block.slice(ganttIdx + 1);
}

// --- Gantt parsing ----------------------------------------------------------

function parseGantt(lines) {
  const chart = { title: 'Gantt', dateFormat: 'YYYY-MM-DD', tasks: [] };
  let parseDate = makeDateParser(chart.dateFormat);
  let section = '';
  const tasksById = {};
  let lastEnd = null;

  for (let raw of lines) {
    const line = raw.replace(/%%.*$/, '').trim(); // strip %% comments
    if (!line) continue;

    let m;
    if ((m = /^dateFormat\s+(.+)$/i.exec(line))) {
      chart.dateFormat = m[1].trim();
      parseDate = makeDateParser(chart.dateFormat);
      continue;
    }
    if ((m = /^title\s+(.+)$/i.exec(line))) { chart.title = m[1].trim(); continue; }
    if ((m = /^section\s+(.+)$/i.exec(line))) { section = m[1].trim(); continue; }
    // Directives we intentionally ignore for spreadsheet output.
    if (/^(excludes|todayMarker|axisFormat|tickInterval|weekday|includes)\b/i.test(line)) continue;

    // Task line:  Label : metadata
    const colon = line.indexOf(':');
    if (colon === -1) continue;
    const label = line.slice(0, colon).trim();
    const meta = line.slice(colon + 1).trim();
    if (!label) continue;

    const task = parseTask(label, meta, { parseDate, tasksById, lastEnd, section });
    if (!task) continue;
    chart.tasks.push(task);
    if (task.id) tasksById[task.id] = task;
    if (task.end) lastEnd = task.end;
  }
  return chart;
}

function parseTask(label, meta, ctx) {
  let fields = meta.split(',').map(s => s.trim()).filter(Boolean);

  // Leading tags.
  const tags = [];
  while (fields.length && TASK_TAGS.includes(fields[0].toLowerCase())) {
    tags.push(fields.shift().toLowerCase());
  }

  const isDate = s => ctx.parseDate(s) !== null;
  const isAfter = s => /^after\b/i.test(s);
  const isDur = s => parseDuration(s) !== null;

  // Optional leading id: a token that is not a date, not "after ...", not a duration.
  let id = null;
  if (fields.length > 1 && !isDate(fields[0]) && !isAfter(fields[0]) && !isDur(fields[0])) {
    id = fields.shift();
  }

  let startField = null, endField = null;
  if (fields.length >= 2) {
    startField = fields[0];
    endField = fields[1];
  } else if (fields.length === 1) {
    if (isAfter(fields[0])) startField = fields[0];
    else endField = fields[0];
  }

  // Resolve start.
  let start = null;
  if (startField) {
    if (isAfter(startField)) {
      const refs = startField.replace(/^after\s+/i, '').split(/\s+/);
      let maxEnd = null;
      for (const r of refs) {
        const t = ctx.tasksById[r];
        if (t && t.end && (!maxEnd || t.end > maxEnd)) maxEnd = t.end;
      }
      start = maxEnd || ctx.lastEnd;
    } else {
      start = ctx.parseDate(startField);
    }
  } else {
    start = ctx.lastEnd; // implicit: begins when the previous task ends
  }

  // Resolve end.
  let end = null;
  if (endField) {
    if (isDur(endField)) {
      end = start ? addDays(start, parseDuration(endField)) : null;
    } else {
      end = ctx.parseDate(endField);
    }
  } else {
    end = start; // milestone or start-only task
  }

  const durationDays = (start && end) ? (end.getTime() - start.getTime()) / MS_PER_DAY : null;

  return { section: ctx.section, label, id, tags, start, end, durationDays };
}

// --- Excel output -----------------------------------------------------------

function writeWorkbook(chart, outPath) {
  const header = ['Section', 'Task', 'Start Date', 'End Date', 'Duration (days)', 'Tags'];
  const rows = chart.tasks.map(t => [
    t.section,
    t.label,
    t.start,
    t.end,
    t.durationDays,
    t.tags.join(', '),
  ]);

  const ws = XLSX.utils.aoa_to_sheet([header, ...rows], { cellDates: true });

  // Apply a date number format to the Start/End columns (C and D).
  for (let r = 1; r <= rows.length; r++) {
    for (const col of ['C', 'D']) {
      const cell = ws[col + (r + 1)];
      if (cell && cell.t === 'd') cell.z = 'yyyy-mm-dd';
    }
  }

  ws['!cols'] = [
    { wch: 18 }, { wch: 36 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 20 },
  ];

  const wb = XLSX.utils.book_new();
  // Sheet names are limited to 31 chars and may not contain : \ / ? * [ ]
  const sheetName = (chart.title || 'Gantt').replace(/[:\\/?*[\]]/g, ' ').slice(0, 31) || 'Gantt';
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, outPath, { cellDates: true });
}

// --- Main -------------------------------------------------------------------

function main() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    fail('usage: node convert.js <input.md> [output.xlsx]');
  }
  const inPath = args[0];
  const outPath = args[1] || inPath.replace(/\.[^.]+$/, '') + '.xlsx';

  if (!fs.existsSync(inPath)) fail('input file not found: ' + inPath);

  const text = fs.readFileSync(inPath, 'utf8');
  const chart = parseGantt(extractGantt(text));

  if (chart.tasks.length === 0) fail('no tasks found in the gantt chart');

  writeWorkbook(chart, outPath);
  console.log(`Wrote ${chart.tasks.length} task(s) to ${path.resolve(outPath)}`);
}

main();
