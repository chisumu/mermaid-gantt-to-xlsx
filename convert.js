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
const ExcelJS = require('exceljs');

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

  // FIXME: Graduate these to actual utilities
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

  return { id: id, section: ctx.section, label: label, tags: tags, start: start, end: end, durationDays: durationDays };
}

// --- Excel output -----------------------------------------------------------

// Bar fill colours (ARGB) keyed by tag, in priority order. The first tag a
// task carries that appears here wins; tagless tasks use the default.
const BAR_COLORS = {
  crit:      'FFE53935', // red
  done:      'FF43A047', // green
  active:    'FF1E88E5', // blue
  milestone: 'FFFB8C00', // orange
};
const DEFAULT_BAR_COLOR = 'FF26A69A'; // teal
const TAG_PRIORITY = ['crit', 'active', 'done'];

function barColor(task) {
  if (task.tags.includes('milestone')) return BAR_COLORS.milestone;
  for (const tag of TAG_PRIORITY) {
    if (task.tags.includes(tag)) return BAR_COLORS[tag];
  }
  if (task.tags.includes('done')) return BAR_COLORS.done;
  return DEFAULT_BAR_COLOR;
}

// The five fixed info columns shown to the left of the calendar grid.
const INFO_COLUMNS = [
  { header: 'Section',         width: 18 },
  { header: 'Task',            width: 34 },
  { header: 'Start',           width: 12 },
  { header: 'End',             width: 12 },
  { header: 'Days',            width: 7  },
];
const FIRST_GRID_COL = INFO_COLUMNS.length + 1; // 1-based column of first date

const DAY = MS_PER_DAY;
const WEEK = 7 * DAY;

// Choose a calendar granularity so the grid stays a reasonable width.
function chooseUnit(spanDays) {
  if (spanDays <= 70) return 'day';
  if (spanDays <= 540) return 'week';
  return 'month';
}

// Floor a date to the start of its unit (UTC).
function unitStart(date, unit) {
  const d = new Date(date.getTime());
  d.setUTCHours(0, 0, 0, 0);
  if (unit === 'week') {
    // Snap back to Monday.
    const dow = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - dow);
  } else if (unit === 'month') {
    d.setUTCDate(1);
  }
  return d;
}

// Advance a date by one unit (UTC).
function unitNext(date, unit) {
  const d = new Date(date.getTime());
  if (unit === 'day') d.setUTCDate(d.getUTCDate() + 1);
  else if (unit === 'week') d.setUTCDate(d.getUTCDate() + 7);
  else d.setUTCMonth(d.getUTCMonth() + 1);
  return d;
}

// Build the ordered list of calendar columns spanning [min, max].
function buildPeriods(min, max, unit) {
  const periods = [];
  let cur = unitStart(min, unit);
  while (cur <= max) {
    periods.push({ start: cur, end: unitNext(cur, unit) });
    cur = unitNext(cur, unit);
  }
  return periods;
}

function periodLabel(start, unit) {
  const y = start.getUTCFullYear();
  const mo = String(start.getUTCMonth() + 1).padStart(2, '0');
  const d = String(start.getUTCDate()).padStart(2, '0');
  if (unit === 'month') return `${y}-${mo}`;
  return `${mo}-${d}`; // day / week: month-day
}

async function writeWorkbook(chart, outPath) {
  // Determine the overall date span from tasks that have dates.
  let min = null, max = null;
  for (const t of chart.tasks) {
    if (t.start && (!min || t.start < min)) min = t.start;
    if (t.end && (!max || t.end > max)) max = t.end;
  }
  const hasGrid = !!(min && max);

  const unit = hasGrid ? chooseUnit((max - min) / DAY) : 'day';
  const periods = hasGrid ? buildPeriods(min, max, unit) : [];

  const wb = new ExcelJS.Workbook();
  const sheetName = (chart.title || 'Gantt').replace(/[:\\/?*[\]]/g, ' ').slice(0, 31) || 'Gantt';
  const ws = wb.addWorksheet(sheetName, {
    views: [{ state: 'frozen', xSplit: INFO_COLUMNS.length, ySplit: 2 }],
  });

  // Column widths.
  ws.columns = [
    ...INFO_COLUMNS.map(c => ({ width: c.width })),
    ...periods.map(() => ({ width: unit === 'month' ? 9 : 4 })),
  ];

  // Row 1: title banner across all columns.
  const lastCol = INFO_COLUMNS.length + periods.length;
  ws.mergeCells(1, 1, 1, Math.max(lastCol, INFO_COLUMNS.length));
  const titleCell = ws.getCell(1, 1);
  titleCell.value = chart.title || 'Gantt';
  titleCell.font = { bold: true, size: 14 };
  titleCell.alignment = { vertical: 'middle' };

  // Row 2: header (info columns + date labels).
  const headerRow = ws.getRow(2);
  INFO_COLUMNS.forEach((c, i) => { headerRow.getCell(i + 1).value = c.header; });
  periods.forEach((p, i) => {
    const cell = headerRow.getCell(FIRST_GRID_COL + i);
    cell.value = periodLabel(p.start, unit);
    cell.alignment = { textRotation: 90, horizontal: 'center' };
  });
  headerRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECEFF1' } };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFB0BEC5' } } };
  });

  // Task rows.
  chart.tasks.forEach((t, idx) => {
    const r = ws.getRow(3 + idx);
    r.getCell(1).value = t.section || '';
    r.getCell(2).value = t.label;
    if (t.start) { r.getCell(3).value = t.start; r.getCell(3).numFmt = 'yyyy-mm-dd'; }
    if (t.end)   { r.getCell(4).value = t.end;   r.getCell(4).numFmt = 'yyyy-mm-dd'; }
    if (t.durationDays != null) r.getCell(5).value = t.durationDays;

    if (!hasGrid || !t.start) return;
    const color = barColor(t);
    const isMilestone = t.tags.includes('milestone') || (t.end && +t.end === +t.start);

    periods.forEach((p, i) => {
      // A period is part of the bar if it overlaps [start, end). Milestones
      // (zero-length) mark the single period that contains their date.
      const overlaps = isMilestone
        ? (t.start >= p.start && t.start < p.end)
        : (t.start < p.end && t.end > p.start);
      if (!overlaps) return;
      const cell = r.getCell(FIRST_GRID_COL + i);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
      if (isMilestone) {
        cell.value = '◆'; // ◆
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
        cell.alignment = { horizontal: 'center' };
      }
    });
  });

  await wb.xlsx.writeFile(outPath);
}

// --- Main -------------------------------------------------------------------

async function main() {
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

  await writeWorkbook(chart, outPath);
  console.log(`Wrote ${chart.tasks.length} task(s) to ${path.resolve(outPath)}`);
}

main().catch(err => fail(err && err.message ? err.message : String(err)));
