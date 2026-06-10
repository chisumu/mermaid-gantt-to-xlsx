# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

A zero-build Node.js tool that converts Mermaid gantt charts into Excel (.xlsx) spreadsheets.

## Commands

```bash
npm install                      # install the single dependency (xlsx / SheetJS)
node convert.js <input.md>       # writes <input>.xlsx alongside the input
node convert.js in.md out.xlsx   # explicit output path
```

There is no build, lint, or test setup — run `convert.js` directly with Node.

## Architecture

Everything lives in `convert.js`, structured as a pipeline:

1. **extractGantt** — pulls the gantt definition out of a markdown ```` ```mermaid ```` fenced block (or treats the whole file as the definition for bare `.mmd`), then drops everything before the `gantt` keyword.
2. **parseGantt** — line-by-line scan. Handles `dateFormat`, `title` (→ sheet name), `section`, and task lines (`Label : metadata`). Ignores non-spreadsheet directives (`excludes`, `todayMarker`, `axisFormat`, etc.) and `%%` comments.
3. **parseTask** — splits the post-colon metadata by commas and classifies fields: leading tags (`done`/`active`/`crit`/`milestone`), an optional task id, then start/end. Start can be an absolute date or `after <id…>` (resolved to the max end date of referenced tasks). End can be an absolute date or a duration (`3d`, `24h`, `2w`). An absent start defaults to the previous task's end date (implicit chaining).
4. **writeWorkbook** — emits columns Section, Task, Start Date, End Date, Duration (days), Tags. Dates are written as real Excel date cells (`cellDates`, `yyyy-mm-dd` number format), not strings.

### Key conventions

- **Dates are parsed in UTC** to avoid timezone shifting in the spreadsheet. `makeDateParser` builds a regex from the Mermaid `dateFormat` string (tokens: `YYYY MM DD HH mm ss`; anything else is a literal).
- Task id detection is heuristic: the first metadata field that is *not* a date, *not* `after …`, and *not* a duration is treated as the id.
- `test.md` is a sample input covering tags, `after` dependencies, durations, and a milestone.
