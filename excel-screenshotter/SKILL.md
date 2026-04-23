---
name: excel-screenshotter
slug: excel-screenshotter
version: 1.0.0
homepage: https://clawic.com/skills/excel-screenshotter
description: "Capture screenshots from Microsoft Excel workbooks by exporting chart sheets, embedded charts, worksheet ranges, or Excel tables to PNG. Use when (1) the user needs an Excel screenshot artifact, preview, or evidence image; (2) the task involves `.xlsx`, `.xlsm`, or `.xls` files on Windows with Excel installed; (3) the agent should inspect sheet objects first, then return a concrete screenshot path or structured JSON result."
metadata: {"clawdbot":{"emoji":"📸","requires":{"bins":["python","pwsh"],"os":["win32"]}}}
---

## When to Use

Use when the user wants a screenshot from an Excel workbook rather than a data edit. This skill is for exporting a chart, a table-like range, or a worksheet region into a PNG that another agent step can return as an artifact.

Use [excel_capture.py](./excel_capture.py) as the primary entrypoint. It wraps [export_excel_range_image.ps1](./export_excel_range_image.ps1), adds timeout handling, auto-generates output paths when needed, and can return structured JSON for downstream automation.

By default, screenshots are written into [screenshots](./screenshots/) under this skill folder. This is the required default location for this skill: do not place screenshot outputs in the workspace root or elsewhere unless the user explicitly asks for a different folder.

## Prerequisites

- Windows with Microsoft Excel installed.
- `python` available on `PATH`.
- `pwsh` or `powershell` available on `PATH`.
- The workbook must be accessible on the local machine.

## Default Workflow

### 1. Inspect before exporting

Do not guess whether a sheet is a chart, table, or plain cell range. First inspect the sheet:

```bash
python ./excel_capture.py "<workbook>" "<sheet>" --list-objects --json
```

This returns structured metadata such as:

- `sheet_type`
- `used_range`
- `used_range_columns`
- `used_range_rows`
- `charts`
- `tables`

### 2. Choose the export mode explicitly

Prefer explicit modes over `auto` for reliable automation:

- `--type chart`: chart sheet or first embedded chart on a worksheet
- `--type table`: first Excel `ListObject` or a named table via `--table-name`
- `--type range`: an explicit worksheet range via `--range "A1:K30"`
- `--type auto`: only when interactive use is fine and ambiguity is acceptable

### 3. Export with JSON enabled

For skill and agent integration, prefer structured results:

```bash
python ./excel_capture.py "<workbook>" "<sheet>" --type chart --json
python ./excel_capture.py "<workbook>" "<sheet>" --type table --table-name "Table1" --json
python ./excel_capture.py "<workbook>" "<sheet>" --type range --range "A1:AE4" --json
```

If `output_path` is omitted, the Python wrapper creates a temporary PNG automatically and returns its absolute path.

If `output_path` is a relative path, it is resolved under [screenshots](./screenshots/). Use an absolute path only when you intentionally need to override that behavior.

## Output Contract

When `--json` is used, expect machine-readable output.

## Storage Rule

- All screenshot outputs from this skill must live under this skill's `screenshots/` folder.
- Do not place screenshot files in the workspace root.
- If `output_path` is omitted, the wrapper must generate a file inside `./screenshots/` and return that path.
- If a caller supplies a relative `output_path`, resolve it under `./screenshots/`.
- Only use a different folder when the user explicitly requests it.

Successful export returns fields like:

- `status: ok`
- `mode: export`
- `output_path`
- `output_exists`
- `output_size_bytes`
- `image.width`
- `image.height`

List mode returns fields like:

- `status: ok`
- `mode: list`
- `sheet_type`
- `sheet_name`
- `used_range`
- `charts`
- `tables`

Failures return:

- `status: error`
- `error_code`
- `message`

## Recommended Decision Rules

### Chart requests

- If `sheet_type` is `ChartSheet`, use `--type chart`.
- If `charts` is non-empty on a worksheet and the user asked for a chart screenshot, use `--type chart`.

### Table requests

- Use `--type table` only when `tables` contains real Excel table objects.
- If `tables` is empty, the visible table is probably just a cell range. Use `--type range` instead.

### Range requests

- Use `--type range` when the target is a plain worksheet region.
- Prefer explicit `--range` values over `UsedRange` when the user cares about exact framing.

## High-Value Patterns

- Inspect first, export second.
- Prefer `--json` for skill chaining.
- Prefer `--type range --range ...` for dashboard-like sheets with no real Excel table objects.
- Prefer `--type chart` for chart workbooks or sheets with embedded charts.
- Use `--timeout <seconds>` when the workbook is large or Excel startup is slow.

## Common Traps

- A sheet that looks like a table may not be an Excel `ListObject`.
- `Range.CopyPicture()` exports cell content, not floating chart objects.
- `auto` mode is convenient for manual use but weaker than explicit mode selection for agents.
- PowerShell script execution policy can block direct `.ps1` invocation; the Python wrapper is the safer entrypoint.
- This workflow depends on desktop Excel COM automation and is Windows-only.

## Suggested Agent Procedure

1. Run `--list-objects --json`.
2. Pick `chart`, `table`, or `range` explicitly.
3. Export with `--json`.
4. Read `output_path` from the JSON result.
5. Return the PNG artifact to the caller.

## Artifact Location

- Prefer omitting `output_path` and let the wrapper create the PNG inside [screenshots](./screenshots/).
- If you pass a relative `output_path`, it will still be placed under [screenshots](./screenshots/).
- Keep screenshot artifacts inside this skill folder unless there is a specific reason to write elsewhere.

## Image Handling (Discord)
 
When you need to send an image to Discord:
 
1. NEVER return local file paths (e.g. /workspace/xxx.png)
2. NEVER return base64 strings
3. ALWAYS ensure the image is accessible to Discord via one of the following:
 
### Preferred: Upload as attachment
- If a tool is available to send Discord messages with files:
  - Use that tool
  - Attach the image file directly
  - Do not convert it to text
 
### Alternative: Public URL
- If attachment is not available:
  - Upload the image to a public storage service
  - Return a direct image URL (must end with .png/.jpg)
  - Ensure Content-Type is image/*

