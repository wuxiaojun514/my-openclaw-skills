import argparse
import json
import os
import struct
import shutil
import subprocess
import sys
import uuid
from pathlib import Path


class ExcelCaptureError(RuntimeError):
    pass


def _screenshots_dir():
    screenshots_path = Path(__file__).with_name("screenshots")
    screenshots_path.mkdir(parents=True, exist_ok=True)
    return screenshots_path


def _find_powershell() -> str:
    if os.name == "nt":
        candidates = ("powershell", "pwsh")
    else:
        candidates = ("pwsh", "powershell")

    for command in candidates:
        resolved = shutil.which(command)
        if resolved:
            return resolved
    raise ExcelCaptureError(
        "PowerShell was not found on PATH. Install PowerShell 7 (pwsh) or ensure powershell.exe is available."
    )


def _normalize_powershell_error(raw_error, sheet_name):
    error_text = (raw_error or "").strip()
    lowered = error_text.lower()
    worksheet_markers = (
        "worksheets.item",
        "subscript out of range",
        "unable to get the item property of the worksheets class",
        "索引超出范围",
        "下标越界",
    )

    if any(marker in lowered for marker in worksheet_markers):
        return f"Worksheet not found: {sheet_name}"

    if not error_text:
        return "Excel export failed without output."

    return f"Excel export failed: {error_text}"


def _create_temp_output_path(sheet_name):
    safe_sheet_name = "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in sheet_name).strip("_")
    if not safe_sheet_name:
        safe_sheet_name = "sheet"

    return str(
        (_screenshots_dir() / f"excel_capture_{safe_sheet_name}_{uuid.uuid4().hex}.png").resolve()
    )


def _resolve_output_path(output_path):
    if not output_path:
        return None

    output = Path(output_path).expanduser()
    if output.is_absolute():
        output.parent.mkdir(parents=True, exist_ok=True)
        return str(output.resolve())

    resolved = _screenshots_dir() / output
    resolved.parent.mkdir(parents=True, exist_ok=True)
    return str(resolved.resolve())


def _read_png_size(image_path):
    with open(image_path, "rb") as image_file:
        signature = image_file.read(8)
        if signature != b"\x89PNG\r\n\x1a\n":
            return None

        length = struct.unpack(">I", image_file.read(4))[0]
        chunk_type = image_file.read(4)
        if length != 13 or chunk_type != b"IHDR":
            return None

        width, height = struct.unpack(">II", image_file.read(8))
        return {"width": width, "height": height}


def _build_success_result(mode, output_path=None, metadata=None):
    result = {
        "status": "ok",
        "mode": mode,
    }

    if metadata:
        result.update(metadata)

    if output_path:
        output_file = Path(output_path)
        result.update(
            {
                "output_path": str(output_file),
                "output_exists": output_file.exists(),
            }
        )

        if output_file.exists():
            result["output_size_bytes"] = output_file.stat().st_size
            image_size = _read_png_size(output_file)
            if image_size:
                result["image"] = image_size

    return result


def _build_error_result(message, error_code, raw_error=None, details=None):
    result = {
        "status": "error",
        "error_code": error_code,
        "message": message,
    }

    if raw_error:
        result["raw_error"] = raw_error

    if details:
        result["details"] = details

    return result


def _classify_error(message):
    lowered = message.lower()

    if "worksheet not found" in lowered:
        return "worksheet_not_found"
    if "excel workbook not found" in lowered:
        return "workbook_not_found"
    if "powershell was not found" in lowered:
        return "powershell_not_found"
    if "timed out" in lowered:
        return "timeout"
    if "table not found" in lowered or "no excel table" in lowered:
        return "table_not_found"
    if "no chart was found" in lowered:
        return "chart_not_found"
    if "outputpath is required" in lowered:
        return "missing_output_path"

    return "export_failed"


def _parse_json_output(stdout_text):
    output = (stdout_text or "").strip()
    if not output:
        return None

    try:
        return json.loads(output)
    except json.JSONDecodeError:
        return None


def capture_excel_sheet(
    file_path,
    sheet_name,
    output_path=None,
    range_address=None,
    visible=False,
    export_type="auto",
    table_name=None,
    list_objects=False,
    timeout_seconds=120,
):
    script_path = Path(__file__).with_name("export_excel_range_image.ps1")
    if not script_path.exists():
        raise ExcelCaptureError(f"PowerShell helper script not found: {script_path}")

    workbook = Path(file_path).expanduser()
    if not workbook.exists():
        raise ExcelCaptureError(f"Excel workbook not found: {workbook}")

    workbook_path = str(workbook.resolve())
    image_path = _resolve_output_path(output_path)
    if not list_objects and not image_path:
        image_path = _create_temp_output_path(sheet_name)

    command = [
        _find_powershell(),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
        "-WorkbookPath",
        workbook_path,
        "-SheetName",
        sheet_name,
        "-ExportType",
        export_type.capitalize(),
        "-Json",
    ]

    if image_path:
        command.extend(["-OutputPath", image_path])

    if range_address:
        command.extend(["-RangeAddress", range_address])

    if table_name:
        command.extend(["-TableName", table_name])

    if list_objects:
        command.append("-ListObjects")

    if visible:
        command.append("-Visible")

    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        raise ExcelCaptureError(
            f"Excel export timed out after {timeout_seconds} seconds."
        ) from exc

    json_output = _parse_json_output(completed.stdout)

    if completed.returncode != 0:
        error_output = (completed.stderr or completed.stdout).strip()
        raise ExcelCaptureError(_normalize_powershell_error(error_output, sheet_name))

    if list_objects:
        if json_output is not None:
            return _build_success_result(
                mode="list",
                metadata=json_output,
            )
        return _build_success_result(
            mode="list",
            metadata={"raw_output": completed.stdout.strip()},
        )

    if json_output is not None:
        return _build_success_result(
            mode="export",
            output_path=image_path,
            metadata=json_output,
        )

    return _build_success_result(mode="export", output_path=image_path)


def _build_parser():
    parser = argparse.ArgumentParser(description="Export an Excel sheet or range to a PNG image via PowerShell.")
    parser.add_argument("workbook_path", help="Path to the Excel workbook.")
    parser.add_argument("sheet_name", help="Worksheet name to export.")
    parser.add_argument("output_path", nargs="?", help="PNG output path.")
    parser.add_argument(
        "--type",
        dest="export_type",
        choices=("auto", "chart", "table", "range"),
        default="auto",
        help="Export mode: auto, chart, table, or range.",
    )
    parser.add_argument("--range", dest="range_address", help="Optional Excel range, for example A1:K30.")
    parser.add_argument("--table-name", help="Optional Excel table name when --type table is used.")
    parser.add_argument("--list-objects", action="store_true", help="List charts, tables, and used range on the sheet.")
    parser.add_argument("--timeout", dest="timeout_seconds", type=int, default=120, help="Timeout in seconds for the Excel export process.")
    parser.add_argument("--visible", action="store_true", help="Show the Excel window while exporting.")
    parser.add_argument("--json", action="store_true", help="Print structured JSON output for skill integration.")
    return parser


if __name__ == "__main__":
    args = _build_parser().parse_args()

    if not args.list_objects and not args.output_path:
        args.output_path = _create_temp_output_path(args.sheet_name)

    try:
        result = capture_excel_sheet(
            args.workbook_path,
            args.sheet_name,
            args.output_path,
            range_address=args.range_address,
            visible=args.visible,
            export_type=args.export_type,
            table_name=args.table_name,
            list_objects=args.list_objects,
            timeout_seconds=args.timeout_seconds,
        )
    except Exception as exc:
        error_result = _build_error_result(
            message=str(exc),
            error_code=_classify_error(str(exc)),
        )
        if args.json:
            print(json.dumps(error_result, ensure_ascii=False))
        else:
            print(error_result["message"], file=sys.stderr)
        sys.exit(1)

    if args.json:
        print(json.dumps(result, ensure_ascii=False))
    elif args.list_objects:
        if "raw_output" in result:
            print(result["raw_output"])
        else:
            print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Image exported to {result['output_path']}")