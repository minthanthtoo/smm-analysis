from __future__ import annotations

import calendar
import html
import os
import re
import socket
import time
from collections import Counter, OrderedDict
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, render_template, request
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = ROOT_DIR / "uploads"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

MONTH_LOOKUP = {
    name.lower(): idx
    for idx, name in enumerate(calendar.month_name)
    if idx
}
MONTH_LOOKUP.update(
    {
        name.lower(): idx
        for idx, name in enumerate(calendar.month_abbr)
        if idx
    }
)

TOWNSHIP_ALIAS = {
    "meiktila": "meiktila",
    "tharzi": "tharzi",
    "tarzi": "tharzi",
    "pyawbwe": "pyawbwe",
    "pyawbwe": "pyawbwe",
    "wanttwin": "wantwin",
    "wantwin": "wantwin",
    "mahaling": "mahlaing",
    "mahlaing": "mahlaing",
    "yamethin": "yamethin",
    "kyaukpadaung": "kyaukpandaung",
    "kyaukpandaung": "kyaukpandaung",
    "taungthar": "taungthar",
    "taugthar": "taungthar",
    "myingyan": "myingan",
    "myingan": "myingan",
    "bagan": "bagan",
    "pakokku": "pakokku",
    "pakoku": "pakokku",
}


app = Flask(__name__)


def normalize_token(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def canonical_sheet_name(sheet_name: str) -> str:
    token = normalize_token(sheet_name)
    return TOWNSHIP_ALIAS.get(token, token)


def parse_month_from_text(text: str) -> int | None:
    tokens = [t for t in re.split(r"[^a-z]+", text.lower()) if t]
    matched = [MONTH_LOOKUP[token] for token in tokens if token in MONTH_LOOKUP]
    if not matched:
        return None
    if len(set(matched)) > 1:
        return None
    return matched[0]


def extract_year_from_text(text: str) -> int | None:
    match = re.search(r"(20\d{2})", text)
    if not match:
        return None
    return int(match.group(1))


def parse_month_header(value: Any, fallback_year: int | None) -> tuple[int, int] | None:
    if isinstance(value, datetime):
        return value.year, value.month
    if isinstance(value, date):
        return value.year, value.month

    if isinstance(value, str):
        month_num = parse_month_from_text(value)
        if month_num is None:
            return None
        year = extract_year_from_text(value) or fallback_year
        if year is None:
            return None
        return year, month_num

    return None


def parse_month_header_loose(
    value: Any, fallback_year: int | None
) -> tuple[int | None, int] | None:
    if isinstance(value, datetime):
        return value.year, value.month
    if isinstance(value, date):
        return value.year, value.month

    if isinstance(value, str):
        month_num = parse_month_from_text(value)
        if month_num is None:
            return None
        year = extract_year_from_text(value) or fallback_year
        return year, month_num

    return None


def parse_numeric(value: Any) -> int | float | None:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        try:
            number = float(stripped)
        except ValueError:
            return None
    elif isinstance(value, (int, float)):
        number = float(value)
    else:
        return None

    if number.is_integer():
        return int(number)
    return round(number, 4)


def parse_year_context(worksheet, header_row: int) -> int | None:
    for row_idx in range(1, header_row):
        for col_idx in range(1, min(worksheet.max_column, 30) + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if isinstance(value, (int, float)) and 2000 <= int(value) <= 2100:
                return int(value)
            if isinstance(value, str):
                year = extract_year_from_text(value)
                if year:
                    return year
    return None


def find_header_layout(worksheet) -> dict[str, int] | None:
    max_scan_col = min(worksheet.max_column, 60)
    for row_idx in range(1, min(worksheet.max_row, 12) + 1):
        columns: dict[str, int] = {}
        for col_idx in range(1, max_scan_col + 1):
            token = normalize_token(worksheet.cell(row=row_idx, column=col_idx).value)
            if not token:
                continue
            if token == "sr":
                columns["sr"] = col_idx
            elif token == "productname":
                columns["product_name"] = col_idx
            elif token == "ml":
                columns["ml"] = col_idx
            elif token == "packing":
                columns["packing"] = col_idx

        required = {"product_name", "ml", "packing"}
        if required.issubset(columns):
            columns["header_row"] = row_idx
            columns["metrics_row"] = row_idx + 1
            return columns
    return None


def parse_month_groups(worksheet, layout: dict[str, int]) -> list[dict[str, Any]]:
    header_row = layout["header_row"]
    metrics_row = layout["metrics_row"]
    fallback_year = parse_year_context(worksheet, header_row)
    start_col = layout["packing"] + 1

    month_groups: list[dict[str, Any]] = []
    seen_keys: set[str] = set()

    for col_idx in range(start_col, worksheet.max_column - 1):
        metric_tokens = [
            normalize_token(worksheet.cell(row=metrics_row, column=col_idx + offset).value)
            for offset in range(3)
        ]
        if metric_tokens[:2] != ["pk", "bottle"]:
            continue
        if metric_tokens[2] not in {"liter", "litre"}:
            continue

        header_value = worksheet.cell(row=header_row, column=col_idx).value
        parsed_month = parse_month_header(header_value, fallback_year)
        if not parsed_month:
            continue

        year, month = parsed_month
        month_key = f"{year:04d}-{month:02d}"
        if month_key in seen_keys:
            continue

        seen_keys.add(month_key)
        month_groups.append(
            {
                "key": month_key,
                "label": f"{calendar.month_abbr[month]} {year}",
                "year": year,
                "month": month,
                "col": col_idx,
            }
        )

    month_groups.sort(key=lambda item: item["key"])
    return month_groups


def row_identity_key(product_name: str, ml: Any, packing: Any) -> str:
    ml_token = "" if ml is None else str(ml).strip()
    packing_token = normalize_token(packing)
    return f"{normalize_token(product_name)}|{ml_token}|{packing_token}"


def parse_rows(worksheet, layout: dict[str, int], month_groups: list[dict[str, Any]]) -> list[dict[str, Any]]:
    sr_col = layout.get("sr")
    product_col = layout["product_name"]
    ml_col = layout["ml"]
    packing_col = layout["packing"]
    data_start = layout["metrics_row"] + 1

    parsed_rows: list[dict[str, Any]] = []
    blank_streak = 0

    for row_idx in range(data_start, worksheet.max_row + 1):
        sr_value = worksheet.cell(row=row_idx, column=sr_col).value if sr_col else None
        product_value = worksheet.cell(row=row_idx, column=product_col).value
        ml_value = worksheet.cell(row=row_idx, column=ml_col).value
        packing_value = worksheet.cell(row=row_idx, column=packing_col).value

        values_by_month: dict[str, dict[str, int | float | None]] = {}
        has_month_data = False

        for group in month_groups:
            col_idx = group["col"]
            pk_val = parse_numeric(worksheet.cell(row=row_idx, column=col_idx).value)
            bottle_val = parse_numeric(worksheet.cell(row=row_idx, column=col_idx + 1).value)
            liter_val = parse_numeric(worksheet.cell(row=row_idx, column=col_idx + 2).value)
            values_by_month[group["key"]] = {
                "pk": pk_val,
                "bottle": bottle_val,
                "liter": liter_val,
            }
            if pk_val is not None or bottle_val is not None or liter_val is not None:
                has_month_data = True

        has_identity = any(
            value not in (None, "") for value in (sr_value, product_value, ml_value, packing_value)
        )

        if not has_identity and not has_month_data:
            blank_streak += 1
            if blank_streak >= 20:
                break
            continue

        blank_streak = 0
        if product_value in (None, ""):
            continue

        product_name = str(product_value).strip()
        parsed_rows.append(
            {
                "sr": sr_value,
                "product_name": product_name,
                "ml": ml_value,
                "packing": packing_value,
                "row_key": row_identity_key(product_name, ml_value, packing_value),
                "values": values_by_month,
            }
        )

    return parsed_rows


def parse_fixed_sheet(worksheet, sheet_name: str) -> dict[str, Any] | None:
    layout = find_header_layout(worksheet)
    if not layout:
        return None

    month_groups = parse_month_groups(worksheet, layout)
    if len(month_groups) < 2:
        return None

    parsed_rows = parse_rows(worksheet, layout, month_groups)
    if not parsed_rows:
        return None

    return {
        "sheet_name": sheet_name,
        "canonical": canonical_sheet_name(sheet_name),
        "months": [
            {
                "key": group["key"],
                "label": group["label"],
                "year": group["year"],
                "month": group["month"],
            }
            for group in month_groups
        ],
        "rows": parsed_rows,
    }


def parse_workbook(path: Path) -> OrderedDict[str, dict[str, Any]]:
    # Use normal mode because this parser relies on random cell access.
    workbook = load_workbook(path, read_only=False, data_only=True)
    parsed_sheets: OrderedDict[str, dict[str, Any]] = OrderedDict()

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        parsed_sheet = parse_fixed_sheet(worksheet, sheet_name)
        if not parsed_sheet:
            continue

        canonical = parsed_sheet["canonical"]
        if canonical in parsed_sheets:
            # Prefer the first matching sheet in workbook order.
            continue
        parsed_sheets[canonical] = parsed_sheet

    return parsed_sheets


@lru_cache(maxsize=16)
def parse_workbook_cached(path_str: str, mtime_ns: int) -> OrderedDict[str, dict[str, Any]]:
    _ = mtime_ns
    return parse_workbook(Path(path_str))


@lru_cache(maxsize=8)
def load_workbook_cached(path_str: str, mtime_ns: int):
    _ = mtime_ns
    return load_workbook(Path(path_str), read_only=False, data_only=True)


@lru_cache(maxsize=8)
def workbook_sheetnames_cached(path_str: str, mtime_ns: int) -> tuple[str, ...]:
    workbook = load_workbook_cached(path_str, mtime_ns)
    return tuple(workbook.sheetnames)


def discover_workbooks() -> list[Path]:
    workbooks: list[Path] = []
    search_dirs = [ROOT_DIR]
    if UPLOAD_DIR.exists():
        search_dirs.append(UPLOAD_DIR)

    for folder in search_dirs:
        for path in folder.iterdir():
            if not path.is_file():
                continue
            if path.name.startswith("~$"):
                continue
            if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            workbooks.append(path)

    workbooks.sort(key=lambda p: p.name.lower())
    return workbooks


def workbook_name_score(path: Path, role: str) -> int:
    token = normalize_token(path.stem)
    if role == "main":
        keywords = ("overview", "main", "master", "mhl", "summary", "analysis")
    else:
        keywords = ("detail", "reference", "township", "breakdown", "summary", "mtl")

    score = 0
    for idx, keyword in enumerate(keywords):
        if keyword in token:
            score += max(1, 12 - (idx * 2))
    return score


def default_workbook(workbooks: list[Path], role: str, avoid_name: str | None = None) -> Path:
    candidates = [path for path in workbooks if path.name != avoid_name] or workbooks
    ranked = sorted(
        candidates,
        key=lambda path: (
            workbook_name_score(path, role),
            path.stat().st_mtime_ns,
            path.name.lower(),
        ),
        reverse=True,
    )
    return ranked[0]


def resolve_workbook_pair(
    main_name: str | None, reference_name: str | None
) -> dict[str, Any]:
    workbooks = discover_workbooks()
    if not workbooks:
        raise FileNotFoundError("No supported Excel files found in this folder.")

    by_name = {path.name: path for path in workbooks}
    default_main = default_workbook(workbooks, "main")
    default_reference = default_workbook(
        workbooks,
        "reference",
        avoid_name=default_main.name if len(workbooks) > 1 else None,
    )

    if main_name and main_name not in by_name:
        raise ValueError(f"Unknown main workbook: {main_name}")
    if reference_name and reference_name not in by_name:
        raise ValueError(f"Unknown reference workbook: {reference_name}")

    main_path = by_name.get(main_name) if main_name else default_main
    reference_path = by_name.get(reference_name) if reference_name else default_reference

    return {
        "available": [path.name for path in workbooks],
        "default_main": default_main.name,
        "default_reference": default_reference.name,
        "main_path": main_path,
        "reference_path": reference_path,
    }


def load_viewer_data(main_path: Path, reference_path: Path) -> dict[str, Any]:
    main_stat = main_path.stat()
    ref_stat = reference_path.stat()
    main_data = parse_workbook_cached(str(main_path), main_stat.st_mtime_ns)
    reference_data = parse_workbook_cached(str(reference_path), ref_stat.st_mtime_ns)
    version = f"{main_stat.st_mtime_ns}:{ref_stat.st_mtime_ns}"
    main_sheet_names = workbook_sheetnames_cached(str(main_path), main_stat.st_mtime_ns)
    reference_sheet_names = workbook_sheetnames_cached(str(reference_path), ref_stat.st_mtime_ns)

    sheet_index = []
    fixed_by_sheet_name: dict[str, dict[str, Any]] = {}
    for canonical, main_sheet in main_data.items():
        reference_sheet = reference_data.get(canonical)
        fixed_by_sheet_name[main_sheet["sheet_name"]] = {
            "canonical": canonical,
            "has_reference": bool(reference_sheet),
        }
        sheet_index.append(
            {
                "canonical": canonical,
                "main_sheet_name": main_sheet["sheet_name"],
                "reference_sheet_name": (
                    reference_sheet["sheet_name"] if reference_sheet else None
                ),
                "main_month_count": len(main_sheet["months"]),
                "reference_month_count": (
                    len(reference_sheet["months"]) if reference_sheet else 0
                ),
                "main_row_count": len(main_sheet["rows"]),
            }
        )

    main_sheet_tabs = []
    for sheet_name in main_sheet_names:
        fixed_info = fixed_by_sheet_name.get(sheet_name)
        main_sheet_tabs.append(
            {
                "sheet_name": sheet_name,
                "canonical": fixed_info["canonical"] if fixed_info else None,
                "filterable": bool(fixed_info),
                "has_reference": fixed_info["has_reference"] if fixed_info else False,
            }
        )

    fixed_main_by_sheet_name: dict[str, dict[str, Any]] = {}
    for canonical, main_sheet in main_data.items():
        fixed_main_by_sheet_name[main_sheet["sheet_name"]] = {"canonical": canonical}

    fixed_reference_by_sheet_name: dict[str, dict[str, Any]] = {}
    for canonical, reference_sheet in reference_data.items():
        fixed_reference_by_sheet_name[reference_sheet["sheet_name"]] = {
            "canonical": canonical,
            "has_main": canonical in main_data,
        }

    reference_sheet_tabs = []
    for sheet_name in reference_sheet_names:
        fixed_info = fixed_reference_by_sheet_name.get(sheet_name)
        main_info = fixed_main_by_sheet_name.get(sheet_name)
        canonical = fixed_info["canonical"] if fixed_info else (main_info["canonical"] if main_info else None)
        reference_sheet_tabs.append(
            {
                "sheet_name": sheet_name,
                "canonical": canonical,
                "filterable": bool(fixed_info),
                "has_main": fixed_info["has_main"] if fixed_info else bool(main_info),
            }
        )

    return {
        "main_workbook": main_path.name,
        "reference_workbook": reference_path.name,
        "version": version,
        "main": main_data,
        "reference": reference_data,
        "sheet_index": sheet_index,
        "main_sheet_tabs": main_sheet_tabs,
        "reference_sheet_tabs": reference_sheet_tabs,
    }


def color_to_css_hex(color) -> str | None:
    if not color:
        return None
    rgb = getattr(color, "rgb", None)
    if not rgb:
        return None
    rgb = str(rgb)
    if len(rgb) == 8:
        rgb = rgb[2:]
    if len(rgb) != 6:
        return None
    return f"#{rgb.lower()}"


def border_side_css(side) -> str | None:
    if side is None or side.style is None:
        return None
    width_map = {
        "hair": "1px",
        "thin": "1px",
        "medium": "2px",
        "thick": "3px",
        "dotted": "1px",
        "dashDot": "1px",
        "dashDotDot": "1px",
        "dashed": "1px",
        "double": "3px",
    }
    line_map = {
        "dotted": "dotted",
        "dashDot": "dashed",
        "dashDotDot": "dashed",
        "dashed": "dashed",
        "double": "double",
    }
    width = width_map.get(side.style, "1px")
    line_style = line_map.get(side.style, "solid")
    color = color_to_css_hex(side.color) or "#6f8797"
    return f"{width} {line_style} {color}"


def cell_css(cell) -> str:
    rules: list[str] = []
    alignment = cell.alignment
    if alignment:
        if alignment.horizontal:
            rules.append(f"text-align:{alignment.horizontal}")
        if alignment.vertical:
            rules.append(f"vertical-align:{alignment.vertical}")
        if alignment.wrap_text:
            rules.append("white-space:pre-wrap")

    font = cell.font
    if font:
        if font.bold:
            rules.append("font-weight:700")
        if font.italic:
            rules.append("font-style:italic")
        if font.underline and font.underline != "none":
            rules.append("text-decoration:underline")
        if font.size:
            rules.append(f"font-size:{font.size:.1f}pt")
        if font.name:
            rules.append(f"font-family:'{font.name}', sans-serif")
        font_color = color_to_css_hex(font.color)
        if font_color:
            rules.append(f"color:{font_color}")

    fill = cell.fill
    if fill and fill.fill_type == "solid":
        fill_color = color_to_css_hex(fill.fgColor) or color_to_css_hex(fill.start_color)
        if fill_color:
            rules.append(f"background-color:{fill_color}")

    border = cell.border
    if border:
        left = border_side_css(border.left)
        right = border_side_css(border.right)
        top = border_side_css(border.top)
        bottom = border_side_css(border.bottom)
        if left:
            rules.append(f"border-left:{left}")
        if right:
            rules.append(f"border-right:{right}")
        if top:
            rules.append(f"border-top:{top}")
        if bottom:
            rules.append(f"border-bottom:{bottom}")

    return ";".join(rules)


def display_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)


def join_unique(items: list[int]) -> list[int]:
    seen: set[int] = set()
    ordered: list[int] = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return ordered


def parse_clamped_int(raw_value: str | None, default: int, min_value: int, max_value: int) -> int:
    try:
        parsed = int(raw_value) if raw_value is not None else default
    except (TypeError, ValueError):
        parsed = default
    return max(min_value, min(max_value, parsed))


def detect_month_column_groups(worksheet) -> list[dict[str, Any]]:
    header_scan_rows = min(4, worksheet.max_row)
    if header_scan_rows < 1:
        return []

    fallback_year = parse_year_context(worksheet, header_scan_rows + 1)
    anchors: list[dict[str, int | None]] = []

    for col_idx in range(1, worksheet.max_column + 1):
        parsed_month: tuple[int | None, int] | None = None
        for row_idx in range(1, header_scan_rows + 1):
            parsed_month = parse_month_header_loose(
                worksheet.cell(row=row_idx, column=col_idx).value,
                fallback_year,
            )
            if parsed_month:
                break
        if not parsed_month:
            continue
        year, month = parsed_month
        anchors.append({"col": col_idx, "year": year, "month": month})

    if not anchors:
        return []

    known_years = [item["year"] for item in anchors if item["year"] is not None]
    if known_years:
        first_year = int(known_years[0])
        for item in anchors:
            if item["year"] is None:
                item["year"] = first_year
    else:
        synthetic_year = 2000
        previous_month = None
        for item in anchors:
            current_month = int(item["month"])
            if previous_month is not None and current_month < previous_month:
                synthetic_year += 1
            item["year"] = synthetic_year
            previous_month = current_month

    diffs = [
        int(anchors[idx + 1]["col"]) - int(anchors[idx]["col"])
        for idx in range(len(anchors) - 1)
        if int(anchors[idx + 1]["col"]) > int(anchors[idx]["col"])
    ]
    default_width = Counter(diffs).most_common(1)[0][0] if diffs else 1
    if default_width < 1:
        default_width = 1

    groups_by_key: dict[str, dict[str, Any]] = {}
    for idx, anchor in enumerate(anchors):
        start_col = int(anchor["col"])
        if idx + 1 < len(anchors):
            end_col = int(anchors[idx + 1]["col"]) - 1
        else:
            end_col = min(worksheet.max_column, start_col + default_width - 1)
        if end_col < start_col:
            end_col = start_col

        year = int(anchor["year"])
        month = int(anchor["month"])
        key = f"{year:04d}-{month:02d}"
        label = f"{calendar.month_abbr[month]} {year}"

        group = groups_by_key.get(key)
        if group is None:
            group = {
                "key": key,
                "label": label,
                "year": year,
                "month": month,
                "cols": [],
                "start_col": start_col,
            }
            groups_by_key[key] = group

        group["cols"].extend(range(start_col, end_col + 1))
        group["start_col"] = min(group["start_col"], start_col)

    groups = sorted(groups_by_key.values(), key=lambda item: item["key"])
    for group in groups:
        group["cols"] = join_unique(
            [col for col in group["cols"] if 1 <= col <= worksheet.max_column]
        )
    return groups


def pick_month_keys_for_mode(
    month_groups: list[dict[str, Any]],
    mode: str,
    n_value: int,
    month_value: int | None,
) -> list[str]:
    if not month_groups:
        return []

    sorted_groups = sorted(month_groups, key=lambda item: item["key"])
    n_value = max(1, min(60, n_value))
    selected = sorted_groups

    if mode == "same_month_years":
        if month_value is None or month_value < 1 or month_value > 12:
            month_value = int(sorted_groups[-1]["month"])
        selected = [group for group in sorted_groups if int(group["month"]) == month_value]

    if not selected:
        return []
    return [item["key"] for item in selected[-n_value:]]


@lru_cache(maxsize=128)
def build_main_sheet_html_cached(
    path_str: str,
    mtime_ns: int,
    sheet_name: str,
    month_keys_csv: str,
    mode: str,
    n_value: int,
    month_value: int,
) -> dict[str, Any]:
    workbook = load_workbook_cached(path_str, mtime_ns)
    worksheet = workbook[sheet_name]

    layout = find_header_layout(worksheet)
    detected_month_groups = detect_month_column_groups(worksheet)
    selected_month_labels: list[str] = []
    available_months = sorted({int(group["month"]) for group in detected_month_groups})
    frozen_count = 0

    if detected_month_groups:
        month_by_key = {item["key"]: item for item in detected_month_groups}
        selected_month_keys = [key for key in month_keys_csv.split(",") if key in month_by_key]

        if not selected_month_keys:
            selected_month_keys = pick_month_keys_for_mode(
                detected_month_groups,
                mode,
                n_value,
                month_value if month_value > 0 else None,
            )

        if not selected_month_keys:
            selected_month_keys = [detected_month_groups[-1]["key"]]

        first_month_col = min(
            int(group["start_col"])
            for group in detected_month_groups
            if group.get("cols")
        )
        frozen_count = max(0, first_month_col - 1)
        selected_columns: list[int] = list(range(1, first_month_col))

        for key in selected_month_keys:
            group = month_by_key.get(key)
            if not group:
                continue
            selected_columns.extend(group["cols"])
            selected_month_labels.append(group["label"])

        selected_columns = join_unique(
            [col for col in selected_columns if 1 <= col <= worksheet.max_column]
        )
        min_required_row = 1
    elif layout:
        # Sheet matches fixed format layout but no month groups were detected.
        selected_columns = list(range(1, layout["packing"] + 1))
        frozen_count = len(selected_columns)
        min_required_row = layout["metrics_row"] + 1
    else:
        used_columns: list[int] = []
        for col_idx in range(1, worksheet.max_column + 1):
            has_value = any(
                worksheet.cell(row=row_idx, column=col_idx).value not in (None, "")
                for row_idx in range(1, worksheet.max_row + 1)
            )
            if has_value:
                used_columns.append(col_idx)

        if not used_columns:
            return {
                "sheet_name": sheet_name,
                "row_count": 0,
                "col_count": 0,
                "frozen_count": 0,
                "selected_month_labels": [],
                "available_months": [],
                "html": '<div class="empty">No data in this sheet.</div>',
            }

        if len(used_columns) > 70:
            first_cols = used_columns[:20]
            last_cols = used_columns[-20:]
            header_cols = [
                col_idx
                for col_idx in used_columns
                if any(
                    worksheet.cell(row=row_idx, column=col_idx).value not in (None, "")
                    for row_idx in range(1, min(worksheet.max_row, 4) + 1)
                )
            ]
            selected_columns = join_unique(first_cols + header_cols + last_cols)
        else:
            selected_columns = used_columns

        min_required_row = 1

    selected_col_set = set(selected_columns)

    last_row = 1
    for row_idx in range(1, worksheet.max_row + 1):
        row_has_value = False
        for col_idx in selected_columns:
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value not in (None, ""):
                row_has_value = True
                break
        if row_has_value:
            last_row = row_idx
    last_row = max(last_row, min_required_row)

    merge_top_left: dict[tuple[int, int], tuple[int, int]] = {}
    merge_skip: set[tuple[int, int]] = set()
    for merged_range in worksheet.merged_cells.ranges:
        if merged_range.max_row > last_row:
            continue
        fully_selected = all(
            col in selected_col_set
            for col in range(merged_range.min_col, merged_range.max_col + 1)
        )
        if not fully_selected:
            continue
        top_left = (merged_range.min_row, merged_range.min_col)
        merge_top_left[top_left] = (
            merged_range.max_row - merged_range.min_row + 1,
            merged_range.max_col - merged_range.min_col + 1,
        )
        for row_idx in range(merged_range.min_row, merged_range.max_row + 1):
            for col_idx in range(merged_range.min_col, merged_range.max_col + 1):
                if (row_idx, col_idx) != top_left:
                    merge_skip.add((row_idx, col_idx))

    style_to_class: dict[int, str] = {}
    style_rules: list[str] = []

    def style_class_for(cell) -> str:
        style_id = cell.style_id
        if style_id in style_to_class:
            return style_to_class[style_id]
        class_name = f"sx{len(style_to_class)}"
        style_to_class[style_id] = class_name
        css = cell_css(cell)
        style_rules.append(f".main-source-table td.{class_name}{{{css}}}")
        return class_name

    colgroup_html: list[str] = []
    for col_idx in selected_columns:
        col_dim = worksheet.column_dimensions.get(get_column_letter(col_idx))
        width = getattr(col_dim, "width", None)
        if width:
            pixel_width = max(44, int(width * 7))
            colgroup_html.append(f'<col style="width:{pixel_width}px">')
        else:
            colgroup_html.append("<col>")

    rows_html: list[str] = []
    for row_idx in range(1, last_row + 1):
        height = worksheet.row_dimensions[row_idx].height
        row_style = f' style="height:{int(height)}px"' if height else ""
        cell_html_parts: list[str] = []
        for col_idx in selected_columns:
            if (row_idx, col_idx) in merge_skip:
                continue
            cell = worksheet.cell(row=row_idx, column=col_idx)
            class_name = style_class_for(cell)
            attrs = [f'class="{class_name}"']
            if (row_idx, col_idx) in merge_top_left:
                rowspan, colspan = merge_top_left[(row_idx, col_idx)]
                if rowspan > 1:
                    attrs.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    attrs.append(f'colspan="{colspan}"')

            text = html.escape(display_value(cell.value)).replace("\n", "<br>")
            cell_html_parts.append(f"<td {' '.join(attrs)}>{text}</td>")

        rows_html.append(f"<tr{row_style}>{''.join(cell_html_parts)}</tr>")

    inline_styles = "".join(style_rules)
    base_style = (
        "<style>"
        ".main-source-table{border-collapse:collapse;width:max-content;min-width:100%;}"
        ".main-source-table td{padding:6px 8px;white-space:nowrap;border:1px solid rgba(133,184,208,0.22);}"
        f"{inline_styles}"
        "</style>"
    )
    table_html = (
        f"{base_style}<table class=\"main-source-table\">"
        f"<colgroup>{''.join(colgroup_html)}</colgroup>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        "</table>"
    )

    return {
        "sheet_name": sheet_name,
        "row_count": last_row,
        "col_count": len(selected_columns),
        "frozen_count": min(frozen_count, len(selected_columns)),
        "selected_month_labels": selected_month_labels,
        "available_months": available_months,
        "html": table_html,
    }


def clear_workbook_caches() -> None:
    parse_workbook_cached.cache_clear()
    load_workbook_cached.cache_clear()
    workbook_sheetnames_cached.cache_clear()
    build_main_sheet_html_cached.cache_clear()


def safe_uploaded_filename(original_name: str) -> str:
    path = Path(original_name)
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", path.stem).strip(" ._")
    if not stem:
        stem = "workbook"
    suffix = path.suffix.lower()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return f"{stem}__{timestamp}{suffix}"


@app.route("/")
def index() -> str:
    return render_template("index.html", static_version=int(time.time()))


@app.route("/api/workbooks")
def api_workbooks():
    try:
        selection = resolve_workbook_pair(None, None)
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc), "workbooks": []}), 404

    return jsonify(
        {
            "workbooks": selection["available"],
            "default_main": selection["default_main"],
            "default_reference": selection["default_reference"],
        }
    )


@app.route("/api/upload-workbooks", methods=["POST"])
def api_upload_workbooks():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No Excel files were uploaded."}), 400

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    saved_files: list[str] = []
    skipped_files: list[str] = []

    for file_storage in files:
        original_name = (file_storage.filename or "").strip()
        if not original_name:
            continue

        suffix = Path(original_name).suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            skipped_files.append(original_name)
            continue

        output_name = safe_uploaded_filename(original_name)
        file_storage.save(UPLOAD_DIR / output_name)
        saved_files.append(output_name)

    if not saved_files:
        return (
            jsonify(
                {
                    "error": "No supported Excel files were uploaded.",
                    "skipped_files": skipped_files,
                }
            ),
            400,
        )

    clear_workbook_caches()
    selection = resolve_workbook_pair(None, None)
    return jsonify(
        {
            "uploaded_files": saved_files,
            "skipped_files": skipped_files,
            "workbooks": selection["available"],
            "default_main": selection["default_main"],
            "default_reference": selection["default_reference"],
        }
    )


@app.route("/api/sheets")
def api_sheets():
    try:
        selection = resolve_workbook_pair(
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = load_viewer_data(selection["main_path"], selection["reference_path"])
    return jsonify(
        {
            "main_workbook": data["main_workbook"],
            "reference_workbook": data["reference_workbook"],
            "version": data["version"],
            "available_workbooks": selection["available"],
            "sheets": data["sheet_index"],
            "main_sheet_tabs": data["main_sheet_tabs"],
            "reference_sheet_tabs": data["reference_sheet_tabs"],
        }
    )


@app.route("/api/sheet/<canonical>")
def api_sheet(canonical: str):
    try:
        selection = resolve_workbook_pair(
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = load_viewer_data(selection["main_path"], selection["reference_path"])
    if canonical not in data["main"] and canonical not in data["reference"]:
        abort(404, description=f"Unknown sheet: {canonical}")

    return jsonify(
        {
            "canonical": canonical,
            "version": data["version"],
            "main": data["main"].get(canonical),
            "reference": data["reference"].get(canonical),
        }
    )


@app.route("/api/main-styled/<canonical>")
def api_main_styled(canonical: str):
    try:
        selection = resolve_workbook_pair(
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = load_viewer_data(selection["main_path"], selection["reference_path"])
    if canonical not in data["main"]:
        abort(404, description=f"Unknown sheet: {canonical}")

    month_keys_csv = request.args.get("month_keys", "")
    mode = request.args.get("mode", "past_months")
    if mode not in {"past_months", "same_month_years"}:
        mode = "past_months"
    n_value = parse_clamped_int(request.args.get("n"), default=6, min_value=1, max_value=60)
    month_value = parse_clamped_int(request.args.get("month"), default=0, min_value=0, max_value=12)
    main_stat = selection["main_path"].stat()
    sheet_name = data["main"][canonical]["sheet_name"]
    html_payload = build_main_sheet_html_cached(
        str(selection["main_path"]),
        main_stat.st_mtime_ns,
        sheet_name,
        month_keys_csv,
        mode,
        n_value,
        month_value,
    )
    return jsonify(
        {
            "version": data["version"],
            "canonical": canonical,
            "filterable": True,
            "sheet_name": html_payload["sheet_name"],
            "row_count": html_payload["row_count"],
            "col_count": html_payload["col_count"],
            "frozen_count": html_payload["frozen_count"],
            "selected_month_labels": html_payload["selected_month_labels"],
            "available_months": html_payload["available_months"],
            "html": html_payload["html"],
        }
    )


@app.route("/api/main-styled-sheet")
def api_main_styled_sheet():
    try:
        selection = resolve_workbook_pair(
            request.args.get("main"),
            request.args.get("reference"),
        )
    except FileNotFoundError as exc:
        return jsonify({"error": str(exc)}), 404
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    data = load_viewer_data(selection["main_path"], selection["reference_path"])
    sheet_name = request.args.get("sheet")
    if not sheet_name:
        return jsonify({"error": "Missing required query parameter: sheet"}), 400

    tab_match = next(
        (tab for tab in data["main_sheet_tabs"] if tab["sheet_name"] == sheet_name),
        None,
    )
    if not tab_match:
        abort(404, description=f"Unknown main sheet: {sheet_name}")

    month_keys_csv = request.args.get("month_keys", "")
    mode = request.args.get("mode", "past_months")
    if mode not in {"past_months", "same_month_years"}:
        mode = "past_months"
    n_value = parse_clamped_int(request.args.get("n"), default=6, min_value=1, max_value=60)
    month_value = parse_clamped_int(request.args.get("month"), default=0, min_value=0, max_value=12)
    main_stat = selection["main_path"].stat()
    html_payload = build_main_sheet_html_cached(
        str(selection["main_path"]),
        main_stat.st_mtime_ns,
        sheet_name,
        month_keys_csv,
        mode,
        n_value,
        month_value,
    )
    return jsonify(
        {
            "version": data["version"],
            "canonical": tab_match["canonical"],
            "filterable": tab_match["filterable"],
            "sheet_name": html_payload["sheet_name"],
            "row_count": html_payload["row_count"],
            "col_count": html_payload["col_count"],
            "frozen_count": html_payload["frozen_count"],
            "selected_month_labels": html_payload["selected_month_labels"],
            "available_months": html_payload["available_months"],
            "html": html_payload["html"],
        }
    )


def pick_available_port(candidates: list[int]) -> int:
    for port in candidates:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("No available port found in candidate list.")


def resolve_runtime_host_port() -> tuple[str, int]:
    render_port = os.getenv("PORT")
    if render_port:
        try:
            return "0.0.0.0", int(render_port)
        except ValueError as exc:
            raise RuntimeError("PORT environment variable must be an integer.") from exc

    return "127.0.0.1", pick_available_port([5055, 8000, 8080, 5000])


if __name__ == "__main__":
    host, port = resolve_runtime_host_port()
    print(f"Starting viewer at http://{host}:{port}")
    app.run(host=host, port=port, debug=False)
