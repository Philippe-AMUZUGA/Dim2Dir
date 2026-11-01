#!/usr/bin/env python3
from __future__ import annotations

import argparse
import math
import shutil
from numbers import Number
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Paramètres à modifier directement dans Spyder.
EXCEL_FILE = "./PLEX2_dim2dir.xlsx"
SHEET_NAME = "Feuil1"

SPATIAL_MAX_RENDERED_CELLS = 90_000


SPATIAL_CELL_PIXELS = 34
SPATIAL_AXIS_MIN_COLUMN_WIDTH = 9
SPATIAL_AXIS_MAX_COLUMN_WIDTH = 30
SPATIAL_AXIS_MIN_ROW_HEIGHT = 30
SPATIAL_AXIS_MAX_ROW_HEIGHT = 90
DEFAULT_BASE_COLOR = "335C81"
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_ROWS = 1_048_576


def _is_run_column(column: Any) -> bool:
    return str(column).strip().casefold() == "run"


def _value_key(value: Any) -> tuple[str, str]:
    if isinstance(value, Number) and not isinstance(value, bool):
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            pass
        else:
            if not math.isnan(numeric_value):
                return "numeric", repr(numeric_value)
    return type(value).__name__, repr(value)


def _is_blank_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return False


def _format_visual_value(value: Any, max_len: int = 36) -> str:
    if _is_blank_value(value):
        text = "(vide)"
    elif isinstance(value, float):
        text = str(int(value)) if math.isclose(value, round(value)) else f"{value:.4g}"
    else:
        text = str(value)
    return text if len(text) <= max_len else f"{text[: max_len - 3]}..."


def _unique_nonblank_values(values: list[Any]) -> list[Any]:
    seen: set[tuple[str, str]] = set()
    unique_values: list[Any] = []
    for value in values:
        if _is_blank_value(value):
            continue
        key = _value_key(value)
        if key in seen:
            continue
        seen.add(key)
        unique_values.append(value)
    return unique_values or ["(vide)"]


def _all_numeric(values: list[Any]) -> bool:
    return all(isinstance(value, Number) and not isinstance(value, bool) for value in values)


def _build_visual_axis(column: str, values: list[Any]) -> dict[str, Any]:
    unique_values = _unique_nonblank_values(values)
    if _all_numeric(unique_values):
        unique_values = sorted(unique_values, key=lambda value: float(value))
    return {
        "column": column,
        "name": str(column),
        "kind": "discrete",
        "levels": unique_values,
        "labels": [_format_visual_value(value) for value in unique_values],
        "index": {_value_key(value): idx for idx, value in enumerate(unique_values)},
    }


def _axis_level_count(axis: dict[str, Any]) -> int:
    return max(1, len(axis["labels"]))


def _axis_value_index(axis: dict[str, Any], value: Any) -> int:
    if _is_blank_value(value):
        return 0
    return int(axis.get("index", {}).get(_value_key(value), 0))


def _axis_strides(axes: list[dict[str, Any]]) -> tuple[list[int], int]:
    strides: list[int] = []
    size = 1
    for axis in axes:
        strides.append(size)
        size *= _axis_level_count(axis)
    return strides, max(1, size)


def _visualization_skip_message(width: int, height: int) -> str | None:
    cell_count = width * height
    if cell_count > SPATIAL_MAX_RENDERED_CELLS:
        return (
            "Visualisation non générée : trop de cellules dans la grille exacte. "
            f"Grille exacte calculée = {width} x {height} ({cell_count:,} cellules). "
            "Aucun niveau n'est fusionné dans ce script standalone."
        )
    return None


def _build_visual_axes(headers: list[str], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    factor_columns = [header for header in headers if not _is_run_column(header)]
    if not factor_columns:
        return []
    return [_build_visual_axis(column, [row[column] for row in rows]) for column in factor_columns]


def _run_label(value: Any, fallback: int) -> str:
    if value is None:
        return f"E{fallback}"
    try:
        number = float(value)
        if math.isnan(number):
            return f"E{fallback}"
        if math.isclose(number, round(number)):
            return f"E{int(round(number))}"
    except (TypeError, ValueError):
        pass
    return f"E{value}"


def project_fractal_2d(headers: list[str], rows: list[dict[str, Any]]) -> dict[str, Any]:
    n_runs = len(rows)
    if n_runs <= 0:
        return {"placements": [], "x_axes": [], "y_axes": [], "x_strides": [], "y_strides": [], "width": 0, "height": 0}

    axes = _build_visual_axes(headers, rows)
    x_axes = axes[0::2]
    y_axes = axes[1::2]
    x_strides, width = _axis_strides(x_axes)
    y_strides, height = _axis_strides(y_axes)
    run_column = next((header for header in headers if _is_run_column(header)), None)

    placements: list[tuple[int, int, str, int]] = []
    for index, row in enumerate(rows, start=1):
        x = sum(_axis_value_index(axis, row[axis["column"]]) * stride for axis, stride in zip(x_axes, x_strides))
        y = sum(_axis_value_index(axis, row[axis["column"]]) * stride for axis, stride in zip(y_axes, y_strides))
        run_value = row[run_column] if run_column else index
        placements.append((height - 1 - y, x, _run_label(run_value, index), index))

    return {
        "placements": placements,
        "x_axes": x_axes,
        "y_axes": y_axes,
        "x_strides": x_strides,
        "y_strides": y_strides,
        "width": width,
        "height": height,
    }


def _contrast_font_color(hex_color: str) -> str:
    red = int(hex_color[0:2], 16)
    green = int(hex_color[2:4], 16)
    blue = int(hex_color[4:6], 16)
    luminance = (0.299 * red) + (0.587 * green) + (0.114 * blue)
    return "000000" if luminance > 150 else "FFFFFF"


def _blend_hex_colors(start_color: str, end_color: str, amount: float) -> str:
    amount = max(0.0, min(1.0, amount))
    start_rgb = [int(start_color[i:i + 2], 16) for i in (0, 2, 4)]
    end_rgb = [int(end_color[i:i + 2], 16) for i in (0, 2, 4)]
    blended = [round(start + (end - start) * amount) for start, end in zip(start_rgb, end_rgb)]
    return "".join(f"{value:02X}" for value in blended)


def _trial_sequence_color(base_color: str, order: int, n_runs: int) -> str:
    ratio = 1.0 if n_runs <= 1 else (order - 1) / (n_runs - 1)
    return _blend_hex_colors("FFFFFF", base_color, 0.22 + (0.78 * ratio))


def _axis_text_width(text_len: int) -> float:
    return min(
        SPATIAL_AXIS_MAX_COLUMN_WIDTH,
        max(SPATIAL_AXIS_MIN_COLUMN_WIDTH, text_len * 0.9 + 2),
    )


def _axis_text_height(text_len: int) -> float:
    return min(
        SPATIAL_AXIS_MAX_ROW_HEIGHT,
        max(SPATIAL_AXIS_MIN_ROW_HEIGHT, text_len * 3.2 + 10),
    )


def _hex_from_fill(cell) -> str:
    color = cell.fill.fgColor
    if color is not None and color.type == "rgb" and color.rgb:
        rgb = color.rgb[-6:]
        if rgb != "000000":
            return rgb
    return DEFAULT_BASE_COLOR


def _detect_table_bounds(ws) -> tuple[int, int]:
    max_col = 0
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value in (None, ""):
            break
        max_col = col
    if max_col == 0:
        raise ValueError("La feuille ne contient pas d'en-têtes en ligne 1.")

    max_row = 1
    for row in range(2, ws.max_row + 1):
        has_value = any(ws.cell(row=row, column=col).value not in (None, "") for col in range(1, max_col + 1))
        if has_value:
            max_row = row
    if max_row <= 1:
        raise ValueError("La feuille ne contient aucun essai sous les en-têtes.")
    return max_row, max_col


def _read_plan_table(ws, max_row: int, max_col: int) -> tuple[list[str], list[dict[str, Any]]]:
    headers = [str(ws.cell(row=1, column=col).value) for col in range(1, max_col + 1)]
    rows: list[dict[str, Any]] = []
    for row_idx in range(2, max_row + 1):
        row = {
            header: ws.cell(row=row_idx, column=col_idx).value
            for col_idx, header in enumerate(headers, start=1)
        }
        if any(value not in (None, "") for value in row.values()):
            rows.append(row)
    return headers, rows


def add_visualization_to_sheet(ws) -> None:
    table_max_row, table_max_column = _detect_table_bounds(ws)
    headers, rows = _read_plan_table(ws, table_max_row, table_max_column)
    layout = project_fractal_2d(headers, rows)
    placements = layout["placements"]
    if not placements:
        raise ValueError("Impossible de générer la visualisation : aucun placement calculé.")

    base_color = _hex_from_fill(ws.cell(row=1, column=1))
    start_row = 2
    start_col = table_max_column + 2
    x_axes = layout["x_axes"]
    y_axes = layout["y_axes"]
    x_strides = layout["x_strides"]
    y_strides = layout["y_strides"]
    width = layout["width"]
    height = layout["height"]
    left_cols = max(1, len(y_axes))
    title_row = start_row
    grid_top = start_row + 1
    data_left = start_col + left_cols
    grid_bottom = grid_top + height - 1
    x_label_start = grid_bottom + 1
    end_row = x_label_start + max(1, len(x_axes)) - 1
    end_col = data_left + width - 1
    skip_message = _visualization_skip_message(width, height)
    if skip_message is None and (end_col > EXCEL_MAX_COLUMNS or end_row > EXCEL_MAX_ROWS):
        skip_message = (
            "La visualisation exacte contient trop de niveaux distincts pour les limites d'Excel. "
            f"Largeur calculée = {width}, hauteur calculée = {height}. "
            "Aucun niveau n'est fusionné dans ce script standalone."
        )
    if skip_message is not None:
        note_cell = ws.cell(row=start_row, column=start_col)
        note_cell.value = skip_message
        note_cell.font = Font(name="Arial", bold=True, color="C00000", size=10)
        note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(start_col)].width = 60
        ws.row_dimensions[start_row].height = 36
        return

    cell_height = SPATIAL_CELL_PIXELS * 0.75
    cell_width = max(2.5, (SPATIAL_CELL_PIXELS - 5) / 7)
    center = Alignment(horizontal="center", vertical="center")
    label_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    y_label_alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    x_level_alignment = Alignment(horizontal="center", vertical="center", textRotation=90)
    label_font = Font(name="Arial", bold=True, color="1F1F1F", size=10)
    x_level_font = Font(name="Arial", bold=True, color="1F1F1F", size=9)
    title_font = Font(name="Arial", bold=True, color="1F1F1F", size=11)
    grid_border = Border(
        left=Side(style="thin", color="1F1F1F"),
        right=Side(style="thin", color="1F1F1F"),
        top=Side(style="thin", color="1F1F1F"),
        bottom=Side(style="thin", color="1F1F1F"),
    )

    ws.sheet_view.showGridLines = False
    ws.row_dimensions[title_row].height = 24
    for row in range(grid_top, grid_bottom + 1):
        ws.row_dimensions[row].height = cell_height
    for axis_idx, axis in enumerate(x_axes):
        row = x_label_start + axis_idx
        max_label_len = max([len(axis["name"])] + [len(str(label)) for label in axis["labels"]])
        ws.row_dimensions[row].height = _axis_text_height(max_label_len)
    for axis_idx, axis in enumerate(y_axes):
        col = start_col + (len(y_axes) - 1 - axis_idx)
        max_label_len = max(len(f"{axis['name']}: {label}") for label in axis["labels"])
        ws.column_dimensions[get_column_letter(col)].width = _axis_text_width(max_label_len)
    if not y_axes:
        max_name_len = max((len(axis["name"]) for axis in x_axes), default=10)
        ws.column_dimensions[get_column_letter(start_col)].width = _axis_text_width(max_name_len)
    for col in range(data_left, end_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = cell_width

    if end_col > start_col:
        ws.merge_cells(start_row=title_row, start_column=start_col, end_row=title_row, end_column=end_col)
    title_cell = ws.cell(row=title_row, column=start_col)
    title_cell.value = f"Visualisation - {ws.title}"
    title_cell.font = title_font
    title_cell.alignment = label_alignment
    title_cell.border = grid_border

    for row in range(grid_top, grid_bottom + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = label_alignment
            cell.border = grid_border
    for row in range(x_label_start, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = label_alignment
            cell.border = grid_border

    def merge_if_needed(row_start: int, col_start: int, row_end: int, col_end: int) -> None:
        if row_start != row_end or col_start != col_end:
            ws.merge_cells(start_row=row_start, start_column=col_start, end_row=row_end, end_column=col_end)

    for axis_idx, axis in enumerate(y_axes):
        col = start_col + (len(y_axes) - 1 - axis_idx)
        stride = y_strides[axis_idx]
        level_count = _axis_level_count(axis)
        period = stride * level_count
        for group_start in range(0, height, period):
            for level_idx, level_label in enumerate(axis["labels"]):
                logical_start = group_start + level_idx * stride
                logical_end = min(logical_start + stride - 1, height - 1)
                row_start = grid_top + (height - 1 - logical_end)
                row_end = grid_top + (height - 1 - logical_start)
                merge_if_needed(row_start, col, row_end, col)
                cell = ws.cell(row=row_start, column=col)
                cell.value = f"{axis['name']}: {level_label}"
                cell.font = label_font
                cell.alignment = y_label_alignment
                cell.border = grid_border

    for axis_idx, axis in enumerate(x_axes):
        row = x_label_start + axis_idx
        name_end_col = data_left - 1
        merge_if_needed(row, start_col, row, name_end_col)
        name_cell = ws.cell(row=row, column=start_col)
        name_cell.value = axis["name"]
        name_cell.font = label_font
        name_cell.alignment = label_alignment
        name_cell.border = grid_border

        stride = x_strides[axis_idx]
        level_count = _axis_level_count(axis)
        period = stride * level_count
        for group_start in range(0, width, period):
            for level_idx, level_label in enumerate(axis["labels"]):
                col_start = data_left + group_start + level_idx * stride
                col_end = min(col_start + stride - 1, end_col)
                merge_if_needed(row, col_start, row, col_end)
                cell = ws.cell(row=row, column=col_start)
                cell.value = level_label
                cell.font = x_level_font
                cell.alignment = x_level_alignment
                cell.border = grid_border

    labels_by_cell: dict[tuple[int, int], list[str]] = {}
    first_order_by_cell: dict[tuple[int, int], int] = {}
    for grid_row, grid_col, label, order in placements:
        key = (grid_row, grid_col)
        labels_by_cell.setdefault(key, []).append(label)
        first_order_by_cell[key] = min(first_order_by_cell.get(key, order), order)

    n_runs = len(rows)
    for (grid_row, grid_col), labels in labels_by_cell.items():
        color = _trial_sequence_color(base_color, first_order_by_cell[(grid_row, grid_col)], n_runs)
        cell = ws.cell(row=grid_top + grid_row, column=data_left + grid_col)
        cell.value = "\n".join(labels)
        cell.fill = PatternFill("solid", start_color=color, end_color=color)
        cell.font = Font(name="Arial", bold=True, color=_contrast_font_color(color), size=9)
        cell.alignment = center
        cell.border = grid_border


def build_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_visu{input_path.suffix}")


def run_visualization(excel_file: str = EXCEL_FILE, sheet_name: str = SHEET_NAME) -> Path:
    input_path = Path(excel_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {input_path}")

    output_path = build_output_path(input_path)
    shutil.copy2(input_path, output_path)

    wb = load_workbook(output_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Feuille introuvable : {sheet_name}. Feuilles disponibles : {', '.join(wb.sheetnames)}")

    add_visualization_to_sheet(wb[sheet_name])
    wb.save(output_path)
    print(f"Fichier généré : {output_path}")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Ajoute une visualisation PLEX2 à une feuille Excel de plan d'expériences.")
    parser.add_argument("excel_file", nargs="?", default=EXCEL_FILE, help="Chemin du fichier Excel source")
    parser.add_argument("sheet_name", nargs="?", default=SHEET_NAME, help="Nom de la feuille contenant la matrice d'essais")
    args = parser.parse_args()
    run_visualization(args.excel_file, args.sheet_name)


if __name__ == "__main__":
    run_visualization()
