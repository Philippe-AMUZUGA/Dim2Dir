from __future__ import annotations

import argparse
import contextlib
import math
import shutil
import json
import os
import re
import sys
import tempfile
import time
import threading
import warnings as python_warnings
import string
from io import BytesIO
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, fields
from numbers import Integral, Number
from pathlib import Path
from typing import Any
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


__version__ = "2.1.0"

# Paramètres à modifier directement dans Spyder.
EXCEL_FILE = "./Dim2Dir.xlsx"
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

class Dim2DirError(ValueError):
    """Entrée ou option invalide pour Dim2Dir."""


class _ImageLimitError(Dim2DirError):
    """Dépassement de pixels : permet un résultat skipped avec nettoyage."""


def _hex_color(value):
    if not isinstance(value, str) or not re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        raise Dim2DirError(f"Couleur invalide : {value!r} ; utiliser RRGGBB ou #RRGGBB.")
    return value.lstrip("#").upper()


@dataclass(frozen=True)
class _LevelOrders(Mapping):
    entries: tuple

    def __iter__(self):
        return (name for name, _ in self.entries)

    def __len__(self):
        return len(self.entries)

    def __getitem__(self, name):
        for key, value in self.entries:
            if key == name:
                return value
        raise KeyError(name)


@dataclass(frozen=True)
class Style:
    """Personnalisation commune au classeur et à ses images ; les défauts reproduisent le rendu historique."""

    palette: str | tuple[str, ...] | None = None
    base_color: str | None = None
    palette_reverse: bool = False
    font_family: str | None = None
    font_size: float | None = None
    font_color: str | None = None
    label_scale: float = 1.0
    cell_font_color: str | None = None
    cell_size: int = 34
    line_color: str = "1F1F1F"
    line_style: str = "thin"
    title: str | None = None
    legend: bool = False
    label_format: str = "{name}: {level}"
    run_prefix: str = "E"
    level_orders: Mapping[str, Sequence] | None = None

    def __post_init__(self):
        for name in ("base_color", "font_color", "cell_font_color", "line_color"):
            value = getattr(self, name)
            if value is not None:
                object.__setattr__(self, name, _hex_color(value))
            elif name == "line_color":
                raise Dim2DirError("line_color doit être une couleur RRGGBB.")
        for name in ("palette_reverse", "legend"):
            if not isinstance(getattr(self, name), bool):
                raise Dim2DirError(f"{name} doit être un booléen.")
        if self.font_family is not None and (not isinstance(self.font_family, str) or not self.font_family.strip()):
            raise Dim2DirError("font_family doit être un nom de police non vide.")
        if self.font_size is not None and (isinstance(self.font_size, bool)
                or not isinstance(self.font_size, (int, float)) or not 4 <= self.font_size <= 72):
            raise Dim2DirError("font_size doit être compris entre 4 et 72.")
        if (isinstance(self.label_scale, bool) or not isinstance(self.label_scale, (int, float))
                or not 0.5 <= self.label_scale <= 3.0):
            raise Dim2DirError("label_scale doit être compris entre 0.5 et 3.0.")
        if (isinstance(self.cell_size, bool) or not isinstance(self.cell_size, Integral)
                or not 12 <= self.cell_size <= 120):
            raise Dim2DirError("cell_size doit être un entier entre 12 et 120.")
        object.__setattr__(self, "cell_size", int(self.cell_size))
        if self.line_style not in ("none", "thin", "medium", "thick"):
            raise Dim2DirError("line_style : choisir none, thin, medium ou thick.")
        if self.title is not None and not isinstance(self.title, str):
            raise Dim2DirError("title doit être un texte ou None.")
        if not isinstance(self.run_prefix, str) or not isinstance(self.label_format, str):
            raise Dim2DirError("run_prefix et label_format doivent être des textes.")
        try:
            for _, field, spec, conversion in string.Formatter().parse(self.label_format):
                if field is not None and (field not in ("name", "level") or spec or conversion):
                    raise ValueError("champs permis : name, level, sans conversion ni format")
            self.label_format.format(name="A", level="1")
        except (ValueError, KeyError, IndexError) as exc:
            raise Dim2DirError(f"label_format invalide : {exc}") from exc
        if self.palette is not None:
            if isinstance(self.palette, str):
                if self.palette not in _matplotlib().colormaps:
                    raise Dim2DirError(f"Palette inconnue : {self.palette!r}. Exemples : viridis, cividis, Greys.")
            elif isinstance(self.palette, Sequence) and len(self.palette) >= 2:
                object.__setattr__(self, "palette", tuple(_hex_color(c) for c in self.palette))
            else:
                raise Dim2DirError("palette doit être un nom de colormap ou au moins deux couleurs.")
        if self.level_orders is not None:
            if not isinstance(self.level_orders, Mapping):
                raise Dim2DirError("level_orders doit associer chaque facteur à une liste de niveaux.")
            entries = []
            for name, levels in self.level_orders.items():
                if not isinstance(name, str) or not isinstance(levels, Sequence) or isinstance(levels, (str, bytes)):
                    raise Dim2DirError("level_orders : nom de facteur et séquence de niveaux requis.")
                if any(isinstance(v, (Mapping, list, set)) for v in levels):
                    raise Dim2DirError("level_orders : les niveaux doivent être des valeurs scalaires.")
                if len({_value_key(v) for v in levels}) != len(levels):
                    raise Dim2DirError(f"level_orders : niveaux en double pour {name}.")
                entries.append((name, tuple(levels)))
            object.__setattr__(self, "level_orders", _LevelOrders(tuple(entries)))

    @classmethod
    def from_dict(cls, mapping):
        if not isinstance(mapping, Mapping):
            raise Dim2DirError("Le style doit être un dictionnaire.")
        unknown = set(mapping) - {field.name for field in fields(cls)}
        if unknown:
            raise Dim2DirError(f"Clés de Style inconnues : {', '.join(map(str, unknown))}.")
        return cls(**mapping)

    def to_dict(self):
        result = {field.name: getattr(self, field.name) for field in fields(self)}
        if self.level_orders is not None:
            result["level_orders"] = {name: list(levels) for name, levels in self.level_orders.items()}
        if isinstance(self.palette, tuple):
            result["palette"] = list(self.palette)
        return result

    @classmethod
    def from_json(cls, path):
        try:
            return cls.from_dict(json.loads(Path(path).read_text(encoding="utf-8-sig")))
        except (OSError, ValueError, TypeError) as exc:
            raise Dim2DirError(f"Style JSON invalide : {exc}") from exc


@dataclass(frozen=True)
class Limits:
    max_cells: int = 90_000
    max_pixels: int = 60_000_000
    max_runs: int = 20_000
    excel_timeout_s: float = 120.0

    def __post_init__(self):
        for name in ("max_cells", "max_pixels", "max_runs"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, Integral) or value <= 0:
                raise Dim2DirError(f"{name} doit être un entier strictement positif.")
            object.__setattr__(self, name, int(value))
        if self.max_cells > 1_048_576 * 16_384:
            raise Dim2DirError(f"max_cells = {self.max_cells} dépasse la borne dure Excel {1_048_576 * 16_384}.")
        if (isinstance(self.excel_timeout_s, bool) or not isinstance(self.excel_timeout_s, (int, float))
                or not math.isfinite(self.excel_timeout_s) or self.excel_timeout_s <= 0):
            raise Dim2DirError("excel_timeout_s doit être un nombre fini strictement positif.")


def _options(style=None, limits=None):
    if style is not None and not isinstance(style, Style):
        raise Dim2DirError("style doit être un Style.")
    if limits is not None and not isinstance(limits, Limits):
        raise Dim2DirError("limits doit être un Limits.")
    return style or Style(), limits or Limits()


def _check_runs(count, limits):
    if count > limits.max_runs:
        raise Dim2DirError(f"{count} essais dépassent la limite {limits.max_runs} ; relever limits.max_runs / --max-runs.")


def _style_color(style, base_color, order, n_runs):
    if style.palette is None:
        if style.palette_reverse:
            order = n_runs + 1 - order
        return _trial_sequence_color(base_color, order, n_runs)
    ratio = 0.5 if n_runs <= 1 else (order - 1) / (n_runs - 1)
    if style.palette_reverse:
        ratio = 1 - ratio
    if isinstance(style.palette, str):
        return "".join(f"{round(v * 255):02X}" for v in _matplotlib().colormaps[style.palette](ratio)[:3])
    position = ratio * (len(style.palette) - 1)
    index = min(int(position), len(style.palette) - 2)
    return _blend_hex_colors(style.palette[index], style.palette[index + 1], position - index)


def _apply_layout_style(layout, rows, run_column, style, warnings):
    axes = {axis["name"]: axis for axis in (*layout["x_axes"], *layout["y_axes"])}
    for name, requested in (style.level_orders or {}).items():
        if name not in axes:
            raise Dim2DirError(f"level_orders : facteur inconnu {name!r}.")
        axis = axes[name]
        observed = {_value_key(v): v for v in axis["levels"]}
        ordered = [observed[_value_key(v)] for v in requested if _value_key(v) in observed]
        if len(ordered) < len(requested):
            warnings.append(f"level_orders ({name}) : niveaux absents du plan ignorés.")
        selected = {_value_key(v) for v in ordered}
        missing = [v for v in axis["levels"] if _value_key(v) not in selected]
        if missing:
            warnings.append(f"level_orders ({name}) : niveaux du plan absents de la liste ajoutés à la fin.")
        axis["levels"] = ordered + missing
        axis["labels"], axis["notes"] = _axis_labels(axis["name"], axis["levels"])
        axis["index"] = {_value_key(v): i for i, v in enumerate(axis["levels"])}
    if style.level_orders:
        run_header = next((key for key in rows[0] if key.strip().casefold() == run_column.casefold()), None)
        layout["placements"] = []
        for order, row in enumerate(rows, 1):
            x, y = [sum(_axis_value_index(axis, row[axis["column"]]) * stride
                        for axis, stride in zip(layout[f"{direction}_axes"], layout[f"{direction}_strides"]))
                    for direction in ("x", "y")]
            layout["placements"].append((layout["height"] - 1 - y, x,
                                         _run_label(row[run_header] if run_header else order, order), order))
    if style.run_prefix != "E":
        layout["placements"] = [(r, c, style.run_prefix + label[1:], order)
                                for r, c, label, order in layout["placements"]]
    for axis in (*layout["x_axes"], *layout["y_axes"]):
        for note in axis.get("notes", ()):
            if note not in warnings:
                warnings.append(note)
    if style.palette is not None and style.base_color is not None:
        warnings.append("base_color ignoré car palette est donnée.")


_EXCEL_LOCK = threading.Lock()



# --- Lecture et validation : utilitaires compatibles avec le rendu historique ---

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
        if math.isinf(value):
            text = "inf" if value > 0 else "-inf"
        else:
            text = str(int(value)) if math.isclose(value, round(value)) else f"{value:.4g}"
    else:
        text = str(value)
    return text if len(text) <= max_len else f"{text[: max_len - 3]}..."


def _precise_label(value: Any, digits: int) -> str:
    """Libellé d'un niveau avec `digits` chiffres significatifs (valeurs finies non entières)."""
    if isinstance(value, float) and math.isfinite(value) and value != int(value):
        return repr(value) if digits >= 17 else f"{value:.{digits}g}"
    if isinstance(value, float) and math.isfinite(value):
        return str(int(value))
    return _format_visual_value(value, max_len=10 ** 9)


def _axis_labels(name: str, levels: list[Any]) -> tuple[list[str], list[str]]:
    """Libellés des niveaux d'un facteur ; ceux d'un facteur sans collision sont inchangés.

    Retourne aussi les avertissements : précision augmentée (jusqu'à 17 chiffres significatifs) ou
    suffixe numérique pour les libellés tronqués identiques."""
    labels = [_format_visual_value(value) for value in levels]
    if len(set(labels)) == len(labels):
        return labels, []
    notes = []
    floats = [value for value in levels if isinstance(value, float)]
    full = [_format_visual_value(value, max_len=10 ** 9) for value in levels]
    if len({_format_visual_value(value) for value in floats}) < len(floats):
        for digits in range(5, 18):
            full = [_precise_label(value, digits) for value in levels]
            if len(set(full)) == len(full) or digits == 17 or len({_precise_label(v, digits) for v in floats}) == len(floats):
                break
        notes.append(f"Facteur {name!r} : des niveaux distincts avaient le même libellé ; "
                     f"précision portée à {digits} chiffres significatifs.")
    labels = [text if len(text) <= 36 else f"{text[:33]}..." for text in full]
    if len(set(labels)) != len(labels):
        seen: dict[str, int] = {}
        counts = {text: labels.count(text) for text in labels}
        result = []
        for text, label in zip(full, labels):
            if counts[label] == 1:
                result.append(label)
                continue
            seen[label] = seen.get(label, 0) + 1
            suffix = f" [{seen[label]}]"
            keep = text if len(text) + len(suffix) <= 36 else f"{text[:36 - 3 - len(suffix)]}..."
            result.append(keep + suffix)
        labels = result
        notes.append(f"Facteur {name!r} : libellés identiques (texte tronqué à 36 caractères) ; "
                     "suffixe numérique ajouté aux niveaux concernés.")
    return labels, notes


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
    labels, notes = _axis_labels(str(column), unique_values)
    return {
        "column": column,
        "name": str(column),
        "kind": "discrete",
        "levels": unique_values,
        "labels": labels,
        "notes": notes,
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


def _build_visual_axes(headers: list[str], rows: list[dict[str, Any]], run_column="Run") -> list[dict[str, Any]]:
    factor_columns = [header for header in headers if header.strip().casefold() != run_column.casefold()]
    if not factor_columns:
        return []
    return [_build_visual_axis(column, [row[column] for row in rows]) for column in factor_columns]


def _run_label(value: Any, fallback: int) -> str:
    if value is None:
        return f"E{fallback}"
    try:
        number = float(value)
        if not math.isfinite(number):
            if math.isnan(number) and not isinstance(value, str):
                return f"E{fallback}"
            return f"E{value}"
        if math.isclose(number, round(number)):
            return f"E{int(round(number))}"
    except (TypeError, ValueError, OverflowError):
        pass
    return f"E{value}"


# --- Projection fractale (algorithme de la version 1 inchangé) ---

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


def _axis_text_width(text_len: int, scale: float = 1.0) -> float:
    return min(
        SPATIAL_AXIS_MAX_COLUMN_WIDTH * scale,
        max(SPATIAL_AXIS_MIN_COLUMN_WIDTH, (text_len * 0.9 + 2) * scale),
    )


def _axis_text_height(text_len: int, scale: float = 1.0) -> float:
    return min(
        SPATIAL_AXIS_MAX_ROW_HEIGHT * scale,
        max(SPATIAL_AXIS_MIN_ROW_HEIGHT, (text_len * 3.2 + 10) * scale),
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
    cached = getattr(ws, "_dim2dir_table", None)
    if cached is not None:
        return cached
    return read_plan_table(ws, ws, getattr(ws, "_dim2dir_warnings", []),
                           limits=getattr(ws, "_dim2dir_limits", None))[0:2]


# --- Écriture Excel : bloc de la version 1 conservé ---

def _set_text(cell, value) -> None:
    """Écrit une valeur ; toute chaîne reste du texte (jamais une formule ni un code d'erreur)."""
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def first_zone_conflict(ws, left, top, right, bottom):
    """Coordonnée de la première cellule non vide de la zone, ou None si elle est libre."""
    found = [(row, col) for (row, col), cell in ws._cells.items()
             if left <= col <= right and top <= row <= bottom and cell.value not in (None, "")]
    if not found:
        return None
    row, col = min(found)
    return f"{get_column_letter(col)}{row}"


def zone_conflict_message(ws, bounds) -> str | None:
    left, top, right, bottom = bounds
    coordinate = first_zone_conflict(ws, left, top, right, bottom)
    if coordinate is None:
        return None
    return (f"Zone de visualisation non libre : la cellule {coordinate} de la feuille {ws.title!r} "
            "n'est pas vide. Aucune cellule n'est écrasée ; vider cette zone (colonnes à droite du "
            "tableau) ou utiliser une copie du plan sans ces données.")


def add_visualization_to_sheet(ws, style=None, limits=None) -> None:
    style, limits = _options(style or getattr(ws, "_dim2dir_style", None),
                             limits or getattr(ws, "_dim2dir_limits", None))
    ws._dim2dir_style = style
    ws._dim2dir_limits = limits
    if limits != Limits() or not hasattr(ws, "_dim2dir_max_cells"):
        ws._dim2dir_max_cells = limits.max_cells
    table_max_row, table_max_column = _detect_table_bounds(ws)
    if getattr(ws, "_dim2dir_table", None) is not None:
        table_max_column = len(ws._dim2dir_table[0])
    headers, rows = _read_plan_table(ws, table_max_row, table_max_column)
    _check_runs(len(rows), limits)
    layout = getattr(ws, "_dim2dir_layout", None) or project_fractal_2d(headers, rows)
    if not layout["placements"]:
        raise Dim2DirError("Impossible de générer la visualisation : aucun placement calculé.")
    if not getattr(ws, "_dim2dir_layout_styled", False):
        _apply_layout_style(layout, rows, "Run", style, getattr(ws, "_dim2dir_warnings", []))
        ws._dim2dir_layout_styled = True
    placements = layout["placements"]
    if not placements:
        raise ValueError("Impossible de générer la visualisation : aucun placement calculé.")

    base_color = style.base_color or getattr(ws, "_dim2dir_color", None) or _hex_from_fill(ws.cell(row=1, column=1))
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
    axes_end_row = end_row
    end_row += int(style.legend)
    skip_message = _sheet_skip_message(ws, width, height)
    if skip_message is None and (end_col > EXCEL_MAX_COLUMNS or end_row > EXCEL_MAX_ROWS):
        skip_message = (
            "La visualisation exacte contient trop de niveaux distincts pour les limites d'Excel. "
            f"Largeur calculée = {width}, hauteur calculée = {height}. "
            "Aucun niveau n'est fusionné dans ce script standalone."
        )
    zone = (start_col, start_row, start_col, start_row) if skip_message is not None else (
        start_col, start_row, end_col, end_row)
    conflict = zone_conflict_message(ws, zone)
    if conflict is not None:
        raise Dim2DirError(conflict)
    if skip_message is not None:
        note_cell = ws.cell(row=start_row, column=start_col)
        note_cell.value = skip_message
        note_cell.font = Font(name="Arial", bold=True, color="C00000", size=10)
        note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(start_col)].width = 60
        ws.row_dimensions[start_row].height = 36
        return

    cell_height = style.cell_size * 0.75
    cell_width = (style.cell_size - 5) / 7
    family = style.font_family or "Arial"
    label_size = style.font_size if style.font_size is not None else 10
    run_size = max(4, style.font_size - 1) if style.font_size is not None else 9
    title_size = style.font_size + 1 if style.font_size is not None else 11
    axis_scale = max(1.0, style.font_size / 10) if style.font_size is not None else 1.0  # largeurs d'axes proportionnelles à la police (version 1 : 1)
    axis_scale = axis_scale * style.label_scale
    font_color = style.font_color or "1F1F1F"
    center = Alignment(horizontal="center", vertical="center")
    label_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    y_label_alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    x_level_alignment = Alignment(horizontal="center", vertical="center", textRotation=90)
    label_font = Font(name=family, bold=True, color=font_color, size=label_size)
    x_level_font = Font(name=family, bold=True, color=font_color, size=label_size if style.font_size is not None else 9)
    title_font = Font(name=family, bold=True, color=font_color, size=title_size)
    grid_border = Border(
        left=Side(style=None if style.line_style == "none" else style.line_style, color=style.line_color),
        right=Side(style=None if style.line_style == "none" else style.line_style, color=style.line_color),
        top=Side(style=None if style.line_style == "none" else style.line_style, color=style.line_color),
        bottom=Side(style=None if style.line_style == "none" else style.line_style, color=style.line_color),
    )

    ws.sheet_view.showGridLines = False
    ws.row_dimensions[title_row].height = 24
    for row in range(grid_top, grid_bottom + 1):
        ws.row_dimensions[row].height = cell_height
    for axis_idx, axis in enumerate(x_axes):
        row = x_label_start + axis_idx
        max_label_len = max([len(axis["name"])] + [len(str(label)) for label in axis["labels"]])
        ws.row_dimensions[row].height = _axis_text_height(max_label_len, axis_scale)
    for axis_idx, axis in enumerate(y_axes):
        col = start_col + (len(y_axes) - 1 - axis_idx)
        max_label_len = max(len(style.label_format.format(name=axis["name"], level=label)) for label in axis["labels"])
        ws.column_dimensions[get_column_letter(col)].width = _axis_text_width(max_label_len, axis_scale)
    if not y_axes:
        max_name_len = max((len(axis["name"]) for axis in x_axes), default=10)
        ws.column_dimensions[get_column_letter(start_col)].width = _axis_text_width(max_name_len, axis_scale)
    for col in range(data_left, end_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = cell_width

    if end_col > start_col:
        ws.merge_cells(start_row=title_row, start_column=start_col, end_row=title_row, end_column=end_col)
    title_cell = ws.cell(row=title_row, column=start_col)
    _set_text(title_cell, f"Visualisation - {ws.title}" if style.title is None else style.title)
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
                _set_text(cell, style.label_format.format(name=axis["name"], level=level_label))
                cell.font = label_font
                cell.alignment = y_label_alignment
                cell.border = grid_border

    for axis_idx, axis in enumerate(x_axes):
        row = x_label_start + axis_idx
        name_end_col = data_left - 1
        merge_if_needed(row, start_col, row, name_end_col)
        name_cell = ws.cell(row=row, column=start_col)
        _set_text(name_cell, axis["name"])
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
                _set_text(cell, level_label)
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
        color = _style_color(style, base_color, first_order_by_cell[(grid_row, grid_col)], n_runs)
        cell = ws.cell(row=grid_top + grid_row, column=data_left + grid_col)
        _set_text(cell, "\n".join(labels))
        cell.fill = PatternFill("solid", start_color=color, end_color=color)
        cell.font = Font(name=family, bold=True, color=style.cell_font_color or _contrast_font_color(color), size=run_size)
        cell.alignment = center
        cell.border = grid_border

    if style.legend:
        _write_legend(ws, axes_end_row + 1, start_col, end_col, base_color, n_runs, style,
                      label_font, grid_border, merge_if_needed)


LEGEND_TEXT = "Couleur : ordre des essais, du premier au dernier"
LEGEND_MIN_TEXT_POINTS = 170  # largeur utile pour garder le texte à 6 points ou plus


def _legend_column_points(ws, col):
    width = ws.column_dimensions[get_column_letter(col)].width or 8.43
    return _column_pixels(width) * 0.75


def _write_legend(ws, row, start_col, end_col, base_color, n_runs, style, label_font, border, merge_if_needed):
    """Une ligne : texte fusionné et, si la largeur le permet, un ou deux échantillons de couleur."""
    points = [_legend_column_points(ws, col) for col in range(start_col, end_col + 1)]
    if len(points) >= 3 and sum(points[1:-1]) >= LEGEND_MIN_TEXT_POINTS:
        samples, text_col, text_end = ((start_col, 1), (end_col, n_runs)), start_col + 1, end_col - 1
    elif len(points) >= 2 and sum(points[1:]) >= LEGEND_MIN_TEXT_POINTS:
        samples, text_col, text_end = ((start_col, 1),), start_col + 1, end_col
    else:
        samples, text_col, text_end = (), start_col, end_col
    for col, order in samples:
        cell = ws.cell(row, col)
        cell.fill = PatternFill("solid", fgColor=_style_color(style, base_color, order, n_runs))
        cell.border = border
    text_points = sum(points[text_col - start_col:text_end - start_col + 1])
    ws.row_dimensions[row].height = 30
    if text_points >= LEGEND_MIN_TEXT_POINTS:
        alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    else:
        # Zone étroite : le texte passe à la ligne à sa taille normale ; la hauteur s'adapte.
        import textwrap
        size = label_font.sz or 10
        lines = len(textwrap.wrap(LEGEND_TEXT, max(1, int((text_points - 4) / (size * 0.55))))) or 1
        ws.row_dimensions[row].height = max(30, lines * size * 1.3 + 6)
        alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    merge_if_needed(row, text_col, row, text_end)
    cell = ws.cell(row, text_col)
    _set_text(cell, LEGEND_TEXT)
    cell.font = label_font
    cell.alignment = alignment


def build_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_visu{input_path.suffix}")



# --- Lecture et validation : valeurs en cache et diagnostics ---

def read_plan_table(ws, values_ws, warnings: list[str], *, run_column="Run", limits=None):
    """Lit les valeurs calculées sans toucher aux cellules source."""
    limits = limits or Limits()
    def value(row, col):
        original = ws.cell(row, col)
        result = values_ws.cell(row, col).value
        if original.data_type == "f" and (values_ws is ws or result is None):
            warnings.append(f"Formule sans valeur en cache : {original.coordinate}, traitée comme vide.")
            return None
        return result

    raw_headers = [value(1, col) for col in range(1, ws.max_column + 1)]
    max_col = next((i for i, item in enumerate(raw_headers) if item in (None, "")), len(raw_headers))
    if max_col < len(raw_headers) and any(item not in (None, "") for item in raw_headers[max_col + 1:]):
        warnings.append(f"En-tête vide en colonne {get_column_letter(max_col + 1)} : tableau tronqué.")
    if not max_col:
        raise ValueError("La feuille ne contient pas d'en-têtes en ligne 1.")
    headers, used = [], set()
    reserved = {str(item) for item in raw_headers[:max_col]}
    for item in raw_headers[:max_col]:
        name, suffix = str(item), 2
        candidate = name
        while candidate in used or (candidate != name and candidate in reserved):
            candidate = f"{name}_{suffix}"
            suffix += 1
        if candidate != name:
            warnings.append(f"En-tête en double {name!r} : renommé {candidate!r} pour la lecture.")
        headers.append(candidate)
        used.add(candidate)
    rows = []
    blanks = {header: 0 for header in headers if header.strip().casefold() != run_column.casefold()}
    for row_idx in range(2, ws.max_row + 1):
        row = {header: value(row_idx, col) for col, header in enumerate(headers, 1)}
        if not any(item not in (None, "") for item in row.values()):
            continue
        _check_runs(len(rows) + 1, limits)
        rows.append(row)
        for header in blanks:
            if _is_blank_value(row[header]):
                blanks[header] += 1
    if not rows:
        raise ValueError("La feuille ne contient aucun essai sous les en-têtes.")
    for header, count in blanks.items():
        if count:
            warnings.append(f"Facteur {header!r} : {count} cellule(s) vide(s), niveau d'indice 0.")
    return headers, rows, max_col


def project_with_assignment(headers, rows, x_factors=None, y_factors=None):
    if x_factors is None and y_factors is None:
        return project_fractal_2d(headers, rows)
    if x_factors is None or y_factors is None:
        raise ValueError("Fournir ensemble --x-factors et --y-factors (liste vide autorisée).")
    factors = [header for header in headers if not _is_run_column(header)]
    combined = x_factors + y_factors
    if len(combined) != len(set(combined)) or set(combined) != set(factors):
        raise ValueError("Affectation x/y : partition exacte des facteurs requise, sans doublon. "
                         f"Facteurs attendus : {', '.join(factors)} ; reçus : {', '.join(combined)}.")
    layout = project_fractal_2d(headers, rows)
    axes = {axis["column"]: axis for axis in _build_visual_axes(headers, rows)}
    for name, names in (("x", x_factors), ("y", y_factors)):
        layout[f"{name}_axes"] = [axes[item] for item in names]
        layout[f"{name}_strides"], layout["width" if name == "x" else "height"] = _axis_strides(layout[f"{name}_axes"])
    run_column = next((header for header in headers if _is_run_column(header)), None)
    layout["placements"] = []
    for order, row in enumerate(rows, 1):
        coords = [sum(_axis_value_index(axis, row[axis["column"]]) * stride
                      for axis, stride in zip(layout[f"{name}_axes"], layout[f"{name}_strides"]))
                  for name in ("x", "y")]
        layout["placements"].append((layout["height"] - 1 - coords[1], coords[0],
                                     _run_label(row[run_column] if run_column else order, order), order))
    return layout


def _sheet_skip_message(ws, width, height):
    limit = getattr(ws, "_dim2dir_max_cells", SPATIAL_MAX_RENDERED_CELLS)
    if limit == SPATIAL_MAX_RENDERED_CELLS:
        message = _visualization_skip_message(width, height)
        if message is not None:
            return message
    elif width * height > limit:
        return ("Visualisation non générée : trop de cellules dans la grille exacte. "
                f"Grille exacte calculée = {width} x {height} ({width * height:,} cellules). "
                f"Limite demandée : {limit} ; relever limits.max_cells / --max-cells. Aucun niveau n'est fusionné dans ce script standalone.")
    # Les limites d'Excel s'appliquent quelle que soit la valeur de --max-cells.
    if width > EXCEL_MAX_COLUMNS or height > EXCEL_MAX_ROWS:
        return ("La visualisation exacte contient trop de niveaux distincts pour les limites d'Excel. "
                f"Largeur calculée = {width}, hauteur calculée = {height}. "
                f"Limites fixes : {EXCEL_MAX_COLUMNS} colonnes, {EXCEL_MAX_ROWS} lignes ; aucun paramètre ne les relève. "
                "Aucun niveau n'est fusionné dans ce script standalone.")
    return None


# --- Géométrie de la scène : uniquement la zone écrite par le bloc de la version 1 ---

def visualization_bounds(table_max_column, layout, style=None):
    start_col = table_max_column + 2
    end_col = start_col + max(1, len(layout["y_axes"])) + layout["width"] - 1
    end_row = 2 + layout["height"] + max(1, len(layout["x_axes"]))
    return start_col, 2, end_col, end_row + int(bool(style and style.legend))


def _column_pixels(width):
    """Pixels d'une colonne de `width` caractères ; la tolérance évite 119 pour 120 (erreur flottante)."""
    return math.floor(width * 7 + 5 + 1e-6)


def scene_geometry(ws, bounds):
    left, top, right, bottom = bounds
    # Conversion Excel (police normale Calibri 11, 96 dpi), en points.
    widths = [_column_pixels(ws.column_dimensions[get_column_letter(col)].width) * 0.75
              for col in range(left, right + 1)]
    heights = [ws.row_dimensions[row].height or ws.sheet_format.defaultRowHeight
               for row in range(top, bottom + 1)]
    return widths, heights


def _scene_rectangles(ws, bounds):
    left, top, right, bottom = bounds
    widths, heights = scene_geometry(ws, bounds)
    xs, ys = [0.0], [0.0]
    for size in widths:
        xs.append(xs[-1] + size)
    for size in heights:
        ys.append(ys[-1] + size)
    merged, covered = {}, set()
    for area in ws.merged_cells.ranges:
        if left <= area.min_col <= area.max_col <= right and top <= area.min_row <= area.max_row <= bottom:
            merged[area.min_row, area.min_col] = (area.max_row, area.max_col)
            covered.update((r, c) for r in range(area.min_row, area.max_row + 1)
                           for c in range(area.min_col, area.max_col + 1)
                           if (r, c) != (area.min_row, area.min_col))
    for row in range(top, bottom + 1):
        for col in range(left, right + 1):
            if (row, col) in covered:
                continue
            end_row, end_col = merged.get((row, col), (row, col))
            yield (ws.cell(row, col), xs[col-left], ys[row-top],
                   xs[end_col-left+1] - xs[col-left], ys[end_row-top+1] - ys[row-top])


# --- Exports image ---

IMAGE_MAX_PIXELS = 60_000_000
IMAGE_MARGIN_POINTS = 3


@contextlib.contextmanager
def _matplotlib_config_dir():
    """Garde la configuration de matplotlib hors du dépôt et ne laisse pas l'environnement modifié."""
    previous = os.environ.get("MPLCONFIGDIR")
    program_dir = Path(__file__).resolve().parent
    if program_dir.name == "dim2dir" and program_dir.parent.name == "src":
        program_dir = program_dir.parent.parent
    if "matplotlib" in sys.modules:
        yield
        return
    if previous:
        try:
            folder = Path(previous).expanduser().resolve()
            inside = folder.is_relative_to(program_dir) or folder.is_relative_to(Path.cwd().resolve())
        except (OSError, ValueError):
            inside = False
        if not inside:
            yield
            return
    folder = Path(tempfile.gettempdir()) / "dim2dir-matplotlib"
    folder.mkdir(parents=True, exist_ok=True)
    os.environ["MPLCONFIGDIR"] = str(folder)
    try:
        yield
    finally:
        if previous is None:
            os.environ.pop("MPLCONFIGDIR", None)
        else:
            os.environ["MPLCONFIGDIR"] = previous


def _matplotlib():
    """Seul point d'import de matplotlib : la garde de configuration est toujours appliquée."""
    with _matplotlib_config_dir():
        import matplotlib
        import matplotlib.backends.backend_agg
        import matplotlib.figure
        import matplotlib.font_manager
        import matplotlib.lines
        import matplotlib.patches
    return matplotlib


def _check_pixel_count(pixels, limit):
    if pixels > limit:
        raise _ImageLimitError(f"Export image refusé : {pixels:,} pixels, limite {limit:,} (60 millions par défaut) ; "
                               "relever limits.max_pixels / --max-pixels ou réduire --dpi.")


def _image_dimensions(ws, bounds, dpi):
    widths, heights = scene_geometry(ws, bounds)
    width, height = sum(widths) + 2 * IMAGE_MARGIN_POINTS, sum(heights) + 2 * IMAGE_MARGIN_POINTS
    pixels = math.ceil(width * dpi / 72) * math.ceil(height * dpi / 72)
    limit = getattr(ws, "_dim2dir_limits", Limits()).max_pixels
    _check_pixel_count(pixels, limit)
    return width, height


def _figure_from_sheet(ws, bounds, dpi=200):
    _matplotlib()
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.patches import Rectangle
    from matplotlib.lines import Line2D
    from matplotlib import font_manager, rc_context

    width, height = _image_dimensions(ws, bounds, dpi)
    fonts = {font.name.casefold(): font.name for font in font_manager.fontManager.ttflist}
    style = getattr(ws, "_dim2dir_style", Style())
    warnings = getattr(ws, "_dim2dir_warnings", [])
    if style.font_family and style.font_family.casefold() not in fonts:
        warning = f"police {style.font_family} introuvable : repli sur DejaVu Sans"
        if warning not in warnings:
            warnings.append(warning)
    family = fonts.get("arial", "DejaVu Sans")
    with rc_context({"font.family": family, "pdf.fonttype": 42, "savefig.bbox": None, "figure.autolayout": False}):
        fig = Figure(figsize=(width / 72, height / 72), dpi=dpi, facecolor="white")
        FigureCanvasAgg(fig)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(-IMAGE_MARGIN_POINTS, width - IMAGE_MARGIN_POINTS)
        ax.set_ylim(height - IMAGE_MARGIN_POINTS, -IMAGE_MARGIN_POINTS)
        ax.set_axis_off()
        for cell, x, y, w, h in _scene_rectangles(ws, bounds):
            fill = "#" + cell.fill.fgColor.rgb[-6:] if cell.fill.patternType == "solid" and cell.fill.fgColor.type == "rgb" else "white"
            rect = Rectangle((x, y), w, h, facecolor=fill, edgecolor="none", linewidth=0)
            ax.add_patch(rect)
            for side, xs, ys in ((cell.border.left, (x, x), (y, y+h)),
                                 (cell.border.right, (x+w, x+w), (y, y+h)),
                                 (cell.border.top, (x, x+w), (y, y)),
                                 (cell.border.bottom, (x, x+w), (y+h, y+h))):
                if side is not None and side.style:
                    color = "#" + side.color.rgb[-6:] if side.color and side.color.type == "rgb" else "#1F1F1F"
                    ax.add_line(Line2D(xs, ys, color=color,
                                      linewidth={"thin": 0.75, "medium": 1.5, "thick": 2.25}.get(side.style, 0.75)))
            if cell.value is not None:
                color = cell.font.color
                color = "#" + color.rgb[-6:] if color and color.type == "rgb" else "#1F1F1F"
                rotation = cell.alignment.textRotation or 0
                text = str(cell.value)
                size = cell.font.sz or 10
                if cell.alignment.shrinkToFit:
                    size = min(size, max(1, (w - 4) / max(1, len(text)) / 0.55))
                elif cell.alignment.wrap_text and not rotation:
                    import textwrap
                    text = "\n".join(textwrap.fill(line, max(1, int((w-4) / (size * 0.55)))) for line in text.split("\n"))
                artist = ax.text(x+w/2, y+h/2, text, ha="center", va="center", fontsize=size,
                                 fontproperties=font_manager.FontProperties(
                                     family=fonts.get((cell.font.name or "Arial").casefold(), "DejaVu Sans"),
                                     size=size, weight="bold" if cell.font.bold else "normal"),
                                 fontweight="bold" if cell.font.bold else "normal", color=color,
                                 rotation=rotation, clip_on=True, linespacing=1.05, parse_math=False)
                artist.set_clip_path(rect)
        return fig


def export_matplotlib(ws, bounds, pdf_path=None, png_path=None, dpi=200, svg_path=None):
    fig = _figure_from_sheet(ws, bounds, dpi)
    try:
        rc_context = _matplotlib().rc_context
        with rc_context({"pdf.fonttype": 42, "savefig.bbox": None}):
            if pdf_path:
                fig.savefig(pdf_path, format="pdf", dpi=dpi)
            if png_path:
                fig.savefig(png_path, format="png", dpi=dpi)
            if svg_path:
                fig.savefig(svg_path, format="svg", dpi=dpi)
    finally:
        fig.clear()

def _excel_worker(source, sheet_name, bounds, pdf_path, png_path, pipe):  # pragma: no cover
    """Tâche COM isolée pour pouvoir borner toute opération bloquante."""
    app = book = None
    pythoncom = None
    try:
        import pythoncom
        import win32com.client
        import win32process
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")
        pipe.send(("pid", win32process.GetWindowThreadProcessId(app.Hwnd)[1]))
        app.Visible = False
        app.DisplayAlerts = False
        app.ErrorCheckingOptions.BackgroundChecking = False  # pas de triangles verts dans l'image
        book = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
        sheet = book.Worksheets(sheet_name)
        left, top, right, bottom = bounds
        area = sheet.Range(sheet.Cells(top, left), sheet.Cells(bottom, right))
        if pdf_path:
            setup = sheet.PageSetup
            setup.PrintArea = area.Address
            setup.Zoom = False
            setup.FitToPagesWide = setup.FitToPagesTall = 1
            setup.LeftMargin = setup.RightMargin = setup.TopMargin = setup.BottomMargin = 3
            setup.HeaderMargin = setup.FooterMargin = 0
            setup.CenterHorizontally = True
            setup.Orientation = 2 if area.Width > area.Height else 1
            # Pas de titres d'impression, d'en-têtes ni de pieds de page hérités du classeur.
            setup.PrintTitleRows = ""
            setup.PrintTitleColumns = ""
            setup.LeftHeader = setup.CenterHeader = setup.RightHeader = ""
            setup.LeftFooter = setup.CenterFooter = setup.RightFooter = ""
            sheet.ExportAsFixedFormat(0, str(pdf_path), IgnorePrintAreas=False)
        if png_path:
            sheet.Activate()
            area.CopyPicture(Appearance=1, Format=2)
            chart = sheet.ChartObjects().Add(0, 0, area.Width, area.Height)
            try:
                chart.Activate()
                chart.Chart.Paste()
                if not chart.Chart.Export(str(png_path), "PNG"):
                    raise RuntimeError("Chart.Export a renvoyé False.")
            finally:
                chart.Delete()
        pipe.send(("ok", None))
    except Exception as exc:
        pipe.send(("error", str(exc)))
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
                pipe.close()


def export_excel(source, sheet_name, bounds, pdf_path=None, png_path=None, timeout=120):
    if sys.platform != "win32":
        raise RuntimeError("Moteur excel indisponible : Windows, Excel et pywin32 sont requis.")
    with _EXCEL_LOCK:
        _export_excel_process(source, sheet_name, bounds, pdf_path, png_path, timeout)


def _kill_excel_process(pid, subprocess):  # pragma: no cover
    """Arrête l'instance Excel de ce traitement, même si le processus enfant est déjà mort."""
    listing = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True)
    if "excel.exe" in listing.stdout.casefold():
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True)


def _export_excel_process(source, sheet_name, bounds, pdf_path, png_path, timeout):  # pragma: no cover
    import multiprocessing
    import subprocess
    context = multiprocessing.get_context("spawn")
    parent, child = context.Pipe(duplex=False)
    pid, result = None, None
    with tempfile.TemporaryDirectory(prefix=".dim2dir-com-", dir=source.parent, ignore_cleanup_errors=True) as tmp:
        copy = Path(tmp) / source.name
        shutil.copy2(source, copy)
        process = context.Process(target=_excel_worker, args=(copy, sheet_name, bounds, pdf_path, png_path, child))
        deadline = time.monotonic() + timeout
        try:
            process.start()
            child.close()
            while time.monotonic() < deadline:
                if parent.poll(min(0.1, max(0, deadline - time.monotonic()))):
                    try:
                        kind, message = parent.recv()
                    except EOFError:
                        process.join(min(1, max(0, deadline - time.monotonic())))
                        break
                    if kind == "pid":
                        pid = message
                    else:
                        result = (kind, message)
                if not process.is_alive():
                    break
            if process.is_alive():
                raise RuntimeError(f"Délai maximal de {timeout} s dépassé pour le moteur excel ; relever limits.excel_timeout_s / --excel-timeout.")
            if result is None or result[0] != "ok":
                raise RuntimeError(f"Moteur excel indisponible ou export impossible : {result[1] if result else 'processus interrompu'}")
        finally:
            if process.pid and process.is_alive():
                process.terminate()
                process.join(2)
            if pid and (result is None or result[0] != "ok"):
                _kill_excel_process(pid, subprocess)
            parent.close()
            child.close()

# --- Rapport de couverture ---

COVERAGE_NOTE = "Couverture calculée sur les niveaux observés dans le tableau, sans jugement d'optimalité statistique."


def coverage_report(sheet_name, headers, rows, layout, warnings, *, run_column="Run"):
    occupied = {}
    for row, col, label, order in layout["placements"]:
        occupied.setdefault((row, col), []).append(label)
    count = layout["width"] * layout["height"]
    return {"sheet": sheet_name, "runs": len(rows),
            "factors": [{"name": axis["name"], "levels": axis["levels"]} for axis in sorted((*layout["x_axes"], *layout["y_axes"]), key=lambda axis: headers.index(axis["column"]))],
            "assignment": {name: [axis["name"] for axis in layout[f"{name}_axes"]] for name in ("x", "y")},
            "grid": {"width": layout["width"], "height": layout["height"], "cells": count},
            "occupied_cells": len(occupied), "empty_cells": count - len(occupied),
            "fill_rate": len(occupied) / count if count else 0,
            "collisions": [{"row": row, "column": col, "runs": labels} for (row, col), labels in occupied.items() if len(labels) > 1],
            "image_engine": None, "files": [], "warnings": warnings, "status": "ok"}


def _locked_message(path):
    return f"Impossible de remplacer {Path(path).name} : fichier ouvert dans une autre application : {Path(path).name}."


def _atomic_save(wb, path):
    descriptor, name = tempfile.mkstemp(prefix=".dim2dir-", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    try:
        wb.save(name)
        try:
            os.replace(name, path)
        except PermissionError as exc:
            raise Dim2DirError(_locked_message(path)) from exc
    finally:
        Path(name).unlink(missing_ok=True)


def _write_report_file(report, path):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2, default=str)
        handle.write("\n")


def _commit_outputs(pairs):
    """Remplace les destinations ensemble : en cas d'échec, les fichiers déjà remplacés sont restaurés.

    `pairs` : liste de (fichier préparé, destination). Rien n'est touché si une destination est un dossier."""
    for _, destination in pairs:
        if Path(destination).is_dir():
            raise Dim2DirError(f"Impossible d'écrire {Path(destination).name} : un dossier porte déjà ce nom.")
    done, folders, current = [], [], None
    try:
        for source, destination in pairs:
            destination = current = Path(destination)
            destination.parent.mkdir(parents=True, exist_ok=True)
            backup = None
            if destination.exists():
                folder = tempfile.mkdtemp(prefix=".dim2dir-sauvegarde-", dir=destination.parent)
                folders.append(folder)
                backup = Path(folder) / destination.name
                os.replace(destination, backup)
            done.append((destination, backup))
            os.replace(source, destination)
    except OSError as exc:
        restored = True
        for destination, backup in reversed(done):
            try:
                if backup is not None:
                    os.replace(backup, destination)
                else:
                    destination.unlink(missing_ok=True)
            except OSError:
                restored = False
        if restored:
            for folder in folders:
                shutil.rmtree(folder, ignore_errors=True)
        if isinstance(exc, PermissionError):
            raise Dim2DirError(_locked_message(current)) from exc
        raise Dim2DirError(f"Impossible d'écrire {current.name} : {exc.strerror or exc}") from exc
    for folder in folders:
        shutil.rmtree(folder, ignore_errors=True)


def _lost_objects_warning(wb):
    """Avertissement si une feuille contient des images ou des graphiques que openpyxl ne réécrit pas."""
    sheets = [ws.title for ws in wb.worksheets
              if getattr(ws, "_images", None) or getattr(ws, "_charts", None)]
    if not sheets:
        return None
    return ("openpyxl ne conserve pas les images, graphiques et formes : le classeur de sortie les perd "
            f"(feuille{'s' if len(sheets) > 1 else ''} : {', '.join(sheets)}).")


def _validate_input(excel_file):
    path = Path(excel_file).expanduser().resolve()
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Seuls les fichiers .xlsx sont acceptés.")
    if not path.is_file():
        raise FileNotFoundError(f"Fichier introuvable : {path}")
    return path


def _safe_sheet_name(name):
    return re.sub(r'[^\w.-]+', "_", name, flags=re.UNICODE).strip(". ") or "feuille"


def _path_key(path):
    return os.path.normcase(os.path.abspath(str(path)))


def _check_output_collisions(args, output_dir, output_path, report_path, prepared):
    """Refuse, avant toute écriture, qu'une sortie en écrase une autre ou le rapport."""
    planned = []
    if not args.no_xlsx and any(entry[3]["status"] != "error" for entry in prepared):
        planned.append(output_path)
    for ws, bounds, stem, item in prepared:
        if item["status"] in ("skipped", "error"):
            continue
        for disabled, suffix in ((args.no_pdf, "pdf"), (args.no_png, "png"), (not getattr(args, "_svg", False), "svg")):
            if not disabled:
                planned.append(output_dir / f"{stem}.{suffix}")
    seen = set()
    for path in planned:
        key = _path_key(path)
        if key in seen:
            raise ValueError(f"Deux fichiers de sortie désigneraient le même chemin : {path}")
        seen.add(key)
    if report_path is not None and _path_key(report_path) in seen:
        raise ValueError(f"Le rapport ne peut pas désigner un fichier écrit par le programme : {report_path}")


def _default_sheet(wb):
    return "Feuil1" if "Feuil1" in wb.sheetnames else wb.sheetnames[0]


def process_workbook(args):
    style, limits = _cli_options(args)
    run_column = getattr(args, "run_column", "Run")
    if not isinstance(run_column, str) or not run_column.strip():
        raise Dim2DirError("run_column doit être un nom de colonne non vide.")
    run_column = run_column.strip()
    input_path = _validate_input(args.excel_file)
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else input_path.parent
    output_path = (output_dir / build_output_path(input_path).name).resolve()
    if _path_key(output_path) == _path_key(input_path):
        raise ValueError("Le chemin de sortie doit être différent du fichier d'entrée.")
    report_path = Path(args.report).expanduser().resolve() if args.report else None
    if report_path and _path_key(report_path) in (_path_key(input_path), _path_key(output_path)):
        raise ValueError("Le rapport doit avoir un chemin distinct de l'entrée et du classeur de sortie.")
    wb = load_workbook(input_path)
    values = None
    report = {"schema_version": 1, "version": __version__, "input": input_path.name,
              "coverage_note": COVERAGE_NOTE, "sheets": []}
    try:
        if args.list_sheets:
            if not args.quiet:
                print("Feuilles disponibles : " + ", ".join(wb.sheetnames))
            return report
        sheet_name = getattr(args, "sheet_name", None) or _default_sheet(wb)
        names = wb.sheetnames if args.all_sheets else [sheet_name]
        if not args.all_sheets and sheet_name not in wb.sheetnames:
            raise ValueError(f"Feuille introuvable : {sheet_name}. Feuilles disponibles : {', '.join(wb.sheetnames)}")
        values = load_workbook(input_path, data_only=True)
        prepared, image_names = [], set()
        for name in names:
            warnings = []
            ws = wb[name]
            try:
                headers, rows, max_col = read_plan_table(ws, values[name], warnings, run_column=run_column,
                                                         limits=limits)
            except ValueError as exc:
                if isinstance(exc, Dim2DirError) or not args.all_sheets:
                    raise
                warnings.append(str(exc))
                report["sheets"].append({"sheet": name, "status": "skipped", "runs": 0,
                                         "warnings": warnings, "image_engine": None, "files": []})
                continue
            layout = _api_layout(headers, rows, args.x_factors, args.y_factors, run_column)
            _apply_layout_style(layout, rows, run_column, style, warnings)
            item = coverage_report(name, headers, rows, layout, warnings, run_column=run_column)
            report["sheets"].append(item)
            ws._dim2dir_table = (headers, rows)
            ws._dim2dir_layout = layout
            ws._dim2dir_color = args.color
            ws._dim2dir_max_cells = limits.max_cells
            ws._dim2dir_style = style
            ws._dim2dir_limits = limits
            ws._dim2dir_layout_styled = True
            ws._dim2dir_warnings = warnings
            bounds = visualization_bounds(max_col, layout, style)
            skip = _sheet_skip_message(ws, layout["width"], layout["height"])
            if skip is None and (bounds[2] > EXCEL_MAX_COLUMNS or bounds[3] > EXCEL_MAX_ROWS):
                skip = (f"La visualisation exacte dépasse les limites d'Excel : {bounds[2]} colonnes, {bounds[3]} lignes ; "
                        f"limites fixes {EXCEL_MAX_COLUMNS} et {EXCEL_MAX_ROWS}, aucun paramètre ne les relève.")
            if skip:
                item["status"] = "skipped"
                warnings.append(skip)
                if layout["width"] * layout["height"] > limits.max_cells:
                    warnings.append(f"{layout['width'] * layout['height']} cellules, limite {limits.max_cells} ; "
                                    "relever limits.max_cells / --max-cells.")
            conflict = zone_conflict_message(
                ws, (bounds[0], bounds[1], bounds[0], bounds[1]) if skip else bounds)
            if conflict:
                item["status"] = "error"
                warnings.append(conflict)
                print(f"Erreur ({name}) : {conflict}", file=sys.stderr)
            stem = f"{input_path.stem}_visu_{_safe_sheet_name(name)}"
            base_stem, suffix = stem, 2
            while stem.casefold() in image_names:
                stem = f"{base_stem}_{suffix}"
                suffix += 1
            image_names.add(stem.casefold())
            prepared.append((ws, bounds, stem, item))
        if not prepared:
            raise ValueError("Aucune feuille valide contenant un tableau d'essais.")
        _check_output_collisions(args, output_dir, output_path, report_path, prepared)
        lost = _lost_objects_warning(wb) if not args.no_xlsx else None
        if lost:
            prepared[0][3]["warnings"].append(lost)
        if not args.dry_run:
            output_dir.mkdir(parents=True, exist_ok=True)
            written = [entry for entry in prepared if entry[3]["status"] != "error"]
            for ws, bounds, stem, item in written:
                add_visualization_to_sheet(ws)
            staged, report_temp = [], None
            try:
                with tempfile.TemporaryDirectory(prefix=".dim2dir-", dir=output_dir) as temp_dir:
                    stage = Path(temp_dir)
                    if not args.no_xlsx and written:
                        wb.save(stage / output_path.name)
                        staged.append((stage / output_path.name, output_path))
                        for ws, bounds, stem, item in written:
                            item["files"].append(output_path.name)
                    for ws, bounds, stem, item in prepared:
                        if item["status"] in ("skipped", "error") or (args.no_pdf and args.no_png and not args._svg):
                            continue
                        pdf = stage / f"{stem}.pdf" if not args.no_pdf else None
                        png = stage / f"{stem}.png" if not args.no_png else None
                        svg = stage / f"{stem}.svg" if args._svg else None
                        try:
                            paths = {kind: path for kind, path in (("pdf", pdf), ("png", png), ("svg", svg)) if path is not None}
                            engine = _save_images(ws, bounds, paths, args.image_engine, args.dpi, wb,
                                                  stage, limits.excel_timeout_s, item["warnings"])
                            item["image_engine"] = engine
                            for path in (pdf, png, svg):
                                if path is not None:
                                    staged.append((path, output_dir / path.name))
                                    item["files"].append(path.name)
                        except Exception as exc:
                            item["status"] = "error"
                            item["warnings"].append(str(exc))
                            print(f"Erreur ({ws.title}) : {exc}", file=sys.stderr)
                    if not args.no_xlsx or not args.no_pdf or not args.no_png or args._svg:
                        for item in report["sheets"]:
                            if not item["files"] and item["status"] == "skipped":
                                item["warnings"].append("skipped : aucune sortie demandée n'a pu être produite.")
                    if report_path:
                        report_path.parent.mkdir(parents=True, exist_ok=True)
                        descriptor, report_name = tempfile.mkstemp(prefix=".dim2dir-", suffix=".json",
                                                                   dir=report_path.parent)
                        os.close(descriptor)
                        report_temp = Path(report_name)
                        _write_report_file(report, report_temp)
                        staged.append((report_temp, report_path))
                    _commit_outputs(staged)
            finally:
                if report_temp is not None:
                    report_temp.unlink(missing_ok=True)
        if not args.quiet:
            print(_report_text(report), end="")
        return report
    finally:
        wb.close()
        if values is not None:
            values.close()


def run_visualization(excel_file: str = EXCEL_FILE, sheet_name: str = SHEET_NAME) -> Path:
    args = _parser().parse_args([str(excel_file), sheet_name, "--no-pdf", "--no-png"])
    process_workbook(args)
    return build_output_path(Path(excel_file).expanduser().resolve())


# --- API Python : même lecture, projection et bloc de rendu que le CLI ---

class _ReadOnlyDict(dict):
    """Instantané récursivement immuable du dictionnaire de projection."""

    def _deny(self, *args, **kwargs):
        raise TypeError("Le layout de Projection est en lecture seule.")

    __setitem__ = __delitem__ = clear = pop = popitem = setdefault = update = __ior__ = _deny

    def __reduce__(self):
        return _ReadOnlyDict, (dict(self),)


class _ReadOnlyList(list):
    """Conserve la structure et l'égalité du layout historique, sans mutation."""

    _deny = _ReadOnlyDict._deny
    __setitem__ = __delitem__ = append = clear = extend = insert = pop = remove = _deny
    reverse = sort = __iadd__ = __imul__ = _deny

    def __reduce__(self):
        return _ReadOnlyList, (list(self),)


def _freeze(value):
    if isinstance(value, dict):
        return _ReadOnlyDict({key: _freeze(item) for key, item in value.items()})
    if isinstance(value, list):
        return _ReadOnlyList(_freeze(item) for item in value)
    if isinstance(value, tuple):
        return tuple(_freeze(item) for item in value)
    return value


@dataclass(frozen=True)
class Projection:
    """Projection exacte ; coordonnées de grille depuis le haut, à partir de zéro."""

    sheet_name: str
    headers: tuple[str, ...]
    n_runs: int
    x_factors: tuple[str, ...]
    y_factors: tuple[str, ...]
    width: int
    height: int
    occupied: int
    collisions: tuple[tuple[str, ...], ...]
    placements: tuple[tuple[int, int, str], ...]
    warnings: tuple[str, ...]
    layout: dict
    style: Style


@dataclass(frozen=True)
class Result:
    """Fichiers absolus écrits, rapport console et projection."""

    files: dict[str, Path]
    report: str
    warnings: tuple[str, ...]
    projection: Projection
    data: dict[str, bytes]
    style: Style
    engine: str | None


def _report_text(report):
    lines = [COVERAGE_NOTE]
    for item in report["sheets"]:
        grid = item.get("grid")
        if grid:
            lines.append(f"{item['sheet']} : {item['runs']} essais ; grille {grid['width']} × {grid['height']} ; "
                         f"{item['occupied_cells']}/{grid['cells']} cellules occupées ({item['fill_rate']:.1%}) ; "
                         f"{len(item['collisions'])} collision(s) ; statut {item['status']}.")
        for warning in item["warnings"]:
            lines.append(f"Avertissement ({item['sheet']}) : {warning}")
    return "\n".join(lines) + "\n"


def _memory_value(value):
    # Aucun import de pandas/numpy : NA, NaT et scalaires sont reconnus par protocole.
    if value is None or type(value).__name__ in ("NAType", "NaTType"):
        return None
    try:
        if bool(value != value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        if getattr(getattr(value, "dtype", None), "kind", None) == "M":
            value = value.astype("datetime64[us]").item()
        else:
            value = value.item()
        if value is None or _is_blank_value(value):
            return None
    if isinstance(value, float) and math.isinf(value):
        return "inf" if value > 0 else "-inf"  # Excel ne stocke pas l'infini : niveau textuel
    return value


def _memory_workbook(data, sheet_name, run_column, color, limits):
    if hasattr(data, "columns") and hasattr(data, "itertuples"):
        if getattr(data.columns, "nlevels", 1) > 1:
            raise Dim2DirError("Colonnes MultiIndex non prises en charge : aplatir les colonnes avant l'appel.")
        try:
            count = len(data)
        except TypeError:
            count = None
        if count is not None:
            _check_runs(count, limits)
        columns = list(data.columns)
        records = []
        for row in data.itertuples(index=False, name=None):
            _check_runs(len(records) + 1, limits)
            records.append(row)
        index = getattr(data, "index", None)
        if (getattr(index, "name", None) is not None
                and str(index.name).strip().casefold() == run_column.casefold()
                and not any(str(c).strip().casefold() == run_column.casefold() for c in columns)):
            columns.insert(0, run_column)
            records = [(label, *row) for label, row in zip(index, records)]
    elif isinstance(data, list) and all(isinstance(row, dict) for row in data):
        _check_runs(len(data), limits)
        columns = list(dict.fromkeys(key for row in data for key in row))
        records = [tuple(row.get(key) for key in columns) for row in data]
    elif isinstance(data, dict):
        columns = list(data)
        if not all(isinstance(values, (list, tuple)) for values in data.values()):
            raise Dim2DirError("Le dictionnaire doit contenir des listes de colonnes.")
        lengths = {len(values) for values in data.values()}
        if len(lengths) > 1:
            raise Dim2DirError("Les colonnes doivent avoir la même longueur.")
        _check_runs(next(iter(lengths), 0), limits)
        records = list(zip(*data.values()))
    else:
        raise Dim2DirError("data doit être un chemin .xlsx, un DataFrame, une liste de dictionnaires ou un dictionnaire de listes.")
    title = sheet_name if sheet_name is not None else "Feuil1"
    if not isinstance(title, str) or not title or len(title) > 31 or re.search(r"[\\/*?:\[\]:]", title):
        raise Dim2DirError("sheet_name : nom Excel non vide de 31 caractères maximum, sans antislash, / * ? : [ ].")
    if len(columns) > EXCEL_MAX_COLUMNS:
        raise Dim2DirError(f"{len(columns)} colonnes dépassent la limite fixe Excel {EXCEL_MAX_COLUMNS}, non paramétrable.")
    if len(records) + 1 > EXCEL_MAX_ROWS:
        raise Dim2DirError(f"{len(records) + 1} lignes dépassent la limite fixe Excel {EXCEL_MAX_ROWS}, non paramétrable.")
    wb = Workbook()
    try:
        ws = wb.active
        ws.title = sheet_name if sheet_name is not None else "Feuil1"
        for col, name in enumerate(columns, 1):
            _set_text(ws.cell(1, col), str(name).strip())
        for row_idx, row in enumerate(records, 2):
            for col, value in enumerate(row, 1):
                _set_text(ws.cell(row_idx, col), _memory_value(value))
        ws["A1"].fill = PatternFill("solid", fgColor=color or DEFAULT_BASE_COLOR)
        return wb
    except Exception:
        wb.close()
        raise


def _api_layout(headers, rows, x_factors, y_factors, run_column):
    # Alias internes pour réutiliser sans modification la projection historique
    # (qui reconnaît Run), y compris lorsqu'un autre nom d'identifiant est demandé.
    aliases = {header: "Run" + " " * i if header.strip().casefold() == run_column.casefold()
               else f"__factor_{i}" for i, header in enumerate(headers)}
    original = {value: key for key, value in aliases.items()}

    def restore(text):
        return re.sub(r"__factor_(\d+)", lambda match: headers[int(match.group(1))], text)

    def factors(names):
        if names is None:
            return None
        if not isinstance(names, (list, tuple)) or not all(isinstance(name, str) for name in names):
            raise Dim2DirError("Les facteurs x/y doivent être des listes de noms de colonnes.")
        unknown = [name for name in names if name not in aliases]
        if unknown:
            raise Dim2DirError(f"Facteurs inconnus : {', '.join(unknown)}.")
        return [aliases[name] for name in names]

    x_alias, y_alias = factors(x_factors), factors(y_factors)
    if x_alias is not None and y_alias is not None:
        expected = [header for header in headers if header.strip().casefold() != run_column.casefold()]
        combined = list(x_factors) + list(y_factors)
        if len(combined) != len(set(combined)) or set(combined) != set(expected):
            raise Dim2DirError("Affectation x/y : partition exacte des facteurs requise, sans doublon. "
                               f"Facteurs attendus : {', '.join(expected)} ; reçus : {', '.join(combined)}.")
    try:
        layout = project_with_assignment(
            list(aliases.values()), [{aliases[key]: value for key, value in row.items()} for row in rows],
            x_alias, y_alias)
    except ValueError as exc:
        raise (exc if isinstance(exc, Dim2DirError) else type(exc)(restore(str(exc)))) from None
    for axis in (*layout["x_axes"], *layout["y_axes"]):
        axis["column"] = axis["name"] = original[axis["column"]]
        axis["notes"] = [restore(note) for note in axis.get("notes", ())]
    return layout


@contextlib.contextmanager
def _prepared_plan(data, x_factors, y_factors, run_column, sheet_name, style=None, limits=None):
    style, limits = _options(style, limits)
    wb = values = None
    ready = False
    try:
        if not isinstance(run_column, str) or not run_column.strip():
            raise Dim2DirError("run_column doit être un nom de colonne non vide.")
        run_column = run_column.strip()
        if sheet_name is not None and not isinstance(sheet_name, str):
            raise Dim2DirError("sheet_name doit être un nom de feuille.")
        source = _validate_input(data) if isinstance(data, (str, Path)) else None
        if source is not None:
            wb = load_workbook(source)
            values = load_workbook(source, data_only=True)
        else:
            wb = _memory_workbook(data, sheet_name, run_column, style.base_color, limits)
            values = wb
        name = sheet_name if sheet_name is not None else wb.sheetnames[0]
        if name not in wb.sheetnames:
            raise Dim2DirError(f"Feuille introuvable : {name}. Feuilles disponibles : {', '.join(wb.sheetnames)}")
        ws = wb[name]
        warnings = []
        headers, rows, max_col = read_plan_table(ws, values[name], warnings, run_column=run_column, limits=limits)
        layout = _api_layout(headers, rows, x_factors, y_factors, run_column)
        _apply_layout_style(layout, rows, run_column, style, warnings)
        ws._dim2dir_table = (headers, rows)
        ws._dim2dir_layout = layout
        ws._dim2dir_style = style
        ws._dim2dir_limits = limits
        ws._dim2dir_layout_styled = True
        ws._dim2dir_warnings = warnings
        ws._dim2dir_max_cells = limits.max_cells
        bounds = visualization_bounds(max_col, layout, style)
        skip = _sheet_skip_message(ws, layout["width"], layout["height"])
        if skip is None and (bounds[2] > EXCEL_MAX_COLUMNS or bounds[3] > EXCEL_MAX_ROWS):
            skip = (f"La visualisation exacte dépasse les limites d'Excel : {bounds[2]} colonnes, {bounds[3]} lignes ; "
                        f"limites fixes {EXCEL_MAX_COLUMNS} et {EXCEL_MAX_ROWS}, aucun paramètre ne les relève.")
        if skip:
            warnings.append(skip)
            if layout["width"] * layout["height"] > limits.max_cells:
                warnings.append(f"{layout['width'] * layout['height']} cellules, limite {limits.max_cells} ; "
                                "relever limits.max_cells / --max-cells.")
        item = coverage_report(name, headers, rows, layout, warnings, run_column=run_column)
        if skip:
            item["status"] = "skipped"
        projection = Projection(
            name, tuple(headers), len(rows), tuple(a["name"] for a in layout["x_axes"]),
            tuple(a["name"] for a in layout["y_axes"]), layout["width"], layout["height"],
            item["occupied_cells"], tuple(tuple(c["runs"]) for c in item["collisions"]),
            tuple((row, col, label) for row, col, label, _ in layout["placements"]),
            tuple(warnings), _freeze(layout), style)
        ready = True
        yield wb, ws, projection, item, source, bounds
    except FileNotFoundError:
        raise
    except Dim2DirError:
        raise
    except (ValueError, TypeError, KeyError, OverflowError, BadZipFile) as exc:
        raise Dim2DirError(str(exc)) from exc
    except Exception as exc:
        if not ready:
            raise Dim2DirError(f"Plan invalide : {exc}") from exc
        raise Dim2DirError(f"Production impossible : {exc}") from exc
    finally:
        if values is not None and values is not wb:
            values.close()
        if wb is not None:
            wb.close()


def project(data, *, x_factors=None, y_factors=None, run_column="Run", sheet_name=None,
            style=None, limits=None) -> Projection:
    """Calcule une projection sans créer de fichier ni de grille dense."""
    with _prepared_plan(data, x_factors, y_factors, run_column, sheet_name, style, limits) as prepared:
        return prepared[2]


def figure(data, *, style=None, limits=None, x_factors=None, y_factors=None,
           run_column="Run", sheet_name=None) -> "matplotlib.figure.Figure":
    """Retourne une Figure Agg, sans pyplot, sauvegarde ni affichage."""
    with _prepared_plan(data, x_factors, y_factors, run_column, sheet_name, style, limits) as prepared:
        _, ws, _, item, _, bounds = prepared
        if item["status"] == "skipped":
            raise Dim2DirError(item["warnings"][0])
        add_visualization_to_sheet(ws)
        before = len(item["warnings"])
        fig = _figure_from_sheet(ws, bounds)
        for warning in item["warnings"][before:]:
            python_warnings.warn(warning, UserWarning, stacklevel=2)
        return fig


def _output_formats(outputs, source):
    if outputs is None:
        return ("xlsx", "png", "pdf") if source else ("png",)
    if isinstance(outputs, str):
        outputs = (outputs,)
    if (not isinstance(outputs, (tuple, list)) or not outputs
            or any(not isinstance(kind, str) or kind not in ("xlsx", "png", "pdf", "svg") for kind in outputs)
            or len(set(outputs)) != len(outputs)):
        raise Dim2DirError("outputs doit être un sous-ensemble ordonné non vide de xlsx, png, pdf, svg, sans doublon.")
    return tuple(outputs)


def _save_images(ws, bounds, paths, engine, dpi, wb, temporary, timeout, warnings):
    _image_dimensions(ws, bounds, dpi)
    actual = engine
    if "svg" in paths and engine == "auto":
        warnings.append("Repli sur matplotlib : SVG nécessite le moteur matplotlib.")
        actual = "matplotlib"
    if actual in ("excel", "auto"):
        try:
            source = temporary / "classeur.xlsx"
            _atomic_save(wb, source)
            export_excel(source, ws.title, bounds, paths.get("pdf"), paths.get("png"), timeout)
            if "png" in paths:
                from PIL import Image
                with Image.open(paths["png"]) as image:
                    _check_pixel_count(image.width * image.height,
                                       getattr(ws, "_dim2dir_limits", Limits()).max_pixels)
            actual = "excel"
        except _ImageLimitError:
            raise
        except Exception as exc:
            if engine == "excel":
                raise Dim2DirError(str(exc)) from exc
            warnings.append(f"Repli sur matplotlib : {exc}")
            actual = "matplotlib"
    if actual == "matplotlib":
        if "svg" in paths:
            export_matplotlib(ws, bounds, paths.get("pdf"), paths.get("png"), dpi, paths["svg"])
        else:
            export_matplotlib(ws, bounds, paths.get("pdf"), paths.get("png"), dpi)
    return actual


def visualize(data, *, output_dir=None, name=None, outputs=None, in_memory=None,
              style=None, limits=None, engine="matplotlib", dpi=200,
              x_factors=None, y_factors=None, run_column="Run", sheet_name=None, report=False) -> Result:
    """Produit des fichiers atomiques ou des octets en mémoire, sans modifier la source."""
    style, limits = _options(style, limits)
    source_path = isinstance(data, (str, Path))
    outputs = _output_formats(outputs, source_path)
    if engine not in ("matplotlib", "excel", "auto"):
        raise Dim2DirError("Moteurs acceptés : matplotlib, excel, auto.")
    if "svg" in outputs and engine == "excel":
        raise Dim2DirError("SVG exige le moteur matplotlib ; engine='excel' ne prend pas en charge svg.")
    if not isinstance(dpi, Integral) or isinstance(dpi, bool) or not 50 <= dpi <= 600:
        shown = int(dpi) if isinstance(dpi, Integral) and not isinstance(dpi, bool) else repr(dpi)
        raise Dim2DirError(f"dpi = {shown} ; limites 50 à 600, régler dpi / --dpi.")
    dpi = int(dpi)
    if in_memory is not None and not isinstance(in_memory, bool):
        raise Dim2DirError("in_memory doit être None, True ou False.")
    if not isinstance(report, bool):
        raise Dim2DirError("report doit être True ou False.")
    if name is not None and (not isinstance(name, str) or not name.strip()
                             or name in (".", "..") or re.search(r'[/\\:*?"<>|]', name)):
        raise Dim2DirError("name doit être un radical de fichier, sans chemin.")
    memory = (output_dir is None and not source_path) if in_memory is None else in_memory
    if not memory and output_dir is None and not source_path:
        raise Dim2DirError("output_dir est obligatoire avec in_memory=False pour une source en mémoire.")
    with _prepared_plan(data, x_factors, y_factors, run_column, sheet_name, style, limits) as prepared:
        wb, ws, projection, item, source, bounds = prepared
        try:
            target = None if memory else (Path(output_dir).expanduser().resolve()
                                         if output_dir is not None else source.parent)
        except (TypeError, ValueError) as exc:
            raise Dim2DirError("output_dir doit être un chemin de dossier.") from exc
        stem = name if name is not None else source.stem if source else "plan"
        names = {kind: f"{stem}_visu.xlsx" if kind == "xlsx" else
                 f"{stem}_visu_{_safe_sheet_name(ws.title)}.{kind}" for kind in outputs}
        if report:
            names["report"] = f"{stem}_rapport.json"
        if source and target and any(_path_key(target / filename) == _path_key(source) for filename in names.values()):
            raise Dim2DirError("Le chemin de sortie doit être différent du fichier d'entrée.")
        if target is not None:
            if target.exists() and not target.is_dir():
                raise Dim2DirError("output_dir doit désigner un dossier.")
            target.mkdir(parents=True, exist_ok=True)
        document = {"schema_version": 1, "version": __version__,
                    "input": source.name if source else None, "coverage_note": COVERAGE_NOTE, "sheets": [item]}
        files, payload, staged = {}, {}, {}
        actual_engine = None
        if "xlsx" in outputs or item["status"] != "skipped":
            add_visualization_to_sheet(ws)
        if "xlsx" in outputs and source is not None:
            lost = _lost_objects_warning(wb)
            if lost:
                item["warnings"].append(lost)
        images = [kind for kind in outputs if kind != "xlsx"]
        # Tout est préparé dans un dossier temporaire ; les destinations ne sont remplacées qu'à la fin.
        in_buffers = memory and (engine == "matplotlib" or (engine == "auto" and "svg" in images))
        needs_staging = not memory or (bool(images) and not in_buffers)
        with (tempfile.TemporaryDirectory(prefix=".dim2dir-", dir=target) if needs_staging
              else contextlib.nullcontext(None)) as staging_name:
            staging = Path(staging_name) if staging_name else None
            if "xlsx" in outputs:
                if memory:
                    buffer = BytesIO()
                    wb.save(buffer)
                    payload["xlsx"] = buffer.getvalue()
                else:
                    wb.save(staging / names["xlsx"])
                    staged["xlsx"] = staging / names["xlsx"]
            if images and item["status"] != "skipped":
                try:
                    _image_dimensions(ws, bounds, dpi)
                except Dim2DirError as exc:
                    item["status"] = "skipped"
                    item["warnings"].append(str(exc))
            if images and item["status"] != "skipped":
                # Matplotlib en mémoire utilise BytesIO ; seul COM requiert des fichiers internes.
                if in_buffers:
                    streams = {kind: BytesIO() for kind in images}
                    actual_engine = _save_images(ws, bounds, streams, engine, dpi, wb, None,
                                                 limits.excel_timeout_s, item["warnings"])
                    payload.update((kind, stream.getvalue()) for kind, stream in streams.items())
                else:
                    try:
                        paths = {kind: staging / names[kind] for kind in images}
                        actual_engine = _save_images(ws, bounds, paths, engine, dpi, wb, staging,
                                                     limits.excel_timeout_s, item["warnings"])
                        for kind, path in paths.items():
                            if memory:
                                payload[kind] = path.read_bytes()
                            else:
                                staged[kind] = path
                    except _ImageLimitError as exc:
                        item["status"] = "skipped"
                        item["warnings"].append(str(exc))
                item["image_engine"] = actual_engine
            produced = [kind for kind in outputs if kind in staged or kind in payload]
            item["files"].extend(names[kind] for kind in produced)
            if not produced:
                item["warnings"].append("skipped : aucune sortie demandée n'a pu être produite.")
            if report:
                if memory:
                    payload["report"] = (json.dumps(document, ensure_ascii=False, indent=2, default=str) + "\n").encode("utf-8")
                else:
                    _write_report_file(document, staging / names["report"])
                    staged["report"] = staging / names["report"]
            if staged:
                _commit_outputs([(path, target / names[kind]) for kind, path in staged.items()])
                files = {kind: target / names[kind] for kind in staged}
        # Respecter l'ordre demandé, y compris lorsque xlsx vient après les images.
        files = {kind: files[kind] for kind in names if kind in files}
        payload = {kind: payload[kind] for kind in names if kind in payload}
        return Result(files, _report_text(document), tuple(item["warnings"]), projection, payload, style, actual_engine)

# --- CLI ---

def _bounded_dpi(value):
    try:
        result = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError("Le dpi doit être un entier entre 50 et 600.") from None
    if not 50 <= result <= 600:
        raise argparse.ArgumentTypeError("Le dpi doit être compris entre 50 et 600.")
    return result


def _positive_int(value):
    try:
        result = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError("Un entier strictement positif est requis.") from None
    if result < 1:
        raise argparse.ArgumentTypeError("Un entier strictement positif est requis.")
    return result


def _color(value):
    if not re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        raise argparse.ArgumentTypeError("La couleur doit être au format RRGGBB ou #RRGGBB.")
    return value.lstrip("#").upper()


def _factor_list(value):
    return [item.strip() for item in value.split(",") if item.strip()]


def _level_order(value):
    if "=" not in value or not value.split("=", 1)[0].strip():
        raise argparse.ArgumentTypeError("Ordre attendu : Facteur=a,b,c.")
    name, text = value.split("=", 1)
    levels = []
    for token in text.split(",") if text else []:
        token = token.strip()
        try:
            levels.append(json.loads(token))
        except ValueError:
            levels.append(token)
    return name.strip(), levels


def _cli_options(args):
    mapping = Style.from_json(args.style_json).to_dict() if getattr(args, "style_json", None) else {}
    for field in fields(Style):
        if hasattr(args, "style_" + field.name):
            mapping[field.name] = getattr(args, "style_" + field.name)
    if args.color is not None:
        mapping["base_color"] = args.color
    if getattr(args, "level_order", None):
        orders = dict(mapping.get("level_orders") or {})
        orders.update(args.level_order)
        mapping["level_orders"] = orders
    style = Style.from_dict(mapping)
    limits = Limits(max_cells=args.max_cells, max_pixels=getattr(args, "max_pixels", 60_000_000),
                    max_runs=getattr(args, "max_runs", 20_000),
                    excel_timeout_s=getattr(args, "excel_timeout", 120.0))
    outputs = getattr(args, "outputs", None)
    if outputs is not None:
        outputs = _output_formats(outputs.split(","), True)
        args.no_xlsx, args.no_png, args.no_pdf = (kind not in outputs for kind in ("xlsx", "png", "pdf"))
    args._svg = outputs is not None and "svg" in outputs
    if args._svg and args.image_engine == "excel":
        raise Dim2DirError("SVG exige matplotlib ; le moteur excel ne prend pas en charge svg.")
    return style, limits


class _FrenchArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        for original, translated in (
            ("unrecognized arguments:", "arguments non reconnus :"),
            ("invalid choice:", "choix invalide :"),
            ("choose from", "choisir parmi"),
            ("expected one argument", "une valeur est requise"),
        ):
            message = message.replace(original, translated)
        self.print_usage(sys.stderr)
        self.exit(2, f"{self.prog} : erreur : {message}\n")


def _parser():
    parser = _FrenchArgumentParser(prog="dim2dir", add_help=False,
                                   description="Dim2Dir : ajoute une visualisation 2D à un plan d'expériences Excel.")
    parser._positionals.title = "Arguments positionnels"
    parser._optionals.title = "Options"
    parser.add_argument("-h", "--help", action="help", help="Afficher cette aide et terminer")
    parser.add_argument("excel_file", nargs="?", default=EXCEL_FILE, help="Fichier source .xlsx")
    parser.add_argument("sheet_name", nargs="?", default=None,
                        help="Feuille (défaut : Feuil1 si elle existe, sinon la première)")
    parser.add_argument("--all-sheets", action="store_true", help="Traiter toutes les feuilles")
    parser.add_argument("--list-sheets", action="store_true", help="Lister les feuilles sans écrire")
    parser.add_argument("--output-dir", help="Dossier de sortie (défaut : celui de l'entrée)")
    for kind in ("xlsx", "pdf", "png"):
        parser.add_argument(f"--no-{kind}", action="store_true", help=f"Ne pas produire de {kind.upper()}")
    parser.add_argument("--image-engine", choices=("matplotlib", "excel", "auto"), default="matplotlib", help="Moteur d'image (défaut : matplotlib)")
    parser.add_argument("--dpi", type=_bounded_dpi, default=200, help="Résolution PNG, moteur matplotlib uniquement : 50..600 (défaut : 200)")
    parser.add_argument("--color", type=_color, help="Couleur RRGGBB ou #RRGGBB (défaut : A1 puis 335C81)")
    parser.add_argument("--run-column", default="Run", help="Nom de la colonne des numéros d'essai (défaut : Run)")
    parser.add_argument("--max-cells", type=_positive_int, default=SPATIAL_MAX_RENDERED_CELLS, help="Plafond de grille (défaut : 90000)")
    parser.add_argument("--outputs", help="Formats ordonnés xlsx,png,pdf,svg ; remplace les options --no-*")
    parser.add_argument("--style", dest="style_json", help="Fichier JSON Style ; les options explicites priment")
    parser.add_argument("--palette", dest="style_palette", type=lambda v: v.split(",") if "," in v else v,
                        default=argparse.SUPPRESS, help="Colormap (cividis, viridis...) ou HEX,HEX,...")
    for option, field, help_text in (
        ("--palette-reverse", "palette_reverse", "Inverser la palette"),
        ("--legend", "legend", "Ajouter la légende dans la feuille"),
    ):
        parser.add_argument(option, dest="style_" + field, action="store_true", default=argparse.SUPPRESS, help=help_text)
    for option, field, kind, help_text in (
        ("--font", "font_family", str, "Nom de police"),
        ("--font-size", "font_size", float, "Taille des étiquettes : 4..72 points"),
        ("--font-color", "font_color", str, "Couleur des étiquettes RRGGBB ou #RRGGBB"),
        ("--label-scale", "label_scale", float, "Échelle des colonnes d'étiquettes y et de la hauteur des étiquettes x : 0.5..3.0, défaut 1"),
        ("--cell-size", "cell_size", int, "Cellule : 12..120 pixels, défaut 34"),
        ("--line-color", "line_color", str, "Couleur des bordures RRGGBB ou #RRGGBB"),
        ("--title", "title", str, "Texte du titre"),
        ("--label-format", "label_format", str, "Étiquette y ; champs {name} et {level}"),
        ("--run-prefix", "run_prefix", str, "Préfixe des essais, défaut E"),
    ):
        parser.add_argument(option, dest="style_" + field, type=kind, default=argparse.SUPPRESS, help=help_text)
    parser.add_argument("--line-style", dest="style_line_style", choices=("none", "thin", "medium", "thick"),
                        default=argparse.SUPPRESS, help="Style des bordures, défaut thin")
    parser.add_argument("--no-title", dest="style_title", action="store_const", const="",
                        default=argparse.SUPPRESS, help="Supprimer le texte du titre")
    parser.add_argument("--level-order", type=_level_order, action="append", help="Facteur=a,b,c ; répétable")
    parser.add_argument("--max-pixels", type=_positive_int, default=60_000_000, help="Plafond image en pixels, défaut 60000000")
    parser.add_argument("--max-runs", type=_positive_int, default=20_000, help="Maximum d'essais avant conversion, défaut 20000")
    parser.add_argument("--excel-timeout", type=float, default=120.0, help="Délai Excel COM en secondes, défaut 120")
    parser.add_argument("--x-factors", type=_factor_list, help="Facteurs x séparés par virgules, premier = varie le plus vite")
    parser.add_argument("--y-factors", type=_factor_list, help="Facteurs y séparés par virgules, partition exacte avec x")
    parser.add_argument("--report", help="Rapport JSON (ne peut pas désigner un fichier de sortie ; noms de fichiers relatifs au dossier de sortie)")
    parser.add_argument("--strict", action="store_true",
                        help="Code de sortie 1 si une feuille est ignorée ou si une sortie demandée manque")
    parser.add_argument("--dry-run", action="store_true", help="Calculer le résumé sans rien écrire (ni sorties, ni rapport, ni dossier)")
    parser.add_argument("--quiet", action="store_true", help="Masquer le résumé et les avertissements")
    parser.add_argument("--version", action="version", version=f"Dim2Dir {__version__}", help="Afficher la version et terminer")
    return parser


def _console_tolerant():
    """Remplace les caractères hors de la page de code de la console au lieu de planter."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(errors="replace")
            except (OSError, ValueError):
                pass


def main(argv=None) -> int:
    _console_tolerant()
    try:
        args = _parser().parse_args(argv)
    except SystemExit as exc:
        return int(exc.code)
    try:
        report = process_workbook(args)
        if any(item["status"] == "error" for item in report["sheets"]):
            return 1
        if args.strict:
            ignored = [item["sheet"] for item in report["sheets"] if item["status"] == "skipped"]
            if ignored:
                print(f"Option --strict : feuille(s) ignorée(s) ou sans sortie : {', '.join(ignored)}.", file=sys.stderr)
                return 1
        return 0
    except Exception as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
