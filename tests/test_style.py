from copy import copy
from dataclasses import FrozenInstanceError
from io import BytesIO
import json

import pytest
from openpyxl import Workbook, load_workbook
from PIL import Image

from dim2dir import Style, Limits, Dim2DirError, figure, project, visualize
import dim2dir.core as app
from conftest import ROOT, import_file

ROWS = [{"Run": 1, "A": 1, "B": "bas"}, {"Run": 2, "A": 2, "B": "haut"}]


def rendered(style):
    result = visualize(ROWS, style=style, outputs=("xlsx", "png"), dpi=96)
    wb = load_workbook(BytesIO(result.data["xlsx"]))
    return result, wb


def grid_cells(ws):
    return [ws["F4"], ws["G3"]]


def pixel_for_cell(ws, image_bytes, coordinate):
    bounds = app.visualization_bounds(3, app.project_fractal_2d(list(ROWS[0]), ROWS), ws._dim2dir_style if hasattr(ws, "_dim2dir_style") else None)
    with Image.open(BytesIO(image_bytes)) as image:
        width, height = app._image_dimensions(ws, bounds, 96)
        for cell, x, y, w, h in app._scene_rectangles(ws, bounds):
            if cell.coordinate == coordinate:
                # À l'écart du texte central et de la bordure.
                px = int((x + w * .2 + app.IMAGE_MARGIN_POINTS) * image.width / width)
                py = int((y + h * .2 + app.IMAGE_MARGIN_POINTS) * image.height / height)
                return image.convert("RGB").getpixel((px, py))
    raise AssertionError(coordinate)


@pytest.mark.parametrize("name", ["tournage_L9", "soudage_2p4", "boulonnerie_L18", "materiaux_textuels", "plan_avec_collisions", "plan_trop_grand"])
def test_explicit_defaults_equal_v1(examples, name):
    v1 = import_file("style_v1", ROOT / "tests/reference/Dim2Dir_V1.py")
    a, b = load_workbook(examples[name]), load_workbook(examples[name])
    try:
        v1.add_visualization_to_sheet(a.active)
        app.add_visualization_to_sheet(b.active, style=Style(), limits=Limits())
        x, y = a.active, b.active
        assert (x.max_row, x.max_column) == (y.max_row, y.max_column)
        assert set(map(str, x.merged_cells.ranges)) == set(map(str, y.merged_cells.ranges))
        assert x.sheet_view.showGridLines == y.sheet_view.showGridLines
        for attr in ("column_dimensions", "row_dimensions"):
            xd, yd = getattr(x, attr), getattr(y, attr)
            assert set(xd) == set(yd)
            assert {k: dict(v) for k, v in xd.items()} == {k: dict(v) for k, v in yd.items()}
        for row in x:
            for cell in row:
                other = y[cell.coordinate]
                assert cell.value == other.value
                for attr in ("fill", "font", "border", "alignment", "protection"):
                    assert copy(getattr(cell, attr)) == copy(getattr(other, attr)), (cell.coordinate, attr)
                assert cell.number_format == other.number_format
    finally:
        a.close()
        b.close()


@pytest.mark.parametrize("palette", [None, "viridis", "cividis", "Greys", ("#ff0000", "00ff00", "0000ff")])
@pytest.mark.parametrize("reverse", [False, True])
def test_colors_in_sheet_and_png(palette, reverse):
    style = Style(palette=palette, palette_reverse=reverse, base_color="#224466")
    result, wb = rendered(style)
    try:
        for cell, order in zip(grid_cells(wb.active), (1, 2)):
            color = app._style_color(style, "224466", order, 2)
            assert cell.fill.fgColor.rgb[-6:] == color
            assert cell.font.color.rgb[-6:] == app._contrast_font_color(color)
            assert pixel_for_cell(wb.active, result.data["png"], cell.coordinate) == tuple(bytes.fromhex(color))
        assert bool(result.warnings) is (palette is not None)
    finally:
        wb.close()


def test_palette_single_run_midpoint():
    style = Style(palette=("000000", "FFFFFF"))
    result = visualize([{"A": 1}], outputs="xlsx", style=style)
    wb = load_workbook(BytesIO(result.data["xlsx"]))
    try:
        assert wb.active["D3"].fill.fgColor.rgb[-6:] == "808080"
    finally:
        wb.close()


@pytest.mark.parametrize("line_style,width", [("none", 0), ("thin", .75), ("medium", 1.5), ("thick", 2.25)])
def test_fonts_lines_labels_and_title(line_style, width):
    style = Style(font_family="DejaVu Sans", font_size=12, font_color="#112233", cell_font_color="FF0000",
                  line_color="ABCDEF", line_style=line_style, title="Mon plan", label_format="{level} / {name}", run_prefix="Test ")
    result, wb = rendered(style)
    fig = figure(ROWS, style=style)
    try:
        ws = wb.active
        cell = ws["F4"]
        assert (cell.value, cell.font.name, cell.font.sz, cell.font.color.rgb[-6:]) == ("Test 1", "DejaVu Sans", 11, "FF0000")
        assert ws["E2"].value == "Mon plan" and ws["E2"].font.sz == 13
        assert ws["E4"].value == "bas / B" and ws["E4"].font.sz == 12
        assert ws["E4"].font.color.rgb[-6:] == "112233"
        assert ws["F5"].font.sz == 12
        assert cell.border.left.style == (None if line_style == "none" else line_style)
        assert cell.border.left.color.rgb[-6:] == "ABCDEF"
        texts = {text.get_text(): text for text in fig.axes[0].texts}
        assert texts["Mon plan"].get_fontsize() == 13
        assert texts["Test 1"].get_fontsize() == 11
        assert texts["1"].get_fontsize() == 12
        assert texts["Test 1"].get_color() == "#FF0000"
        assert texts["bas / B"].get_color() == "#112233"
        assert texts["Test 1"].get_fontproperties().get_name() == "DejaVu Sans"
        if width:
            assert all(line.get_linewidth() == width for line in fig.axes[0].lines)
            assert all(line.get_color() == "#ABCDEF" for line in fig.axes[0].lines)
        else:
            assert not fig.axes[0].lines
        assert result.data["png"].startswith(b"\x89PNG")
    finally:
        fig.clear()
        wb.close()


def test_cell_size_dimensions_title_legend_and_order():
    small = visualize(ROWS, dpi=96)
    style = Style(cell_size=60, title="", legend=True, run_prefix="", level_orders={"A": [2, 1], "B": ["haut", "bas"]})
    result, wb = rendered(style)
    fig = figure(ROWS, style=style)
    try:
        ws = wb.active
        assert ws["E2"].value is None  # openpyxl recharge une chaîne vide comme cellule vide.
        assert ws.row_dimensions[3].height == 45
        assert ws.column_dimensions["F"].width == pytest.approx(55 / 7)
        assert ws["G3"].value == "1" and ws["F4"].value == "2"
        assert project(ROWS, style=style).layout["x_axes"][0]["levels"] == [2, 1]
        # Plan étroit (largeur 2) : texte fusionné sur toute la zone, sans échantillon qui le comprime.
        assert ws["E6"].value == "Couleur : ordre des essais, du premier au dernier"
        assert "E6:G6" in {str(area) for area in ws.merged_cells.ranges}
        text = {t.get_text() for t in fig.axes[0].texts}
        assert "Visualisation - Feuil1" not in text and "1" in text and "2" in text
        assert any(t.startswith("Couleur :") for t in text)
        with Image.open(BytesIO(small.data["png"])) as a, Image.open(BytesIO(result.data["png"])) as b:
            assert b.width > a.width and b.height > a.height
    finally:
        wb.close()
        fig.clear()


def test_level_orders_warnings_and_unknown_factor():
    p = project(ROWS, style=Style(level_orders={"A": [2, 999]}))
    assert p.layout["x_axes"][0]["levels"] == [2, 1]
    assert len(p.warnings) == 2
    with pytest.raises(Dim2DirError, match="facteur inconnu"):
        project(ROWS, style=Style(level_orders={"absent": []}))


def test_missing_font_warning_once_and_default_silent():
    result = visualize(ROWS, style=Style(font_family="PoliceCertainementAbsenteDim2Dir"), outputs=("png", "pdf", "svg"), dpi=50)
    assert result.warnings == ("police PoliceCertainementAbsenteDim2Dir introuvable : repli sur DejaVu Sans",)
    assert not visualize(ROWS, dpi=50).warnings


@pytest.mark.parametrize("kwargs", [
    {"base_color": "red"}, {"font_color": "123"}, {"cell_font_color": 123}, {"line_color": None},
    {"palette": "introuvable"}, {"palette": ["FFFFFF"]}, {"palette": ["bad", "FFFFFF"]}, {"palette": 1},
    {"font_size": 3}, {"font_size": 73}, {"font_size": float("nan")}, {"font_size": True},
    {"label_scale": 0.49}, {"label_scale": 3.01}, {"label_scale": float("nan")}, {"label_scale": True},
    {"label_scale": "2"}, {"label_scale": None},
    {"cell_size": 11}, {"cell_size": 121}, {"cell_size": 34.0}, {"line_style": "dotted"},
    {"label_format": "{bad}"}, {"label_format": "{name.foo}"}, {"label_format": "{}"},
    {"label_format": "{name!r}"}, {"label_format": "{level:bad}"}, {"label_format": "{"},
    {"title": 1}, {"run_prefix": 1}, {"font_family": ""}, {"legend": 1}, {"palette_reverse": 1},
    {"level_orders": []}, {"level_orders": {"A": "abc"}}, {"level_orders": {"A": [1, 1.0]}},
    {"level_orders": {"A": [[1]]}},
])
def test_style_validation(kwargs):
    with pytest.raises(Dim2DirError):
        Style(**kwargs)


def test_label_scale_widens_label_axes():
    rows = [{"Run": 1, "A": 1, "B": "niveau_long"}, {"Run": 2, "A": 2, "B": "autre_niveau"}]

    def book(style):
        return load_workbook(BytesIO(visualize(rows, style=style, outputs="xlsx").data["xlsx"]))
    base, wide, narrow = book(Style()), book(Style(label_scale=1.5)), book(Style(label_scale=0.5))
    try:
        w0 = base.active.column_dimensions["E"].width
        assert w0 > app.SPATIAL_AXIS_MIN_COLUMN_WIDTH
        assert wide.active.column_dimensions["E"].width == pytest.approx(w0 * 1.5)
        assert narrow.active.column_dimensions["E"].width < w0
    finally:
        base.close()
        wide.close()
        narrow.close()


def test_label_scale_bounds_and_json(tmp_path):
    assert Style().label_scale == 1.0
    assert Style(label_scale=0.5).label_scale == 0.5 and Style(label_scale=3).label_scale == 3
    path = tmp_path / "style.json"
    path.write_text(json.dumps({"label_scale": 2.0}), encoding="utf-8")
    style = Style.from_json(path)
    assert style.label_scale == 2.0 and style.to_dict()["label_scale"] == 2.0
    assert Style.from_dict({"label_scale": 1.25}).label_scale == 1.25
    with pytest.raises(Dim2DirError):
        Style.from_dict({"label_scale": 4})


def test_style_dict_json_and_immutability(tmp_path):
    original = {"A": [2, 1]}
    style = Style(palette=["#ff0000", "0000ff"], level_orders=original)
    original["A"].append(3)
    assert style.level_orders["A"] == (2, 1)
    assert isinstance(hash(style), int)
    assert list(style.level_orders) == ["A"] and len(style.level_orders) == 1
    with pytest.raises(KeyError):
        style.level_orders["bad"]
    with pytest.raises(FrozenInstanceError):
        style.title = "x"
    with pytest.raises(TypeError):
        style.level_orders["A"] = (1,)
    mapping = style.to_dict()
    mapping["level_orders"]["A"].append(4)
    assert style.level_orders["A"] == (2, 1)
    path = tmp_path / "style.json"
    path.write_text(json.dumps(style.to_dict()), encoding="utf-8")
    assert Style.from_json(path) == style == Style.from_dict(style.to_dict())
    with pytest.raises(Dim2DirError, match="inconnues"):
        Style.from_dict({"bad": 1})
    with pytest.raises(Dim2DirError):
        Style.from_dict([])
    path.write_text("{bad", encoding="utf-8")
    with pytest.raises(Dim2DirError, match="JSON"):
        Style.from_json(path)


def test_figure_reports_font_fallback_to_caller():
    with pytest.warns(UserWarning, match="police PoliceAbsenteDim2Dir introuvable : repli sur DejaVu Sans") as caught:
        fig = figure(ROWS, style=Style(font_family="PoliceAbsenteDim2Dir"))
    try:
        assert len(caught) == 1
        assert all(t.get_fontproperties().get_name() == "DejaVu Sans" for t in fig.axes[0].texts)
    finally:
        fig.clear()
