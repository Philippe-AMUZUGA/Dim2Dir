import math
from collections import Counter
from itertools import combinations
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

import dim2dir.core as app
from conftest import ROOT, import_file


def test_projection():
    headers = ["rUn", "A", "B", "C", "D"]
    rows = [dict(zip(headers, row)) for row in [(1,20,"haut",2,9), (2,10,"bas",1,8),
                                               (3,None,"haut",1,9), (4,10,"bas",1,8)]]
    layout = app.project_fractal_2d(headers, rows)
    assert layout["x_strides"] == [1, 2]
    assert layout["y_strides"] == [1, 2]
    assert (layout["width"], layout["height"]) == (4, 4)
    assert layout["x_axes"][0]["levels"] == [10, 20]
    assert layout["y_axes"][0]["levels"] == ["haut", "bas"]
    assert layout["placements"] == [(1,3,"E1",1), (2,0,"E2",2), (1,0,"E3",3), (2,0,"E4",4)]
    report = app.coverage_report("test", headers, rows, layout, [])
    assert report["occupied_cells"] == 3
    assert report["empty_cells"] == 13
    assert report["collisions"][0]["runs"] == ["E2", "E4"]
    assert app.project_fractal_2d([], [])["width"] == 0
    assert app.project_fractal_2d(["Run"], [{"Run": 1}])["placements"] == [(0,0,"E1",1)]


def test_assignment():
    headers = ["Run", "A", "B", "C"]
    rows = [dict(zip(headers, row)) for row in [(1,1,10,"a"), (2,2,20,"b")]]
    layout = app.project_with_assignment(headers, rows, ["C","A"], ["B"])
    assert [axis["name"] for axis in layout["x_axes"]] == ["C","A"]
    assert layout["x_strides"] == [1,2]
    assert layout["placements"][1][:2] == (0,3)
    assert app.project_with_assignment(headers, rows, [], ["A","B","C"])["width"] == 1
    for x, y in [(["A"],None), (None,["B"]), (["A"],["B"]), (["A","B"],["B","C"]), (["X"],["B","C"])]:
        with pytest.raises(ValueError):
            app.project_with_assignment(headers, rows, x, y)


def test_value_helpers():
    assert app._value_key(1) == app._value_key(1.0)
    assert app._value_key(True)[0] == "bool"
    assert app._value_key(float("nan"))[0] == "float"
    assert app._is_blank_value(float("nan"))
    assert app._format_visual_value(None) == "(vide)"
    assert app._format_visual_value(1.0) == "1"
    assert app._format_visual_value(0.125) == "0.125"
    assert app._format_visual_value("x" * 50).endswith("...")
    assert app._unique_nonblank_values([None]) == ["(vide)"]
    assert app._unique_nonblank_values([1,1.0,2]) == [1,2]
    for value, result in [(None,"E7"), (float("nan"),"E7"), ("x","Ex"), (1.5,"E1.5")]:
        assert app._run_label(value,7) == result
    assert app._contrast_font_color("FFFFFF") == "000000"
    assert app._contrast_font_color("000000") == "FFFFFF"
    assert app._blend_hex_colors("FFFFFF","000000",2) == "000000"
    assert app._trial_sequence_color("335C81",1,1) == "335C81"


def test_read_duplicates_and_gap():
    ws = Workbook().active
    ws.append(["Run","A","A","A_2","A",None,"ignoré"])
    ws.append([1,2,None,4,5,6,7])
    warnings = []
    headers, rows, col = app.read_plan_table(ws,ws,warnings)
    assert headers == ["Run","A","A_3","A_2","A_4"]
    assert col == 5
    assert rows[0]["A_3"] is None
    assert any("colonne F" in message for message in warnings)
    assert sum("double" in message for message in warnings) == 2
    assert any("1 cellule(s) vide(s)" in message for message in warnings)
    assert app._read_plan_table(ws,2,5) == (headers,rows)
    empty = Workbook().active
    with pytest.raises(ValueError, match="en-têtes"):
        app.read_plan_table(empty,empty,[])
    empty = Workbook().active
    empty.append(["Run","A"])
    with pytest.raises(ValueError, match="aucun essai"):
        app.read_plan_table(empty,empty,[])
    with pytest.raises(ValueError):
        app._detect_table_bounds(empty)
    with pytest.raises(ValueError):
        app._detect_table_bounds(Workbook().active)


def test_formula_cache_and_preservation(tmp_path):
    path = tmp_path / "formules.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    ws.append(["Run","A","B"])
    ws.append([1,"=2+3","=1+1"])
    ws.append([2,7,9])
    wb.save(path)
    # openpyxl n'écrit pas les caches : injection d'une valeur numérique XML synthétique.
    with ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    root = ET.fromstring(files["xl/worksheets/sheet1.xml"])
    root.find(".//s:c[@r='B2']/s:v", ns).text = "5"
    files["xl/worksheets/sheet1.xml"] = ET.tostring(root)
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    normal, values = load_workbook(path), load_workbook(path,data_only=True)
    warnings = []
    headers, rows, col = app.read_plan_table(normal.active,values.active,warnings)
    assert rows[0]["A"] == 5
    assert rows[0]["B"] is None
    assert any("C2" in warning for warning in warnings)
    direct_warnings = []
    assert app.read_plan_table(normal.active,normal.active,direct_warnings)[1][0]["A"] is None
    normal.close()
    values.close()
    out = app.run_visualization(str(path),"Feuil1")
    wb = load_workbook(out)
    assert wb.active["B2"].value == "=2+3"
    assert wb.active["C2"].value == "=1+1"
    wb.close()


def test_guard_and_empty_axes():
    ws = Workbook().active
    assert app._visualization_skip_message(300,300) is None
    assert "trop de cellules" in app._visualization_skip_message(301,300)
    ws._dim2dir_max_cells = 2
    assert app._sheet_skip_message(ws,2,1) is None
    assert "Limite demandée : 2" in app._sheet_skip_message(ws,2,2)
    ws.append(["Run","A"])
    ws.append([1,1])
    app.add_visualization_to_sheet(ws)
    assert ws.sheet_view.showGridLines is False
    assert app._hex_from_fill(ws["A1"]) == "335C81"
    ws["A1"].fill = PatternFill("solid",fgColor="AABBCC")
    assert app._hex_from_fill(ws["A1"]) == "AABBCC"


def test_standard_orthogonal_tables():
    generator = import_file("generator_check", ROOT / "examples/generate_examples.py")
    for table in (generator.L9,generator.L18):
        for i, j in combinations(range(len(table[0])),2):
            counts = Counter((row[i],row[j]) for row in table)
            assert len(set(counts.values())) == 1
            assert len(counts) == len({row[i] for row in table}) * len({row[j] for row in table})


def test_direct_sheet_edge_cases(monkeypatch):
    ws = Workbook().active
    ws.append(["Run","A"])
    ws.append([1,1])
    original = app.project_fractal_2d
    monkeypatch.setattr(app,"project_fractal_2d",lambda headers,rows:{"placements":[]})
    with pytest.raises(ValueError,match="aucun placement"):
        app.add_visualization_to_sheet(ws)
    monkeypatch.setattr(app,"project_fractal_2d",original)
    ws._dim2dir_layout = app.project_with_assignment(["Run","A"],[{"Run":1,"A":1}],[],["A"])
    app.add_visualization_to_sheet(ws)
    only_run = Workbook().active
    only_run.append(["Run"])
    only_run.append([1])
    app.add_visualization_to_sheet(only_run)
    assert app._value_key(1+2j)[0] == "complex"
