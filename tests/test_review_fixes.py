"""Régressions issues de la relecture : copie, libellés, sorties atomiques, CLI et légende."""
import copy
import dataclasses
import hashlib
import io
import json
import os
import pickle
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import dim2dir
import dim2dir.core as app
from dim2dir import Dim2DirError, Limits, Style, project, visualize
from conftest import ROOT

ROWS = [{"Run": 1, "A": 10, "B": "bas"}, {"Run": 2, "A": 20, "B": "haut"},
        {"Run": 3, "A": 10, "B": "haut"}]


def cli(*args, env=None, cwd=ROOT, script=None):
    return subprocess.run([sys.executable, str(script or ROOT / "Dim2Dir.py"), *map(str, args)], cwd=cwd,
                          capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180,
                          env={**os.environ, "PYTHONIOENCODING": "utf-8", **(env or {})})


def module(*args, env=None):
    arguments = [str(arg) for arg in args]
    if arguments and arguments[-1].endswith(".xlsx"):
        arguments.insert(0, arguments.pop())
    return subprocess.run([sys.executable, "-m", "dim2dir", *arguments], cwd=ROOT, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=180,
                          env={**os.environ, "PYTHONIOENCODING": "utf-8", **(env or {})})


def plan_file(path, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, table in sheets.items():
        ws = wb.create_sheet(title)
        for row in table:
            ws.append(row)
    wb.save(path)
    wb.close()
    return path


# --- A1 : copie et sérialisation -------------------------------------------------------------

def test_projection_and_result_survive_pickle_copy_replace():
    result = visualize(ROWS, outputs="png", in_memory=True, dpi=50, style=Style(level_orders={"A": [20, 10]}))
    for original in (result.projection, result):
        for clone in (pickle.loads(pickle.dumps(original)), copy.copy(original), copy.deepcopy(original)):
            assert clone == original
    projection = result.projection
    clones = [pickle.loads(pickle.dumps(projection)), copy.deepcopy(projection),
              dataclasses.replace(projection, sheet_name="Autre")]
    assert clones[2].sheet_name == "Autre"
    for clone in clones:
        assert clone.layout == projection.layout
        with pytest.raises(TypeError):
            clone.layout["width"] = 99
        with pytest.raises(TypeError):
            clone.layout["x_axes"].append(1)
    assert dataclasses.replace(result, report="x").report == "x"


# --- A2, A3 : libellés ----------------------------------------------------------------------

def test_infinite_levels_are_labelled_and_nan_stays_blank():
    inf = float("inf")
    p = project({"A": [1.5, inf, -inf, float("nan")], "B": ["a", "b", "c", "d"]})
    assert p.layout["x_axes"][0]["labels"] == ["1.5", "inf", "-inf"]  # niveaux textuels : ordre d'apparition
    assert app._format_visual_value(float("nan")) == "(vide)"
    assert app._format_visual_value(inf) == "inf"


def test_default_labels_unchanged_without_collision():
    p = project({"A": [0.1, 12345.6, 3.14159, 2]})
    assert p.layout["x_axes"][0]["labels"] == ["0.1", "2", "3.142", "1.235e+04"]
    assert not p.warnings


def test_close_levels_get_more_precision_with_warning():
    p = project({"A": [12345.1, 12345.2, 3], "B": [1, 2, 3], "C": [0.3, 0.4, 0.5]})
    axes = {axis["name"]: axis for axis in (*p.layout["x_axes"], *p.layout["y_axes"])}
    assert len(set(axes["A"]["labels"])) == 3
    assert axes["A"]["labels"] == ["3", "12345.1", "12345.2"]
    assert axes["C"]["labels"] == ["0.3", "0.4", "0.5"]
    assert any("Facteur 'A'" in warning and "chiffres significatifs" in warning for warning in p.warnings)
    assert not any("'C'" in warning for warning in p.warnings)
    assert not any("__factor_" in warning for warning in p.warnings)
    very_close = project({"A": [1.0000000001, 1.0000000002]})
    assert len(set(very_close.layout["x_axes"][0]["labels"])) == 2


def test_long_labels_with_same_prefix_are_distinguished():
    prefix = "x" * 50
    p = project({"A": [prefix + "1", prefix + "2", prefix + "3"], "B": [1, 2, 3]})
    labels = p.layout["x_axes"][0]["labels"]
    assert len(set(labels)) == 3 and all(len(label) <= 36 for label in labels)
    assert any("tronqué" in warning for warning in p.warnings)


# --- A4 : texte forcé ------------------------------------------------------------------------

HOSTILE = ["#N/A", "#DIV/0!", "=1+1", "+SUM(A1)", "@x"]


def test_every_written_string_is_text_cell():
    data = {"A": HOSTILE, "B": ["#N/A", "=1+1", "x", "y", "z"], "#N/A": [1, 2, 3, 4, 5]}
    result = visualize(data, outputs="xlsx", in_memory=True,
                       style=Style(title="#DIV/0!", legend=True, run_prefix="@", label_format="=N: {level}"))
    wb = load_workbook(io.BytesIO(result.data["xlsx"]))
    try:
        count = 0
        for row in wb.active.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    count += 1
                    assert cell.data_type == "s", (cell.coordinate, cell.value, cell.data_type)
        assert count > 15
    finally:
        wb.close()


# --- A5 : messages d'affectation ---------------------------------------------------------------

@pytest.mark.parametrize("x, y", [(["A"], ["A", "B"]), (["Run", "A"], ["B"]), (["A", "A"], ["B"]), (["A"], ["Z"])])
def test_assignment_messages_never_show_internal_aliases(x, y):
    with pytest.raises(Dim2DirError) as info:
        project({"Run": [1, 2], "A": [1, 2], "B": [3, 4]}, x_factors=x, y_factors=y)
    assert "__factor_" not in str(info.value)
    if "Run" in x:
        assert "Run" in str(info.value)


def test_assignment_message_names_custom_run_column():
    with pytest.raises(Dim2DirError) as info:
        project({"Essai": [1, 2], "A": [1, 2], "B": [3, 4]}, run_column="Essai", x_factors=["Essai"], y_factors=["A", "B"])
    assert "__factor_" not in str(info.value) and "Essai" in str(info.value)


# --- A6 : console non UTF-8 ----------------------------------------------------------------------

def test_cli_survives_cp1252_console_with_foreign_sheet_name(tmp_path):
    source = plan_file(tmp_path / "plan.xlsx", {"Plan 实验": [["Run", "A", "B"], [1, 1, "x"], [2, 2, "y"]]})
    env = {"PYTHONIOENCODING": "cp1252"}
    listing = cli(source, "--list-sheets", env=env)
    assert listing.returncode == 0, listing.stderr
    result = cli(source, "Plan 实验", "--output-dir", tmp_path / "out", "--dpi", "50", env=env)
    assert result.returncode == 0, result.stderr
    assert len(list((tmp_path / "out").glob("plan_visu*"))) == 3
    assert "Traceback" not in result.stderr


# --- A7 : légende sur plan étroit -------------------------------------------------------------------

@pytest.mark.parametrize("width", [1, 2, 3, 6, 12])
def test_legend_text_stays_readable_on_narrow_plans(width):
    style = Style(legend=True)
    with app._prepared_plan({"A": list(range(1, width + 1))}, None, None, "Run", None, style) as prepared:
        wb, ws, projection, item, source, bounds = prepared
        assert projection.width == width
        app.add_visualization_to_sheet(ws)
        legend_row = bounds[3]
        texts = [(cell, w) for cell, x, y, w, h in app._scene_rectangles(ws, bounds)
                 if cell.row == legend_row and cell.value]
        assert len(texts) == 1
        cell, w = texts[0]
        assert w > 0
        size = cell.font.sz or 10
        if cell.alignment.shrinkToFit:
            size = min(size, (w - 4) / len(cell.value) / 0.55)
        assert size >= 6
        assert cell.value.startswith("Couleur")
    fig = dim2dir.figure({"A": list(range(1, width + 1))}, style=style)
    legend = [t for t in fig.axes[0].texts if t.get_text().startswith("Couleur")]
    assert legend and all(t.get_fontsize() >= 6 for t in legend)


def test_legend_keeps_samples_when_wide_enough():
    style = Style(legend=True)
    with app._prepared_plan({"A": list(range(1, 13))}, None, None, "Run", None, style) as prepared:
        wb, ws, projection, item, source, bounds = prepared
        app.add_visualization_to_sheet(ws)
        left, _, right, bottom = bounds
        assert ws.cell(bottom, left).fill.patternType == "solid"
        assert ws.cell(bottom, right).fill.patternType == "solid"


# --- A8 : sorties atomiques --------------------------------------------------------------------------

def snapshot(folder):
    return {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(folder.iterdir()) if p.is_file()}


def test_visualize_replaces_nothing_when_png_destination_is_a_folder(tmp_path):
    out = tmp_path / "out"
    out.mkdir()
    (out / "plan_visu.xlsx").write_bytes(b"ancien xlsx")
    (out / "plan_visu_Feuil1.pdf").write_bytes(b"ancien pdf")
    (out / "plan_visu_Feuil1.png").mkdir()
    before = snapshot(out)
    with pytest.raises(Dim2DirError) as info:
        visualize(ROWS, output_dir=out, outputs=("xlsx", "pdf", "png"), dpi=50, report=True)
    assert "plan_visu_Feuil1.png" in str(info.value) and ".dim2dir-" not in str(info.value)
    assert snapshot(out) == before
    assert not list(out.glob(".dim2dir-*")) and not (out / "plan_rapport.json").exists()


def test_visualize_restores_files_when_replacement_is_refused(tmp_path, monkeypatch):
    out = tmp_path / "out"
    out.mkdir()
    (out / "plan_visu.xlsx").write_bytes(b"ancien xlsx")
    (out / "plan_visu_Feuil1.pdf").write_bytes(b"ancien pdf")
    before = snapshot(out)
    real = os.replace

    def refuse(source, destination):
        if str(destination).endswith(".png"):
            raise PermissionError("verrouillé")
        return real(source, destination)

    monkeypatch.setattr(app.os, "replace", refuse)
    with pytest.raises(Dim2DirError) as info:
        visualize(ROWS, output_dir=out, outputs=("xlsx", "pdf", "png"), dpi=50)
    assert "ouvert dans une autre application : plan_visu_Feuil1.png" in str(info.value)
    assert ".dim2dir-" not in str(info.value)
    monkeypatch.undo()
    assert snapshot(out) == before
    assert [p.name for p in out.iterdir()] == sorted(before)


def test_cli_replaces_nothing_when_png_destination_is_a_folder(simple_plan, tmp_path):
    out = tmp_path / "out"
    out.mkdir()
    (out / "source_visu.xlsx").write_bytes(b"ancien xlsx")
    (out / "source_visu_Feuil1.pdf").write_bytes(b"ancien pdf")
    (out / "source_visu_Feuil1.png").mkdir()
    before = snapshot(out)
    result = cli(simple_plan, "--output-dir", out, "--dpi", "50", "--report", tmp_path / "rapport.json")
    assert result.returncode == 1
    assert "source_visu_Feuil1.png" in result.stderr and ".dim2dir-" not in result.stderr
    assert snapshot(out) == before
    assert not (tmp_path / "rapport.json").exists()
    assert not list(out.glob(".dim2dir-*")) and not list(tmp_path.glob(".dim2dir-*"))


def test_same_permission_message_for_api_and_cli(simple_plan, tmp_path, monkeypatch):
    real = os.replace

    def refuse(source, destination):
        if str(destination).endswith(".png"):
            raise PermissionError("verrouillé")
        return real(source, destination)

    monkeypatch.setattr(app.os, "replace", refuse)
    with pytest.raises(Dim2DirError) as api:
        visualize(simple_plan, output_dir=tmp_path / "a", outputs=("png",), dpi=50)
    assert app.main([str(simple_plan), "--output-dir", str(tmp_path / "b"), "--dpi", "50"]) == 1
    assert "ouvert dans une autre application : source_visu_Feuil1.png" in str(api.value)
    assert not list((tmp_path / "b").iterdir())


# --- A9 : rapport en mémoire -----------------------------------------------------------------------------

def test_memory_report_lists_logical_file_names():
    result = visualize(ROWS, outputs=("png", "xlsx"), in_memory=True, dpi=50, report=True)
    document = json.loads(result.data["report"].decode("utf-8"))
    assert document["sheets"][0]["files"] == ["plan_visu_Feuil1.png", "plan_visu.xlsx"]


# --- A10 : options de la ligne de commande -----------------------------------------------------------------

def test_color_accepts_hash_prefix(simple_plan, tmp_path):
    assert app._parser().parse_args([str(simple_plan), "--color", "#334455"]).color == "334455"
    result = cli(simple_plan, "--color", "#334455", "--output-dir", tmp_path, "--no-pdf", "--no-png")
    assert result.returncode == 0, result.stderr
    wb = load_workbook(tmp_path / "source_visu.xlsx")
    try:
        colours = {cell.fill.fgColor.rgb[-6:] for row in wb.active.iter_rows() for cell in row
                   if cell.fill.patternType == "solid"}
        assert any(color != "335C81" for color in colours)
    finally:
        wb.close()


def test_run_column_option(tmp_path):
    source = plan_file(tmp_path / "plan.xlsx", {"Feuil1": [["Essai", "A", "B"], [10, 1, "x"], [20, 2, "y"]]})
    cli(source, "--run-column", "Essai", "--report", tmp_path / "r.json", "--no-pdf", "--no-png", "--quiet")
    report = json.loads((tmp_path / "r.json").read_text(encoding="utf-8"))
    assert [factor["name"] for factor in report["sheets"][0]["factors"]] == ["A", "B"]
    cli(source, "--report", tmp_path / "s.json", "--no-pdf", "--no-png", "--quiet")
    other = json.loads((tmp_path / "s.json").read_text(encoding="utf-8"))
    assert [factor["name"] for factor in other["sheets"][0]["factors"]] == ["Essai", "A", "B"]
    wrong = cli(source, "--run-column", "Essai", "--x-factors", "Essai,A", "--y-factors", "B", "--dry-run")
    assert wrong.returncode == 1 and "__factor_" not in wrong.stderr


def test_default_sheet_is_feuil1_else_first(tmp_path):
    table = [["Run", "A"], [1, 1], [2, 2]]
    only = plan_file(tmp_path / "seule.xlsx", {"Plan": table})
    assert cli(only, "--dry-run").returncode == 0
    both = plan_file(tmp_path / "deux.xlsx", {"Autre": [["Run", "X"], [1, 1]], "Feuil1": table})
    cli(both, "--report", tmp_path / "r.json", "--no-pdf", "--no-png", "--no-xlsx", "--quiet")
    report = json.loads((tmp_path / "r.json").read_text(encoding="utf-8"))
    assert report["sheets"][0]["sheet"] == "Feuil1"
    third = plan_file(tmp_path / "trois.xlsx", {"Premiere": table, "Seconde": table})
    cli(third, "--report", tmp_path / "t.json", "--no-pdf", "--no-png", "--no-xlsx", "--quiet")
    assert json.loads((tmp_path / "t.json").read_text(encoding="utf-8"))["sheets"][0]["sheet"] == "Premiere"


def test_strict_mode(tmp_path):
    source = plan_file(tmp_path / "plan.xlsx", {"Feuil1": [["Run", "A"], [1, 1], [2, 2]], "Vide": [["Run", "A"]]})
    assert cli(source, "--all-sheets", "--no-pdf", "--no-png", "--output-dir", tmp_path / "o1").returncode == 0
    strict = cli(source, "--all-sheets", "--strict", "--no-pdf", "--no-png", "--output-dir", tmp_path / "o2")
    assert strict.returncode == 1 and "strict" in strict.stderr.lower()
    assert cli(source, "Feuil1", "--strict", "--no-pdf", "--no-png", "--output-dir", tmp_path / "o3").returncode == 0
    big = cli(source, "Feuil1", "--strict", "--max-cells", "1", "--no-pdf", "--no-png", "--output-dir", tmp_path / "o4")
    assert big.returncode == 1


def test_parser_names():
    parser = app._parser()
    assert parser.prog == "dim2dir"
    assert "PLEX2" not in parser.format_help() and "Dim2Dir" in parser.description


# --- A11 : entiers numpy -----------------------------------------------------------------------------------

def test_numpy_integers_are_accepted_but_not_booleans():
    np = pytest.importorskip("numpy")
    limits = Limits(max_cells=np.int64(500), max_pixels=np.int32(10_000_000), max_runs=np.int64(50))
    assert (limits.max_cells, limits.max_pixels, limits.max_runs) == (500, 10_000_000, 50)
    assert type(limits.max_cells) is int
    style = Style(cell_size=np.int64(40))
    assert style.cell_size == 40 and type(style.cell_size) is int
    result = visualize(ROWS, outputs="png", in_memory=True, dpi=np.int64(60), style=style, limits=limits)
    assert result.data["png"]
    for kwargs in ({"max_cells": True}, {"max_runs": False}):
        with pytest.raises(Dim2DirError):
            Limits(**kwargs)
    with pytest.raises(Dim2DirError):
        Style(cell_size=True)
    with pytest.raises(Dim2DirError) as info:
        visualize(ROWS, outputs="png", in_memory=True, dpi=np.int64(900))
    assert "np.int64" not in str(info.value) and "900" in str(info.value)
    with pytest.raises(Dim2DirError):
        visualize(ROWS, outputs="png", in_memory=True, dpi=True)


# --- A12 : configuration de matplotlib ---------------------------------------------------------------------

def tree(folder):
    return sorted(str(p.relative_to(folder)) for p in folder.rglob("*")
                  if "__pycache__" not in p.parts and ".pytest_cache" not in p.parts and p.suffix != ".pyc")


@pytest.mark.parametrize("inside", [True, False])
def test_matplotlib_configuration_never_lands_in_the_repository(tmp_path, inside):
    configuration = (ROOT / ".mplcfg_sonde") if inside else tmp_path / "mpl-marque"
    before = tree(ROOT)
    script = ("import dim2dir\n"
              "from dim2dir import Style, visualize\n"
              "Style(palette='viridis')\n"
              "visualize({'A': [1, 2, 3]}, style=Style(palette='cividis'), outputs='png', in_memory=True, dpi=50)\n")
    env = {**os.environ, "MPLCONFIGDIR": str(configuration), "PYTHONPATH": str(ROOT / "src"),
           "PYTHONDONTWRITEBYTECODE": "1"}
    result = subprocess.run([sys.executable, "-c", script], cwd=tmp_path, env=env, capture_output=True,
                            text=True, timeout=300)
    assert result.returncode == 0, result.stderr
    assert tree(ROOT) == before
    assert not (ROOT / ".mplcfg_sonde").exists()


# --- A13 : largeur des cellules -----------------------------------------------------------------------------

def test_every_cell_size_is_exact_in_pixels():
    for size in range(12, 121):
        assert app._column_pixels((size - 5) / 7) == size
    with app._prepared_plan({"A": [1, 2]}, None, None, "Run", None, Style(cell_size=120)) as prepared:
        wb, ws, projection, item, source, bounds = prepared
        app.add_visualization_to_sheet(ws)
        widths, _ = app.scene_geometry(ws, bounds)
        assert widths[-1] == 90.0
    with app._prepared_plan({"A": [1, 2]}, None, None, "Run", None, None) as prepared:
        wb, ws, projection, item, source, bounds = prepared
        app.add_visualization_to_sheet(ws)
        assert app.scene_geometry(ws, bounds)[0][-1] == 25.5


# --- A14 : images et graphiques ------------------------------------------------------------------------------

def workbook_with_picture(path):
    from PIL import Image as PilImage
    from openpyxl.drawing.image import Image
    picture = path.with_suffix(".png")
    PilImage.new("RGB", (8, 8), "red").save(picture)
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    for row in ([["Run", "A"], [1, 1], [2, 2]]):
        ws.append(row)
    other = wb.create_sheet("Notes")
    other.add_image(Image(str(picture)), "B2")
    wb.save(path)
    wb.close()
    return path


def test_pictures_and_charts_trigger_a_loss_warning(tmp_path):
    source = workbook_with_picture(tmp_path / "images.xlsx")
    result = visualize(source, output_dir=tmp_path / "o", outputs="xlsx")
    assert any("openpyxl ne conserve pas" in warning and "Notes" in warning for warning in result.warnings)
    images_only = visualize(source, output_dir=tmp_path / "p", outputs="png", dpi=50)
    assert not any("openpyxl ne conserve pas" in warning for warning in images_only.warnings)
    shell = cli(source, "--output-dir", tmp_path / "q", "--no-pdf", "--no-png")
    assert "openpyxl ne conserve pas" in shell.stdout
    plain = cli(source, "--output-dir", tmp_path / "r", "--no-xlsx", "--no-pdf", "--no-png", "--dry-run")
    assert "openpyxl ne conserve pas" not in plain.stdout


# --- parité du fichier autonome et du paquet -----------------------------------------------------------------

def test_standalone_and_package_cli_give_identical_outputs(simple_plan, tmp_path):
    env = {"PYTHONPATH": str(ROOT / "src")}
    standalone = cli(simple_plan, "--output-dir", tmp_path / "a", "--dpi", "50", "--legend", env=env)
    package = module("--output-dir", tmp_path / "b", "--dpi", "50", "--legend", simple_plan, env=env)
    assert standalone.returncode == package.returncode == 0
    assert standalone.stdout == package.stdout
    assert sorted(p.name for p in (tmp_path / "a").iterdir()) == sorted(p.name for p in (tmp_path / "b").iterdir())
    assert (tmp_path / "a" / "source_visu_Feuil1.png").read_bytes() == (tmp_path / "b" / "source_visu_Feuil1.png").read_bytes()
    first, second = load_workbook(tmp_path / "a" / "source_visu.xlsx"), load_workbook(tmp_path / "b" / "source_visu.xlsx")
    try:
        dump = lambda wb: [[cell.value for cell in row] for row in wb.active.iter_rows()]
        assert dump(first) == dump(second)
    finally:
        first.close()
        second.close()
    assert cli("--help", env=env).stdout == module("--help", env=env).stdout


def test_inf_dans_classeur_ecrit_en_texte(tmp_path):
    import openpyxl
    result = dim2dir.visualize({"A": [1.0, float("inf")], "B": [1, 2]}, output_dir=tmp_path, outputs=("xlsx",), name="inf")
    ws = openpyxl.load_workbook(result.files["xlsx"]).active
    values = {ws.cell(r, 1).value for r in (2, 3)}
    assert "inf" in values and all(v == v for v in values if isinstance(v, float))
    assert not any(isinstance(c.value, float) and c.value == float("inf") for row in ws.iter_rows() for c in row)
