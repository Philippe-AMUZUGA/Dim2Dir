import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import dim2dir.core as app
from conftest import ROOT


def make_plan(path, rows, sheet="Feuil1", text_cells=()):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for row in rows:
        ws.append(row)
    for coordinate in text_cells:
        ws[coordinate].data_type = "s"
    wb.save(path)
    wb.close()
    return path


def run(*args):
    return app.main([str(item) for item in args])


# 1. Libellés commençant par "=" : texte, jamais formule.
def test_equal_sign_labels_stay_text(tmp_path):
    source = make_plan(tmp_path / "eq.xlsx",
                       [["Run", "=A", "B"], [1, "=niv1", "x"], [2, "=niv2", "y"]],
                       text_cells=("B1", "B2", "B3"))
    assert run(source, "--no-pdf", "--no-png", "--quiet") == 0
    wb = load_workbook(tmp_path / "eq_visu.xlsx")
    try:
        ws = wb.active
        zone = [cell for row in ws.iter_rows(min_col=5) for cell in row if cell.value is not None]
        by_value = {cell.value: cell for cell in zone}
        for value in ("=A", "=niv1", "=niv2"):
            assert by_value[value].data_type == "s", value
        assert not [cell for cell in zone if cell.data_type == "f"]
    finally:
        wb.close()


def test_equal_sign_in_y_label_and_title(tmp_path):
    source = make_plan(tmp_path / "eqy.xlsx",
                       [["A", "=B"], [1, "=p"], [2, "=q"]], sheet="=feuille", text_cells=("B1", "B2", "B3"))
    args = app._parser().parse_args([str(source), "=feuille", "--no-pdf", "--no-png", "--quiet"])
    app.process_workbook(args)
    wb = load_workbook(tmp_path / "eqy_visu.xlsx")
    try:
        cells = [c for row in wb.active.iter_rows(min_col=4) for c in row if c.value is not None]
        assert any(c.value == "=B: =p" for c in cells)
        assert all(c.data_type != "f" for c in cells)
    finally:
        wb.close()


# 2. matplotlib : pas d'interprétation mathématique.
def test_text_has_parse_math_disabled(tmp_path, monkeypatch):
    from matplotlib.axes import Axes
    calls = []
    original = Axes.text

    def spy(self, *args, **kwargs):
        calls.append(kwargs.get("parse_math"))
        return original(self, *args, **kwargs)

    monkeypatch.setattr(Axes, "text", spy)
    source = make_plan(tmp_path / "dollar.xlsx", [["A", "B"], ["$a^", "$"], ["$b_", "x $ y"]])
    assert run(source, "--output-dir", tmp_path / "out", "--dpi", "50", "--quiet") == 0
    assert calls and all(value is False for value in calls)


# 3. Limites d'Excel même avec --max-cells personnalisé.
def test_excel_limits_always_checked():
    ws = Workbook().active
    ws._dim2dir_max_cells = 10 ** 12
    assert "limites d'Excel" in app._sheet_skip_message(ws, app.EXCEL_MAX_COLUMNS + 1, 1)
    assert "limites d'Excel" in app._sheet_skip_message(ws, 1, app.EXCEL_MAX_ROWS + 1)
    assert app._sheet_skip_message(ws, 10, 10) is None
    del ws._dim2dir_max_cells
    assert "trop de cellules" in app._sheet_skip_message(ws, 400, 400)


# 4. --report ne peut pas désigner une sortie.
@pytest.mark.parametrize("name", ["source_visu.xlsx", "source_visu_Feuil1.pdf", "source_visu_Feuil1.png"])
def test_report_cannot_be_an_output(simple_plan, tmp_path, capsys, name):
    out = tmp_path / "out"
    assert run(simple_plan, "--output-dir", out, "--report", out / name, "--quiet") == 1
    assert "rapport" in capsys.readouterr().err.lower()
    assert not out.exists()


@pytest.mark.skipif(sys.platform != "win32", reason="chemins insensibles à la casse sous Windows")
def test_report_collision_case_insensitive(simple_plan, tmp_path):
    out = tmp_path / "out"
    assert run(simple_plan, "--output-dir", out, "--report", out / "SOURCE_VISU.XLSX", "--quiet") == 1
    assert not out.exists()


def test_report_unwritten_output_is_allowed(simple_plan, tmp_path):
    out = tmp_path / "out"
    assert run(simple_plan, "--output-dir", out, "--no-pdf",
               "--report", out / "source_visu_Feuil1.pdf", "--quiet") == 0


def test_output_collision_between_outputs(simple_plan, tmp_path):
    args = app._parser().parse_args([str(simple_plan)])
    out = tmp_path / "o"
    prepared = [(None, None, "s", {"status": "ok"}), (None, None, "s", {"status": "ok"})]
    with pytest.raises(ValueError, match="même chemin"):
        app._check_output_collisions(args, out, out / "a.xlsx", None, prepared)


# 5. Zone cible occupée : refus explicite, rien n'est écrasé.
def test_target_zone_conflict(tmp_path, capsys):
    source = make_plan(tmp_path / "busy.xlsx", [["Run", "A", "B"], [1, 1, 1], [2, 2, 2], [3, 1, 2]])
    wb = load_workbook(source)
    wb.active["F4"] = "note"
    wb.active["E3"] = "autre"
    wb.save(source)
    wb.close()
    args = app._parser().parse_args([str(source), "--output-dir", str(tmp_path / "out"), "--quiet"])
    report = app.process_workbook(args)
    item = report["sheets"][0]
    assert item["status"] == "error"
    assert any("E3" in warning for warning in item["warnings"])
    assert "E3" in capsys.readouterr().err
    assert not (tmp_path / "out" / "busy_visu.xlsx").exists()
    assert run(source, "--output-dir", tmp_path / "out2", "--quiet") == 1


def test_target_zone_conflict_direct():
    ws = Workbook().active
    ws.append(["Run", "A"])
    ws.append([1, 1])
    ws["D2"] = "x"
    with pytest.raises(ValueError, match="D2"):
        app.add_visualization_to_sheet(ws)
    assert app.first_zone_conflict(ws, 4, 2, 6, 5) == "D2"
    assert app.first_zone_conflict(ws, 5, 2, 6, 5) is None


def test_target_zone_conflict_on_skipped_note():
    ws = Workbook().active
    ws.append(["Run", "A"])
    ws.append([1, 1])
    ws._dim2dir_max_cells = 1
    ws.append([2, 2])
    ws["D2"] = "x"
    with pytest.raises(ValueError, match="D2"):
        app.add_visualization_to_sheet(ws)


def test_free_zone_is_unchanged(simple_plan, tmp_path):
    assert run(simple_plan, "--output-dir", tmp_path / "out", "--no-pdf", "--no-png", "--quiet") == 0


# 6. Identifiants Run non finis.
@pytest.mark.parametrize("value, expected", [("inf", "Einf"), ("-inf", "E-inf"), ("nan", "Enan"),
                                              ("1e309", "E1e309"), (float("nan"), "E7"),
                                              (float("inf"), "Einf"), (10 ** 400, f"E{10 ** 400}"),
                                              (3, "E3"), (3.0, "E3"), (None, "E7"), ("a", "Ea")])
def test_run_label_non_finite(value, expected):
    assert app._run_label(value, 7) == expected


def test_run_column_with_text_infinity(tmp_path):
    source = make_plan(tmp_path / "inf.xlsx", [["Run", "A"], ["inf", 1], ["nan", 2], ["1e309", 3]])
    assert run(source, "--no-pdf", "--no-png", "--quiet") == 0


# 8. MPLCONFIGDIR : ni dépôt pollué, ni environnement modifié.
@pytest.mark.parametrize("inside", [True, False])
def test_mplconfigdir_stays_out_of_repository(tmp_path, inside):
    folder = ROOT / ".mpl_test_inside"
    code = (
        "import os, sys, pathlib\n"
        f"sys.path.insert(0, {str(ROOT / 'src')!r})\n"
        "import dim2dir.core as app\n"
        "from openpyxl import Workbook\n"
        "ws = Workbook().active; ws['A1'] = 'x'\n"
        f"app.export_matplotlib(ws, (1, 1, 1, 1), None, pathlib.Path({str(tmp_path / 'a.png')!r}), 50)\n"
        "import matplotlib\n"
        "print(matplotlib.get_configdir()); print(os.environ.get('MPLCONFIGDIR'))\n"
    )
    env = {key: item for key, item in os.environ.items() if key != "MPLCONFIGDIR"}
    if inside:
        env["MPLCONFIGDIR"] = str(folder)
    result = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, env=env, timeout=180)
    assert result.returncode == 0, result.stderr
    configdir, after = result.stdout.strip().splitlines()
    assert not Path(configdir).resolve().is_relative_to(ROOT)
    assert after == ("None" if not inside else str(folder))
    assert not folder.exists() and not (ROOT / ".matplotlib").exists()


# 9. Rapport sans chemin absolu.
def test_report_has_no_absolute_paths(simple_plan, tmp_path):
    out, report = tmp_path / "out", tmp_path / "r.json"
    assert run(simple_plan, "--output-dir", out, "--dpi", "50", "--report", report, "--quiet") == 0
    data = json.loads(report.read_text(encoding="utf-8"))
    assert data["input"] == "source.xlsx"
    names = data["sheets"][0]["files"]
    assert sorted(names) == ["source_visu.xlsx", "source_visu_Feuil1.pdf", "source_visu_Feuil1.png"]
    assert str(tmp_path) not in json.dumps(data)
