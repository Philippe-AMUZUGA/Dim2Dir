import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest
from openpyxl import Workbook, load_workbook

import dim2dir.core as app
from conftest import ROOT


def cli(*args,cwd=ROOT):
    return subprocess.run([sys.executable,str(ROOT/"Dim2Dir.py"),*map(str,args)],
                          cwd=cwd,capture_output=True,text=True,encoding="utf-8",errors="replace",
                          env={**os.environ,"PYTHONIOENCODING":"utf-8"},timeout=120)


def test_cli_success_and_hash(simple_plan,tmp_path):
    before = hashlib.sha256(simple_plan.read_bytes()).hexdigest()
    result = cli(simple_plan,"--output-dir",tmp_path/"out","--dpi","50","--report",tmp_path/"report.json")
    assert result.returncode == 0, result.stderr
    assert hashlib.sha256(simple_plan.read_bytes()).hexdigest() == before
    assert "niveaux observés" in result.stdout
    assert len(list((tmp_path/"out").iterdir())) == 3
    assert json.loads((tmp_path/"report.json").read_text(encoding="utf-8"))["schema_version"] == 1


@pytest.mark.parametrize("args", [["--dpi","49"],["--dpi","601"],["--dpi","abc"],
                                 ["--max-cells","0"],["--max-cells","x"],["--color","bad"],
                                 ["--image-engine","unknown"],["--unknown"]])
def test_cli_usage(args):
    assert cli(*args).returncode == 2


def test_cli_errors(simple_plan,tmp_path):
    for args, message in [([tmp_path/"absent.xlsx"],"introuvable"),
                          ([tmp_path/"source.xls"],".xlsx"),
                          ([simple_plan,"Absente"],"Feuilles disponibles"),
                          ([simple_plan,"--x-factors","A"],"ensemble"),
                          ([simple_plan,"--x-factors","A,B","--y-factors","B,C"],"partition")]:
        result = cli(*args)
        assert result.returncode == 1
        assert message in result.stderr
    assert not (tmp_path/"source_visu.xlsx").exists()


def test_cli_dry_run_list_version(simple_plan,tmp_path):
    out = tmp_path/"ne_pas_creer"
    report = tmp_path/"absent.json"
    result = cli(simple_plan,"--dry-run","--output-dir",out,"--report",report)
    assert result.returncode == 0
    assert not out.exists() and not report.exists()
    assert "grille" in result.stdout
    result = cli(simple_plan,"--list-sheets","--output-dir",out)
    assert result.returncode == 0 and "Feuil1" in result.stdout
    assert not out.exists()
    assert "2.1.0" in cli("--version").stdout
    assert cli("--help").returncode == 0


def test_all_sheets(simple_plan,tmp_path):
    wb = load_workbook(simple_plan)
    wb.create_sheet("Vide")
    ws = wb.create_sheet("En-têtes seuls")
    ws.append(["Run","Facteur"])
    ws = wb.create_sheet("Essais bis")
    ws.append(["Run","A"])
    ws.append([1,7])
    wb.save(simple_plan)
    wb.close()
    result = cli(simple_plan,"--all-sheets","--no-pdf","--no-png","--report",tmp_path/"all.json")
    assert result.returncode == 0, result.stderr
    items = json.loads((tmp_path/"all.json").read_text(encoding="utf-8"))["sheets"]
    assert [item["status"] for item in items] == ["ok","skipped","skipped","ok"]
    assert "Avertissement" in result.stdout
    empty = tmp_path/"empty.xlsx"
    Workbook().save(empty)
    assert cli(empty,"--all-sheets").returncode == 1


def test_default_entrypoint(simple_plan,tmp_path):
    source = tmp_path/"Dim2Dir.xlsx"
    source.write_bytes(simple_plan.read_bytes())
    assert cli(cwd=tmp_path).returncode == 0
    assert (tmp_path/"Dim2Dir_visu.xlsx").exists()
    assert (tmp_path/"Dim2Dir_visu_Feuil1.pdf").exists()
    assert (tmp_path/"Dim2Dir_visu_Feuil1.png").exists()


def test_main_flags_and_manual_axes(simple_plan,tmp_path,capsys):
    args = [str(simple_plan),"--output-dir",str(tmp_path/"manual"),"--x-factors","C,A","--y-factors","B",
            "--color","AA44bb","--quiet","--no-pdf","--no-png"]
    assert app.main(args) == 0
    assert capsys.readouterr().out == ""
    wb = load_workbook(tmp_path/"manual/source_visu.xlsx")
    assert wb.active["J3"].fill.fgColor.rgb[-6:] == app._trial_sequence_color("AA44BB",2,3)
    wb.close()
    assert app.main([str(simple_plan),"--list-sheets","--quiet"]) == 0
    assert app.main(["--version"]) == 0
    assert app.main(["--dpi","x"]) == 2
    assert app.main(["--dpi","1"]) == 2
    assert app.main(["--max-cells","x"]) == 2
    assert app.main(["--max-cells","0"]) == 2
    assert app.main(["--color","#33445"]) == 2
    assert app.main([str(tmp_path/"missing.xlsx")]) == 1
    assert app.main([str(simple_plan),"--report",str(simple_plan)]) == 1
    assert app.main([str(simple_plan),"--report",str(app.build_output_path(simple_plan))]) == 1
    assert app.main([str(tmp_path/"input.xls")]) == 1
    assert app.main([str(simple_plan),"Absent"]) == 1


def test_atomic_failures(simple_plan,tmp_path,monkeypatch):
    output = tmp_path/"out.xlsx"
    output.write_bytes(b"original")
    wb = Workbook()
    def fail(*args,**kwargs):
        raise OSError("écriture refusée synthétique")
    monkeypatch.setattr(wb,"save",fail)
    with pytest.raises(OSError):
        app._atomic_save(wb,output)
    assert output.read_bytes() == b"original"
    assert not list(tmp_path.glob(".dim2dir-*"))
    monkeypatch.setattr(app,"build_output_path",lambda path:path)
    assert app.main([str(simple_plan)]) == 1


def test_sheet_name_sanitization():
    assert app._safe_sheet_name("Essais / été") == "Essais_été"
    assert app._safe_sheet_name("...") == "feuille"
    assert app._factor_list(" A, B , ") == ["A","B"]


def test_direct_reports_and_limits(simple_plan,tmp_path,monkeypatch):
    def args(*extra):
        return app._parser().parse_args([str(simple_plan),"--output-dir",str(tmp_path/"direct"),*extra])
    assert app.process_workbook(args("--dry-run"))["sheets"][0]["files"] == []
    assert not (tmp_path/"direct").exists()
    assert app.process_workbook(args("--list-sheets"))["sheets"] == []
    wb = load_workbook(simple_plan)
    wb.create_sheet("Vide")
    wb.save(simple_plan)
    wb.close()
    report = app.process_workbook(args("--all-sheets","--no-xlsx","--no-pdf","--no-png"))
    assert report["sheets"][1]["status"] == "skipped"
    monkeypatch.setattr(app,"EXCEL_MAX_COLUMNS",2)
    report = app.process_workbook(args("--report",str(tmp_path/"limit.json")))
    assert report["sheets"][0]["status"] == "skipped"
    assert not list((tmp_path/"direct").glob("*.png"))
    wb = load_workbook(tmp_path/"direct/source_visu.xlsx")
    assert "limites d'Excel" in wb.active["F2"].value
    wb.close()
    monkeypatch.setattr(app,"EXCEL_MAX_COLUMNS",16384)
    report = app.process_workbook(args("--max-cells","1","--no-pdf","--no-png"))
    assert report["sheets"][0]["status"] == "skipped"
    empty = tmp_path/"empty_direct.xlsx"
    Workbook().save(empty)
    assert app.main([str(empty),"--all-sheets"]) == 1
    assert app.main([str(empty)]) == 1


def test_sanitized_name_collisions(simple_plan,tmp_path,monkeypatch):
    wb = load_workbook(simple_plan)
    wb.active.title = "Essais +"
    ws = wb.copy_worksheet(wb.active)
    ws.title = "Essais !"
    wb.save(simple_plan)
    wb.close()
    paths = []
    def record_export(ws,bounds,pdf,png,dpi):
        paths.append(png)
        png.write_bytes(b"fichier de test")
    monkeypatch.setattr(app,"export_matplotlib",record_export)
    assert app.main([str(simple_plan),"--all-sheets","--no-pdf","--no-xlsx","--output-dir",str(tmp_path/"names")]) == 0
    assert paths[0].name == "source_visu_Essais_.png"
    assert paths[1].name == "source_visu_Essais__2.png"


def test_export_failure_report(simple_plan,tmp_path,monkeypatch,capsys):
    def fail(*args,**kwargs):
        raise RuntimeError("échec export synthétique")
    monkeypatch.setattr(app,"export_matplotlib",fail)
    report = tmp_path/"error.json"
    assert app.main([str(simple_plan),"--quiet","--no-xlsx","--output-dir",str(tmp_path/"err"),"--report",str(report)]) == 1
    assert json.loads(report.read_text(encoding="utf-8"))["sheets"][0]["status"] == "error"
    assert "échec export" in capsys.readouterr().err


NEW_OPTIONS = ("--outputs", "--palette", "--palette-reverse", "--font", "--font-size", "--font-color",
               "--label-scale", "--cell-size", "--line-color", "--line-style", "--title", "--no-title", "--legend",
               "--label-format", "--run-prefix", "--level-order", "--style", "--max-pixels", "--max-runs", "--excel-timeout")


def test_new_help_options():
    result = cli("--help")
    assert result.returncode == 0
    assert all(option in result.stdout for option in NEW_OPTIONS)


def test_cli_custom_style_and_outputs(simple_plan, tmp_path):
    out = tmp_path / "styled"
    style = tmp_path / "style.json"
    style.write_text(json.dumps({"palette": "cividis", "font_size": 9, "title": "JSON",
                                 "cell_font_color": "FF0000", "level_orders": {"B": ["bas", "haut"]}}), encoding="utf-8")
    args = [str(simple_plan), "--output-dir", str(out), "--style", str(style),
            "--outputs", "svg,png,pdf,xlsx", "--no-png", "--no-pdf", "--no-xlsx",
            "--palette", "FF0000,0000FF", "--palette-reverse", "--font", "DejaVu Sans", "--font-size", "12",
            "--font-color", "112233", "--cell-size", "50", "--line-color", "ABCDEF", "--line-style", "thick",
            "--title", "Titre CLI", "--legend", "--label-format", "{level} / {name}", "--run-prefix", "T",
            "--level-order", "A=20,10", "--level-order", "B=haut,bas",
            "--max-pixels", "60000000", "--max-runs", "10", "--excel-timeout", "2.5", "--dpi", "50"]
    assert app.main(args) == 0
    assert {p.suffix for p in out.iterdir()} == {".svg", ".png", ".pdf", ".xlsx"}
    wb = load_workbook(out / "source_visu.xlsx")
    try:
        ws = wb.active
        assert ws["F2"].value == "Titre CLI" and ws["F2"].font.sz == 13
        assert ws["F3"].value == "bas / B"
        assert ws.row_dimensions[3].height == 37.5
        assert ws["H3"].value == "T1\nT3"
        assert ws["H3"].font.name == "DejaVu Sans" and ws["H3"].font.sz == 11
        assert ws["H3"].font.color.rgb[-6:] == "FF0000"
        assert ws["H3"].fill.fgColor.rgb[-6:] == "0000FF"
        assert ws["H3"].border.left.style == "thick"
        assert ws["H3"].border.left.color.rgb[-6:] == "ABCDEF"
        assert ws["F3"].font.color.rgb[-6:] == "112233"
        assert any(cell.value == "Couleur : ordre des essais, du premier au dernier" for row in ws for cell in row)
    finally:
        wb.close()


def test_cli_no_title_and_colormap(simple_plan, tmp_path):
    out = tmp_path / "without_title"
    assert app.main([str(simple_plan), "--output-dir", str(out), "--outputs", "xlsx", "--palette", "viridis",
                     "--title", "Retiré", "--no-title", "--line-style", "none", "--quiet"]) == 0
    wb = load_workbook(out / "source_visu.xlsx")
    assert wb.active["F2"].value is None
    assert wb.active["G4"].border.left.style is None
    assert wb.active["G4"].fill.fgColor.rgb[-6:] == "440154"
    wb.close()


def test_cli_label_scale_option_and_json(simple_plan, tmp_path):
    def width(out, *args):
        assert app.main([str(simple_plan), "--output-dir", str(out), "--outputs", "xlsx", "--quiet", *args]) == 0
        wb = load_workbook(out / "source_visu.xlsx")
        try:
            return wb.active.column_dimensions["F"].width
        finally:
            wb.close()
    base = width(tmp_path / "a")
    scaled = width(tmp_path / "b", "--label-scale", "1.5")
    assert scaled > base
    path = tmp_path / "style.json"
    path.write_text('{"label_scale": 1.5}', encoding="utf-8")
    assert width(tmp_path / "c", "--style", str(path)) == pytest.approx(scaled)
    assert width(tmp_path / "d", "--style", str(path), "--label-scale", "1") == pytest.approx(base)


def test_cli_style_color_precedence(simple_plan, tmp_path):
    path = tmp_path / "style.json"
    path.write_text('{"base_color": "112233", "font_size": 14}', encoding="utf-8")
    assert app.main([str(simple_plan), "--outputs", "xlsx", "--style", str(path), "--color", "AABBCC", "--quiet"]) == 0
    wb = load_workbook(tmp_path / "source_visu.xlsx")
    assert wb.active["G4"].fill.fgColor.rgb[-6:] == app._trial_sequence_color("AABBCC", 1, 3)
    assert wb.active["G4"].font.sz == 13
    wb.close()


@pytest.mark.parametrize("options,message", [
    (["--outputs", "png,png"], "outputs"), (["--outputs", ""], "outputs"), (["--outputs", "gif"], "outputs"),
    (["--outputs", "svg", "--image-engine", "excel"], "SVG"),
    (["--max-runs", "2"], "3 essais"), (["--excel-timeout", "0"], "excel_timeout_s"),
    (["--font-size", "3"], "font_size"), (["--label-scale", "3.5"], "label_scale"), (["--palette", "bad"], "Palette inconnue"),
])
def test_cli_new_validation(simple_plan, tmp_path, capsys, options, message):
    out = tmp_path / "unwritten"
    assert app.main([str(simple_plan), "--output-dir", str(out), *options]) == 1
    assert message in capsys.readouterr().err
    assert not out.exists()


def test_cli_level_order_usage_error():
    assert app.main(["--level-order", "sans_egal"]) == 2


def test_cli_max_pixels_and_timeout(simple_plan, tmp_path, monkeypatch):
    assert app.main([str(simple_plan), "--outputs", "png", "--max-pixels", "1", "--output-dir", str(tmp_path / "pixels")]) == 1
    assert not list((tmp_path / "pixels").iterdir())
    def export(source, sheet, bounds, pdf, png, timeout):
        assert timeout == .25
        from PIL import Image
        Image.new("RGB", (10, 10)).save(png)
    monkeypatch.setattr(app, "export_excel", export)
    assert app.main([str(simple_plan), "--outputs", "png", "--image-engine", "excel", "--excel-timeout", ".25",
                     "--output-dir", str(tmp_path / "com")]) == 0