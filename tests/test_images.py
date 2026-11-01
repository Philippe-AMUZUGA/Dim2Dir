import importlib.util
import os
import sys

import pytest
from PIL import Image, ImageChops
from openpyxl import load_workbook, Workbook

import dim2dir.core as app


@pytest.mark.parametrize("name", ["tournage_L9", "soudage_2p4", "boulonnerie_L18", "materiaux_textuels", "plan_avec_collisions"])
def test_matplotlib_exports(examples, tmp_path, name):
    path = examples[name]
    report_path = tmp_path / "rapport.json"
    assert app.main([str(path),"--output-dir",str(tmp_path),"--dpi","50","--report",str(report_path)]) == 0
    import json
    report = json.loads(report_path.read_text(encoding="utf-8"))
    item = report["sheets"][0]
    assert report["schema_version"] == 1
    assert item["image_engine"] == "matplotlib"
    png = tmp_path / f"{name}_visu_Feuil1.png"
    pdf = tmp_path / f"{name}_visu_Feuil1.pdf"
    assert pdf.read_bytes().startswith(b"%PDF")
    wb = load_workbook(tmp_path / f"{name}_visu.xlsx")
    source = load_workbook(path,data_only=True)
    try:
        headers, rows, col = app.read_plan_table(source.active,source.active,[])
        bounds = app.visualization_bounds(col,app.project_fractal_2d(headers,rows))
        widths, heights = app.scene_geometry(wb.active,bounds)
        expected = sum(widths) / sum(heights)
        with Image.open(png) as image:
            image.load()
            rgb = image.convert("RGB")
            assert ImageChops.difference(rgb, Image.new("RGB",image.size,"white")).getbbox()
            assert abs(image.width / image.height / expected - 1) < 0.05
    finally:
        wb.close()
        source.close()


def test_too_large_no_images(examples,tmp_path):
    args = app._parser().parse_args([str(examples["plan_trop_grand"]),"--output-dir",str(tmp_path)])
    report = app.process_workbook(args)
    item = report["sheets"][0]
    assert item["status"] == "skipped"
    assert item["grid"]["cells"] == 262144
    assert not list(tmp_path.glob("*.png")) and not list(tmp_path.glob("*.pdf"))
    wb = load_workbook(tmp_path / "plan_trop_grand_visu.xlsx")
    assert wb.active["L2"].font.color.rgb[-6:] == "C00000"
    wb.close()


def test_pixel_limit(tmp_path):
    ws = Workbook().active
    ws.column_dimensions["A"].width = 10000
    ws.row_dimensions[1].height = 10000
    with pytest.raises(ValueError,match="60 millions"):
        app.export_matplotlib(ws,(1,1,1,1),tmp_path/"no.pdf",tmp_path/"no.png",600)
    assert not list(tmp_path.iterdir())


def test_images_without_xlsx(simple_plan,tmp_path):
    out = tmp_path / "images"
    assert app.main([str(simple_plan),"--no-xlsx","--no-pdf","--dpi","50","--output-dir",str(out)]) == 0
    assert [path.suffix for path in out.iterdir()] == [".png"]


def test_auto_fallback_and_excel_error(simple_plan,tmp_path,monkeypatch,capsys):
    def unavailable(*args,**kwargs):
        raise RuntimeError("Excel synthétiquement indisponible")
    monkeypatch.setattr(app,"export_excel",unavailable)
    base = [str(simple_plan),"--output-dir",str(tmp_path / "out"),"--dpi","50","--no-pdf"]
    assert app.main(base + ["--image-engine","auto"]) == 0
    assert "Repli sur matplotlib" in capsys.readouterr().out
    assert app.main(base + ["--image-engine","excel"]) == 1
    assert "indisponible" in capsys.readouterr().out
    monkeypatch.undo()
    monkeypatch.setattr(app.sys,"platform","linux")
    with pytest.raises(RuntimeError,match="Windows"):
        app.export_excel(simple_plan,"Feuil1",(1,1,2,2))


def _excel_available():
    if sys.platform != "win32" or importlib.util.find_spec("win32com") is None:
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,r"Excel.Application\CLSID"):
            return True
    except OSError:
        return False


@pytest.mark.excel
@pytest.mark.skipif(not _excel_available() or os.environ.get("DIM2DIR_TEST_EXCEL") != "1",
                    reason="Excel/pywin32 requis ; activation explicite DIM2DIR_TEST_EXCEL=1")
def test_excel_export(simple_plan,tmp_path):
    assert app.main([str(simple_plan),"--output-dir",str(tmp_path / "excel"),"--image-engine","excel"]) == 0
    assert (tmp_path / "excel/source_visu_Feuil1.pdf").read_bytes().startswith(b"%PDF")
    with Image.open(tmp_path / "excel/source_visu_Feuil1.png") as image:
        image.load()
        assert image.width > 0
