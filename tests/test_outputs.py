from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from itertools import combinations
import hashlib
import json
from pathlib import Path
import threading
import time

import pytest
from openpyxl import load_workbook
from PIL import Image

from dim2dir import Dim2DirError, visualize
import dim2dir.core as app

ROWS = [{"Run": 1, "A": 10, "B": "bas", "C": 5},
        {"Run": 2, "A": 20, "B": "haut", "C": 8},
        {"Run": 3, "A": 10, "B": "bas", "C": 5}]
FORMATS = [combo for n in range(1, 5) for combo in combinations(("svg", "png", "xlsx", "pdf"), n)]


def decode(kind, value):
    assert isinstance(value, bytes) and value
    if kind == "png":
        with Image.open(BytesIO(value)) as image:
            image.load()
            assert image.format == "PNG" and image.width > 0
    elif kind == "pdf":
        assert value.startswith(b"%PDF")
    elif kind == "svg":
        assert value.startswith((b"<?xml", b"<svg")) and b"<svg" in value
    elif kind == "xlsx":
        wb = load_workbook(BytesIO(value))
        assert wb.active.sheet_view.showGridLines is False
        wb.close()
    else:
        assert json.loads(value)["schema_version"] == 1


@pytest.fixture(params=["xlsx", "records", "pandas"])
def source(request, simple_plan):
    if request.param == "xlsx":
        return simple_plan
    if request.param == "pandas":
        return pytest.importorskip("pandas").DataFrame(ROWS)
    return ROWS


@pytest.mark.parametrize("outputs", FORMATS)
@pytest.mark.parametrize("memory", [True, False, None])
def test_formats_destinations_sources(source, outputs, memory, tmp_path):
    before = hashlib.sha256(source.read_bytes()).digest() if isinstance(source, Path) else None
    folder = tmp_path / "out"
    result = visualize(source, output_dir=folder, outputs=outputs, in_memory=memory, report=True, dpi=50)
    expected = (*outputs, "report")
    if memory:
        assert list(result.data) == list(expected)
        assert result.files == {} and not folder.exists()
        for kind, value in result.data.items():
            decode(kind, value)
    else:
        assert list(result.files) == list(expected) and result.data == {}
        assert len(list(folder.iterdir())) == len(expected)
        for kind, path in result.files.items():
            assert path.is_absolute()
            decode(kind, path.read_bytes())
    if before:
        assert hashlib.sha256(source.read_bytes()).digest() == before
    assert result.engine == ("matplotlib" if any(k != "xlsx" for k in outputs) else None)


def test_automatic_destinations_and_defaults(simple_plan, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    result = visualize(ROWS, dpi=50)
    assert list(result.data) == ["png"] and result.files == {}
    result = visualize(simple_plan, dpi=50)
    assert list(result.files) == ["xlsx", "png", "pdf"]
    assert all(path.parent == simple_plan.parent for path in result.files.values())
    result = visualize(simple_plan, in_memory=True, dpi=50)
    assert list(result.data) == ["xlsx", "png", "pdf"] and not result.files
    assert list(visualize(ROWS, outputs="svg").data) == ["svg"]


@pytest.mark.parametrize("outputs", [(), [], "bad", ("png", "bad"), ("png", "png"), {"png"}, ([],)])
def test_invalid_outputs(outputs):
    with pytest.raises(Dim2DirError, match="outputs"):
        visualize(ROWS, outputs=outputs)


def test_svg_excel_and_missing_destination(tmp_path):
    with pytest.raises(Dim2DirError, match="SVG.*matplotlib"):
        visualize(ROWS, outputs="svg", engine="excel")
    with pytest.raises(Dim2DirError, match="output_dir"):
        visualize(ROWS, in_memory=False)
    result = visualize(ROWS, outputs="svg", engine="auto")
    assert result.engine == "matplotlib"
    assert "SVG" in result.warnings[0]


def test_skipped_no_phantom_output(examples, tmp_path):
    result = visualize(examples["plan_trop_grand"], outputs=("png", "pdf", "svg"), output_dir=tmp_path / "out")
    assert not result.files and not result.data
    assert "skipped" in result.report
    assert any("aucune sortie" in warning for warning in result.warnings)
    assert not list((tmp_path / "out").iterdir())
    result = visualize(examples["plan_trop_grand"], outputs="png", in_memory=True, report=True)
    assert set(result.data) == {"report"} and result.files == {}
    assert json.loads(result.data["report"])["sheets"][0]["status"] == "skipped"


@pytest.mark.parametrize("fail", [False, True])
@pytest.mark.parametrize("memory", [False, True])
def test_com_internal_directory_cleanup(tmp_path, monkeypatch, fail, memory):
    sources = []
    def export(source, sheet, bounds, pdf, png, timeout):
        assert source.is_file() and timeout == 120
        sources.append(source)
        wb = load_workbook(source)
        try:
            assert wb[sheet]["G4"].value == "E1\nE3"
            if png:
                app.export_matplotlib(wb[sheet], bounds, None, png, 50)
            if fail:
                raise RuntimeError("exception injectée")
            if pdf:
                app.export_matplotlib(wb[sheet], bounds, pdf, None, 50)
        finally:
            wb.close()
    monkeypatch.setattr(app, "export_excel", export)
    kwargs = dict(engine="excel", in_memory=memory, outputs=("png", "pdf"), output_dir=tmp_path / "out", dpi=50)
    if fail:
        with pytest.raises(Dim2DirError, match="injectée"):
            visualize(ROWS, **kwargs)
    else:
        result = visualize(ROWS, **kwargs)
        assert result.engine == "excel"
        for kind in ("png", "pdf"):
            decode(kind, result.data[kind] if memory else result.files[kind].read_bytes())
    assert sources and all(not path.parent.exists() for path in sources)
    assert not list(tmp_path.rglob(".dim2dir-*"))
    if memory:
        assert not (tmp_path / "out").exists()


def test_matplotlib_memory_has_no_output_temporary_directory(tmp_path, monkeypatch):
    def unexpected(*args, **kwargs):
        pytest.fail("matplotlib en mémoire ne requiert aucun dossier temporaire de sortie")
    monkeypatch.setattr(app.tempfile, "TemporaryDirectory", unexpected)
    result = visualize(ROWS, outputs=("png", "pdf", "svg", "xlsx"), in_memory=True, dpi=50)
    assert not result.files
    for kind, value in result.data.items():
        decode(kind, value)


def test_com_calls_are_serialized(monkeypatch, tmp_path):
    active = maximum = 0
    guard = threading.Lock()
    def export(*args):
        nonlocal active, maximum
        with guard:
            active += 1
            maximum = max(maximum, active)
        time.sleep(.02)
        with guard:
            active -= 1
    monkeypatch.setattr(app.sys, "platform", "win32")
    monkeypatch.setattr(app, "_export_excel_process", export)
    with ThreadPoolExecutor(max_workers=4) as pool:
        list(pool.map(lambda _: app.export_excel(tmp_path / "source.xlsx", "Feuil1", (1, 1, 2, 2)), range(4)))
    assert maximum == 1


def test_permission_error_explains_open_workbook(tmp_path, monkeypatch):
    def blocked(*args):
        raise PermissionError("verrouillé")
    monkeypatch.setattr(app.os, "replace", blocked)
    with pytest.raises(Dim2DirError, match="ouvert dans une autre application"):
        visualize(ROWS, outputs="xlsx", output_dir=tmp_path)
    assert not list(tmp_path.iterdir())