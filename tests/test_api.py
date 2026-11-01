from dataclasses import FrozenInstanceError
from datetime import date, datetime
import hashlib
import json
import os
import subprocess
import sys

import pytest
from openpyxl import load_workbook

import dim2dir
import dim2dir.core as app
from conftest import ROOT

ROWS = [{"Run": 1, "A": 10, "B": "bas", "C": 5},
        {"Run": 2, "A": 20, "B": "haut", "C": 8},
        {"Run": 3, "A": 10, "B": "bas", "C": 5}]


def test_input_forms(simple_plan, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    before = set(tmp_path.iterdir())
    expected = dim2dir.project(simple_plan)
    columns = {key: [row[key] for row in ROWS] for key in ROWS[0]}
    for data in (str(simple_plan), ROWS, columns):
        p = dim2dir.project(data)
        assert isinstance(p, dim2dir.Projection)
        assert p == expected
        assert p.collisions == (("E1", "E3"),)
        assert p.occupied == 2
    assert set(tmp_path.iterdir()) == before
    with pytest.raises(FrozenInstanceError):
        expected.width = 7
    assert expected.layout == app.project_fractal_2d(list(expected.headers), ROWS)
    with pytest.raises(TypeError):
        expected.layout["placements"].append((0, 0, "new", 4))
    with pytest.raises(TypeError):
        expected.layout["width"] = 7
    with pytest.raises(TypeError):
        expected.layout["x_axes"][0]["name"] = "changed"


def test_dataframe_forms(simple_plan):
    pd = pytest.importorskip("pandas")
    assert dim2dir.project(pd.DataFrame(ROWS)) == dim2dir.project(simple_plan)


def test_excel_dataframe_equivalence_on_every_example(examples):
    pd = pytest.importorskip("pandas")
    for path in examples.values():
        excel = dim2dir.project(path)
        frame = dim2dir.project(pd.read_excel(path))
        assert frame.placements == excel.placements, path.name
        assert frame.layout == excel.layout, path.name
        assert (frame.width, frame.height, frame.collisions) == (excel.width, excel.height, excel.collisions)
        assert frame.warnings == excel.warnings


def test_normalization_and_missing_values():
    pd = pytest.importorskip("pandas")
    np = pytest.importorskip("numpy")
    df = pd.DataFrame({" Run ": [np.int64(11), np.int64(12), np.int64(13), np.int64(14)],
                       " A ": [np.float64(2), np.nan, pd.NA, None],
                       7: [True, 1, np.bool_(False), np.int64(0)],
                       "Date": pd.Series([date(2026, 1, 1), pd.NaT, datetime(2026, 1, 2), None], dtype=object)})
    p = dim2dir.project(df)
    assert p.headers == ("Run", "A", "7", "Date")
    assert p.n_runs == 4
    assert p.placements[0][2] == "E11"
    assert len(p.layout["y_axes"][0]["levels"]) == 4
    assert any("'A' : 3 cellule(s)" in warning for warning in p.warnings)
    assert any("'Date' : 2 cellule(s)" in warning for warning in p.warnings)
    assert p.layout["x_axes"][1]["levels"][0] == date(2026, 1, 1)
    assert dim2dir.project([{"A": np.int64(1)}, {"A": np.float64(2)}]).width == 2


def test_dataframe_index_and_multiindex():
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"A": [1, 2]}, index=pd.Index([40, 50], name="rUN"))
    assert [p[2] for p in dim2dir.project(df).placements] == ["E40", "E50"]
    df["Run"] = [7, 8]
    assert [p[2] for p in dim2dir.project(df).placements] == ["E7", "E8"]
    df.index.name = "other"
    assert dim2dir.project(df).headers == ("A", "Run")
    df.columns = pd.MultiIndex.from_tuples([("A", "a"), ("Run", "b")])
    with pytest.raises(dim2dir.Dim2DirError, match="aplatir"):
        dim2dir.project(df)


def test_duplicate_and_empty_headers():
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame([[1, 2, 3, 4, 5]], columns=["Run", " A ", "A", "A_2", "A"])
    p = dim2dir.project(df)
    assert p.headers == ("Run", "A", "A_3", "A_2", "A_4")
    assert len(p.warnings) == 2
    df = pd.DataFrame([[1, 2, 3, 4]], columns=["Run", "A", "  ", "B"])
    p = dim2dir.project(df)
    assert p.headers == ("Run", "A")
    assert any("colonne C" in warning for warning in p.warnings)


def test_custom_run_and_axes(simple_plan):
    p = dim2dir.project([{"ID": 9, "Run": "r", "A": 3}], run_column="id",
                        x_factors=("A",), y_factors=("Run",))
    assert p.x_factors == ("A",)
    assert p.y_factors == ("Run",)
    assert p.placements == ((0, 0, "E9"),)
    assert dim2dir.project(simple_plan, x_factors=("C", "A"), y_factors=("B",)).x_factors == ("C", "A")
    assert dim2dir.project({"Run": [1]}).placements == ((0, 0, "E1"),)


def test_figure_is_detached_and_restores_state(tmp_path, monkeypatch):
    import matplotlib
    import matplotlib.pyplot as plt
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    plt.close("all")
    backend = matplotlib.get_backend()
    rc = dict(matplotlib.rcParams)
    env = dict(os.environ)
    monkeypatch.chdir(tmp_path)
    fig = dim2dir.figure(ROWS, style=dim2dir.Style(base_color="AA44bb"))
    try:
        assert isinstance(fig, Figure)
        assert isinstance(fig.canvas, FigureCanvasAgg)
        assert fig.axes[0].patches
        assert plt.get_fignums() == []
        assert matplotlib.get_backend() == backend
        assert dict(matplotlib.rcParams) == rc
        assert dict(os.environ) == env
        assert not list(tmp_path.iterdir())
    finally:
        fig.clear()


def test_visualize_files_and_hash(simple_plan, tmp_path, capsys):
    before = hashlib.sha256(simple_plan.read_bytes()).hexdigest()
    out = tmp_path / "out"
    result = dim2dir.visualize(simple_plan, output_dir=out, report=True, dpi=50)
    assert isinstance(result, dim2dir.Result)
    assert set(result.files) == {"xlsx", "pdf", "png", "report"}
    assert all(path.is_absolute() and path.is_file() for path in result.files.values())
    assert result.files["pdf"].read_bytes().startswith(b"%PDF")
    assert result.files["png"].read_bytes().startswith(b"\x89PNG")
    assert result.files["report"].name == "source_rapport.json"
    assert hashlib.sha256(simple_plan.read_bytes()).hexdigest() == before
    assert capsys.readouterr() == ("", "")
    args = app._parser().parse_args([str(simple_plan), "--dry-run"])
    app.process_workbook(args)
    assert capsys.readouterr().out == result.report
    doc = json.loads(result.files["report"].read_text(encoding="utf-8"))
    assert doc["sheets"][0]["status"] == "ok"
    assert not list(out.glob(".dim2dir-*"))


@pytest.mark.parametrize("outputs", [("xlsx",), ("png",), ("png", "pdf", "xlsx")])
@pytest.mark.parametrize("memory", [True, False])
def test_xlsx_policy(simple_plan, tmp_path, outputs, memory):
    data = ROWS if memory else simple_plan
    result = dim2dir.visualize(data, output_dir=tmp_path / "out", outputs=outputs, dpi=50)
    assert set(result.files) == set(outputs)
    if "xlsx" in outputs:
        wb = load_workbook(result.files["xlsx"])
        assert wb.active.sheet_view.showGridLines is False
        wb.close()


def test_output_defaults_and_no_output(simple_plan, tmp_path):
    result = dim2dir.visualize(simple_plan, outputs="xlsx")
    assert result.files["xlsx"] == tmp_path / "source_visu.xlsx"
    assert set(dim2dir.visualize(ROWS, dpi=50).data) == {"png"}
    with pytest.raises(dim2dir.Dim2DirError, match="outputs"):
        dim2dir.visualize(ROWS, outputs=())
    with pytest.raises(dim2dir.Dim2DirError, match="output_dir"):
        dim2dir.visualize(ROWS, in_memory=False)


def test_memory_images_match_excel(simple_plan, tmp_path):
    a = dim2dir.visualize(simple_plan, output_dir=tmp_path / "excel", outputs=("png", "pdf"), dpi=50)
    b = dim2dir.visualize(ROWS, output_dir=tmp_path / "memory", outputs=("png", "pdf"), dpi=50)
    assert a.files["png"].read_bytes() == b.files["png"].read_bytes()
    assert set(b.files) == {"png", "pdf"}


def test_large_grid_is_sparse(examples, tmp_path, monkeypatch):
    def no_dense_allocation(*args, **kwargs):
        pytest.fail("Une grille refusée ne doit pas être allouée")
    monkeypatch.setattr(app, "add_visualization_to_sheet", no_dense_allocation)
    p = dim2dir.project(examples["plan_trop_grand"])
    assert (p.width, p.height) == (1024, 256)
    assert p.warnings
    result = dim2dir.visualize(examples["plan_trop_grand"], output_dir=tmp_path,
                               outputs=("png", "pdf"), report=True)
    assert set(result.files) == {"report"}
    assert "skipped" in result.report
    with pytest.raises(dim2dir.Dim2DirError, match="trop de cellules"):
        dim2dir.figure(examples["plan_trop_grand"])
    boundary = dim2dir.project({"A": list(range(300)), "B": list(range(300))})
    assert boundary.width * boundary.height == 90_000
    assert not boundary.warnings
    # Un axe vertical permet de tester 90001 sans dépasser la limite de colonnes.
    over = dim2dir.project({"A": list(range(90_001))}, x_factors=[], y_factors=["A"], limits=dim2dir.Limits(max_runs=100_000))
    assert over.width * over.height == 90_001
    assert any("trop de cellules" in w for w in over.warnings)
    result = dim2dir.visualize({"A": list(range(90_001))}, x_factors=[], y_factors=["A"],
                               output_dir=tmp_path / "over", limits=dim2dir.Limits(max_runs=100_000))
    assert result.files == {}
    assert "skipped" in result.report


def test_excel_limits_and_skip_note(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "EXCEL_MAX_COLUMNS", 6)
    p = dim2dir.project(ROWS, limits=dim2dir.Limits(max_cells=10**6))
    assert any("limites d'Excel" in w for w in p.warnings)
    result = dim2dir.visualize(ROWS, output_dir=tmp_path, outputs=("xlsx", "png"))
    assert set(result.files) == {"xlsx"}
    wb = load_workbook(result.files["xlsx"])
    assert wb.active["F2"].font.color.rgb[-6:] == "C00000"
    wb.close()


def test_exception_cleans_temporary_files(tmp_path, monkeypatch):
    def fail(ws, bounds, pdf, png, dpi):
        png.write_bytes(b"partial")
        raise RuntimeError("injected")
    monkeypatch.setattr(app, "export_matplotlib", fail)
    with pytest.raises(dim2dir.Dim2DirError, match="injected"):
        dim2dir.visualize(ROWS, output_dir=tmp_path)
    assert not list(tmp_path.iterdir())


def test_auto_fallback_and_internal_excel_workbook(tmp_path, monkeypatch, capsys):
    seen = []
    def unavailable(source, *args):
        assert source.is_file()
        seen.append(source)
        raise RuntimeError("Excel unavailable")
    monkeypatch.setattr(app, "export_excel", unavailable)
    result = dim2dir.visualize(ROWS, output_dir=tmp_path, outputs=("png", "pdf"), engine="auto", dpi=50)
    assert set(result.files) == {"pdf", "png"}
    assert any("Repli sur matplotlib" in w for w in result.warnings)
    assert all(not path.exists() for path in seen)
    assert not list(tmp_path.glob(".dim2dir-*"))
    assert capsys.readouterr() == ("", "")
    with pytest.raises(dim2dir.Dim2DirError, match="unavailable"):
        dim2dir.visualize(ROWS, output_dir=tmp_path, outputs=("png", "pdf"), engine="excel")
    assert not list(tmp_path.glob(".dim2dir-*"))


def test_excel_engine_success(tmp_path, monkeypatch):
    sources = []
    def fake_excel(source, sheet, bounds, pdf, png, timeout):
        sources.append(source)
        wb = load_workbook(source)
        assert wb[sheet].sheet_view.showGridLines is False
        wb.close()
        pdf.write_bytes(b"%PDF fake")
        from PIL import Image
        Image.new("RGB", (10, 10)).save(png)
    monkeypatch.setattr(app, "export_excel", fake_excel)
    result = dim2dir.visualize(ROWS, output_dir=tmp_path, outputs=("png", "pdf"), engine="excel")
    assert set(result.files) == {"pdf", "png"}
    assert all(not path.exists() for path in sources)


@pytest.mark.parametrize("data", [42, object(), [], {}, [1], {"A": [1], "B": [2, 3]},
                                  {"A": "not a list"}, {"": [1]}, {"A": [object()]}])
def test_invalid_data(data):
    with pytest.raises(dim2dir.Dim2DirError):
        dim2dir.project(data)


@pytest.mark.parametrize("options", [
    {"outputs": ()}, {"outputs": ("bad",)}, {"outputs": ("png", "png")}, {"engine": "bad"},
    {"dpi": 49}, {"dpi": True}, {"in_memory": 1}, {"report": "yes"}, {"name": "../x"},
    {"name": ""}, {"style": {}}, {"limits": 0}, {"run_column": ""}, {"sheet_name": 42},
    {"x_factors": ["A"]}, {"x_factors": ["bad"], "y_factors": ["B", "C"]},
    {"x_factors": "A", "y_factors": []},
])
def test_invalid_options(tmp_path, options):
    with pytest.raises(dim2dir.Dim2DirError):
        dim2dir.visualize(ROWS, output_dir=tmp_path / "unwritten", **options)
    assert not (tmp_path / "unwritten").exists()


def test_missing_and_invalid_workbooks(simple_plan, tmp_path):
    with pytest.raises(FileNotFoundError):
        dim2dir.project(tmp_path / "missing.xlsx")
    with pytest.raises(dim2dir.Dim2DirError, match=".xlsx"):
        dim2dir.project(tmp_path / "wrong.xls")
    with pytest.raises(dim2dir.Dim2DirError, match="Feuille introuvable"):
        dim2dir.project(simple_plan, sheet_name="missing")
    bad = tmp_path / "invalid.xlsx"
    bad.write_bytes(b"invalid zip")
    with pytest.raises(dim2dir.Dim2DirError):
        dim2dir.project(bad)


def test_output_cannot_overwrite_input(simple_plan, tmp_path):
    source = tmp_path / "plan_visu.xlsx"
    source.write_bytes(simple_plan.read_bytes())
    before = source.read_bytes()
    with pytest.raises(dim2dir.Dim2DirError, match="différent"):
        dim2dir.visualize(source, name="plan", outputs="xlsx")
    assert source.read_bytes() == before


def test_import_is_lightweight(tmp_path):
    code = (
        "import time, sys; start=time.perf_counter(); import dim2dir;"
        "assert 'pandas' not in sys.modules;"
        "assert 'win32com' not in sys.modules and 'pythoncom' not in sys.modules;"
        "assert 'matplotlib' not in sys.modules;"
        "print(time.perf_counter()-start)"
    )
    result = subprocess.run([sys.executable, "-c", code], cwd=tmp_path,
                            env={**os.environ, "PYTHONPATH": str(ROOT / "src")},
                            capture_output=True, text=True, timeout=30)
    assert result.returncode == 0, result.stderr
    assert float(result.stdout) < 1.5, f"Import trop lent : {result.stdout}s"

def test_numpy_dates_and_literal_text():
    np = pytest.importorskip("numpy")
    p = dim2dir.project([{"A": np.datetime64("2026-01-02"), "B": "=niveau"},
                         {"A": np.datetime64("NaT"), "B": "x"}])
    assert p.layout["x_axes"][0]["levels"][0].year == 2026
    assert p.layout["y_axes"][0]["levels"][0] == "=niveau"
    assert any("cellule(s) vide(s)" in w for w in p.warnings)


def test_duck_dataframe_errors_are_domain_errors():
    class BrokenFrame:
        columns = ["A"]
        def itertuples(self, **kwargs):
            raise RuntimeError("invalid table")
    with pytest.raises(dim2dir.Dim2DirError, match="invalid table"):
        dim2dir.project(BrokenFrame())


def test_case_variant_run_matches_cli(tmp_path):
    from openpyxl import Workbook
    source = tmp_path / "runs.xlsx"
    wb = Workbook()
    wb.active.title = "Feuil1"
    wb.active.append(["A", "Run", "run"])
    wb.active.append([2, 10, 20])
    wb.save(source)
    wb.close()
    p = dim2dir.project(source)
    assert p.x_factors == ("A",)
    assert p.placements == ((0, 0, "E10"),)


def test_failed_xlsx_is_atomic(tmp_path, monkeypatch):
    from openpyxl import Workbook
    output = tmp_path / "plan_visu.xlsx"
    output.write_bytes(b"original")
    def fail(self, path):
        raise OSError("save failed")
    monkeypatch.setattr(Workbook, "save", fail)
    with pytest.raises(dim2dir.Dim2DirError, match="save failed"):
        dim2dir.visualize(ROWS, output_dir=tmp_path, outputs="xlsx")
    assert output.read_bytes() == b"original"
    assert not list(tmp_path.glob(".dim2dir-*"))

def test_output_directory_must_be_a_directory(tmp_path):
    output = tmp_path / "file"
    output.write_text("keep", encoding="utf-8")
    with pytest.raises(dim2dir.Dim2DirError, match="dossier"):
        dim2dir.visualize(ROWS, output_dir=output)
    assert output.read_text(encoding="utf-8") == "keep"