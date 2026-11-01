from io import BytesIO
from unittest.mock import Mock
import math

import pytest
from openpyxl import Workbook

from dim2dir import Limits, Style, Dim2DirError, project, figure, visualize
import dim2dir.core as app

ROWS = [{"A": 1, "B": 1}, {"A": 2, "B": 2}]


@pytest.mark.parametrize("kwargs", [
    {"max_cells": 0}, {"max_cells": True}, {"max_cells": 1.5}, {"max_cells": 1_048_576 * 16_384 + 1},
    {"max_pixels": -1}, {"max_pixels": False}, {"max_runs": 0}, {"max_runs": 2.0},
    {"excel_timeout_s": 0}, {"excel_timeout_s": -1}, {"excel_timeout_s": True},
    {"excel_timeout_s": "120"}, {"excel_timeout_s": math.inf}, {"excel_timeout_s": math.nan},
])
def test_invalid_limits(kwargs):
    with pytest.raises(Dim2DirError):
        Limits(**kwargs)


def test_hard_cell_bound_is_inclusive():
    assert Limits(max_cells=1_048_576 * 16_384).max_cells == 17_179_869_184


def test_exact_grid_boundary_before_dense_allocation(monkeypatch):
    dense = Mock(side_effect=AssertionError("allocation dense interdite"))
    monkeypatch.setattr(app, "_scene_rectangles", dense)
    accepted = project({"A": list(range(300)), "B": list(range(300))})
    assert accepted.width * accepted.height == 90_000 and not accepted.warnings
    data = {"A": list(range(90_001))}
    limits = Limits(max_runs=90_001)
    with app._prepared_plan(data, [], ["A"], "Run", None, Style(), limits) as prepared:
        _, ws, projection, item, _, _ = prepared
        assert projection.width * projection.height == 90_001
        assert item["status"] == "skipped"
        assert any("90001 cellules, limite 90000" in warning and "max_cells" in warning for warning in projection.warnings)
        before = len(ws._cells)
        app.add_visualization_to_sheet(ws, limits=limits)
        assert len(ws._cells) == before + 1  # Uniquement la note rouge, aucune grille.
    dense.assert_not_called()
    assert not project(data, x_factors=[], y_factors=["A"], limits=Limits(max_runs=90_001, max_cells=90_001)).warnings


@pytest.mark.parametrize("source", ["records", "columns", "dataframe", "xlsx"])
def test_max_runs_before_memory_conversion(source, simple_plan, monkeypatch):
    if source == "records":
        data = [{"A": 1}, {"A": 2}, {"A": 3}]
    elif source == "columns":
        data = {"A": [1, 2, 3]}
    elif source == "dataframe":
        data = pytest.importorskip("pandas").DataFrame({"A": [1, 2, 3]})
        monkeypatch.setattr(data, "itertuples", Mock(side_effect=AssertionError("conversion prématurée")))
    else:
        data = simple_plan
    if source != "xlsx":
        monkeypatch.setattr(app, "Workbook", Mock(side_effect=AssertionError("classeur prématuré")))
    with pytest.raises(Dim2DirError, match="3 essais.*limite 2.*max_runs.*--max-runs"):
        project(data, limits=Limits(max_runs=2))


def test_dataframe_protocol_without_len_is_checked_during_iteration(monkeypatch):
    class Frame:
        columns = ["A"]
        def itertuples(self, **kwargs):
            yield (1,)
            yield (2,)
    monkeypatch.setattr(app, "Workbook", Mock(side_effect=AssertionError("classeur prématuré")))
    with pytest.raises(Dim2DirError, match="2 essais.*limite 1"):
        project(Frame(), limits=Limits(max_runs=1))


@pytest.mark.parametrize("dpi", [49, 601, 50.0, True])
def test_dpi_rejected(dpi):
    with pytest.raises(Dim2DirError, match="dpi.*50.*600"):
        visualize(ROWS, dpi=dpi)


@pytest.mark.parametrize("dpi", [50, 600])
def test_dpi_endpoints(dpi):
    assert visualize([{"A": 1}], dpi=dpi).data["png"].startswith(b"\x89PNG")


def test_pixel_guard_before_figure_allocation(monkeypatch, tmp_path):
    from matplotlib.figure import Figure
    constructor = Mock(side_effect=AssertionError("figure prématurée"))
    monkeypatch.setattr(Figure, "__init__", constructor)
    result = visualize(ROWS, outputs="png", limits=Limits(max_pixels=1), report=True)
    assert set(result.data) == {"report"} and not result.files
    assert "skipped" in result.report
    assert any("pixels, limite 1" in warning and "max_pixels" in warning for warning in result.warnings)
    with pytest.raises(Dim2DirError, match="max_pixels"):
        figure(ROWS, limits=Limits(max_pixels=1))
    constructor.assert_not_called()


def test_pixel_limit_can_be_raised():
    result = visualize(ROWS, outputs="png", limits=Limits(max_pixels=10_000_000), dpi=50)
    assert result.data["png"] and not result.warnings


def test_timeout_is_forwarded_and_cleanup(tmp_path, monkeypatch):
    sources = []
    def export(source, sheet, bounds, pdf, png, timeout):
        sources.append(source)
        assert timeout == .125
        raise RuntimeError(f"Délai maximal de {timeout} s dépassé ; relever excel_timeout_s / --excel-timeout.")
    monkeypatch.setattr(app, "export_excel", export)
    with pytest.raises(Dim2DirError, match="0.125.*excel_timeout_s"):
        visualize(ROWS, engine="excel", limits=Limits(excel_timeout_s=.125))
    assert all(not p.parent.exists() for p in sources)


@pytest.mark.parametrize("name", ["", "a" * 32, "a/b", "a\\b", "a*b", "a?b", "a:b", "a[b", "a]b"])
def test_excel_sheet_name_constraints(name):
    with pytest.raises(Dim2DirError, match="31 caractères"):
        project(ROWS, sheet_name=name)


def test_excel_dimensions_always_apply(monkeypatch):
    monkeypatch.setattr(app, "EXCEL_MAX_COLUMNS", 4)
    p = project(ROWS, limits=Limits(max_cells=1_000_000))
    assert any("limites d'Excel" in warning and "4" in warning for warning in p.warnings)
    with pytest.raises(Dim2DirError, match="colonnes.*limite fixe Excel 4"):
        project([{str(i): i for i in range(5)}])
    monkeypatch.setattr(app, "EXCEL_MAX_ROWS", 2)
    with pytest.raises(Dim2DirError, match="3 lignes.*limite fixe Excel 2"):
        project(ROWS)


def test_actual_excel_raster_pixels_are_checked(tmp_path, monkeypatch):
    from PIL import Image
    sources = []
    def export(source, sheet, bounds, pdf, png, timeout):
        sources.append(source)
        Image.new("RGB", (1000, 1000)).save(png)
    monkeypatch.setattr(app, "export_excel", export)
    result = visualize([{"A": 1}], outputs="png", engine="excel", in_memory=True,
                       limits=Limits(max_pixels=100_000), dpi=50)
    assert not result.files and not result.data and result.engine is None
    assert "skipped" in result.report
    assert any("1,000,000 pixels, limite 100,000" in w for w in result.warnings)
    assert all(not p.parent.exists() for p in sources)



def test_direct_writer_checks_runs_before_projection(monkeypatch):
    ws = Workbook().active
    ws.append(["A"])
    ws.append([1])
    ws.append([2])
    projection = Mock(side_effect=AssertionError("projection prématurée"))
    monkeypatch.setattr(app, "project_fractal_2d", projection)
    with pytest.raises(Dim2DirError, match="2 essais.*limite 1"):
        app.add_visualization_to_sheet(ws, limits=Limits(max_runs=1))
    projection.assert_not_called()
