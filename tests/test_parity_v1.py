from copy import copy
import shutil

import pytest
from openpyxl import load_workbook

import dim2dir.core as v2
from conftest import ROOT, import_file


@pytest.mark.parametrize("name", ["tournage_L9", "soudage_2p4", "boulonnerie_L18", "materiaux_textuels",
                                 "plan_avec_collisions", "plan_trop_grand"])
def test_parity_v1(examples, tmp_path, name):
    v1 = import_file("v1", ROOT / "tests/reference/Dim2Dir_V1.py")
    old, new = tmp_path / "v1.xlsx", tmp_path / "v2.xlsx"
    shutil.copy2(examples[name], old)
    shutil.copy2(examples[name], new)
    old_out = v1.run_visualization(str(old), "Feuil1")
    new_out = v2.run_visualization(str(new), "Feuil1")
    a, b = load_workbook(old_out), load_workbook(new_out)
    try:
        assert a.sheetnames == b.sheetnames
        for sheet_name in a.sheetnames:
            x, y = a[sheet_name], b[sheet_name]
            assert (x.max_row, x.max_column) == (y.max_row, y.max_column)
            assert set(map(str, x.merged_cells.ranges)) == set(map(str, y.merged_cells.ranges))
            assert x.sheet_view.showGridLines == y.sheet_view.showGridLines
            assert set(x.column_dimensions) == set(y.column_dimensions)
            assert set(x.row_dimensions) == set(y.row_dimensions)
            for key in x.column_dimensions:
                assert dict(x.column_dimensions[key]) == dict(y.column_dimensions[key])
            for key in x.row_dimensions:
                assert dict(x.row_dimensions[key]) == dict(y.row_dimensions[key])
            for row in x:
                for cell in row:
                    other = y[cell.coordinate]
                    assert cell.value == other.value, cell.coordinate
                    for attr in ("fill", "font", "border", "alignment", "protection"):
                        assert copy(getattr(cell, attr)) == copy(getattr(other, attr)), (cell.coordinate, attr)
                    assert cell.number_format == other.number_format
    finally:
        a.close()
        b.close()
