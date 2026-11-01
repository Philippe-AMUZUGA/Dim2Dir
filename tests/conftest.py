import importlib.util
import os
from pathlib import Path
import sys

import pytest
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def import_file(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="session")
def examples(tmp_path_factory):
    generator = import_file("generate_examples", ROOT / "examples" / "generate_examples.py")
    return {path.stem: path for path in generator.generate_examples(tmp_path_factory.mktemp("examples"))}


@pytest.fixture
def simple_plan(tmp_path):
    path = tmp_path / "source.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    ws.append(["Run", "A", "B", "C"])
    ws.append([1, 10, "bas", 5])
    ws.append([2, 20, "haut", 8])
    ws.append([3, 10, "bas", 5])
    wb.save(path)
    wb.close()
    return path
