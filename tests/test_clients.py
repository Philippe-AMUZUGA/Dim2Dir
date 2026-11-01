import os
from pathlib import Path
import shutil
import subprocess
import sys

import pytest

from conftest import ROOT, import_file


def test_client_style_and_memory_api():
    pytest.importorskip("pandas")
    client = import_file("client_pandas", ROOT / "examples/client_project/depuis_pandas.py")
    import dim2dir
    frame = client.welding_plan()
    style = client.personalized_style()
    assert style.palette == "cividis" and style.legend and style.cell_size == 44
    result = dim2dir.visualize(frame, style=style, outputs="png", in_memory=True, dpi=50)
    assert result.data["png"].startswith(b"\x89PNG") and not result.files
    assert result.style == style


@pytest.mark.parametrize("client", ["depuis_excel.py", "depuis_pandas.py"])
def test_client_scripts_from_another_directory(client, examples, tmp_path):
    if "pandas" in client:
        pytest.importorskip("pandas")
    folder = tmp_path / "examples" / "client_project"
    shutil.copytree(ROOT / "examples/client_project", folder)
    shutil.copyfile(examples["tournage_L9"], folder.parent / "tournage_L9.xlsx")
    outside = tmp_path / "outside"
    outside.mkdir()
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src"), "PYTHONIOENCODING": "utf-8"}
    result = subprocess.run([sys.executable, str(folder / client)], cwd=outside, env=env,
                            capture_output=True, text=True, encoding="utf-8", timeout=120)
    assert result.returncode == 0, result.stderr
    if "pandas" in client:
        assert "PNG en mémoire :" in result.stdout
        assert {p.suffix for p in (folder / "resultats_pandas").iterdir()} == {".png", ".pdf"}
    else:
        assert {p.suffix for p in (folder / "resultats_excel").iterdir()} == {".png", ".xlsx"}
    assert not list(outside.iterdir())