import os
import subprocess
import sys
from conftest import ROOT


def test_standalone_is_exact_copy():
    assert (ROOT / "Dim2Dir.py").read_bytes() == (ROOT / "src/dim2dir/core.py").read_bytes()


def test_sync_check():
    result = subprocess.run([sys.executable, str(ROOT / "tools/sync_standalone.py"), "--check"],
                            capture_output=True, text=True, timeout=30)
    assert result.returncode == 0, result.stderr


def test_module_help_and_version(tmp_path):
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src"), "PYTHONIOENCODING": "utf-8"}
    for flag, expected in (("--help", "--all-sheets"), ("--version", "2.1.0")):
        result = subprocess.run([sys.executable, "-m", "dim2dir", flag], cwd=tmp_path,
                                env=env, capture_output=True, text=True, timeout=30)
        assert result.returncode == 0, result.stderr
        assert expected in result.stdout
    assert not list(tmp_path.iterdir())


def test_module_visualization(simple_plan, tmp_path):
    env = {**os.environ, "PYTHONPATH": str(ROOT / "src")}
    result = subprocess.run([sys.executable, "-m", "dim2dir", str(simple_plan),
                             "--output-dir", str(tmp_path / "module"), "--no-pdf", "--dpi", "50"],
                            env=env, capture_output=True, text=True, timeout=120)
    assert result.returncode == 0, result.stderr
    assert (tmp_path / "module/source_visu.xlsx").is_file()
    assert (tmp_path / "module/source_visu_Feuil1.png").is_file()
