"""Vérifie sdist, roue et clients dans un venv vierge (réseau utilisé uniquement par pip)."""
from pathlib import Path
import os
import shutil
import subprocess
import sys
import tarfile
import tempfile
import venv
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]


def run(command, *, cwd, env):
    print("+", " ".join(map(str, command)), flush=True)
    subprocess.run(list(map(str, command)), cwd=cwd, env=env, check=True)


def check_wheel(wheel):
    with ZipFile(wheel) as archive:
        names = archive.namelist()
        assert names, "Roue vide"
        assert all(name.startswith("dim2dir/") or name.split("/")[0].endswith(".dist-info")
                   for name in names), names
        for name in ("__init__.py", "core.py", "__main__.py", "py.typed"):
            assert f"dim2dir/{name}" in names, name
        assert not any(name.startswith(("tests/", "examples/")) or name == "Dim2Dir.py" for name in names)


def check_sdist(sdist):
    with tarfile.open(sdist) as archive:
        names = {name.split("/", 1)[-1] for name in archive.getnames()}
    for name in ("README.md", "LICENSE", "CHANGELOG.md", "Dim2Dir.py",
                 "tests/test_api.py", "tests/reference/Dim2Dir_V1.py",
                 "examples/generate_examples.py", "examples/client_project/depuis_pandas.py",
                 "tools/sync_standalone.py"):
        assert name in names, f"Manquant dans la sdist : {name}"


def main():
    env = {key: value for key, value in os.environ.items() if key not in ("PYTHONPATH", "PYTHONHOME")}
    env.update(PYTHONNOUSERSITE="1", PYTHONIOENCODING="utf-8")
    with tempfile.TemporaryDirectory(prefix="dim2dir-install-") as temporary:
        work = Path(temporary)
        dist = work / "dist"
        run([sys.executable, "-m", "build", "--outdir", dist, ROOT], cwd=work, env=env)
        wheels, sdists = list(dist.glob("*.whl")), list(dist.glob("*.tar.gz"))
        assert len(wheels) == len(sdists) == 1
        wheel, sdist = wheels[0], sdists[0]
        check_wheel(wheel)
        check_sdist(sdist)
        run([sys.executable, "-m", "twine", "check", wheel, sdist], cwd=work, env=env)
        print("OK : construction, contenu des distributions et métadonnées.", flush=True)

        venv_dir = work / "venv"
        venv.EnvBuilder(with_pip=True).create(venv_dir)
        binaries = venv_dir / ("Scripts" if os.name == "nt" else "bin")
        python = binaries / ("python.exe" if os.name == "nt" else "python")
        console = binaries / ("dim2dir.exe" if os.name == "nt" else "dim2dir")
        outside = work / "outside"
        outside.mkdir()
        run([python, "-m", "pip", "install", wheel], cwd=outside, env=env)
        code = (
            "import sys, dim2dir; from pathlib import Path;"
            "assert 'pandas' not in sys.modules;"
            "assert 'pythoncom' not in sys.modules and 'win32com' not in sys.modules;"
            "r=dim2dir.visualize([{'Run':1,'A':1,'B':'bas'},"
            "{'Run':2,'A':2,'B':'haut'}],output_dir='api',dpi=50);"
            "assert set(r.files)=={'png'};"
            "assert all(p.is_absolute() and p.is_file() for p in r.files.values())"
        )
        run([python, "-c", code], cwd=outside, env=env)
        run([console, "--version"], cwd=outside, env=env)
        run([python, "-m", "dim2dir", "--help"], cwd=outside, env=env)
        print("OK : roue sans pandas, API, console et python -m.", flush=True)

        # Retire le paquet tout en conservant ses dépendances : le fichier seul
        # est testé sans aucune installation de dim2dir dans cet interpréteur.
        run([python, "-m", "pip", "uninstall", "-y", "dim2dir"], cwd=outside, env=env)
        standalone = outside / "standalone"
        standalone.mkdir()
        shutil.copyfile(ROOT / "Dim2Dir.py", standalone / "Dim2Dir.py")
        run([python, "-c",
             "import importlib.util; assert importlib.util.find_spec('dim2dir') is None;"
             "from openpyxl import Workbook;"
             "w=Workbook(); w.active.title='Feuil1'; w.active.append(['Run','A','B']);"
             "w.active.append([1,1,'bas']); w.active.append([2,2,'haut']);"
             "w.save('source.xlsx'); w.close();"
             "import Dim2Dir; assert Dim2Dir.project('source.xlsx').n_runs==2"],
            cwd=standalone, env=env)
        run([python, standalone / "Dim2Dir.py", standalone / "source.xlsx",
             "--output-dir", standalone / "out", "--dpi", "50"], cwd=outside, env=env)
        assert (standalone / "out/source_visu_Feuil1.png").is_file()
        assert (standalone / "out/source_visu_Feuil1.pdf").is_file()
        assert (standalone / "out/source_visu.xlsx").is_file()
        print("OK : fichier autonome sans paquet installé.", flush=True)

        run([python, "-m", "pip", "install", str(wheel) + "[pandas]"], cwd=outside, env=env)
        copied_examples = work / "examples"
        shutil.copytree(ROOT / "examples", copied_examples,
                        ignore=shutil.ignore_patterns("__pycache__", "resultats_*"))
        # Génère les entrées dans la copie, jamais dans le dépôt.
        run([python, copied_examples / "generate_examples.py"], cwd=outside, env=env)
        clients = copied_examples / "client_project"
        for name in ("depuis_pandas.py", "depuis_excel.py"):
            run([python, clients / name], cwd=outside, env=env)
        code = (
            "import dim2dir; from depuis_pandas import welding_plan;"
            "r=dim2dir.visualize(welding_plan(), style=dim2dir.Style(palette='viridis',font_size=11),"
            "outputs=('png','svg'),in_memory=True);"
            "assert r.files=={} and list(r.data)==['png','svg'];"
            "assert r.data['png'].startswith(bytes.fromhex('89504e47'));"
            "assert b'<svg' in r.data['svg'];"
            "assert r.engine=='matplotlib'"
        )
        run([python, "-c", code], cwd=clients, env=env)
        for name in ("plan_visu_Feuil1.png", "plan_visu_Feuil1.pdf"):
            assert (clients / "resultats_pandas" / name).is_file(), name
        assert (clients / "resultats_excel/tournage_L9_visu_Feuil1.png").is_file()
        assert (clients / "resultats_excel/tournage_L9_visu.xlsx").is_file()
        print("OK : extra pandas et deux clients hors dépôt, sans PYTHONPATH.", flush=True)
    print("Vérification d'installation réussie ; aucun environnement temporaire conservé.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (subprocess.CalledProcessError, AssertionError, OSError) as exc:
        print(f"ÉCHEC de vérification : {exc}", file=sys.stderr)
        raise SystemExit(1)