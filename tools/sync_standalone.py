"""Copie binaire de la source de vérité ; --check ne modifie rien."""
import argparse
from pathlib import Path
import shutil


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Retourner 1 si la copie diffère")
    args = parser.parse_args(argv)
    root = Path(__file__).resolve().parents[1]
    source, target = root / "src/dim2dir/core.py", root / "Dim2Dir.py"
    if args.check:
        if not target.is_file() or source.read_bytes() != target.read_bytes():
            print("Dim2Dir.py diffère de src/dim2dir/core.py : lancer tools/sync_standalone.py")
            return 1
        print("Dim2Dir.py est synchronisé.")
        return 0
    shutil.copyfile(source, target)
    print("Dim2Dir.py synchronisé depuis src/dim2dir/core.py.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
