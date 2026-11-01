"""Client Excel ; les chemins ne dépendent pas du dossier courant."""
from pathlib import Path
import dim2dir

HERE = Path(__file__).resolve().parent

if __name__ == "__main__":
    result = dim2dir.visualize(
        HERE / "../tournage_L9.xlsx",
        output_dir=HERE / "resultats_excel",
        outputs=("xlsx", "png"),
    )
    print(result.report, end="")
