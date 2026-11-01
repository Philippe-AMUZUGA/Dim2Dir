"""Plan factoriel 2³ de soudage ; aucune figure n'est affichée par le script."""
from itertools import product
from pathlib import Path
import pandas as pd
import dim2dir

HERE = Path(__file__).resolve().parent


def welding_plan():
    df = pd.DataFrame(
        product([140, 180], [4, 6], [12, 16]),
        columns=["Courant A", "Avance mm/s", "Débit gaz L/min"],
    )
    df.insert(0, "Run", range(1, len(df) + 1))
    return df


def personalized_style():
    return dim2dir.Style(
        palette="cividis",  # Palette adaptée aux déficiences de vision des couleurs.
        font_family="DejaVu Sans",
        font_size=11,
        cell_size=44,
        title="Plan de soudage : courant, avance et gaz",
        legend=True,
    )


if __name__ == "__main__":
    df = welding_plan()
    style = personalized_style()
    projection = dim2dir.project(df, style=style)
    print(f"{projection.occupied}/{projection.width * projection.height} cellules occupées")
    result = dim2dir.visualize(df, style=style, outputs=("png", "pdf"), output_dir=HERE / "resultats_pandas")
    print(result.report, end="")
    memory = dim2dir.visualize(df, style=style, outputs="png", in_memory=True)
    png_bytes = memory.data["png"]  # À transmettre à une application, sans fichier intermédiaire.
    print(f"PNG en mémoire : {len(png_bytes)} octets")
    fig = dim2dir.figure(df, style=style)  # Dans un notebook : placer fig en dernière expression.
    fig.clear()  # Le script ne l'affiche pas ; l'appelant gère sa durée de vie.
