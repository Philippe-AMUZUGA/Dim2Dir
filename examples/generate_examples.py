"""Plans de génie mécanique synthétiques, reproductibles (openpyxl seulement)."""
from itertools import product
from pathlib import Path
import argparse

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# Tables orthogonales standard : indices des niveaux à partir de 1.
L9 = [
    [1, 1, 1, 1], [1, 2, 2, 2], [1, 3, 3, 3],
    [2, 1, 2, 3], [2, 2, 3, 1], [2, 3, 1, 2],
    [3, 1, 3, 2], [3, 2, 1, 3], [3, 3, 2, 1],
]
L18 = [
    [1,1,1,1,1,1,1,1], [1,1,2,2,2,2,2,2], [1,1,3,3,3,3,3,3],
    [1,2,1,1,2,2,3,3], [1,2,2,2,3,3,1,1], [1,2,3,3,1,1,2,2],
    [1,3,1,2,1,3,2,3], [1,3,2,3,2,1,3,1], [1,3,3,1,3,2,1,2],
    [2,1,1,3,3,2,2,1], [2,1,2,1,1,3,3,2], [2,1,3,2,2,1,1,3],
    [2,2,1,2,3,1,3,2], [2,2,2,3,1,2,1,3], [2,2,3,1,2,3,2,1],
    [2,3,1,3,2,3,1,2], [2,3,2,1,3,1,2,3], [2,3,3,2,1,2,3,1],
]


def generate_examples(output_dir=None):
    target = Path(output_dir) if output_dir is not None else Path(__file__).resolve().parent
    target.mkdir(parents=True, exist_ok=True)
    paths = []

    def save(name, headers, records):
        wb = Workbook()
        ws = wb.active
        ws.title = "Feuil1"
        ws.append(["Run", *headers])
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="335C81")
            cell.font = Font(color="FFFFFF", bold=True)
        for run, values in enumerate(records, 1):
            ws.append([run, *values])
        path = target / name
        wb.save(path)
        wb.close()
        paths.append(path)

    def decode(table, levels):
        return [[levels[col][level - 1] for col, level in enumerate(row)] for row in table]

    save("tournage_L9.xlsx", ["Vc m/min", "f mm/tr", "ap mm", "Lubrification"],
         decode(L9, [[120,180,240], [0.1,0.2,0.3], [0.5,1.0,1.5], ["sec","émulsion","huile"]]))
    save("soudage_2p4.xlsx", ["I A", "Avance mm/s", "Gaz L/min", "Tension V"],
         product([140,180], [4,6], [12,16], [22,26]))
    save("boulonnerie_L18.xlsx", ["Classe qualité", "Couple N.m", "Lubrifiant", "État de surface",
                                  "Filet mm", "Précharge %", "Fréquence Hz", "R contrainte"],
         decode(L18, [[8.8,10.9], [40,60,80], ["sec","huile","MoS2"], ["brut","zingué","phosphaté"],
                      [16,24,32], [50,65,80], [5,10,20], [0.1,0.3,0.5]]))
    save("materiaux_textuels.xlsx", ["Matériau", "Traitement thermique", "Épaisseur mm"],
         product(["acier S355","aluminium 6061","titane TA6V"], ["aucun","trempe","revenu"], [2,4,6]))
    # Neuf facteurs : 4**8 = 65536 serait inférieur au plafond de 90000.
    save("plan_trop_grand.xlsx", [f"Facteur {i}" for i in range(1,10)], [[level] * 9 for level in range(1,5)])
    save("plan_avec_collisions.xlsx", ["Vc m/min", "f mm/tr", "Lubrification"],
         [[120,0.1,"sec"], [120,0.1,"sec"], [180,0.2,"huile"],
          [240,0.3,"émulsion"], [None,0.2,"sec"], [180,0.1,"huile"]])
    return paths


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Génère six plans mécaniques synthétiques.")
    parser.add_argument("output_dir", nargs="?", help="Dossier cible (défaut : examples)")
    for path in generate_examples(parser.parse_args().output_dir):
        print(path)
