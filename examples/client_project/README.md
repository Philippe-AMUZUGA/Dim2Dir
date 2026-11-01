# Utiliser Dim2Dir dans son projet

Installer Python 3.10+ puis `python -m pip install "dim2dir[pandas]"`.

- `depuis_excel.py` lit `../tournage_L9.xlsx` et produit un classeur et un PNG dans `resultats_excel`.
- `depuis_pandas.py` construit un plan factoriel 2³ de soudage. Son `Style` utilise la palette cividis, la police DejaVu Sans, une taille de police de 11 points, des cellules de 44 pixels, un titre et une légende. Il écrit PNG et PDF dans `resultats_pandas`, puis génère un PNG en mémoire dans `Result.data["png"]`. Il calcule aussi la couverture et construit une figure pour un notebook sans l'afficher.

Les chemins sont calculés depuis chaque script : on peut les lancer depuis n'importe quel dossier.
Pour créer les six classeurs d'exemple si nécessaire : `python ../generate_examples.py`.
Les sorties existantes sont remplacées atomiquement ; la source reste intacte.

`visualize(df)` retourne par défaut des octets PNG sans fichier. Utiliser `outputs=("xlsx", "png", "pdf", "svg")` pour choisir les formats ; `output_dir` active l'écriture de fichiers, sauf si `in_memory=True` est explicite. Le même style s'applique à toutes les sorties car les images lisent la feuille écrite.
