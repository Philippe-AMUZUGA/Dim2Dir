# Dim2Dir 2.1.0

Dim2Dir ajoute une visualisation 2D à un plan d'expériences Excel. Les niveaux des facteurs sont emboîtés sur chaque axe ; par défaut les facteurs impairs sont sur x et les pairs sur y. Le premier facteur de chaque axe varie le plus vite. La couleur suit l'ordre des essais et les essais superposés sont réunis dans la même cellule.

La couverture est calculée sur les niveaux observés dans le tableau, sans jugement d'optimalité statistique. Une cellule vide de la grille représente une combinaison de niveaux observés qui n'a pas été essayée ; les collisions signalent plusieurs essais au même emplacement.

Dim2Dir est la brique de visualisation du projet PLEX2 ; il s'utilise aussi seul, sur n'importe quel plan d'expériences Excel. Il accompagne [PLEX2 DOE Explorer](https://github.com/Philippe-AMUZUGA/plex2-doe-explorer). Licence GNU AGPLv3, voir `LICENSE`.

Dans ce document, « version 1 » désigne la première version du programme, sans paquet ni API Python ; son rendu Excel est le rendu historique, conservé à l'identique par défaut.

## Installation

Python 3.10 ou ultérieur est requis. Depuis le dossier du programme :

Windows :

```console
python -m venv .runtime
.runtime\Scripts\python -m pip install -r requirements.txt
```

Linux et macOS :

```console
python3 -m venv .runtime
source .runtime/bin/activate
pip install -r requirements.txt
```

Sous Windows, `Lancer_Dim2Dir.bat` prépare automatiquement `.runtime` au premier lancement. On peut glisser un fichier `.xlsx` sur ce lanceur. Les arguments supplémentaires sont transmis au programme. Sans argument, le programme lit `./Dim2Dir.xlsx`, feuille `Feuil1` si elle existe, sinon la première feuille.

Les dépendances principales sont openpyxl, matplotlib et Pillow. Le moteur Excel optionnel nécessite Windows, Excel de bureau et :

```console
.runtime\Scripts\python -m pip install -r requirements-excel.txt
```

## Deux voies d'usage

### A. Le fichier autonome

Télécharger uniquement `Dim2Dir.py`, puis installer ses dépendances :

~~~console
python -m pip install "openpyxl>=3.1" "matplotlib>=3.7" "pillow>=10"
python Dim2Dir.py mon_plan.xlsx
~~~

On peut également placer ce fichier à côté de son programme :

~~~python
import Dim2Dir
projection = Dim2Dir.project("mon_plan.xlsx")
~~~

Attention : `import Dim2Dir` (fichier autonome) et `import dim2dir` (paquet) désignent deux modules distincts qui ne diffèrent que par la casse. Sur les systèmes de fichiers insensibles à la casse (Windows, macOS par défaut), ne pas placer `Dim2Dir.py` et le paquet `dim2dir` dans le même dossier, au risque d'importer l'un à la place de l'autre.

Il contient la même API et la même version que le paquet. Sans argument, le CLI conserve le traitement de `./Dim2Dir.xlsx`, feuille `Feuil1` si elle existe, sinon la première. Les constantes Spyder `EXCEL_FILE` et `SHEET_NAME` restent disponibles pour le CLI historique ; elles ne configurent pas l'API.

### B. Installation avec pip

~~~console
python -m pip install dim2dir
dim2dir mon_plan.xlsx
python -m dim2dir mon_plan.xlsx
python -m pip install "dim2dir[pandas]"
python -m pip install "dim2dir[excel]"
~~~

Python 3.10 ou ultérieur est requis. L'extra `pandas` ajoute pandas >= 2.0 ; l'extra `excel` ajoute pywin32 uniquement sous Windows et exige Excel de bureau. L'import de `dim2dir` ne charge ni pandas ni pywin32 ; matplotlib n'est chargé que pour construire ou exporter une figure.

## API Python

| Élément | Usage et retour |
| --- | --- |
| `project(data, *, x_factors=None, y_factors=None, run_column="Run", sheet_name=None, style=None, limits=None)` | `Projection` figée, sans fichier ni allocation de grille dense |
| `figure(data, *, style=None, limits=None, x_factors=None, y_factors=None, run_column="Run", sheet_name=None)` | `matplotlib.figure.Figure` Agg ; sans pyplot, sauvegarde ni affichage ; lève `Dim2DirError` si le rendu dépasse les limites |
| `visualize(data, *, output_dir=None, name=None, outputs=None, in_memory=None, style=None, limits=None, engine="matplotlib", dpi=200, x_factors=None, y_factors=None, run_column="Run", sheet_name=None, report=False)` | Produit les formats demandés et retourne un `Result`, sans imprimer |
| `Dim2DirError` | Sous-classe de `ValueError` pour les entrées et options invalides ; chemin source absent : `FileNotFoundError` |
| `Projection` | Feuille, en-têtes, essais, facteurs x/y, dimensions, cellules occupées, collisions, placements, avertissements, `style` et `layout` récursivement en lecture seule |
| `Result` | `files` : chemins absolus ; `data` : octets ; `report` : résumé texte ; `warnings` ; `projection` ; `style` effectif ; `engine` réellement utilisé après repli, ou `None` sans image |

`data` accepte un chemin `str` ou `Path` vers un `.xlsx`, un DataFrame pandas, une liste de dictionnaires (un par essai) ou un dictionnaire de listes de même longueur (une liste par colonne). La première feuille est utilisée par défaut ; `sheet_name` permet de la choisir. Toutes les sources utilisent la même lecture, projection et écriture openpyxl. Les images matplotlib lisent cette feuille écrite ; le moteur COM exporte le même classeur.

Pour les entrées en mémoire, les noms de colonnes sont convertis en texte et dépouillés des espaces aux extrémités. Les colonnes MultiIndex doivent être aplaties. L'index est ignoré, sauf s'il porte le nom `run_column` (sans distinction de casse) et qu'aucune colonne de ce nom n'existe. Les scalaires numpy sont convertis en valeurs Python ; `None`, NaN, `pd.NA` et NaT deviennent des cellules vides. Les booléens restent distincts des nombres et les dates sont conservées. `run_column` personnalise l'identifiant des essais. Les listes `x_factors` et `y_factors` doivent être fournies ensemble et partitionner exactement les autres colonnes ; les listes vides sont acceptées.

## Personnaliser

`Style` et `Limits` sont des dataclasses immuables et validées à la construction ; une valeur invalide lève `Dim2DirError`. `Style()` et `Limits()` conservent exactement le rendu Excel historique (version 1). Les dictionnaires et séquences d'ordre des niveaux sont copiés dans des structures immuables ; les styles de niveaux scalaires sont hashables.

| Champ `Style` | Défaut | Effet / valeurs |
| --- | --- | --- |
| `palette` | `None` | Dégradé historique blanc vers la couleur de base ; sinon colormap matplotlib (`viridis`, `cividis`, `Greys`…) ou séquence d'au moins deux couleurs `RRGGBB` / `#RRGGBB`, interpolées linéairement |
| `base_color` | `None` | Couleur de A1, sinon `335C81` ; couleur explicite au même format. Ignorée avec avertissement si une palette est donnée |
| `palette_reverse` | `False` | Inverse l'ordre des couleurs |
| `font_family` | `None` | Arial ; police explicite introuvable : avertissement et repli DejaVu Sans une fois par appel de visualisation. Repli de la police par défaut silencieux |
| `font_size` | `None` | Tailles historiques : essais 9, étiquettes 10, titre 11 points. Valeur 4..72 : étiquettes = taille, essais = max(4, taille − 1), titre = taille + 1 |
| `font_color` | `None` | Étiquettes et titre : `1F1F1F` ; sinon couleur explicite |
| `label_scale` | `1.0` | Facteur 0.5..3.0 ; élargit les colonnes d'étiquettes de l'axe y et la hauteur des étiquettes de l'axe x ; utile avec une police large comme Verdana |
| `cell_font_color` | `None` | Texte des essais : noir/blanc automatique selon la luminance ; sinon couleur explicite |
| `cell_size` | `34` | Pixels de cellule à 96 dpi, entier 12..120 ; dimensions Excel déduites des formules historiques |
| `line_color` | `"1F1F1F"` | Couleur des bordures |
| `line_style` | `"thin"` | `none`, `thin`, `medium`, `thick` ; matplotlib : 0, 0,75, 1,5, 2,25 points |
| `title` | `None` | `Visualisation - {feuille}` ; `""` retire le texte |
| `legend` | `False` | Une ligne sous les axes x avec couleurs du premier et du dernier essai et texte « Couleur : ordre des essais, du premier au dernier », écrite dans la feuille ; sur un plan étroit, le texte occupe toute la largeur et les échantillons de couleur sont réduits à un seul, ou omis |
| `label_format` | `"{name}: {level}"` | Étiquettes y ; seuls les champs `name` et `level` sont permis, sans conversion ni spécification de format |
| `run_prefix` | `"E"` | Préfixe des essais (`E1`, `E2`…) ; chaîne vide acceptée |
| `level_orders` | `None` | Facteur → liste de niveaux ; niveaux absents du plan ignorés, niveaux observés absents de la liste ajoutés à la fin, avec avertissements ; facteur inconnu ou niveaux dupliqués : erreur |

Une palette associe les rangs 1..N aux positions 0..1 ; pour un seul essai, la couleur est prise au milieu. Le dégradé historique par défaut conserve sa formule historique. Pour une collision, la couleur est celle du premier essai présent dans la cellule.

| Champ `Limits` | Défaut | Contrôle |
| --- | --- | --- |
| `max_cells` | `90_000` | Nombre maximal de cellules de grille, contrôlé avant toute allocation dense ; entier positif, au plus 17 179 869 184 (borne dure Excel) |
| `max_pixels` | `60_000_000` | Plafond de largeur × hauteur en pixels, entier positif ; protège également la taille de scène PDF/SVG |
| `max_runs` | `20_000` | Essais lus, entier positif ; dépassement refusé avant de convertir les données en mémoire en classeur |
| `excel_timeout_s` | `120.0` | Délai du moteur COM, nombre fini strictement positif |

| Argument de sortie | Comportement |
| --- | --- |
| `outputs` | Sous-ensemble **ordonné non vide**, sans doublon, de `xlsx`, `png`, `pdf`, `svg`. Une chaîne seule, par exemple `"png"`, est acceptée. Inconnu, doublon ou `()` : erreur |
| `outputs=None` | Source `.xlsx` : `("xlsx", "png", "pdf")` ; source en mémoire : `("png",)` |
| `in_memory=True` | Aucun fichier de sortie ; `Result.files == {}` et `Result.data` contient des octets aux clés demandées, plus `report` si activé. COM utilise un dossier temporaire interne, nettoyé même après exception |
| `in_memory=False` | Fichiers dans `output_dir`, obligatoire sauf pour une source `.xlsx` (défaut : son dossier) ; `Result.data == {}` |
| `in_memory=None` | Mémoire si `output_dir is None` et source en mémoire ; fichiers sinon |
| `output_dir`, `name` | Dossier et radical ; défaut du radical : nom de la source ou `plan`. Avec `in_memory=True`, aucun dossier de sortie n'est créé |
| `engine` | `matplotlib`, `excel` ou `auto` ; `auto` essaie Excel puis signale le repli. SVG exige matplotlib ; `excel` + SVG est une erreur, `auto` + SVG choisit matplotlib avec avertissement |
| `dpi` | Entier 50..600, défaut 200 ; résolution matplotlib. Le PNG COM conserve la résolution fournie par Excel |
| `report` | `False` par défaut ; `True` ajoute un rapport JSON UTF-8 (clé `report`, octets ou fichier) en plus du résumé texte toujours présent dans `Result.report` |

Les noms de fichiers sont `<name>_visu.xlsx`, `<name>_visu_<feuille>.png` / `.pdf` / `.svg` et `<name>_rapport.json`. Ils sont absolus dans `Result.files`. Un classeur peut être généré depuis un DataFrame avec `outputs="xlsx"`. Toutes les sorties sont d'abord préparées dans un dossier temporaire du dossier de sortie ; les fichiers existants ne sont remplacés qu'une fois toutes les sorties prêtes, et restaurés si un remplacement échoue (par exemple fichier ouvert dans une autre application). La source reste intacte et les temporaires sont nettoyés. En mémoire, `files` du rapport JSON liste les noms logiques des sorties produites.

Depuis Excel :

```python
from dim2dir import visualize, Style, Limits

result = visualize("mon_plan.xlsx", output_dir="resultats",
                   outputs=("xlsx", "png", "pdf", "svg"),
                   style=Style(palette="viridis", font_size=11),
                   limits=Limits(max_cells=120_000), report=True)
print(result.files["xlsx"], result.engine)
```

Depuis pandas, en mémoire :

```python
import pandas as pd
from dim2dir import visualize, Style

df = pd.DataFrame({"Run": [1, 2], "Courant A": [140, 180], "Gaz L/min": [12, 16]})
style = Style(palette="cividis", font_family="DejaVu Sans", cell_size=44,
              title="Plan de soudage", legend=True,
              level_orders={"Courant A": [180, 140]})
result = visualize(df, style=style, outputs=("png", "svg"), in_memory=True)
png_bytes = result.data["png"]
```

Le fichier JSON contient les champs de `Style`, sans enveloppe supplémentaire :

```json
{
  "palette": ["FFFFFF", "009E73", "0072B2"],
  "font_family": "DejaVu Sans",
  "font_size": 11,
  "cell_size": 44,
  "legend": true,
  "label_format": "{name} = {level}",
  "run_prefix": "Essai ",
  "level_orders": {"Gaz L/min": [16, 12]}
}
```

`Style.from_json("style.json")`, `Style.from_dict(mapping)` et `style.to_dict()` permettent de charger, valider et sérialiser le style. Toute clé inconnue est refusée.

```console
python Dim2Dir.py mon_plan.xlsx --outputs xlsx,png,svg --style style.json --palette cividis --font-size 12 --legend
python Dim2Dir.py mon_plan.xlsx --palette FF0000,FFFFFF,0000FF --palette-reverse --no-title --max-runs 30000
```

Les clients complets, avec personnalisation et PNG en mémoire, sont dans `examples/client_project/` et fonctionnent depuis n'importe quel dossier.

## Format d'entrée et sorties

Le fichier doit être un `.xlsx`. La ligne 1 contient les en-têtes contigus, et les lignes suivantes les essais. La colonne `Run`, quelle que soit sa casse, contient les identifiants des essais et est exclue des facteurs. Sans `Run`, les essais sont numérotés suivant l'ordre des lignes. Les lignes entièrement vides sont ignorées. Tous les autres en-têtes désignent des facteurs.

Le classeur source reste intact (il n'est jamais modifié). La sortie atomique `<nom>_visu.xlsx` est écrite par openpyxl : elle conserve le tableau, les formules et la mise en forme d'origine et ajoute la visualisation à droite, mais openpyxl peut perdre les images, graphiques et formes du classeur source. Pour un classeur qui en contient, travailler sur une copie et vérifier le résultat. Les fichiers `<nom>_visu_<feuille_assainie>.pdf` et `.png` ne montrent que le titre, les étiquettes des axes et la grille. Les caractères inadaptés des noms de feuilles sont remplacés par `_` ; les collisions de noms reçoivent un suffixe numérique.

Les niveaux numériques sont triés ; les textes gardent leur ordre d'apparition. L'axe y est inversé pour la lecture dans Excel. Les cellules vides d'un facteur utilisent l'indice de niveau 0, comme dans la version historique, avec avertissement. Un en-tête vide tronque le tableau ; un avertissement désigne la colonne si d'autres en-têtes suivent. Les doublons d'en-têtes sont distingués par `_2`, `_3`, etc., en évitant les noms existants, sans modifier le tableau source.

Les formules sont lues depuis leurs valeurs en cache. Une formule sans cache est signalée et traitée comme vide. Le programme ne calcule pas les formules ; recalculer et enregistrer le fichier dans un tableur avant traitement si nécessaire. Les formules restent présentes dans le classeur écrit, dont les caches peuvent être invalidés par openpyxl.

## Utilisation

```console
python Dim2Dir.py
python Dim2Dir.py examples/tournage_L9.xlsx Feuil1
python Dim2Dir.py examples/soudage_2p4.xlsx --output-dir resultats --report resultats/rapport.json
python Dim2Dir.py exemples.xlsx --all-sheets --dpi 150
python Dim2Dir.py exemples.xlsx --list-sheets
python Dim2Dir.py exemples.xlsx --dry-run --max-cells 60000
python Dim2Dir.py exemples.xlsx --x-factors "Vc m/min,ap mm" --y-factors "f mm/tr,Lubrification"
python Dim2Dir.py exemples.xlsx --image-engine auto --no-png
```

| Argument / option | Effet |
| --- | --- |
| `excel_file` | Chemin source, défaut `./Dim2Dir.xlsx` |
| `sheet_name` | Feuille, défaut `Feuil1` si elle existe, sinon la première du classeur |
| `--all-sheets` | Toutes les feuilles ; celles sans tableau sont signalées et ignorées. Erreur si aucune n'est valide |
| `--list-sheets` | Liste les feuilles puis termine sans écrire |
| `--output-dir DIR` | Dossier des fichiers de sortie, défaut : dossier d'entrée |
| `--no-xlsx` | Ne conserve pas de classeur de sortie |
| `--no-pdf` | Désactive le PDF |
| `--no-png` | Désactive le PNG |
| `--image-engine {matplotlib,excel,auto}` | Défaut `matplotlib`. `auto` essaie Excel puis explique le repli éventuel |
| `--dpi N` | Résolution PNG, entier de 50 à 600, défaut 200. Ne s'applique qu'au moteur matplotlib : le PNG du moteur excel a la résolution fournie par Excel |
| `--color RRGGBB` | Couleur de base, `RRGGBB` ou `#RRGGBB`, défaut : remplissage de A1 puis `335C81` |
| `--run-column NOM` | Colonne des numéros d'essai, défaut `Run` (sans distinction de casse) |
| `--strict` | Code de sortie 1 si une feuille est ignorée (`skipped`) ou si une sortie demandée n'a pas pu être produite |
| `--max-cells N` | Maximum de cellules dans la grille exacte, défaut 90000 |
| `--outputs xlsx,png,pdf,svg` | Formats demandés ; remplace `--no-xlsx`, `--no-pdf`, `--no-png`, qui restent acceptés |
| `--style FICHIER.json` | Charge un Style ; les options explicites du CLI priment |
| `--palette NOM_OU_HEX,HEX,...` | Colormap ou couleurs interpolées |
| `--palette-reverse` | Inverse la palette |
| `--font NOM`, `--font-size N`, `--font-color HEX` | Police, taille et couleur des étiquettes et du titre |
| `--label-scale N` | Facteur 0.5..3.0 sur la largeur des colonnes d'étiquettes y et la hauteur des étiquettes x |
| `--cell-size N` | Cellule de grille : 12..120 pixels |
| `--line-color HEX`, `--line-style {none,thin,medium,thick}` | Couleur et épaisseur des bordures |
| `--title TEXTE`, `--no-title` | Titre personnalisé ou suppression du texte ; dernière option explicite prioritaire |
| `--legend` | Ajoute une ligne de légende dans la feuille |
| `--label-format TEXTE`, `--run-prefix TEXTE` | Format y et préfixe des essais ; chaîne vide acceptée pour le préfixe |
| `--level-order "Facteur=a,b,c"` | Répétable ; nombres reconnus, sinon texte ; complète l'ordre chargé du JSON par facteur |
| `--max-pixels N`, `--max-runs N`, `--excel-timeout S` | Plafonds d'image, d'essais et délai COM ; mêmes défauts que Limits |
| `--x-factors A,B` | Affectation ordonnée de x, premier facteur = plus petit pas |
| `--y-factors C,D` | Affectation ordonnée de y ; les deux listes sont obligatoires ensemble et doivent partitionner exactement les facteurs |
| `--report FICHIER.json` | Rapport JSON UTF-8, `schema_version: 1`. Le chemin ne peut pas désigner un fichier écrit par le programme (xlsx, pdf, png) : erreur avant toute écriture. Le champ `input` ne contient que le nom du fichier et les fichiers écrits sont désignés par leur nom, relatif au dossier de sortie |
| `--dry-run` | Calcule et affiche le résumé sans écrire aucun fichier (rapport JSON compris) ni créer le dossier de sortie |
| `--quiet` | Masque le résumé et les avertissements ; les erreurs fatales restent sur stderr |
| `--version` | Affiche `2.1.0` et termine |
| `-h`, `--help` | Aide de la ligne de commande |

Une liste d'axes vide s'écrit `--x-factors ""` ou `--y-factors ""`. Les noms contenant une virgule ne sont pas représentables dans ces options ; l'alternance par défaut reste utilisable. Codes de sortie : 0 succès, 1 erreur de données ou d'exécution (avec `--strict`, aussi feuille ignorée ou sortie manquante), 2 usage incorrect. La sortie standard tolère une console non UTF-8 : les caractères impossibles à afficher sont remplacés, sans interrompre le traitement.

Le rapport par feuille donne les essais, les niveaux observés, l'affectation x/y, les dimensions de grille, les cellules occupées et vides, le taux de remplissage entre 0 et 1, les collisions (coordonnées à partir de zéro, y déjà inversé), le moteur réellement utilisé, les fichiers écrits, les avertissements et le statut `ok`, `skipped` ou `error`. Sans export d'image, le moteur du rapport JSON vaut `null`. En simulation aucun rapport n'est écrit, la liste des fichiers du rapport reste vide et le moteur vaut `null`.

## Écarts par rapport à la version 1

Sur une entrée valide, le bloc d'écriture Excel est identique à celui de la version 1 (cellules, fusions, styles, dimensions : le test de parité le vérifie). Les écarts voulus sont les suivants :

- les formules sont lues à partir de leurs valeurs en cache (la version 1 lisait l'expression), et traitées comme vides, avec avertissement, si le cache manque ;
- les en-têtes en double sont renommés (`_2`, `_3`, ...) pour la lecture, sans modifier le tableau source ;
- des avertissements explicites signalent les en-têtes vides et les facteurs vides ;
- les libellés de la visualisation commençant par `=` (niveaux, en-têtes de facteurs, étiquettes Run, texte des axes) sont écrits comme texte et non comme formules ;
- la visualisation refuse d'écraser des cellules non vides : si la zone à droite du tableau n'est pas libre, la feuille passe en statut `error` avec un message nommant la première cellule en conflit ;
- un identifiant `Run` textuel non fini (`inf`, `nan`, `1e309`) est conservé tel quel au lieu de lever une exception ;
- toute chaîne écrite dans la visualisation reste du texte (`#N/A`, `=1+1`, `+SUM(A1)`) ;
- des niveaux distincts dont le libellé serait identique (`12345.1` et `12345.2` en `.4g`, ou texte tronqué à 36 caractères) reçoivent une précision plus grande ou un suffixe `[n]`, avec avertissement ; les entrées sans collision gardent leurs libellés ;
- une valeur infinie dans un facteur forme le niveau `inf` ou `-inf`.

## Limites et cas particuliers

- Les entiers supérieurs à 2^53 sont comparés comme des nombres flottants : deux entiers voisins de cet ordre de grandeur se confondent en un seul niveau.
- Une cellule vide d'un facteur retombe sur le premier niveau (indice 0), avec avertissement, comme dans la version historique.
- Le moteur Excel (`engine="excel"`, `--image-engine excel`) ne fonctionne que sous Windows, avec Excel de bureau et pywin32. Aucun appel COM ne s'exécute en parallèle : un verrou les sérialise.
- Un classeur ouvert dans Excel ne peut pas être remplacé : le message désigne le fichier à fermer, et aucune autre sortie n'est modifiée.
- Les proportions de l'image peuvent différer de quelques pour cent entre le moteur Excel et matplotlib (polices et papier installés).
- openpyxl ne conserve pas les images, graphiques et formes du classeur source : le classeur de sortie les perd, avec un avertissement ; travailler sur une copie et vérifier le résultat.
- La grille exacte est contrôlée avant allocation dense. À 90 000 cellules elle est acceptée ; au-delà du plafond, `project` retourne des avertissements, `figure` lève `Dim2DirError`, `visualize` retourne un résultat `skipped` sans image. Un XLSX demandé reçoit la note rouge de la version 1. Si aucune sortie demandée ne peut être produite, un avertissement le dit explicitement ; aucune sortie inexistante n'est annoncée.
- Les limites Excel restent fixes : 16 384 colonnes et 1 048 576 lignes, tableau, titre, axes et légende compris. Aucun paramètre ne les relève. `max_cells` peut être relevé uniquement jusqu'à leur produit ; cela ne garantit pas que les dimensions ni la mémoire disponible conviennent.
- `max_runs` est vérifié avant conversion des données en mémoire ; les lectures Excel sont interrompues dès que le nombre d'essais dépasse ce plafond. Les grandes scènes et les classeurs proches des limites peuvent consommer beaucoup de mémoire.
- `max_pixels` borne l'image et la scène vectorielle avant de créer la figure ; diminuer `dpi` ou les niveaux, ou relever explicitement ce plafond. Les dimensions du PNG réellement exporté par COM sont également vérifiées. Un export refusé dans `visualize` est signalé par `skipped`. La limite de DPI est toujours 50..600.
- Pour un DataFrame ou des listes, `sheet_name` doit être un nom Excel non vide de 31 caractères maximum, sans `\`, `/`, `*`, `?`, `:`, `[` ni `]`. Les noms des fichiers images sont assainis ; les collisions reçoivent un suffixe numérique.
- Le moteur Excel COM exige Windows, Excel de bureau et pywin32. Un verrou de module sérialise les appels COM d'un même processus : jamais en parallèle. Ce moteur n'est pas adapté à un serveur multi-processus. Il ouvre une copie interne en lecture seule dans une instance dédiée invisible, neutralise titres d'impression, en-têtes et pieds de page, puis exporte et ferme l'instance.
- Le délai COM est réglable avec `excel_timeout_s` / `--excel-timeout`. L'instance Excel dédiée est arrêtée après échec ou dépassement ; les temporaires sont nettoyés dans un `finally`, en mémoire comme sur disque. `auto` signale le repli éventuel sur matplotlib dans `warnings` et `Result.engine` indique le moteur utilisé. SVG est réservé à matplotlib.
- Les refus d'accès à un fichier de sortie sont signalés par le même message pour l'API et pour le CLI ; la source n'est jamais écrasée.
- Le rendu Excel par défaut conserve valeurs, fusions, styles, dimensions et rotations historiques. Matplotlib lit la même feuille (polices, tailles, couleurs et bordures comprises), sans modifier les `rcParams` globaux. Police explicite absente : avertissement puis DejaVu Sans ; Arial par défaut absent : repli silencieux. `figure()` transmet un avertissement de police absente par le mécanisme Python `warnings`.
- Un écart de proportions jusqu'à environ 10 % est possible entre COM et matplotlib à cause des polices et du papier installés. La conversion des largeurs Excel suppose Calibri 11, soit 7 pixels par caractère à 96 dpi ; les dimensions matplotlib sont en points, avec une marge de 3 points. L'enveloppe et la réduction des textes restent approchées.
- `figure()` ne crée jamais de figure pyplot et ne change pas le backend. L'appelant gère la durée de vie de la figure : la fermer après affichage, ou appeler `fig.clear()` pour libérer ses artistes quand elle n'est plus utilisée. Un enregistrement ultérieur avec d'autres paramètres relève de l'appelant.
- En-têtes contigus, doublons, cellules vides et formules suivent les règles de lecture décrites plus haut. Une formule sans cache est signalée ; Dim2Dir ne la calcule pas. Les libellés qui commencent par `=` restent du texte. Une cellule non vide dans la zone de rendu empêche son écrasement et nomme la première cellule en conflit.

`project_fractal_2d`, `run_visualization`, `process_workbook` et `build_output_path` conservent leur interface historique. `add_visualization_to_sheet(ws, style=None, limits=None)` ajoute deux paramètres facultatifs, avec comportement historique par défaut. `run_visualization` conserve la production du seul classeur. Utiliser `visualize` pour les formats et les octets, ou `main(argv)` pour le CLI.

## Exemples et vérification

```console
python examples/generate_examples.py
python -m pip install -r requirements-dev.txt
python -m pytest --cov=dim2dir --cov-report=term-missing -m "not excel"
python Dim2Dir.py examples/tournage_L9.xlsx --output-dir resultats --report resultats/rapport.json
```

Le dossier `examples/sorties/` contient les sorties des plans d'exemple (classeur, PDF, PNG et rapport JSON). Elles sont régénérables avec `python Dim2Dir.py examples/<nom>.xlsx --output-dir examples/sorties`.

Le générateur fournit six plans purement synthétiques de tournage, soudage, boulonnerie, matériaux, grande grille et collisions. Les tableaux L9 et L18 standard sont écrits en dur et leur orthogonalité est testée. Les tests comparent chaque cellule et chaque dimension au rendu historique sur entrée valide, vérifient les formules en cache, le rapport et les images, ainsi que les codes CLI et l'empreinte du fichier source.

L'intégration Excel est désactivée par défaut. Pour un contrôle volontaire sur une machine équipée, définir `DIM2DIR_TEST_EXCEL=1` puis lancer `python -m pytest -m excel`. L'objectif de couverture est de 85 % hors fonctions COM isolées.

Le cas `plan_trop_grand.xlsx` utilise neuf facteurs à quatre niveaux, soit 262 144 cellules : huit facteurs à quatre niveaux ne donneraient que 65 536 cellules et ne déclencheraient pas le seuil par défaut de 90 000.


## Contribuer

La source de vérité est `src/dim2dir/core.py`. Modifier ce fichier, puis régénérer `Dim2Dir.py` ; ne jamais éditer la copie autonome à la main. Voir aussi `CONTRIBUTING.md` et `SECURITY.md`. Le rendu Excel historique est protégé par les tests de parité avec la version 1.

~~~console
python -m pip install ".[dev]"
python tools/sync_standalone.py
python -m pytest --cov=dim2dir --cov-report=term-missing --cov-fail-under=85 -m "not excel"
python tools/sync_standalone.py --check
python tools/verify_install.py
~~~

Le vérificateur construit sdist et roue, contrôle leur contenu et les métadonnées avec twine, puis teste le paquet sans pandas, le fichier autonome sans paquet installé et les deux projets clients avec l'extra pandas dans un environnement virtuel temporaire, sans PYTHONPATH. Il utilise le réseau uniquement pour pip. La roue contient uniquement `dim2dir/*` et ses métadonnées ; la sdist inclut aussi tests, exemples, outils et fichier autonome. La CI exécute ces vérifications sous Windows et Ubuntu avec Python 3.10, 3.12 et 3.13.
