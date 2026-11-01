# Historique

Dans ce fichier, « version 1 » désigne la première version du programme, sans paquet ni API Python.

## Non publié

- Copie et sérialisation : `Projection` et `Result` acceptent `pickle`, `copy.copy`, `copy.deepcopy` et `dataclasses.replace` ; l'objet reconstruit reste en lecture seule.
- Libellés : une valeur infinie forme le niveau `inf` ou `-inf` ; des niveaux distincts de même libellé reçoivent une précision plus grande (jusqu'à 17 chiffres significatifs) ou un suffixe, avec avertissement ; toute chaîne écrite dans la visualisation reste du texte (`#N/A`, `=1+1`).
- Affectation x/y : les messages d'erreur utilisent les noms d'origine des colonnes, jamais d'alias internes.
- CLI : console non UTF-8 tolérée ; `--run-column`, `--strict`, `--color` avec `#`, feuille par défaut `Feuil1` sinon la première ; nom de programme `dim2dir` et aide mentionnant Dim2Dir.
- Légende lisible sur un plan étroit : texte fusionné sur toute la largeur, échantillons de couleur seulement si la largeur le permet.
- Sorties : aucune destination n'est remplacée tant que toutes les sorties ne sont pas prêtes, restauration en cas d'échec, même message pour l'API et le CLI ; `files` du rapport JSON en mémoire listé ; avertissement si le classeur contient des images ou des graphiques que openpyxl perd.
- Entiers numpy acceptés pour `dpi`, `cell_size`, `max_cells`, `max_pixels`, `max_runs` ; `cell_size` exact de 12 à 120 ; import de matplotlib centralisé avec garde de configuration.
- Dépôt : couverture mesurée sur le paquet `dim2dir`, texte intégral de la licence AGPL 3, métadonnées de distribution complétées, `.gitattributes`, lanceur Windows qui conserve le dossier courant, publication PyPI par « trusted publishing », `CONTRIBUTING.md` et `SECURITY.md`.

## 2.1.0

- API de sorties unifiée : `outputs`, choix fichiers ou octets en mémoire avec `in_memory`, exports XLSX, PNG, PDF et SVG ; `Result.data`, style effectif et moteur réellement utilisé.
- `Style` immuable et validé : palettes matplotlib ou couleurs interpolées, inversion, polices, tailles, couleurs, cellules, bordures, titre, légende, libellés, préfixes et ordre des niveaux ; chargement JSON et priorité des options CLI.
- `Limits` immuable : cellules, pixels, essais avant conversion et délai COM ; contrôles avant allocations, messages explicites et avertissement lorsqu'aucune sortie demandée n'est produite.
- Matplotlib lit les polices et bordures de la feuille écrite ; avertissement de police absente, sans modification globale de configuration. Verrou de processus pour sérialiser COM et nettoyage des sorties internes, même après exception.
- CLI enrichi sans retrait des options historiques ; `--outputs` prime sur `--no-*`. Exemples pandas personnalisés avec cividis et PNG en mémoire, contrôles installés, tests de style, formats et limites.
- Champ `Style.label_scale` et option `--label-scale` (0.5 à 3.0, défaut 1) : multiplie la largeur des colonnes d'étiquettes de l'axe y et la hauteur des étiquettes de l'axe x, utile avec une police large comme Verdana.
- Rendu Excel de la version 1 inchangé avec les valeurs par défaut, y compris `Style()` et `Limits()` explicites ; source autonome synchronisée.

- Distribution pip avec paquet src/dim2dir, commande dim2dir et python -m dim2dir ; extras pandas, excel et dev.
- API project, figure, visualize, résultats Projection et Result, exceptions Dim2DirError ; mêmes lecture, projection et rendu pour Excel, DataFrame, liste de dictionnaires et dictionnaire de listes.
- Normalisation des données en mémoire, valeurs manquantes, index Run, colonnes en double et identifiants personnalisés ; aucune dépendance obligatoire à pandas.
- Figures Agg sans pyplot ni changement de backend ; exports silencieux, chemins absolus, rapport et nettoyage atomique des fichiers temporaires.
- Source unique dans src/dim2dir/core.py, copie autonome Dim2Dir.py produite par tools/sync_standalone.py et vérifiée octet pour octet.
- Deux clients mécaniques, tests d'API et d'équivalence, vérification réutilisable des distributions en environnement isolé et CI Windows/Ubuntu Python 3.10, 3.12, 3.13.
- Fonctions historiques, CLI sans argument et bloc de rendu Excel de la version 1 conservés.

## 2.0.0

- CLI complète : sélection de feuilles, listes, simulation sans écriture, chemins de sortie, filtres de formats, couleur, résolution, seuil de grille et version.
- Affectation manuelle ordonnée des facteurs à x et y avec validation de partition exacte.
- Exports de la seule zone de visualisation : PDF vectoriel et PNG matplotlib ; moteur Excel optionnel isolé, limité à 120 s, et repli automatique expliqué.
- Rapport de couverture console et JSON versionné : niveaux observés, occupation, cellules vides, collisions, fichiers, moteur et statuts.
- Lecture des valeurs en cache des formules, préservation des formules lors de l'écriture, sauvegarde atomique, validation des feuilles avant toute copie, protection du fichier source.
- Exemples mécaniques synthétiques reproductibles, tests de parité avec la version 1, projection, lecture, images et CLI, lanceur Windows et documentation française.
- Rendu Excel de la version 1 conservé par réutilisation de son bloc d'écriture, mêmes signatures publiques et constantes.

Écarts volontaires de lecture : avertissements explicites pour les en-têtes vides au milieu du tableau et les facteurs vides ; en-têtes dupliqués distingués par suffixes (sans modifier les cellules source) ; formules lues depuis le cache au lieu de leur expression, et traitées comme vides si ce cache manque. Les images et le résumé console sont ajoutés au lancement CLI par défaut ; l'API `run_visualization` conserve la production du seul classeur.

Correction de l'exemple de grande grille : neuf facteurs à quatre niveaux (262 144 cellules), car huit facteurs à quatre niveaux (65 536 cellules) restent sous le plafond par défaut de 90 000.

Écarts volontaires d'écriture : les libellés de la visualisation commençant par `=` (niveaux, en-têtes de facteurs, étiquettes Run, texte des axes) sont écrits comme texte (`data_type = "s"`) au lieu d'être interprétés comme formules par openpyxl ; la visualisation refuse d'écraser une cellule non vide de la zone cible (statut `error`, première cellule en conflit nommée). Les entrées valides donnent un rendu identique au rendu de la version 1.

Corrections de la version 2 :

- matplotlib : `parse_math=False` pour tous les textes (les libellés contenant `$` ne sont plus interprétés).
- Les limites d'Excel (16 384 colonnes, 1 048 576 lignes) sont contrôlées même avec `--max-cells` personnalisé.
- `--report` ne peut plus désigner un fichier écrit par le programme (xlsx, pdf, png), comparaison insensible à la casse sous Windows ; contrôle des sorties entre elles ; erreur avant toute écriture.
- Identifiant `Run` textuel `inf`, `-inf`, `nan`, `1e309` : plus d'exception.
- Moteur Excel : titres d'impression, en-têtes et pieds de page neutralisés avant l'export PDF ; processus Excel arrêté (nom `EXCEL.EXE` vérifié) quand l'export n'a pas réussi, même si le processus enfant est déjà mort ; dossier temporaire tolérant aux erreurs de nettoyage ; délai effectif dans le message.
- `MPLCONFIGDIR` : plus de dossier de configuration dans le dépôt ; dossier temporaire du système, environnement du processus restauré.
- Rapport JSON : `input` réduit au nom du fichier ; fichiers écrits désignés par leur nom relatif au dossier de sortie.
- `.gitignore` complété ; README : section « Écarts par rapport à la version 1 », précisions sur `--dpi`, openpyxl, installation Linux et macOS, `--dry-run`.

## 1.0.0 : version initiale

- Version 1 : fichier autonome `Dim2Dir.py`, sans paquet ni API Python.
