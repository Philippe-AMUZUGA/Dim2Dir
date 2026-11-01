# Contribuer à Dim2Dir

1. Installer le projet avec ses outils de développement :

   ```console
   python -m pip install ".[dev]"
   ```

2. La source de vérité est `src/dim2dir/core.py`. Ne jamais éditer `Dim2Dir.py` à la main : c'est une copie autonome, octet pour octet, régénérée par `python tools/sync_standalone.py` après chaque série de modifications.
3. Lancer les tests et les contrôles avant toute proposition :

   ```console
   python tools/sync_standalone.py
   python -m pytest -q --cov=dim2dir
   python tools/sync_standalone.py --check
   python tools/verify_install.py
   ```

4. Le rendu Excel par défaut doit rester identique à la référence `tests/reference/Dim2Dir_V1.py` (version 1). Tout correctif de comportement reçoit un test de régression ; les messages destinés à l'utilisateur sont en français.
