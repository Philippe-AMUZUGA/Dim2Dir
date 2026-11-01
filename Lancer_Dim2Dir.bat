@echo off
setlocal
rem Le dossier courant n'est pas modifie : les chemins relatifs donnes par l'utilisateur restent valides.
if exist "%~dp0.runtime\Scripts\python.exe" goto dependencies
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 -m venv "%~dp0.runtime"
if errorlevel 1 goto fail
goto dependencies
:use_python
python -m venv "%~dp0.runtime"
if errorlevel 1 goto fail
:dependencies
if exist "%~dp0.runtime\.dependencies-ok" goto run
"%~dp0.runtime\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto fail
type nul > "%~dp0.runtime\.dependencies-ok"
:run
"%~dp0.runtime\Scripts\python.exe" "%~dp0Dim2Dir.py" %*
exit /b %ERRORLEVEL%
:fail
echo Echec de preparation de l'environnement Python local.
exit /b 1
