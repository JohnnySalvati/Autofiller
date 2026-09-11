@echo off
REM Nombre del script que querés convertir a EXE
set SCRIPT=AutoFiller.py



REM Ejecutar pyinstaller con opciones
pyinstaller --onefile --noconsole --icon autofiller.ico --noconfirm %SCRIPT%


pause