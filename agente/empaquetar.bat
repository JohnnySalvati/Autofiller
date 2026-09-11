@echo off
REM Compila el agente y deja el zip listo para subir al servidor.
REM
REM Esto es el deploy del agente entero: compilar, subir el zip, y la web le
REM avisa sola a cada operador que hay una version nueva. No hay que ir PC por
REM PC ni acordarse de quien quedo atrasado.
REM
REM onedir y NO --onefile a proposito: arranca mas rapido (onefile se
REM desempaqueta entero en cada arranque) y los antivirus lo marcan menos.
REM
REM --collect-all playwright es obligatorio: Playwright trae su propio node.exe
REM y PyInstaller no se lo lleva por su cuenta. Sin eso, el agente compila bien
REM y falla recien al intentar abrir Chrome, en la PC del operador.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo No hay entorno. Corre iniciar.bat una vez primero.
    pause
    exit /b 1
)

echo.
echo   Compilando el agente...
echo.

.venv\Scripts\python.exe -m pip install --quiet --upgrade pyinstaller || goto :error

.venv\Scripts\python.exe -m PyInstaller --noconfirm --onedir --console ^
    --name AutoFillerAgente ^
    --collect-all playwright ^
    --collect-all uvicorn ^
    --hidden-import main ^
    --hidden-import navegador ^
    --hidden-import operador ^
    --hidden-import padron ^
    --hidden-import sisalud ^
    --icon ..\insoft.ico ^
    arrancar.py || goto :error

REM Lo que el operador necesita al lado del .exe.
copy /y LEEME.txt "dist\AutoFillerAgente\" >nul || goto :error
copy /y configurar.bat "dist\AutoFillerAgente\" >nul || goto :error

echo.
echo   Verificando el empaquetado...
echo.
REM Ejercita el driver de Playwright y la ruta de Chrome YA COMPILADO. Es la
REM unica forma de saber que el .exe sirve sin cargar un comprobante real.
echo. | "dist\AutoFillerAgente\AutoFillerAgente.exe" --probar
if errorlevel 1 (
    echo.
    echo   El autodiagnostico encontro problemas. No se genera el zip.
    pause
    exit /b 1
)

echo.
echo   Armando el zip...
if exist AutoFillerAgente.zip del AutoFillerAgente.zip
powershell -NoProfile -Command ^
  "Compress-Archive -Path 'dist\AutoFillerAgente\*' -DestinationPath 'AutoFillerAgente.zip' -Force" || goto :error

echo.
echo   Listo: %~dp0AutoFillerAgente.zip
echo   Subilo al servidor para que los operadores lo descarguen.
echo.
pause
goto :eof

:error
echo.
echo   Fallo el empaquetado.
pause
exit /b 1
