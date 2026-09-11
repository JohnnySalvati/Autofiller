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
REM --noconsole: el agente vive en la bandeja del sistema, sin ventana. No es la
REM vuelta al --noconsole del escritorio, que escondia los errores: todo lo que
REM antes iba a la consola queda en %LOCALAPPDATA%\AutoFiller\agente.log y el
REM menu del icono lo muestra en vivo (ver registro.py y visor.py).
REM
REM El icono hace falta adentro del paquete, no solo como icono del .exe: es la
REM imagen que se dibuja en la bandeja.
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

REM El agente abierto tiene tomado su propio .exe y Windows no deja reemplazarlo:
REM PyInstaller falla recien al final, despues de compilar todo, y con un
REM traceback de PermissionError que no dice que hay que cerrar el agente.
tasklist /FI "IMAGENAME eq AutoFillerAgente.exe" | find /I "AutoFillerAgente.exe" >nul
if not errorlevel 1 (
    echo.
    echo   El agente esta abierto y Windows no deja reemplazar su .exe.
    echo   Cerralo desde el icono de la bandeja, al lado del reloj:
    echo   boton derecho, "Salir del agente". Despues volve a correr esto.
    echo.
    pause
    exit /b 1
)

echo.
echo   Compilando el agente...
echo.

.venv\Scripts\python.exe -m pip install --quiet --upgrade pyinstaller || goto :error

.venv\Scripts\python.exe -m PyInstaller --noconfirm --onedir --noconsole ^
    --name AutoFillerAgente ^
    --collect-all playwright ^
    --collect-all uvicorn ^
    --hidden-import main ^
    --hidden-import navegador ^
    --hidden-import operador ^
    --hidden-import padron ^
    --hidden-import sisalud ^
    --hidden-import bandeja ^
    --hidden-import visor ^
    --hidden-import registro ^
    --hidden-import recursos ^
    --hidden-import pystray._win32 ^
    --add-data "..\autofiller.ico;." ^
    --icon ..\autofiller.ico ^
    arrancar.py || goto :error

REM Lo que el operador necesita al lado del .exe.
copy /y LEEME.txt "dist\AutoFillerAgente\" >nul || goto :error
copy /y configurar.bat "dist\AutoFillerAgente\" >nul || goto :error

echo.
echo   Verificando el empaquetado...
echo.
REM Ejercita el driver de Playwright, la ruta de Chrome y el icono de la bandeja
REM YA COMPILADOS. Es la unica forma de saber que el .exe sirve sin cargar un
REM comprobante real.
REM
REM Dos detalles de llamar a un programa que ya no tiene consola (medido el
REM 2026-09-11, no supuesto):
REM  - cmd SI lo espera y SI propaga su errorlevel desde un .bat, asi que la
REM    llamada directa sirve igual que antes. El pipe de echo sigue estando por
REM    el Enter final del autodiagnostico, que con una consola de verdad
REM    esperaria de verdad.
REM  - --sin-carteles porque el .exe sin consola avisa con un MessageBox si no
REM    logra reengancharse a esta consola, y un cartel que nadie clickea deja
REM    este empaquetado colgado para siempre.
echo. | "dist\AutoFillerAgente\AutoFillerAgente.exe" --probar --sin-carteles
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
