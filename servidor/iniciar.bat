@echo off
REM Servidor de extraccion de AutoFiller.
REM La primera vez crea el entorno e instala las dependencias; despues solo arranca.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creando el entorno de Python...
    python -m venv .venv || goto :error
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements.txt || goto :error
)

if "%ANTHROPIC_API_KEY%"=="" (
    echo.
    echo   AVISO: no hay ANTHROPIC_API_KEY configurada.
    echo   El servidor va a leer PDFs, pero no comprobantes en foto.
    echo.
)

if "%AUTOFILLER_PUERTO%"=="" set AUTOFILLER_PUERTO=8000

echo Servidor de AutoFiller en http://localhost:%AUTOFILLER_PUERTO%
.venv\Scripts\python.exe -m uvicorn main:app --host 0.0.0.0 --port %AUTOFILLER_PUERTO%
goto :eof

:error
echo.
echo No se pudo preparar el entorno. Revisa que Python 3.11 o superior este instalado.
pause
