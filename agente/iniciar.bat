@echo off
REM Agente local de AutoFiller: es lo que carga los comprobantes en SISalud.
REM Tiene que estar abierto en la PC del operador mientras se usa la web.
REM La primera vez crea el entorno e instala las dependencias; despues solo arranca.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creando el entorno de Python...
    python -m venv .venv || goto :error
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements.txt || goto :error
)

REM De donde se sirve la web de AutoFiller. El agente solo acepta pedidos de
REM estos origenes: cualquier otra pagina abierta en el navegador no puede
REM hacerle cargar comprobantes. Cambiar al desplegar el servidor.
if "%AUTOFILLER_ORIGENES%"=="" set AUTOFILLER_ORIGENES=http://localhost:8000,http://127.0.0.1:8000

echo Agente de AutoFiller escuchando en 127.0.0.1:8765
echo Origenes permitidos: %AUTOFILLER_ORIGENES%
echo Dejalo abierto mientras cargas comprobantes.
.venv\Scripts\python.exe -m uvicorn main:app --host 127.0.0.1 --port 8765
goto :eof

:error
echo.
echo No se pudo preparar el entorno. Revisa que Python 3.11 o superior este instalado.
pause
