@echo off
REM Configuracion de la PC del operador. Se corre UNA VEZ, despues de
REM descomprimir el agente. No instala nada: solo deja puesta la variable que
REM el agente necesita y el acceso directo para que arranque con Windows.

setlocal

set "ORIGEN=https://autofiller.insoft.net.ar"
if not "%~1"=="" set "ORIGEN=%~1"

echo.
echo   Configurando el agente de AutoFiller
echo   ------------------------------------
echo.

REM De donde se sirve la web. El agente solo acepta pedidos de estos origenes:
REM cualquier otra pagina abierta en el navegador no puede hacerle cargar
REM comprobantes. Queda como variable de usuario, asi no hay que acordarse de
REM ponerla en cada arranque.
setx AUTOFILLER_ORIGENES "%ORIGEN%" >nul
if errorlevel 1 goto :error
echo   [ok] Origen permitido: %ORIGEN%

REM Acceso directo en la carpeta Inicio, para que el agente este abierto sin que
REM el operador tenga que acordarse. Sin esto, el sintoma es "No se detecta el
REM agente" en la web y nadie sabe por que.
REM
REM Desde la 2.1 el agente arranca sin ninguna ventana --queda como icono en la
REM bandeja-- asi que esto ya no le pone una consola adelante al iniciar sesion.
set "DESTINO=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\AutoFiller Agente.lnk"
powershell -NoProfile -Command ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('%DESTINO%');" ^
  "$s.TargetPath = '%~dp0AutoFillerAgente.exe';" ^
  "$s.WorkingDirectory = '%~dp0';" ^
  "$s.Description = 'Agente de AutoFiller';" ^
  "$s.Save()" >nul
if errorlevel 1 goto :error
echo   [ok] Arranca junto con Windows

echo.
echo   Listo. Abri AutoFillerAgente.exe ahora, o reinicia la sesion.
echo.
echo   No se abre ninguna ventana: el agente queda como un icono al lado del
echo   reloj, atras de la flechita de "Iconos ocultos". Clic derecho ahi para
echo   abrir la web o ver la actividad.
echo.
echo   IMPORTANTE: el agente y el AutoFiller.exe viejo manejan el MISMO Chrome,
echo   asi que no se pueden usar los dos a la vez. Termina el lote en uno y
echo   cerralo antes de abrir el otro.
echo.
pause
goto :eof

:error
echo.
echo   No se pudo configurar. Probá abriendo esta ventana como administrador.
pause
