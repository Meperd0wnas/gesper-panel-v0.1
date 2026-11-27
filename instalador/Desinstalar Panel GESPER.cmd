@echo off
setlocal
rem ---------------------------------------------------------------------------
rem Desinstala el Panel GESPER para el usuario actual de Windows: borra el
rem registro en Excel y la copia del manifiesto en %LOCALAPPDATA%\GESPER.
rem No toca ningun libro de Excel.
rem
rem Plantilla: instalador/ del repositorio. Solo caracteres ASCII.
rem ---------------------------------------------------------------------------

set "ID={{ID}}"
set "DESTINO=%LOCALAPPDATA%\GESPER"
set "CLAVE=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo.
echo   Panel GESPER - desinstalacion
echo   ------------------------------------------------
echo.

reg delete "%CLAVE%" /v "%ID%" /f >nul 2>&1
reg delete "%CLAVE%\%ID%" /f >nul 2>&1
if exist "%DESTINO%\manifest.xml" del /q "%DESTINO%\manifest.xml"
rmdir "%DESTINO%" 2>nul

echo   Listo: el Panel GESPER se quito de Excel para este usuario.
echo   Cierra todas las ventanas de Excel y vuelve a abrirlo para que
echo   desaparezca el boton.
echo.
pause
exit /b 0
