@echo off
setlocal
rem ---------------------------------------------------------------------------
rem Instala el Panel GESPER {{VERSION}} para el usuario actual de Windows.
rem
rem No necesita permisos de administrador, Node.js ni certificados: copia el
rem manifiesto a %LOCALAPPDATA%\GESPER y lo registra en Excel con la misma
rem clave que usa office-addin-debugging. El panel se carga desde
rem {{URL}}
rem
rem Plantilla: instalador/ del repositorio. La genera scripts/empaquetar.js;
rem solo caracteres ASCII (cmd.exe lee este archivo con la pagina de codigos
rem de la consola).
rem ---------------------------------------------------------------------------

set "ID={{ID}}"
set "ORIGEN=%~dp0manifest.xml"
set "DESTINO=%LOCALAPPDATA%\GESPER"
set "CLAVE=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo.
echo   Panel GESPER {{VERSION}} - instalacion
echo   ------------------------------------------------
echo.

if not exist "%ORIGEN%" goto :sin_manifiesto

if not exist "%DESTINO%" mkdir "%DESTINO%"
copy /y "%ORIGEN%" "%DESTINO%\manifest.xml" >nul
if errorlevel 1 goto :fallo_copia

reg add "%CLAVE%" /v "%ID%" /t REG_SZ /d "%DESTINO%\manifest.xml" /f >nul
if errorlevel 1 goto :fallo_registro

echo   Listo: el Panel GESPER quedo instalado para este usuario.
echo.
tasklist /fi "imagename eq EXCEL.EXE" 2>nul | findstr /i "EXCEL.EXE" >nul
if not errorlevel 1 (
  echo   Excel esta abierto. Cierra TODAS sus ventanas y vuelve a abrirlo:
  echo   el panel aparece solo al arrancar Excel.
  echo.
)
echo   1. Abre Excel y tu libro GESPER ^(.xlsm^).
echo   2. En la cinta, ficha Inicio, pulsa el boton "Panel GESPER"
echo      ^(grupo GESPER, a la derecha^).
echo.
echo   El panel necesita conexion a internet: se carga desde
echo   {{URL}}
echo.
pause
exit /b 0

:sin_manifiesto
echo   No encuentro manifest.xml junto a este archivo.
echo.
echo   Descomprime el ZIP completo ^(clic derecho sobre el ZIP, "Extraer todo"^)
echo   y ejecuta "Instalar Panel GESPER.cmd" desde la carpeta descomprimida.
goto :fin_con_error

:fallo_copia
echo   No pude copiar el manifiesto a:
echo   %DESTINO%
goto :fin_con_error

:fallo_registro
echo   Windows no dejo registrar el panel en Excel ^(clave %CLAVE%^).
echo   Puede que una politica de la empresa lo impida. Pide a TI que instale
echo   el archivo manifest.xml de esta carpeta desde el centro de
echo   administracion de Microsoft 365 ^(Aplicaciones integradas^).
goto :fin_con_error

:fin_con_error
echo.
echo   El panel NO quedo instalado.
echo.
pause
exit /b 1
