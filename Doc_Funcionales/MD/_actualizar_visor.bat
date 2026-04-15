@echo off
chcp 65001 >nul
title SGP Visor — Modo Vigilancia
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║   SGP Upgrade — Generador de Documentación HTML     ║
echo  ║   Vigilando cambios en archivos .md...               ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
echo  Cada vez que guardes un archivo .md o agregues una
echo  nueva carpeta, el visor_documentos.html se actualizara
echo  automaticamente.
echo.
echo  Deja esta ventana abierta mientras trabajas.
echo  Cierra con Ctrl+C cuando termines.
echo.

cd /d "%~dp0"
node _gen_html.js --watch

echo.
echo  ══════════════════════════════════════════════════════
echo  Proceso detenido.
echo  ══════════════════════════════════════════════════════
echo.

:PREGUNTAR_GIT
set /p RESPUESTA="  Deseas actualizar el repositorio Git? (s/n): "
if /i "%RESPUESTA%"=="s" goto HACER_COMMIT
if /i "%RESPUESTA%"=="n" goto FIN
echo  Respuesta no valida. Escribe s o n.
goto PREGUNTAR_GIT

:HACER_COMMIT
echo.
echo  Agregando archivos al stage...
git add .
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo ejecutar git add.
    echo  Asegurate de estar dentro de un repositorio Git.
    goto FIN
)
echo  OK.
echo.

:PEDIR_MENSAJE
set "MENSAJE="
set /p MENSAJE="  Mensaje del commit: "
if "%MENSAJE%"=="" (
    echo  El mensaje no puede estar vacio. Intentalo de nuevo.
    goto PEDIR_MENSAJE
)

echo.
echo  Creando commit...
git commit -m "%MENSAJE%"
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo crear el commit.
    echo  Es posible que no haya cambios para confirmar.
    goto FIN
)

echo.
echo  Subiendo cambios a origin/main...
git push -u origin main
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo ejecutar git push.
    echo  Verifica tu conexion y credenciales de Git.
    goto FIN
)

echo.
echo  Repositorio actualizado correctamente.

:FIN
echo.
echo  Presiona cualquier tecla para cerrar.
pause >nul
