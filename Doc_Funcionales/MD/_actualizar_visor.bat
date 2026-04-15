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
echo  Proceso detenido. Presiona cualquier tecla para cerrar.
pause >nul
