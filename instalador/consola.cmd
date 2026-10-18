@echo off
rem Consola del Exportador EVM: agrega la carpeta del programa al PATH de esta ventana.
chcp 65001 > nul
title Exportador EVM
set "PATH=%~dp0;%PATH%"
evm_export --version
echo Carpeta de trabajo: %CD%
echo.
echo   evm_export --help        ayuda y comandos (generar, actualizar, analizar)
echo   copy "%~dp0config.ejemplo.toml" config.toml
echo                            configuración de ejemplo en esta carpeta
echo   Documentación: https://github.com/MarceloSalazar/evm-automation
echo.
cmd /k
