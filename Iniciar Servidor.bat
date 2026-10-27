@echo off
chcp 65001 >nul
title Servidor - Distribuicao Cerveja (porta 8098)
cd /d "%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (python servidor.py) else (py servidor.py)
echo.
echo O servidor parou. Feche esta janela ou pressione uma tecla para reiniciar.
pause >nul
"%~f0"
