@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Contagem de Cartas - Abrindo interface...
echo.

python contagem_cartas_gui.py
if errorlevel 1 (
    py contagem_cartas_gui.py
    if errorlevel 1 (
        echo.
        echo Erro: Python nao encontrado ou falha ao executar.
        echo 1. Instale o Python em https://www.python.org/downloads/
        echo 2. Nesta pasta, execute: pip install -r requirements.txt
        echo.
        pause
    )
)
