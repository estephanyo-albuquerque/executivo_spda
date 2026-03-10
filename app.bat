@echo off
title ArthWind Dashboard Launcher
color 0A

echo ==================================================
echo      INICIANDO DASHBOARD ARTHWIND
echo ==================================================
echo.

REM 1. Tenta instalar bibliotecas automaticamente
echo [1/2] Verificando bibliotecas necessarias...
pip install streamlit pandas matplotlib seaborn altair openpyxl fpdf
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERRO] Nao foi possivel instalar as bibliotecas.
    echo Verifique sua internet ou instalacao do Python.
    echo.
    pause
    exit
)

echo.
echo [2/2] Bibliotecas OK. Iniciando o sistema...
echo.
echo --------------------------------------------------
echo O navegador vai abrir em instantes.
echo Para fechar, apenas feche esta janela.
echo --------------------------------------------------
echo.

REM 2. Roda o aplicativo
streamlit run app.py

REM 3. Se o streamlit fechar por erro, o pause segura a tela
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ==================================================
    echo        OCORREU UM ERRO FATAL NO PYTHON
    echo ==================================================
    echo Leia a mensagem de erro acima (geralmente em vermelho).
    echo.
)

pause