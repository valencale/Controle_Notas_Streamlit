@echo off
TITLE Sistema de Gestão Logística - Armazém
echo ===================================================
echo     INICIANDO SISTEMA DE GESTÃO LOGÍSTICA
echo ===================================================
echo.

:: Verifica se o ambiente virtual existe
if not exist "armazem_venv\" (
    echo [ERRO] Ambiente virtual não encontrado!
    echo Criando ambiente e instalando dependências...
    python -m venv armazem_venv
    call armazem_venv\Scripts\activate.bat
    pip install -r requirements.txt
) else (
    echo [INFO] Ativando ambiente virtual...
    call armazem_venv\Scripts\activate.bat
)

echo [INFO] Iniciando o Streamlit...
streamlit run app.py

pause
