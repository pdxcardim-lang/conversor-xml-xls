@echo off
REM ====================================================
REM 🚀 EXECUTOR UNIVERSAL PYTHON
REM 
REM Copie este arquivo para qualquer pasta de projeto Python
REM Ele detecta automaticamente o tipo de projeto e executa
REM 
REM Funciona com:
REM - Flask (main.py, app.py, run.py)
REM - Django (manage.py)
REM - Scripts simples (.py)
REM - Jupyter notebooks (opcional)
REM ====================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo 🚀 ================================
echo    EXECUTOR UNIVERSAL PYTHON
echo    ================================
echo.

REM ===== VERIFICA AMBIENTE VIRTUAL =====
if not exist ".venv" (
    echo ❌ Ambiente virtual não encontrado!
    echo.
    echo 🔧 Deseja criar um ambiente virtual? ^(S/N^)
    set /p "criar="
    if /i "!criar!"=="S" (
        echo 📦 Criando ambiente virtual...
        python -m venv .venv
        if errorlevel 1 (
            echo ❌ Erro! Certifique-se que Python está instalado
            pause
            exit /b 1
        )
        echo ✅ Ambiente virtual criado!
        
        REM Instala dependências se existir requirements.txt
        if exist "requirements.txt" (
            echo 📦 Instalando dependências...
            .venv\Scripts\pip.exe install -r requirements.txt
        )
    ) else (
        echo ❌ Operação cancelada
        pause
        exit /b 1
    )
)

REM ===== DETECÇÃO AUTOMÁTICA DO PROJETO =====
set "PROJETO_TIPO="
set "ARQUIVO_EXEC="
set "COMANDO_EXEC="

echo 🔍 Detectando tipo de projeto...

if exist "manage.py" (
    set "PROJETO_TIPO=🌐 Django"
    set "ARQUIVO_EXEC=manage.py"
    set "COMANDO_EXEC=.venv\Scripts\python.exe manage.py runserver"
) else if exist "main.py" (
    set "PROJETO_TIPO=🔥 Flask/Python"
    set "ARQUIVO_EXEC=main.py"
    set "COMANDO_EXEC=.venv\Scripts\python.exe main.py"
) else if exist "app.py" (
    set "PROJETO_TIPO=🔥 Flask"
    set "ARQUIVO_EXEC=app.py"
    set "COMANDO_EXEC=.venv\Scripts\python.exe app.py"
) else if exist "run.py" (
    set "PROJETO_TIPO=🔥 Flask"
    set "ARQUIVO_EXEC=run.py"
    set "COMANDO_EXEC=.venv\Scripts\python.exe run.py"
) else (
    echo.
    echo 📄 Projetos específicos não detectados.
    echo    Arquivos Python disponíveis:
    echo.
    set contador=0
    for %%f in (*.py) do (
        set /a contador+=1
        echo    !contador!. %%f
        set "arquivo!contador!=%%f"
    )
    
    if !contador! equ 0 (
        echo    ❌ Nenhum arquivo Python encontrado!
        pause
        exit /b 1
    )
    
    echo.
    echo 🔢 Digite o número do arquivo para executar:
    set /p "escolha="
    
    if defined arquivo!escolha! (
        set "PROJETO_TIPO=🐍 Script Python"
        set "ARQUIVO_EXEC=!arquivo%escolha%!"
        set "COMANDO_EXEC=.venv\Scripts\python.exe !arquivo%escolha%!"
    ) else (
        echo ❌ Escolha inválida!
        pause
        exit /b 1
    )
)

REM ===== EXIBE INFORMAÇÕES =====
echo.
echo ✅ Projeto detectado: !PROJETO_TIPO!
echo 📁 Pasta: %CD%
echo 📄 Arquivo: !ARQUIVO_EXEC!
echo 🐍 Python: .venv\Scripts\python.exe
echo.

REM ===== EXECUÇÃO =====
echo 🚀 Iniciando...
echo    Comando: !COMANDO_EXEC!
echo.
echo ============================================
echo.

!COMANDO_EXEC!

echo.
echo ============================================
echo ⏹️  Execução finalizada.
echo.
pause