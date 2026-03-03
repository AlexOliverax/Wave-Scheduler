@echo off
REM ============================================================================
REM Script de Recompilação - Waves Scheduler
REM Usa %~dp0 para localizar o projeto relativo a este script
REM ============================================================================
setlocal

SET PROJECT_DIR=%~dp0
echo.
echo ========================================
echo  Waves Scheduler - Recompilacao
echo ========================================
echo.
echo Diretorio do projeto: %PROJECT_DIR%
echo.

REM Navegção para o diretório do projeto
cd /d "%PROJECT_DIR%"
if errorlevel 1 (
    echo ERRO: Nao foi possivel acessar o diretorio do projeto!
    pause
    exit /b 1
)

echo [1/5] Ativando ambiente virtual...
IF EXIST "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) ELSE IF EXIST ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) ELSE (
    echo AVISO: Ambiente virtual nao encontrado. Usando Python do sistema.
)
echo.

echo [2/5] Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo Build anterior removido.
echo.

echo [3/5] Verificando dependencias...
pip install --quiet pyqt5 pandas openpyxl pytz holidays pywin32 numpy
echo Dependencias verificadas.
echo.

echo [4/5] Compilando com PyInstaller...
echo Isso pode levar alguns minutos...
pyinstaller "Waves Scheduler.spec"
if errorlevel 1 (
    echo.
    echo ERRO: Falha na compilacao!
    pause
    exit /b 1
)
echo.

echo [5/5] Verificando arquivos gerados...
if exist "dist\Waves Scheduler\Waves Scheduler.exe" (
    echo.
    echo ========================================
    echo  COMPILACAO CONCLUIDA COM SUCESSO!
    echo ========================================
    echo.
    echo Executavel gerado em:
    echo %PROJECT_DIR%dist\Waves Scheduler\Waves Scheduler.exe
    echo.
    echo Deseja testar o executavel agora? (S/N)
    choice /c SN /n
    if errorlevel 2 goto fim
    if errorlevel 1 goto testar
) else (
    echo.
    echo ERRO: Executavel nao foi gerado!
    pause
    exit /b 1
)

:testar
echo.
echo Iniciando Waves Scheduler...
start "" "dist\Waves Scheduler\Waves Scheduler.exe"
goto fim

:fim
endlocal
echo.
echo Pressione qualquer tecla para sair...
pause >nul