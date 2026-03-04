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

REM Navegação para o diretório do projeto
cd /d "%PROJECT_DIR%"
if errorlevel 1 (
    echo ERRO: Nao foi possivel acessar o diretorio do projeto!
    pause
    exit /b 1
)

REM Ler versão atual do arquivo VERSION
SET /P APP_VERSION=<VERSION
echo Versao: %APP_VERSION%
echo.

echo [1/6] Ativando ambiente virtual...
IF EXIST "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) ELSE IF EXIST ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) ELSE (
    echo AVISO: Ambiente virtual nao encontrado. Usando Python do sistema.
)
echo.

echo [2/6] Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo Build anterior removido.
echo.

echo [3/6] Verificando dependencias...
pip install --quiet pyqt5 pandas openpyxl pytz holidays pywin32 numpy
echo Dependencias verificadas.
echo.

echo [4/6] Compilando com PyInstaller...
echo Isso pode levar alguns minutos...
pyinstaller "Waves Scheduler.spec"
if errorlevel 1 (
    echo.
    echo ERRO: Falha na compilacao PyInstaller!
    pause
    exit /b 1
)
echo.

echo [5/6] Verificando executavel gerado...
if not exist "dist\Waves Scheduler\Waves Scheduler.exe" (
    echo.
    echo ERRO: Executavel nao foi gerado!
    pause
    exit /b 1
)
echo Executavel OK: dist\Waves Scheduler\Waves Scheduler.exe
echo.

echo [6/6] Gerando instalador com Inno Setup...
REM Procurar Inno Setup em locais padrão
SET ISCC=""
IF EXIST "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" SET ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
IF EXIST "C:\Program Files\Inno Setup 6\ISCC.exe"       SET ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
IF EXIST "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" SET ISCC="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"

IF %ISCC%=="" (
    echo AVISO: Inno Setup nao encontrado. Pulando geracao do instalador.
    echo        Instale em: https://jrsoftware.org/isinfo.php
) ELSE (
    %ISCC% "WavesScheduler_Setup.iss"
    if errorlevel 1 (
        echo.
        echo ERRO: Falha ao gerar o instalador!
        pause
        exit /b 1
    )
    echo.
    echo Instalador gerado em:
    echo %PROJECT_DIR%installer\WavesScheduler_Setup_v%APP_VERSION%.exe
)
echo.

echo ========================================
echo  COMPILACAO CONCLUIDA COM SUCESSO!
echo  Versao: %APP_VERSION%
echo ========================================
echo.

echo Deseja testar o executavel agora? (S/N)
choice /c SN /n
if errorlevel 2 goto fim
if errorlevel 1 goto testar

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