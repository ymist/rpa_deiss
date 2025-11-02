@echo off
chcp 65001 >nul
title RPA Notas Fiscais - Executar

:: Verifica se setup já foi executado
if not exist "setup_concluido.txt" (
    echo ========================================
    echo    PRIMEIRA EXECUÇÃO DETECTADA
    echo ========================================
    echo.
    echo 🔧 Executando configuração inicial...
    echo.
    call setup_python.bat

    :: Cria arquivo de controle
    echo Setup executado em %date% %time% > setup_concluido.txt
    echo.
)

cls
echo ========================================
echo    RPA NOTAS FISCAIS - INICIANDO
echo ========================================
echo.

:: Verifica se Python está disponível
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python não encontrado!
    echo 🔧 Execute setup_python.bat primeiro
    pause
    exit /b 1
)

:: Verifica se o arquivo Python existe
if not exist "rpa_notas_fiscais.py" (
    echo ❌ Arquivo rpa_notas_fiscais.py não encontrado!
    echo 📁 Certifique-se de que todos os arquivos estão na mesma pasta
    pause
    exit /b 1
)

:: Verifica se existe arquivo Excel
set "excel_encontrado=0"
for %%f in (*.xlsx *.xls) do (
    set "excel_encontrado=1"
    echo ✅ Arquivo Excel encontrado: %%f
)

if %excel_encontrado%==0 (
    echo ⚠️  AVISO: Nenhum arquivo Excel (.xlsx/.xls) encontrado na pasta
    echo.
    echo 📋 O programa irá perguntar qual arquivo usar durante a execução
    echo.
)

echo.
echo 🚀 Iniciando RPA...
echo.
echo 📋 Instruções:
echo    • Mantenha o Chrome atualizado
echo    • Faça login no site antes de iniciar
echo    • Não feche esta janela durante a execução
echo.
echo ========================================
echo.

:: Executa o RPA
python rpa_notas_fiscais.py

:: Verifica se houve erro
if %errorlevel% neq 0 (
    echo.
    echo ❌ Erro na execução do RPA
    echo 📋 Verifique os logs acima para mais detalhes
) else (
    echo.
    echo ✅ RPA executado com sucesso!
)

echo.
echo ========================================
pause