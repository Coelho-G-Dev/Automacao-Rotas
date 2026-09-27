@echo off
chcp 65001 > nul
title Roteirização de Frios - Logística
echo ===============================================================================
echo                EXECUTANDO ROTEIRIZAÇÃO DIÁRIA - SETOR DE FRIOS
echo ===============================================================================
echo.

py automacao.py

echo.
echo ===============================================================================
echo Processamento finalizado. Pressione qualquer tecla para fechar esta janela.
echo ===============================================================================
pause > nul
