@echo off
cd /d %~dp0
title Bolao Copa 2026 Pluma - Porta 8080
echo ===========================================
echo  Bolao Copa 2026 Pluma - Servidor Local/Rede
echo ===========================================
echo.
echo Instalando dependencias, se necessario...
python -m pip install openpyxl

echo.
echo Iniciando servidor na porta 8080...
echo Acesse neste computador: http://localhost:8080
echo Para outros usuarios: use http://SEU_IP:8080
echo.
python app.py
pause
