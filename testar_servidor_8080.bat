@echo off
echo Verificando porta 8080...
netstat -aon | findstr :8080
echo.
echo Se aparecer 0.0.0.0:8080 LISTENING, o servidor esta ouvindo na rede.
pause
