@echo off
echo Este arquivo precisa ser executado como ADMINISTRADOR.
echo Liberando porta 8000 no Firewall do Windows...
netsh advfirewall firewall add rule name="Bolao Copa 2026 Pluma Porta 8000" dir=in action=allow protocol=TCP localport=8000
pause
