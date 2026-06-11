@echo off
title Liberar Firewall - Bolao Copa 2026 Pluma 8080
echo ===========================================
echo Liberando porta 8080 no Firewall do Windows
echo Execute este arquivo como Administrador
echo ===========================================
echo.
netsh advfirewall firewall add rule name="Bolao Copa 2026 Pluma 8080" dir=in action=allow protocol=TCP localport=8080
netsh advfirewall firewall add rule name="Bolao Copa 2026 Pluma 8080 Python" dir=in action=allow program="%LocalAppData%\Programs\Python\Python313\python.exe" enable=yes action=allow dir=in
netsh advfirewall firewall add rule name="Bolao Copa 2026 Pluma 8080 Python Launcher" dir=in action=allow program="%SystemRoot%\py.exe" enable=yes action=allow dir=in
echo.
echo Regra criada. Agora rode iniciar_site.bat e teste:
echo http://SEU_IP:8080
echo.
pause
