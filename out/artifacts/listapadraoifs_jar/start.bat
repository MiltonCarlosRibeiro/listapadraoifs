@echo off
cd /d "%~dp0"
start "" javaw -jar listapadraoifs.jar
timeout /t 3 >nul
start http://localhost:8085
exit
