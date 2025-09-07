@echo off
REM Ustawienie folderu, w którym znajduje się BAT
cd /d %~dp0

REM Uruchomienie skryptu w Pythonie
python operatorzy.py

pause
