@echo off
cd /d "%~dp0"
rem Подтянуть свежую версию сайта с GitHub (если установлен git); без интернета или git — запустится текущая версия
git pull --ff-only --quiet 2>nul
python server.py
pause
