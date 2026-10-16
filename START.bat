@echo off
chcp 1251 >nul
title Генератори звітів — запуск
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\launcher.ps1"
pause >nul
