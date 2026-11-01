@echo off
chcp 65001 >nul
title 🤖 Маршрут-бот

echo ==============================
echo   Маршрут-бот по Астане
echo ==============================
echo.

cd /d "%~dp0"

echo [1/2] Обновление кода с GitHub...
git pull
if %errorlevel% neq 0 (
    echo ПРЕДУПРЕЖДЕНИЕ: git pull не удался, запускаю текущую версию...
)

echo.
echo [2/2] Запуск бота...
echo.
.venv\Scripts\python.exe tg_bot.py

echo.
echo Бот остановлен. Нажмите любую клавишу для выхода.
pause >nul

