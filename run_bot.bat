@echo off
REM Быстрый запуск бота автозаполнения Google формы

REM === Настройки (можно менять) ===
set NODE_FILE=form_bot.js
set START_ROW=0
set MAX_ROWS=
set HEADLESS=false

REM === Запуск ===
echo Запуск бота...
node %NODE_FILE% %START_ROW% %MAX_ROWS% %HEADLESS%
pause 