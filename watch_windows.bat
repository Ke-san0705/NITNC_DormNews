@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo OneDriveで自動同期されるForms回答のExcelを指定してください。
set /p FILE=Excelファイルの完全パス: 
set /p VOL=号数: 
python generate.py --input "%FILE%" --volume %VOL% --design both --watch
pause
