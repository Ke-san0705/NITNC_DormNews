@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Microsoft Forms の回答.xlsxをこのフォルダに置いてください。
echo 初めに pip install -r requirements.txt を実行してください。
set /p FILE=FormsのExcelファイル名（例：回答.xlsx）: 
set /p VOL=新聞の号数（例：5）: 
python generate.py --input "%FILE%" --volume %VOL% --design both
pause
