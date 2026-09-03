@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo.
echo  자금판 자동화 - 창이 열립니다.
echo  (이 검은 창을 닫으면 종료됩니다)
echo.
python -m streamlit run app.py
pause
