@echo off
cd /d C:\Users\Administrator\Desktop\codewhale
echo ============================================
echo   AC/AS ????
echo ============================================
echo.
python translator.py --verbose -t "C:\Users\Administrator\Desktop\Pet????.xlsx" -o "C:\Users\Administrator\Desktop\pet6.25.xls" -r "C:\Users\Administrator\Desktop\result_test.xlsx"
echo.
echo ============================================
echo   ?????????????...
echo ============================================
start "" "C:\Users\Administrator\Desktop\result_test.xlsx"
pause
