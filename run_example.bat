@echo off
echo CSV to XLS Template Converter
echo --------------------------
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Please install Python 3.6 or higher.
    pause
    exit /b 1
)

REM Install dependencies if needed
echo Installing required dependencies...
pip install -r requirements.txt

REM Run the conversion
echo.
echo Converting products_export_1.csv to XLS format...
python csv_to_xls.py "products_export_1 (1).csv"

echo.
if %errorlevel% equ 0 (
    echo Conversion completed successfully!
) else (
    echo Conversion failed. Please check the error messages above.
)

echo.
pause