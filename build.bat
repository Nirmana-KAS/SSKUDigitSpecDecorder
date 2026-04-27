@echo off
echo ==========================================
echo   Building ORVA SSKU Code Decoder EXE
echo ==========================================
echo.

echo Installing dependencies...
pip install pandas openpyxl Pillow pyinstaller
echo.

echo Converting icon.png to icon.ico...
python -c "from PIL import Image; Image.open('icon.png').save('icon.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])"
echo.

echo Building EXE (this may take a few minutes)...
pyinstaller --noconfirm --onefile --windowed ^
    --name "ORVA_SSKU_Code_Decoder" ^
    --icon "icon.ico" ^
    --add-data "logo.png;." ^
    --add-data "icon.png;." ^
    --add-data "database.py;." ^
    --add-data "decoder.py;." ^
    main.py

echo.
echo ==========================================
if exist "dist\ORVA_SSKU_Code_Decoder.exe" (
    echo   BUILD SUCCESSFUL!
    echo   EXE location: dist\ORVA_SSKU_Code_Decoder.exe
) else (
    echo   BUILD FAILED - check errors above
)
echo ==========================================
pause
