@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

for /f %%v in ('python -c "from version import APP_VERSION as v; print(v)"') do set VER=%%v
echo === Dong goi ToolNHCS phien ban %VER% ===

python -m PyInstaller --clean --noconsole --onefile --name "ToolNHCS" ^
  --icon "assets\app.ico" --add-data "assets;assets" ^
  --add-data "tools;tools" --collect-submodules tools ^
  --hidden-import openpyxl --collect-all pdfplumber --collect-all pdfminer --collect-all pypdfium2 --collect-data docx ^
  launcher.py
if errorlevel 1 (
  echo.
  echo *** Build LOI. Xem thong bao phia tren. ***
  exit /b 1
)

powershell -NoProfile -Command "(Get-FileHash 'dist\ToolNHCS.exe' -Algorithm SHA256).Hash.ToLower() | Set-Content -Encoding ascii 'dist\ToolNHCS.exe.sha256'"

echo.
echo === XONG phien ban %VER% ===
echo Tao Release tren GitHub voi tag  v%VER%  va dinh kem 2 file:
echo     dist\ToolNHCS.exe
echo     dist\ToolNHCS.exe.sha256
echo.
echo Hoac dung GitHub CLI:
echo     gh release create v%VER% dist\ToolNHCS.exe dist\ToolNHCS.exe.sha256 --title "v%VER%" --notes "Noi dung thay doi"
endlocal
