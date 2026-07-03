@echo off
chcp 65001 >nul
echo ==============================================
echo BÃ¡ÂºÂ¯t Ã„â€˜Ã¡ÂºÂ§u Ã„â€˜ÃƒÂ³ng gÃƒÂ³i DEANTRANS V1.4.3
echo ==============================================

echo [1] Kiem tra va cai dat PyInstaller...
pip install pyinstaller

echo.
echo [2] Dang dong goi DEANTRANS bang PyInstaller (ONEFILE)...
:: --noconfirm: ghi de thu muc build cu
:: --windowed: an man hinh Ã„â€˜en console window khi chay app
pyinstaller --name "DEANTRANS" --windowed --onefile --noconfirm run.py

echo.
echo [3] Don dep va sap xep thu muc phat hanh...
:: Tao thu muc phat hanh
IF EXIST "DEANTRANS_V1.4.3" rmdir /S /Q "DEANTRANS_V1.4.3"
mkdir "DEANTRANS_V1.4.3"

:: Move exe sang thu muc phat hanh
move "dist\DEANTRANS.exe" "DEANTRANS_V1.4.3\" >nul

:: Tao thu muc config vi app can chung, copy nguyen configs/ cu sang
mkdir "DEANTRANS_V1.4.3\configs"
xcopy /E /I /Y "configs" "DEANTRANS_V1.4.3\configs" >nul

:: Copy tai lieu
copy /Y "README.txt" "DEANTRANS_V1.4.3\" >nul
copy /Y "CHANGELOG.txt" "DEANTRANS_V1.4.3\" >nul

:: Xoa cac thu muc build temp (neu muon giu ranh may, co the bo pycache)
rmdir /S /Q build
rmdir /S /Q dist

echo ==============================================
echo [HOAN THANH!] 
echo.
echo Ung dung cua ban da duoc dong goi tai thu muc: DEANTRANS_V1.4.3
echo Ban co extreme the nap/zip thu muc nay va gui cho nguoi khac su dung.
echo.
