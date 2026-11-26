@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ==========================================================
echo   MRP ve PARALEL PLANLAMA
echo ==========================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [HATA] Python bulunamadi.
    echo        https://www.python.org/downloads/ adresinden kurun
    echo        ve kurulumda "Add Python to PATH" secenegini isaretleyin.
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo openpyxl kutuphanesi kuruluyor...
    python -m pip install --quiet openpyxl
    if errorlevel 1 (
        echo [HATA] openpyxl kurulamadi.
        pause
        exit /b 1
    )
)

set "GIRDI=MRP_Girdi.xlsm"
if not exist "%GIRDI%" set "GIRDI=MRP_Girdi.xlsx"
if not exist "%GIRDI%" (
    echo [HATA] MRP_Girdi.xlsm / MRP_Girdi.xlsx bulunamadi.
    pause
    exit /b 1
)

echo [0/3] Cizelgeden kaydedilen ayar dosyalari toplaniyor...
python ayar_topla.py
echo.

echo [1/3] Ham ERP sayfalari plan girdisine cevriliyor...
echo.
python veri_aktar.py "%GIRDI%" MRP_Plan_Girdi.xlsx
if errorlevel 1 (
    echo.
    echo [HATA] Veri aktarimi basarisiz. Yukaridaki mesaji kontrol et.
    pause
    exit /b 1
)

echo.
echo [2/3] Plan olusturuluyor...
echo.
python mrp_planlama.py MRP_Plan_Girdi.xlsx MRP_Sonuc.xlsx
if errorlevel 1 (
    echo.
    echo [HATA] Plan olusturulamadi.
    pause
    exit /b 1
)

echo.
echo [3/3] Haftalik cizelge gorseli hazirlaniyor...
echo.
python cizelge_olustur.py MRP_Sonuc.xlsx MRP_Cizelge.html MRP_Plan_Girdi.xlsx
if errorlevel 1 (
    echo [UYARI] Cizelge olusturulamadi, plan dosyasi yine de hazir.
)

echo.
echo Aciliyor: MRP_Sonuc.xlsx + MRP_Cizelge.html
start "" "MRP_Sonuc.xlsx"
if exist "MRP_Cizelge.html" start "" "MRP_Cizelge.html"
pause
