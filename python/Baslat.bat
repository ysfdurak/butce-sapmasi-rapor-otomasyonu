@echo off
setlocal
cd /d "%~dp0"
py -3.13 --version >nul 2>&1
if not errorlevel 1 (
  py -3.13 bootstrap.py
  goto done
)
py -3.14 --version >nul 2>&1
if not errorlevel 1 (
  py -3.14 bootstrap.py
  goto done
)
py -3 --version >nul 2>&1
if not errorlevel 1 (
  py -3 bootstrap.py
  goto done
)
python --version >nul 2>&1
if not errorlevel 1 (
  python bootstrap.py
  goto done
)
echo.
echo  Python bulunamadi. Uygulama icin Python 3.9 veya uzeri gerekir (onerilen 3.13).
echo.
echo  Kurulum secenekleri:
echo    1) Sirket yazilim merkezinde (Software Center / Sirket Portali) "Python" arayin.
echo    2) Yonetici izniniz varsa komut isteminde:  winget install Python.Python.3.13
echo    3) https://www.python.org/downloads/ adresinden indirin; kurulumda
echo       "Add python.exe to PATH" kutusunu isaretleyin.
echo.
echo  Kurulumdan sonra bu dosyayi yeniden calistirin. Ek paket kurmaniz gerekmez.
echo  Python kurulamiyorsa ayni raporlari makro\ klasorundeki Excel surumuyle uretebilirsiniz.
echo.
pause
exit /b 1
:done
if errorlevel 1 pause
