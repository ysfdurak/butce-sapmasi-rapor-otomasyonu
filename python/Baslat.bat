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
echo Python 3.9 veya uzeri bulunamadi. Python 3.13 / 3.14 kurulumunu kontrol edin.
pause
exit /b 1
:done
if errorlevel 1 pause
