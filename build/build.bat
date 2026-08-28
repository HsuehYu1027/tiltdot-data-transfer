@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

echo ==========================================
echo  tiltdot-data-transfer 打包程序
echo ==========================================

REM 讓 Chromium 裝進 playwright 套件目錄，PyInstaller 才收得到
set PLAYWRIGHT_BROWSERS_PATH=0

echo [1/5] 安裝相依套件...
python -m pip install --upgrade pip || goto :fail
python -m pip install -r requirements.txt || goto :fail
python -m pip install pyinstaller || goto :fail

echo [2/5] 下載 Chromium（會佔數百 MB，第一次較久）...
python -m playwright install chromium || goto :fail

echo [3/5] 執行離線自我驗證...
python tools\selftest.py || goto :fail

echo [4/5] 清除舊的產出...
if exist dist\tiltdot rmdir /s /q dist\tiltdot

echo [5/5] 打包中...
python -m PyInstaller --noconfirm --clean build\tiltdot.spec || goto :fail

echo.
echo ==========================================
echo  完成！產出位置：dist\tiltdot\tiltdot.exe
echo  接著可用 Inno Setup 編譯 build\installer.iss 產生安裝檔。
echo ==========================================
goto :eof

:fail
echo.
echo [失敗] 打包中止，請往上捲動查看錯誤訊息。
exit /b 1
