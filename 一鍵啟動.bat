@echo off
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% equ 0 (
  set "WC_PY=python"
) else (
  py -3 --version >nul 2>&1
  if %errorlevel% equ 0 (
    set "WC_PY=py -3"
  ) else (
    echo [錯誤] 找不到 python 或 py -3，請安裝 Python 並加入 PATH。
    pause
    exit /b 1
  )
)

echo.
echo 正在啟動「網站檢核工具」Streamlit 版…
echo 關閉黑色視窗即停止服務。
echo.

start "WebChecker Streamlit" cmd /k "cd /d ""%~dp0"" && %WC_PY% -m streamlit run app.py"

timeout /t 2 /nobreak >nul
