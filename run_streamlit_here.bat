@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0."

rem 不使用 goto。Streamlit 結束後不強制 pause（Ctrl+C 時 errorlevel 也常非 0，易誤判）。

python --version >nul 2>&1
if not errorlevel 1 (
  python -m streamlit run app.py
) else (
  py -3 --version >nul 2>&1
  if not errorlevel 1 (
    py -3 -m streamlit run app.py
  ) else (
    echo.
    echo [錯誤] 找不到 python 或 py -3，請安裝 Python 3 並勾選「加入 PATH」。
    echo.
    pause
    exit /b 1
  )
)

rem 不關閉視窗：cmd /k 會留在命令提示字元，方便複製上方錯誤訊息。
exit /b 0
