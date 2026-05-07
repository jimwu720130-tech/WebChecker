@echo off
setlocal EnableExtensions
chcp 65001 >nul
rem %~dp0 尾端含 \，若寫成 ""%~dp0"" 在 start 內層會變成 \\" 而觸發「語法不正確」；尾端加 . 可避免
cd /d "%~dp0."

if not exist "%~dp0run_streamlit_here.bat" (
  echo [錯誤] 找不到同資料夾內的 run_streamlit_here.bat，請勿只複製本檔而漏了其他檔案。
  timeout /t 5 /nobreak >nul
  exit /b 1
)

echo 正在啟動「網站檢核小幫手」…
echo.
echo • 將另開一個標題為「WebChecker Streamlit」的黑色視窗，請勿關閉；請在該視窗內查看是否有紅色錯誤訊息。
echo • 若瀏覽器沒自動開啟，請到該視窗內顯示的本機網址手動開啟。
echo.
echo 本視窗將於約 3 秒後自動關閉（不必按鍵）。
echo.

start "WebChecker Streamlit" /D "%~dp0." cmd /k "call run_streamlit_here.bat"

timeout /t 3 /nobreak >nul
exit /b 0
