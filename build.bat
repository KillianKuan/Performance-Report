@echo off
chcp 65001 > nul
echo ============================================
echo  SalesReportTool 打包腳本
echo ============================================

echo [1/4] 安裝依賴套件...
pip install -r requirements.txt
pip install pyinstaller

echo [2/4] 清理舊檔案...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist launcher.spec del launcher.spec

echo [3/4] 打包中（需要幾分鐘）...
pyinstaller ^
    --name "啟動程式" ^
    --onedir ^
    --noconsole ^
    --icon NONE ^
    --add-data "app;app" ^
    launcher.py

echo [4/4] 整理輸出資料夾...
xcopy /e /i /y app dist\啟動程式\app

echo.
echo ============================================
echo  打包完成！輸出位置：dist\啟動程式\
echo  將整個資料夾交給使用者即可。
echo ============================================
pause
