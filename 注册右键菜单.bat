@echo off
chcp 65001 >nul
rem 注册右键菜单：把「DLP 一键解密」加到 Excel/Word/PPT/PDF 文件的右键菜单
set "EXE=%~dp0DLP批量解密工具.exe"
if not exist "%EXE%" (
    echo [错误] 未找到 DLP批量解密工具.exe，请把本脚本与 exe 放在同一目录。
    pause
    exit /b 1
)
"%EXE%" --register
