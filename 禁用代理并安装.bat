@echo off
REM 禁用Windows系统代理
REM 这会清除系统级别的代理设置

echo 【正在禁用系统代理】

REM 使用netsh禁用WinHTTP代理
echo 清除WinHTTP代理...
netsh winhttp reset proxy
netsh winhttp set proxy proxy-server="direct" bypass-list="*"

echo.
echo 【正在尝试安装依赖包】
cd /d "%~dp0"

REM 激活虚拟环境并安装
call .venv\Scripts\activate.bat
pip install pdf2docx pymupdf python-docx

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✅ 安装成功！
    echo.
    echo 现在可以运行脚本：
    echo   python pdf_to_word.py
) else (
    echo.
    echo ❌ 安装仍然失败
    echo.
    echo 手动方案：
    echo 1. 打开 设置 Settings
    echo 2. 搜索 "代理" Proxy  
    echo 3. 禁用所有代理选项
    echo 4. 重启VS Code
    echo 5. 重新运行此脚本
    pause
)
