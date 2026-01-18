@echo off
chcp 65001 >nul
echo ==========================================
echo 🚀 开始 Windows 一键打包流程
echo ==========================================

:: 设置虚拟环境名称
set VENV_NAME=venv_win_pack

:: 1. 检查并创建纯净虚拟环境
if not exist "%VENV_NAME%" (
    echo 📦 正在创建纯净虚拟环境 [%VENV_NAME%]...
    python -m venv %VENV_NAME%
    if errorlevel 1 (
        echo ❌ 创建虚拟环境失败，请检查 python 是否在环境变量中。
        pause
        exit /b 1
    )
) else (
    echo ℹ️  发现已有虚拟环境，将直接使用...
)

:: 2. 激活环境
call %VENV_NAME%\Scripts\activate

:: 3. 安装最小依赖
echo 📥 正在安装/更新最小依赖 (flask, openpyxl, xlrd, pyinstaller, Pillow)...
pip install --disable-pip-version-check --upgrade pip
pip install flask openpyxl xlrd pyinstaller Pillow

:: 4. 执行打包
echo 🏗️  开始打包应用...
echo ------------------------------------------

if exist ExcelMergePro.spec (
    echo 📄 检测到配置文件 ExcelMergePro.spec，正在使用...
    pyinstaller ExcelMergePro.spec --clean --noconfirm
) else (
    echo ⚠️ 未找到 .spec 文件，正在使用命令行参数打包...
    :: 注意：Windows下 add-data 使用分号 ; 分隔
    pyinstaller --noconfirm --clean --onefile --windowed ^
      --add-data "templates;templates" ^
      --add-data "static;static" ^
      --exclude-module tkinter ^
      --exclude-module tcl ^
      --exclude-module tk ^
      --exclude-module matplotlib ^
      --exclude-module numpy ^
      --exclude-module scipy ^
      --exclude-module pandas ^
      --exclude-module notebook ^
      --icon "icon.icns" ^
      --name "ExcelMergePro" ^
      app.py
)

if errorlevel 1 (
    echo ❌ 打包失败！请检查上方错误信息。
    pause
    exit /b 1
)

echo ==========================================
echo ✅ 打包成功！
echo 📂 文件路径: %~dp0dist\ExcelMergePro.exe
echo ==========================================
pause
