#!/bin/bash

# 退出脚本如果任何命令失败
set -e

echo "🚀 开始优化打包流程..."

# 1. 清理旧的构建文件
echo "🧹 清理旧的构建产物..."
rm -rf build dist venv_pack ExcelMergePro.spec

# 2. 创建全新的纯净虚拟环境
echo "📦 创建纯净虚拟环境 (venv_pack)..."
python3 -m venv venv_pack

# 3. 安装运行和打包所需的最小依赖 (严禁安装 pandas/numpy)
echo "📥 安装最小依赖项..."
./venv_pack/bin/python -m pip install --upgrade pip
./venv_pack/bin/python -m pip install flask openpyxl xlrd pyinstaller

# 4. 执行打包命令
# 我们在命令中显式排除了所有大库，以防万一
echo "🏗️ 正在打包应用 (这可能需要几分钟)..."
./venv_pack/bin/pyinstaller --noconfirm --clean --onefile --windowed \
  --add-data "templates:templates" \
  --add-data "static:static" \
  --exclude-module tkinter \
  --exclude-module tcl \
  --exclude-module tk \
  --exclude-module matplotlib \
  --exclude-module numpy \
  --exclude-module scipy \
  --exclude-module pandas \
  --exclude-module notebook \
  --icon "icon.icns" \
  --name "ExcelMergePro" app.py

echo "✅ 打包完成！"
echo "📂 可执行文件位于: $(pwd)/dist/ExcelMergePro"
echo "📊 请检查 dist 目录下的应用体积。"
