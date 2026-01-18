# Universal Excel Merger Pro (Excel Merge Tool)

**Universal Excel Merger Pro** is a modern, web-based tool designed to solve the common headache of merging data between different Excel files. Think of it as a "Visual VLOOKUP" or "Batch Match & Fill" tool that allows you to easily transfer data from a **Source File** to a **Target File** based on a unique identifier (Key).

Whether you are a teacher merging grades, an HR professional updating employee records, or an analyst combining datasets, this tool streamlines the process with a clean, intuitive interface.

## 📸 Screenshots

![Home Page](assets/preview_home.png)
*Modern and Clean Interface*

![Merge Steps](assets/preview_steps.png)
*Easy Step-by-Step Configuration*

![Match Results](assets/preview_results.png)
*Real-time Data Preview & Validation*


## 🌟 Features

- **📂 Multi-Format Support**: Seamlessly handle `.xlsx`, `.xls`, and `.csv` files.
- **🔗 Smart Key Matching**: Link records between files using a unique key (e.g., Student ID, Phone Number, Product Code).
- **🗺️ Flexible Column Mapping**: easy-to-use interface to map multiple columns from source to target (e.g., "Fill 'Grade' in Target with 'Final Score' from Source").
- **👀 Data Preview**: Instant feedback on how many records matched successfully before you download.
- **📊 Sheet Support**: Select specific sheets from multi-sheet Excel workbooks.
- **🎨 Modern UI**: Features a beautiful "Glassmorphism" design for a premium user experience.

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- `pip` (Python package manager)

### Installation

1.  **Clone the repository** (or download the source code):
    ```bash
    git clone <repository-url>
    cd ExcelMergeTool
    ```

2.  **Install dependencies**:
    You will need `Flask` for the web server and `pandas` for data processing.
    ```bash
    pip install flask pandas openpyxl xlrd
    ```

### Usage

1.  **Start the application**:
    ```bash
    python app.py
    ```

2.  **Open in Browser**:
    Visit `http://localhost:5001` in your web browser.

3.  **Merge Data**:
    - **Step 1**: Upload your **Target File** (the file you want to edit) and **Source File** (the file containing new data).
    - **Step 2**: Select the match keys (e.g., "Student ID" in both files).
    - **Step 3**: Add mapping rules (e.g., Map "Midterm_Score" from Source -> "Midterm" in Target).
    - **Step 4**: Click **Preview** to see match stats, then **Merge & Download** to get your result.

### 📦 Packaging as EXE (Windows)

If you want to create a standalone `.exe` file to run this tool without installing Python everywhere:

1.  **Install PyInstaller**:
    ```bash
    pip install pyinstaller
    ```

2.  **Run Build Command**:
    Execute the following command in the project root directory:
    ```bash
    pyinstaller --noconfirm --clean --onefile --console --add-data "templates;templates" --add-data "static;static" --name "ExcelMergePro" app.py
    ```

3.  **Locate EXE**:
    The generated `ExcelMergePro.exe` will be in the `dist` folder.

### 🍎 Packaging for macOS (Apple Silicon M1/M2/M3 & Intel)

**Note**: You must run this **on a Mac**. You cannot build a Mac app from Windows.

1.  **Open Terminal** in the project folder.

2.  **Install PyInstaller**:
    ```bash
    pip3 install pyinstaller
    ```

3.  **Run Build Command**:
    (Note the use of `:` separator instead of `;`)
    ```bash
    pyinstaller --noconfirm --clean --onefile --windowed --add-data "templates:templates" --add-data "static:static" --name "ExcelMergePro" app.py
    ```

4.  **Locate App**:
    The generated `ExcelMergePro.app` will be in the `dist` folder.

## 📂 Project Structure

```
ExcelMergeTool/
├── app.py              # Main Flask application logic
├── templates/
│   └── index.html      # Frontend HTML interface
├── static/
│   └── style.css       # Stylesheets (Glassmorphism design)
├── uploads/            # Temporary directory for file processing
└── README.md           # Project documentation
```

## 🛠️ Technology Stack

- **Backend**: Python, Flask
- **Data Processing**: Pandas, OpenPyXL
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)

## 📝 License

This project is created for personal and educational use. Feel free to modify and improve it!

---

# Universal Excel Merger Pro (Excel 表格合并工具)

**Universal Excel Merger Pro** 是一个现代化的 Web 工具，旨在解决 Excel 文件之间数据合并的繁琐问题。您可以把它看作是一个“可视化 VLOOKUP”或“批量匹配填充”工具，它允许您基于唯一的标识符（Key）轻松地将数据从 **源文件** 传输到 **目标文件**。

无论您是合并成绩的教师、更新员工记录的 HR，还是合并数据集的分析师，此工具都能通过清晰直观的界面简化流程。

## 🌟 功能特性

- **📂 多格式支持**：无缝处理 `.xlsx`, `.xls` 和 `.csv` 文件。
- **🔗 智能主键匹配**：使用唯一键（如学号、手机号、产品代码）关联文件记录。
- **🗺️ 灵活的列映射**：简单易用的界面，可将源文件的多列映射到目标文件（例如：将源文件的 "Final Score" 填充到目标文件的 "成绩" 列）。
- **👀 数据预览**：在下载之前即时反馈匹配成功的记录数。
- **📊 多 Sheet 支持**：支持从多工作表 Excel 文件中选择特定 Sheet。
- **🎨 现代化 UI**：采用漂亮的“玻璃拟态 (Glassmorphism)”设计，提供优质的用户体验。

## 🚀 快速开始

### 环境要求

- Python 3.8+
- `pip` (Python 包管理器)

### 安装

1.  **克隆仓库** (或下载源码)：
    ```bash
    git clone <repository-url>
    cd ExcelMergeTool
    ```

2.  **安装依赖**：
    您需要安装 `Flask` 用于 Web 服务，以及 `pandas` 用于数据处理。
    ```bash
    pip install -r requirements.txt
    ```

### 使用方法

1.  **启动应用**：
    ```bash
    python app.py
    ```

2.  **在浏览器打开**：
    访问 `http://localhost:5001`。

3.  **合并数据**：
    - **第一步**：上传您的 **目标文件**（您想要修改的文件）和 **源文件**（包含新数据的文件）。
    - **第二步**：选择匹配主键（例如：两个文件中的 "学号"）。
    - **第三步**：添加映射规则（例如：将源文件的 "平时分" 映射 -> 目标文件的 "平时成绩"）。
    - **第四步**：点击 **预览** 查看匹配统计，然后点击 **确认合并并下载** 获取结果。

### 📦 打包为 EXE (Windows)

如果您想将工具打包成无需 Python 环境即可独立运行的 `.exe` 文件：

1.  **安装 PyInstaller**：
    ```bash
    pip install pyinstaller
    ```

2.  **执行打包命令**：
    在项目根目录下运行以下命令：
    ```bash
    pyinstaller --noconfirm --clean --onefile --console --add-data "templates;templates" --add-data "static;static" --name "ExcelMergePro" app.py
    ```

3.  **获取程序**：
    生成后的 `ExcelMergePro.exe` 文件位于 `dist` 文件夹中。

### 🍎 打包为 macOS 应用 (M1/M2/Intel)

**注意**: 这些步骤必须**在 Mac 电脑上运行**。您无法从 Windows 生成 Mac 应用程序。

1.  **打开终端 (Terminal)** 并进入项目目录。

2.  **安装 PyInstaller**：
    ```bash
    pip3 install pyinstaller
    ```

3.  **执行打包命令**：
    (注意这里使用冒号 `:` 分隔，而不是分号 `;`)
    ```bash
    pyinstaller --noconfirm --clean --onefile --windowed --add-data "templates:templates" --add-data "static:static" --name "ExcelMergePro" app.py
    ```

4.  **获取程序**：
    生成的 `ExcelMergePro.app` 将在 `dist` 文件夹中。

## 📂 项目结构

```
ExcelMergeTool/
├── app.py              # Flask 主程序逻辑
├── templates/
│   └── index.html      # 前端 HTML 界面
├── static/
│   └── style.css       # 样式表 (玻璃拟态设计)
├── uploads/            # 文件处理临时目录
└── README.md           # 项目文档
```

## 🛠️ 技术栈

- **后端**: Python, Flask
- **数据处理**: Pandas, OpenPyXL
- **前端**: HTML5, CSS3, JavaScript (原生)

## 📝 许可证

本项目仅供个人和教育使用。欢迎随意修改和改进！
