# Windows 构建指南

## 快速开始

### 方法1：使用批处理脚本（推荐）

在 Windows PowerShell 或 CMD 中运行：

```cmd
cd D:\project\xiongjing\excel-compare-tool
build_windows.bat
```

### 方法2：使用 PyInstaller spec 文件

```cmd
cd D:\project\xiongjing\excel-compare-tool
uv sync
uv run pyinstaller ExcelCompare.spec
```

### 方法3：直接使用 PyInstaller 命令

```cmd
cd D:\project\xiongjing\excel-compare-tool
uv sync
uv run pyinstaller --onefile --windowed --name ExcelCompare --add-data "src\excel_compare;excel_compare" --hidden-import "tkinterdnd2" src\excel_compare\main.py
```

## 前置要求

### 安装 uv

```cmd
pip install uv
```

或者从官网下载：https://github.com/astral-sh/uv

### Python 版本

需要 Python 3.12 或更高版本。

## 构建结果

成功构建后，可执行文件将位于：
```
dist\ExcelCompare.exe
```

这是一个单文件可执行程序，无需安装 Python 即可在任何 Windows 机器上运行。

## 构建选项

### 调试版本（显示控制台）

如果需要查看调试信息，可以将 `--windowed` 参数改为 `--console` 或删除该参数：

```cmd
uv run pyinstaller --onefile --console --name ExcelCompare --add-data "src\excel_compare;excel_compare" --hidden-import "tkinterdnd2" src\excel_compare\main.py
```

### 减小文件大小

如果文件太大，可以在 `ExcelCompare.spec` 中添加排除项：

```python
excludes=[
    'numpy.testing',
    'scipy.testing',
    'pandas.testing',
    'pandas.tests',
    'matplotlib',
    # 添加更多不需要的模块
]
```

## 故障排除

### 问题1：tkinterdnd2 找不到

确保在 `hiddenimports` 中包含了 `tkinterdnd2`。

### 问题2：拖拽功能不工作

检查是否正确包含了 `tkinterdnd2` 和相关的 `.dll` 文件。

### 问题3：文件太大

正常的文件大小在 20-50 MB 之间，因为包含了 Python 运行时和所有依赖。

### 问题4：运行时找不到模块

使用 `--debug all` 参数运行 PyInstaller 以获取详细错误信息：

```cmd
uv run pyinstaller --debug all --onefile --windowed --name ExcelCompare --add-data "src\excel_compare;excel_compare" --hidden-import "tkinterdnd2" src\excel_compare\main.py
```

## 测试构建结果

1. 双击 `dist\ExcelCompare.exe`
2. 拖拽两个 Excel 文件到相应区域
3. 点击"Start Comparison"
4. 检查差异是否正确显示

## 分发

`ExcelCompare.exe` 是独立可执行文件，可以直接：
- 通过电子邮件发送
- 上传到文件共享服务
- 放入 U 盘或其他存储设备

无需附带任何其他文件。
