# Excel File Compare Tool

一个用于比较两个Excel文件差异的Python图形化工具。

## 功能特性

- 支持对比两个.xlsx/.xls文件
- 支持多个sheets对比
- 对比所有列和行的差异
- 支持文件拖拽操作
- 可视化显示差异结果
- 跨平台编译为Windows exe单文件

## 安装与运行

### 环境要求

- Python 3.12 或更高版本
- uv (Python包管理器)

### 本地开发运行

```bash
# 克隆或进入项目目录
cd excel-compare-tool

# 安装依赖
uv sync

# 运行应用
uv run python src/excel_compare/main.py
```

## 打包为Windows EXE

### 使用PyInstaller打包

```bash
# 安装依赖（如果尚未安装）
uv sync

# 运行打包脚本
uv run python build.py
```

打包完成后，可执行文件将位于 `dist/ExcelCompare.exe`

### 手动打包命令

```bash
uv run pyinstaller --onefile --windowed --name ExcelCompare \
  --add-data "src/excel_compare;excel_compare" \
  --hidden-import "tkinterdnd2" \
  src/excel_compare/main.py
```

## 使用方法

1. 启动应用程序后，会看到一个主窗口
2. 拖拽第一个Excel文件到"File 1"区域，或点击区域选择文件
3. 拖拽第二个Excel文件到"File 2"区域，或点击区域选择文件
4. 点击"Start Comparison"按钮开始对比
5. 查看结果：
   - Summary标签页显示整体差异统计
   - 各个sheet标签页显示详细差异
   - 绿色：新增内容
   - 红色：删除内容
   - 黄色：修改内容

## 项目结构

```
excel-compare-tool/
├── src/
│   └── excel_compare/
│       ├── __init__.py
│       ├── main.py         # GUI主程序
│       ├── comparator.py    # Excel对比核心逻辑
│       └── diff_view.py    # 差异显示组件
├── build.py                # PyInstaller打包脚本
├── pyproject.toml          # 项目配置
└── README.md
```

## 技术栈

- **GUI框架**: tkinter (内置) + tkinterdnd2 (拖拽支持)
- **Excel处理**: pandas + openpyxl
- **打包工具**: PyInstaller
- **包管理**: uv

## 注意事项

- 拖拽功能在打包为exe时需要正确包含tkinterdnd2的依赖
- 在Linux/WSL环境下测试GUI可能需要X服务器支持
- 打包的exe文件较大，因为包含了Python运行时和所有依赖
