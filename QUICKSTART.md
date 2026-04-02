# Excel Compare Tool - 快速开始指南

## 📋 项目概述

Excel Compare Tool 是一个用于比较两个 Excel 文件差异的图形化工具。

**功能特点：**
- ✅ 支持多个 Sheet 对比
- ✅ 对比所有列和行
- ✅ 支持文件拖拽
- ✅ 可视化显示差异（绿色=新增，红色=删除，黄色=修改）
- ✅ 跨平台支持
- ✅ 可编译为 Windows 单文件 EXE

## 🚀 快速开始

### 本地开发运行（Windows/Linux）

```bash
cd excel-compare-tool
uv sync
uv run python src/excel_compare/main.py
```

### 编译为 Windows EXE

#### 方法1：在 Windows 上编译（推荐）

在 Windows 系统上打开 PowerShell 或 CMD：

```cmd
cd D:\project\xiongjing\excel-compare-tool
build_windows.bat
```

构建完成后，EXE 文件位于：`dist\ExcelCompare.exe`

#### 方法2：使用 spec 文件

```cmd
cd D:\project\xiongjing\excel-compare-tool
uv sync
uv run pyinstaller ExcelCompare.spec
```

## 📁 项目结构

```
excel-compare-tool/
├── src/
│   └── excel_compare/
│       ├── __init__.py       # 包初始化
│       ├── main.py           # GUI 主程序
│       ├── comparator.py      # Excel 对比核心逻辑
│       └── diff_view.py      # 差异显示组件
├── build.py                 # 通用打包脚本
├── build_windows.bat        # Windows 构建脚本
├── compile_all.py           # 跨平台编译脚本
├── ExcelCompare.spec       # PyInstaller 配置
├── pyproject.toml          # 项目配置
└── README.md               # 详细文档
```

## 🎯 使用方法

1. 启动应用后，会看到一个主窗口
2. 拖拽第一个 Excel 文件到 "File 1" 区域
3. 拖拽第二个 Excel 文件到 "File 2" 区域
4. 点击 "Start Comparison" 开始对比
5. 查看结果：
   - Summary 标签页：整体差异统计
   - 各个 Sheet 标签页：详细差异
   - 颜色标记：绿色（新增）、红色（删除）、黄色（修改）

## 🛠️ 技术栈

- **GUI**: tkinter + tkinterdnd2（拖拽支持）
- **Excel 处理**: pandas + openpyxl
- **打包工具**: PyInstaller
- **包管理**: uv

## 📦 依赖项

```
pandas>=2.0.0
openpyxl>=3.1.0
tkinterdnd2>=0.3.0
pyinstaller>=6.0.0
```

## ⚠️ 重要说明

### 关于交叉编译

在 WSL/Linux 环境下直接交叉编译为 Windows EXE 是非常困难的，因为：
- 需要 Windows 特定的系统库和运行时
- tkinterdnd2 在 Windows 上的依赖不同
- PyInstaller 主要是为原生平台设计的

**建议：** 在 Windows 系统上直接运行构建脚本以获得最佳结果。

### EXE 文件大小

生成的 EXE 文件通常在 20-50 MB 之间，因为包含了：
- Python 运行时
- 所有必需的库和依赖
- tkinter 和 GUI 组件

## 🐛 故障排除

### 问题：启动时显示缺少模块

**解决方案：** 在 spec 文件中添加到 `hiddenimports` 列表中

### 问题：拖拽功能不工作

**解决方案：** 确保正确包含了 `tkinterdnd2` 及其依赖

### 问题：构建失败

**解决方案：**
1. 确保安装了 Python 3.12+
2. 运行 `uv sync` 更新依赖
3. 查看错误日志进行调试

## 📞 支持

如果遇到问题，请检查：
1. Python 版本是否符合要求
2. 所有依赖是否正确安装
3. 构建日志中的错误信息

---

**最后更新：** 2026-04-02 (v0.1.1)
**版本：** 0.1.1
