"""差异显示组件。"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Optional
from .comparator import SheetDiff, RowDiff, CellDiff, ColumnDiff


class DiffViewer(ttk.Frame):
    """用于显示Excel文件差异的框架。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.diff_results = None
        self.current_sheet_index = 0
        self._setup_ui()

    def _setup_ui(self):
        """设置UI组件。"""
        # 创建sheet的notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 创建初始空标签页
        self.empty_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.empty_tab, text="结果")

    def show_differences(self, diff_results) -> None:
        """
        显示对比结果。

        Args:
            diff_results: 包含差异的CompareResult对象
        """
        self.diff_results = diff_results
        self.current_sheet_index = 0

        # 清除现有标签页
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        # 首先显示摘要
        self._show_summary(diff_results)

        # 显示每个sheet的差异
        if diff_results.sheet_diffs:
            for sheet_diff in diff_results.sheet_diffs:
                self._add_sheet_tab(sheet_diff)
        else:
            self._show_no_differences()

    def _show_summary(self, diff_results) -> None:
        """显示所有差异的摘要。"""
        summary = diff_results.summary

        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="摘要")

        # 创建摘要文本控件
        text = tk.Text(frame, wrap=tk.WORD, padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True)

        # 添加摘要内容
        summary_text = f"""Excel对比摘要
{'=' * 50}

对比的文件：
  文件1: {diff_results.file1}
  文件2: {diff_results.file2}

总体统计：
  有差异的Sheet总数: {summary['total_sheets']}
    - 新增的Sheet: {summary['added_sheets']}
    - 删除的Sheet: {summary['deleted_sheets']}
    - 修改的Sheet: {summary['modified_sheets']}

  列差异：
    - 新增的列: {summary['added_columns']}
    - 删除的列: {summary['deleted_columns']}

  行差异：
    - 新增的行: {summary['added_rows']}
    - 删除的行: {summary['deleted_rows']}
    - 修改的单元格: {summary['modified_cells']}

  总差异数: {summary['total_differences']}
"""
        text.insert(tk.END, summary_text)
        text.config(state=tk.DISABLED)

    def _add_sheet_tab(self, sheet_diff: SheetDiff) -> None:
        """
        添加一个sheet的差异标签页。

        Args:
            sheet_diff: 要显示的SheetDiff对象
        """
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=sheet_diff.sheet_name)

        # 主容器
        container = ttk.Frame(frame)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Sheet类型指示器
        type_label = ttk.Label(
            container,
            text=f"Sheet: {sheet_diff.sheet_name} ({sheet_diff.diff_type.upper()})",
            font=('Microsoft YaHei UI', 10, 'bold')
        )
        type_label.pack(anchor=tk.W, pady=(0, 10))

        # 创建差异的可滚动框架
        scroll_frame = ttk.Frame(container)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        # 滚动条
        vsb = ttk.Scrollbar(scroll_frame, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        hsb = ttk.Scrollbar(scroll_frame, orient="horizontal")
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # 内容画布
        canvas = tk.Canvas(
            scroll_frame,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            bg='white'
        )
        canvas.pack(fill=tk.BOTH, expand=True)

        vsb.config(command=canvas.yview)
        hsb.config(command=canvas.xview)

        # 内容的内部框架
        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # 配置画布以调整内部框架大小
        inner_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        ))

        # 填充差异
        self._populate_sheet_diffs(inner_frame, sheet_diff)

    def _populate_sheet_diffs(self, parent: ttk.Frame, sheet_diff: SheetDiff) -> None:
        """
        用sheet差异填充框架。

        Args:
            parent: 要添加小部件的父框架
            sheet_diff: 要显示的SheetDiff对象
        """
        row_idx = 0

        # 显示列差异
        if sheet_diff.column_diffs:
            col_frame = ttk.LabelFrame(parent, text="列差异", padding=5)
            col_frame.pack(fill=tk.X, pady=5)

            headers = ["类型", "列名"]
            for col, header in enumerate(headers):
                lbl = ttk.Label(col_frame, text=header, font=('Microsoft YaHei UI', 9, 'bold'))
                lbl.grid(row=0, column=col, padx=5, pady=2, sticky=tk.W)

            for i, col_diff in enumerate(sheet_diff.column_diffs, start=1):
                color = '#90EE90' if col_diff.diff_type == 'added' else '#FFB6C1'

                type_lbl = ttk.Label(col_frame, text=col_diff.diff_type.upper())
                type_lbl.grid(row=i, column=0, padx=5, pady=2, sticky=tk.W)

                name_lbl = ttk.Label(col_frame, text=col_diff.column_name)
                name_lbl.grid(row=i, column=1, padx=5, pady=2, sticky=tk.W)

            row_idx = 1

        # 显示行差异
        if sheet_diff.row_diffs:
            row_frame = ttk.LabelFrame(parent, text="行差异", padding=5)
            row_frame.pack(fill=tk.X, pady=5)

            for row_diff in sheet_diff.row_diffs:
                diff_frame = ttk.Frame(row_frame, relief=tk.SOLID, borderwidth=1)
                diff_frame.pack(fill=tk.X, pady=2, padx=2)

                # 基于差异类型的颜色
                bg_color = self._get_diff_color(row_diff.diff_type)

                # 行信息
                info_text = f"行 {row_diff.row_index} - {row_diff.diff_type.upper()}"
                info_lbl = tk.Label(diff_frame, text=info_text, bg=bg_color, font=('Microsoft YaHei UI', 9, 'bold'))
                info_lbl.pack(fill=tk.X, padx=5, pady=2)

                # 显示修改行的单元格详情
                if row_diff.diff_type == 'modified' and row_diff.cell_diffs:
                    details_frame = ttk.Frame(diff_frame)
                    details_frame.pack(fill=tk.X, padx=10, pady=5)

                    for cell_diff in row_diff.cell_diffs:
                        cell_text = f"  {cell_diff.column}: {cell_diff.old_value} → {cell_diff.new_value}"
                        cell_lbl = tk.Label(details_frame, text=cell_text, bg='#FFFFE0', anchor=tk.W)
                        cell_lbl.pack(fill=tk.X, pady=1)

                # 显示新增/删除行的数据
                if row_diff.diff_type in ['added', 'deleted']:
                    data = row_diff.new_data if row_diff.diff_type == 'added' else row_diff.old_data
                    if data:
                        data_text = "  数据: " + str(data)
                        data_lbl = tk.Label(diff_frame, text=data_text, bg=bg_color, wraplength=600, justify=tk.LEFT)
                        data_lbl.pack(fill=tk.X, padx=5, pady=2)

    def _get_diff_color(self, diff_type: str) -> str:
        """获取差异类型的背景颜色。"""
        colors = {
            'added': '#90EE90',      # 浅绿色
            'deleted': '#FFB6C1',    # 浅红色
            'modified': '#FFD700'    # 金色
        }
        return colors.get(diff_type, '#FFFFFF')

    def _show_no_differences(self) -> None:
        """当未发现差异时显示消息。"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="结果")

        label = ttk.Label(
            frame,
            text="两个文件之间未发现差异！",
            font=('Microsoft YaHei UI', 12),
            foreground='green'
        )
        label.pack(pady=20)
