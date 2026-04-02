"""差异显示组件 - 优化版。"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Optional
from .comparator import SheetDiff, RowDiff, CellDiff, ColumnDiff


class DiffViewer(ttk.Frame):
    """用于显示Excel文件差异的框架（优化版）。"""

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
            text=f"Sheet: {sheet_diff.sheet_name} ({self._get_diff_type_text(sheet_diff.diff_type)})",
            font=('Microsoft YaHei UI', 10, 'bold')
        )
        type_label.pack(anchor=tk.W, pady=(0, 10))

        # 使用Treeview显示差异（性能优化版）
        self._create_diff_treeview(container, sheet_diff)

    def _get_diff_type_text(self, diff_type: str) -> str:
        """获取差异类型的中文文本。"""
        type_map = {
            'added': '新增',
            'deleted': '删除',
            'modified': '修改'
        }
        return type_map.get(diff_type, diff_type.upper())

    def _create_diff_treeview(self, parent: ttk.Frame, sheet_diff: SheetDiff) -> None:
        """
        使用Treeview创建差异显示（性能优化版）。

        Args:
            parent: 父框架
            sheet_diff: SheetDiff对象
        """
        # 创建Notebook用于分类显示
        category_notebook = ttk.Notebook(parent)
        category_notebook.pack(fill=tk.BOTH, expand=True)

        # 列差异标签页
        if sheet_diff.column_diffs:
            col_tab = ttk.Frame(category_notebook)
            category_notebook.add(col_tab, text=f"列差异 ({len(sheet_diff.column_diffs)})")
            self._show_column_diffs(col_tab, sheet_diff.column_diffs)

        # 行差异标签页
        if sheet_diff.row_diffs:
            row_tab = ttk.Frame(category_notebook)
            category_notebook.add(row_tab, text=f"行差异 ({len(sheet_diff.row_diffs)})")
            self._show_row_diffs(row_tab, sheet_diff.row_diffs)

        # 如果没有差异
        if not sheet_diff.column_diffs and not sheet_diff.row_diffs:
            no_diff_tab = ttk.Frame(category_notebook)
            category_notebook.add(no_diff_tab, text="差异详情")
            ttk.Label(no_diff_tab, text="此Sheet没有差异", foreground='green').pack(pady=20)

    def _show_column_diffs(self, parent: ttk.Frame, column_diffs: List[ColumnDiff]) -> None:
        """显示列差异（使用Treeview）。"""
        # 创建Treeview
        tree = ttk.Treeview(parent, columns=('type', 'column'), show='headings')
        tree.heading('type', text='类型')
        tree.heading('column', text='列名')
        tree.column('type', width=100, anchor='center')
        tree.column('column', width=300, anchor='w')

        # 添加滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)

        # 填充所有数据
        for col_diff in column_diffs:
            type_text = self._get_diff_type_text(col_diff.diff_type)
            item_id = tree.insert('', 'end', values=(type_text, col_diff.column_name))

            # 设置行颜色
            if col_diff.diff_type == 'added':
                tree.tag_configure('added', background='#90EE90')
                tree.item(item_id, tags=('added',))
            elif col_diff.diff_type == 'deleted':
                tree.tag_configure('deleted', background='#FFB6C1')
                tree.item(item_id, tags=('deleted',))

    def _show_row_diffs(self, parent: ttk.Frame, row_diffs: List[RowDiff]) -> None:
        """显示行差异（使用Treeview，显示坐标）。"""
        # 创建Treeview
        tree = ttk.Treeview(parent, columns=('type', 'row', 'column', 'details'), show='headings')
        tree.heading('type', text='类型')
        tree.heading('row', text='行号')
        tree.heading('column', text='列名')
        tree.heading('details', text='详情')
        tree.column('type', width=100, anchor='center')
        tree.column('row', width=80, anchor='center')
        tree.column('column', width=150, anchor='w')
        tree.column('details', width=350, anchor='w')

        # 添加滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)

        # 配置标签颜色
        tree.tag_configure('added', background='#90EE90')
        tree.tag_configure('deleted', background='#FFB6C1')
        tree.tag_configure('modified', background='#FFD700')

        # 填充所有数据
        for row_diff in row_diffs:
            type_text = self._get_diff_type_text(row_diff.diff_type)
            row_num = str(row_diff.row_index)

            # 根据差异类型构建详情
            if row_diff.diff_type == 'modified' and row_diff.cell_diffs:
                # 为每个单元格差异创建单独的行
                for cell_diff in row_diff.cell_diffs:
                    details = f"{cell_diff.old_value} → {cell_diff.new_value}"
                    item_id = tree.insert('', 'end', values=(type_text, row_num, cell_diff.column, details))
                    tree.item(item_id, tags=('modified',))
            elif row_diff.diff_type in ['added', 'deleted']:
                # 新增或删除的行
                details = "整行"
                item_id = tree.insert('', 'end', values=(type_text, row_num, "-", details))
                tree.item(item_id, tags=(row_diff.diff_type,))
            else:
                # 其他类型
                details = ""
                item_id = tree.insert('', 'end', values=(type_text, row_num, "-", details))
