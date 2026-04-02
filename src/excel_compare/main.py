"""Excel文件对比工具主程序。"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys

# 处理包内导入和独立运行
try:
    from .comparator import ExcelComparator
    from .diff_view import DiffViewer
except ImportError:
    # PyInstaller 打包后的备用导入
    from excel_compare.comparator import ExcelComparator
    from excel_compare.diff_view import DiffViewer


class ExcelCompareApp:
    """Excel文件对比工具主应用类。"""

    def __init__(self, root: tk.Tk):
        """
        初始化应用程序。

        Args:
            root: Tk 根窗口
        """
        self.root = root
        self.root.title("Excel文件对比工具")
        self.root.geometry("1000x700")

        self.file1_path = None
        self.file2_path = None

        self._setup_ui()

    def _setup_ui(self):
        """设置用户界面。"""
        # 创建主容器
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(
            main_frame,
            text="Excel文件对比工具",
            font=('Microsoft YaHei UI', 16, 'bold')
        )
        title_label.pack(pady=(0, 10))

        # 文件选择区域
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        # 源文件区域
        self._create_file_section(file_frame, "源文件", 1, self._select_file1)

        # 新文件区域
        self._create_file_section(file_frame, "新文件", 2, self._select_file2)

        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        self.compare_btn = ttk.Button(
            control_frame,
            text="开始对比",
            command=self._start_comparison,
            state=tk.DISABLED
        )
        self.compare_btn.pack(side=tk.LEFT, padx=5)

        clear_btn = ttk.Button(
            control_frame,
            text="清除全部",
            command=self._clear_all
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # 状态标签
        self.status_label = ttk.Label(
            control_frame,
            text="就绪 - 请选择要对比的Excel文件",
            foreground='gray'
        )
        self.status_label.pack(side=tk.LEFT, padx=20)

        # 结果查看器
        self.diff_viewer = DiffViewer(main_frame)
        self.diff_viewer.pack(fill=tk.BOTH, expand=True, pady=10)

    def _create_file_section(
        self,
        parent: ttk.Frame,
        title: str,
        file_num: int,
        select_handler
    ):
        """
        创建文件选择区域。

        Args:
            parent: 父框架
            title: 区域标题
            file_num: 文件编号（1或2）
            select_handler: 选择文件的处理函数
        """
        # 创建此文件的框架
        section = ttk.Frame(parent)
        section.pack(fill=tk.X, pady=5)

        # 标签
        label = ttk.Label(section, text=title, font=('Microsoft YaHei UI', 10, 'bold'), width=8)
        label.pack(side=tk.LEFT, padx=5)

        # 创建带按钮的区域
        container = tk.Frame(section)
        container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 选择按钮
        btn = ttk.Button(
            container,
            text="点击选择文件",
            command=select_handler
        )
        btn.pack(fill=tk.X)

        # 文件路径标签
        path_label = ttk.Label(container, text="未选择文件", foreground='gray')
        path_label.pack(fill=tk.X, pady=5)
        setattr(self, f'file{file_num}_label', path_label)

    def _update_file_label(self, label: ttk.Label, file_path: str):
        """更新文件路径标签。"""
        # 只获取文件名
        filename = os.path.basename(file_path)
        # 如果太长则截断
        if len(filename) > 50:
            filename = '...' + filename[-47:]
        label.config(text=filename, foreground='green')

    def _select_file1(self):
        """处理源文件的选择对话框。"""
        file_path = filedialog.askopenfilename(
            title="选择源Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file1_path = file_path
            self._update_file_label(self.file1_label, file_path)
            self._check_ready()

    def _select_file2(self):
        """处理新文件的选择对话框。"""
        file_path = filedialog.askopenfilename(
            title="选择新Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file2_path = file_path
            self._update_file_label(self.file2_label, file_path)
            self._check_ready()

    def _check_ready(self):
        """检查是否两个文件都已选择，启用/禁用对比按钮。"""
        if self.file1_path and self.file2_path:
            self.compare_btn.config(state=tk.NORMAL)
            self.status_label.config(text="文件已选择 - 可以开始对比")
        else:
            self.compare_btn.config(state=tk.DISABLED)
            self.status_label.config(text="请选择两个要对比的文件")

    def _start_comparison(self):
        """开始对比过程。"""
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("警告", "请先选择两个文件！")
            return

        try:
            self.status_label.config(text="正在对比文件... 请稍候...")
            self.root.update()

            # 创建对比器并对比
            comparator = ExcelComparator(self.file1_path, self.file2_path)
            result = comparator.compare()

            # 显示结果
            self.diff_viewer.show_differences(result)

            # 更新状态
            if result.summary['total_differences'] > 0:
                self.status_label.config(
                    text=f"对比完成 - 发现 {result.summary['total_differences']} 处差异"
                )
            else:
                self.status_label.config(
                    text="对比完成 - 未发现差异！",
                    foreground='green'
                )

        except Exception as e:
            messagebox.showerror("错误", f"对比文件时出错:\n{str(e)}")
            self.status_label.config(text="对比过程中发生错误", foreground='red')

    def _clear_all(self):
        """清除所有选择和结果。"""
        self.file1_path = None
        self.file2_path = None

        self.file1_label.config(text="未选择文件", foreground='gray')
        self.file2_label.config(text="未选择文件", foreground='gray')

        self.compare_btn.config(state=tk.DISABLED)
        self.status_label.config(text="就绪 - 请选择要对比的Excel文件", foreground='gray')

        # 清除结果
        self.diff_viewer = DiffViewer(self.root.winfo_children()[0])
        # 需要重新打包
        main_frame = self.root.winfo_children()[0]
        for widget in main_frame.winfo_children():
            if isinstance(widget, DiffViewer):
                widget.destroy()
        self.diff_viewer.pack(fill=tk.BOTH, expand=True, pady=10)


def main():
    """主入口点。"""
    # 配置tkinter以支持高DPI显示
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    root = tk.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
