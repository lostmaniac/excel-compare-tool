"""Main GUI application for Excel file comparison."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys

# Handle both package and standalone execution
try:
    from .comparator import ExcelComparator
    from .diff_view import DiffViewer
except ImportError:
    # Fallback for PyInstaller standalone execution
    from excel_compare.comparator import ExcelComparator
    from excel_compare.diff_view import DiffViewer


class ExcelCompareApp:
    """Main application class for Excel file comparison."""

    def __init__(self, root: tk.Tk):
        """
        Initialize the application.

        Args:
            root: The root Tk window
        """
        self.root = root
        self.root.title("Excel File Compare Tool")
        self.root.geometry("1000x700")

        self.file1_path = None
        self.file2_path = None

        self._setup_ui()

    def _setup_ui(self):
        """Setup the user interface."""
        # Create main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Excel File Compare Tool",
            font=('Arial', 16, 'bold')
        )
        title_label.pack(pady=(0, 10))

        # File selection frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        # File 1 section
        self._create_file_section(file_frame, "File 1", 0, self._on_file1_drop, self._select_file1)

        # File 2 section
        self._create_file_section(file_frame, "File 2", 1, self._on_file2_drop, self._select_file2)

        # Control buttons
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        self.compare_btn = ttk.Button(
            control_frame,
            text="Start Comparison",
            command=self._start_comparison,
            state=tk.DISABLED
        )
        self.compare_btn.pack(side=tk.LEFT, padx=5)

        clear_btn = ttk.Button(
            control_frame,
            text="Clear All",
            command=self._clear_all
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # Status label
        self.status_label = ttk.Label(
            control_frame,
            text="Ready - Select Excel files to compare",
            foreground='gray'
        )
        self.status_label.pack(side=tk.LEFT, padx=20)

        # Result viewer
        self.diff_viewer = DiffViewer(main_frame)
        self.diff_viewer.pack(fill=tk.BOTH, expand=True, pady=10)

    def _create_file_section(
        self,
        parent: ttk.Frame,
        title: str,
        row: int,
        drop_handler,
        select_handler
    ):
        """
        Create a file selection section.

        Args:
            parent: Parent frame
            title: Section title
            row: Grid row number
            drop_handler: Handler for drop events
            select_handler: Handler for select button
        """
        # Frame for this file
        section = ttk.Frame(parent)
        section.pack(fill=tk.X, pady=5)

        # Label
        label = ttk.Label(section, text=title, font=('Arial', 10, 'bold'), width=8)
        label.pack(side=tk.LEFT, padx=5)

        # Drop zone frame
        drop_frame = tk.Frame(
            section,
            height=40,
            borderwidth=2,
            relief=tk.GROOVE,
            bg='#f0f0f0'
        )
        drop_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # File path label
        path_label = ttk.Label(drop_frame, text="Click to select file")
        path_label.pack(expand=True, padx=10)
        setattr(self, f'file{row + 1}_label', path_label)

        # Click to select
        drop_frame.bind('<Button-1>', lambda e: select_handler())

    def _on_file1_drop(self, file_path: str):
        """Handle file drop for file 1."""
        if file_path and os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
            self.file1_path = file_path
            self._update_file_label(self.file1_label, file_path)
            self._check_ready()

    def _on_file2_drop(self, file_path: str):
        """Handle file drop for file 2."""
        if file_path and os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
            self.file2_path = file_path
            self._update_file_label(self.file2_label, file_path)
            self._check_ready()

    def _update_file_label(self, label: ttk.Label, file_path: str):
        """Update the file path label."""
        # Get just the filename
        filename = os.path.basename(file_path)
        # Truncate if too long
        if len(filename) > 50:
            filename = '...' + filename[-47:]
        label.config(text=filename, foreground='green')

    def _select_file1(self):
        """Handle file selection dialog for file 1."""
        file_path = filedialog.askopenfilename(
            title="Select First Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file1_path = file_path
            self._update_file_label(self.file1_label, file_path)
            self._check_ready()

    def _select_file2(self):
        """Handle file selection dialog for file 2."""
        file_path = filedialog.askopenfilename(
            title="Select Second Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file2_path = file_path
            self._update_file_label(self.file2_label, file_path)
            self._check_ready()

    def _check_ready(self):
        """Check if both files are selected and enable/disable compare button."""
        if self.file1_path and self.file2_path:
            self.compare_btn.config(state=tk.NORMAL)
            self.status_label.config(text="Files selected - Ready to compare")
        else:
            self.compare_btn.config(state=tk.DISABLED)
            self.status_label.config(text="Select both files to compare")

    def _start_comparison(self):
        """Start the comparison process."""
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("Warning", "Please select both files first!")
            return

        try:
            self.status_label.config(text="Comparing files... Please wait...")
            self.root.update()

            # Create comparator and compare
            comparator = ExcelComparator(self.file1_path, self.file2_path)
            result = comparator.compare()

            # Display results
            self.diff_viewer.show_differences(result)

            # Update status
            if result.summary['total_differences'] > 0:
                self.status_label.config(
                    text=f"Comparison complete - Found {result.summary['total_differences']} differences"
                )
            else:
                self.status_label.config(
                    text="Comparison complete - No differences found!",
                    foreground='green'
                )

        except Exception as e:
            messagebox.showerror("Error", f"Error comparing files:\n{str(e)}")
            self.status_label.config(text="Error occurred during comparison", foreground='red')

    def _clear_all(self):
        """Clear all selections and results."""
        self.file1_path = None
        self.file2_path = None

        self.file1_label.config(text="Click to select file", foreground='black')
        self.file2_label.config(text="Click to select file", foreground='black')

        self.compare_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Ready - Select Excel files to compare", foreground='gray')

        # Clear results
        self.diff_viewer = DiffViewer(self.root.winfo_children()[0])
        # Need to repack
        main_frame = self.root.winfo_children()[0]
        for widget in main_frame.winfo_children():
            if isinstance(widget, DiffViewer):
                widget.destroy()
        self.diff_viewer.pack(fill=tk.BOTH, expand=True, pady=10)


def main():
    """Main entry point."""
    # Configure tkinter for HiDPI displays
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
