"""Difference view components for GUI."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Optional
from .comparator import SheetDiff, RowDiff, CellDiff, ColumnDiff


class DiffViewer(ttk.Frame):
    """Frame for displaying Excel file differences."""

    def __init__(self, parent):
        super().__init__(parent)
        self.diff_results = None
        self.current_sheet_index = 0
        self._setup_ui()

    def _setup_ui(self):
        """Setup the UI components."""
        # Create notebook for sheets
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create empty initial tab
        self.empty_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.empty_tab, text="Result")

    def show_differences(self, diff_results) -> None:
        """
        Display the comparison results.

        Args:
            diff_results: CompareResult object containing differences
        """
        self.diff_results = diff_results
        self.current_sheet_index = 0

        # Clear existing tabs
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        # Show summary first
        self._show_summary(diff_results)

        # Show each sheet's differences
        if diff_results.sheet_diffs:
            for sheet_diff in diff_results.sheet_diffs:
                self._add_sheet_tab(sheet_diff)
        else:
            self._show_no_differences()

    def _show_summary(self, diff_results) -> None:
        """Show the summary of all differences."""
        summary = diff_results.summary

        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Summary")

        # Create text widget for summary
        text = tk.Text(frame, wrap=tk.WORD, padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True)

        # Add summary content
        summary_text = f"""Excel Comparison Summary
{'=' * 50}

Files Compared:
  File 1: {diff_results.file1}
  File 2: {diff_results.file2}

Overall Statistics:
  Total Sheets with Differences: {summary['total_sheets']}
    - Added Sheets: {summary['added_sheets']}
    - Deleted Sheets: {summary['deleted_sheets']}
    - Modified Sheets: {summary['modified_sheets']}

  Column Differences:
    - Added Columns: {summary['added_columns']}
    - Deleted Columns: {summary['deleted_columns']}

  Row Differences:
    - Added Rows: {summary['added_rows']}
    - Deleted Rows: {summary['deleted_rows']}
    - Modified Cells: {summary['modified_cells']}

  Total Differences: {summary['total_differences']}
"""
        text.insert(tk.END, summary_text)
        text.config(state=tk.DISABLED)

    def _add_sheet_tab(self, sheet_diff: SheetDiff) -> None:
        """
        Add a tab for a sheet's differences.

        Args:
            sheet_diff: SheetDiff object to display
        """
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=sheet_diff.sheet_name)

        # Main container
        container = ttk.Frame(frame)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Sheet type indicator
        type_label = ttk.Label(
            container,
            text=f"Sheet: {sheet_diff.sheet_name} ({sheet_diff.diff_type.upper()})",
            font=('Arial', 10, 'bold')
        )
        type_label.pack(anchor=tk.W, pady=(0, 10))

        # Create scrollable frame for differences
        scroll_frame = ttk.Frame(container)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        vsb = ttk.Scrollbar(scroll_frame, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        hsb = ttk.Scrollbar(scroll_frame, orient="horizontal")
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Canvas for content
        canvas = tk.Canvas(
            scroll_frame,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            bg='white'
        )
        canvas.pack(fill=tk.BOTH, expand=True)

        vsb.config(command=canvas.yview)
        hsb.config(command=canvas.xview)

        # Inner frame for content
        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Configure canvas to resize inner frame
        inner_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        ))

        # Populate with differences
        self._populate_sheet_diffs(inner_frame, sheet_diff)

    def _populate_sheet_diffs(self, parent: ttk.Frame, sheet_diff: SheetDiff) -> None:
        """
        Populate a frame with sheet differences.

        Args:
            parent: Parent frame to add widgets to
            sheet_diff: SheetDiff object to display
        """
        row_idx = 0

        # Show column differences
        if sheet_diff.column_diffs:
            col_frame = ttk.LabelFrame(parent, text="Column Differences", padding=5)
            col_frame.pack(fill=tk.X, pady=5)

            headers = ["Type", "Column"]
            for col, header in enumerate(headers):
                lbl = ttk.Label(col_frame, text=header, font=('Arial', 9, 'bold'))
                lbl.grid(row=0, column=col, padx=5, pady=2, sticky=tk.W)

            for i, col_diff in enumerate(sheet_diff.column_diffs, start=1):
                color = '#90EE90' if col_diff.diff_type == 'added' else '#FFB6C1'

                type_lbl = ttk.Label(col_frame, text=col_diff.diff_type.upper())
                type_lbl.grid(row=i, column=0, padx=5, pady=2, sticky=tk.W)

                name_lbl = ttk.Label(col_frame, text=col_diff.column_name)
                name_lbl.grid(row=i, column=1, padx=5, pady=2, sticky=tk.W)

            row_idx = 1

        # Show row differences
        if sheet_diff.row_diffs:
            row_frame = ttk.LabelFrame(parent, text="Row Differences", padding=5)
            row_frame.pack(fill=tk.X, pady=5)

            for row_diff in sheet_diff.row_diffs:
                diff_frame = ttk.Frame(row_frame, relief=tk.SOLID, borderwidth=1)
                diff_frame.pack(fill=tk.X, pady=2, padx=2)

                # Color based on diff type
                bg_color = self._get_diff_color(row_diff.diff_type)

                # Row info
                info_text = f"Row {row_diff.row_index} - {row_diff.diff_type.upper()}"
                info_lbl = tk.Label(diff_frame, text=info_text, bg=bg_color, font=('Arial', 9, 'bold'))
                info_lbl.pack(fill=tk.X, padx=5, pady=2)

                # Show cell details for modified rows
                if row_diff.diff_type == 'modified' and row_diff.cell_diffs:
                    details_frame = ttk.Frame(diff_frame)
                    details_frame.pack(fill=tk.X, padx=10, pady=5)

                    for cell_diff in row_diff.cell_diffs:
                        cell_text = f"  {cell_diff.column}: {cell_diff.old_value} → {cell_diff.new_value}"
                        cell_lbl = tk.Label(details_frame, text=cell_text, bg='#FFFFE0', anchor=tk.W)
                        cell_lbl.pack(fill=tk.X, pady=1)

                # Show data for added/deleted rows
                if row_diff.diff_type in ['added', 'deleted']:
                    data = row_diff.new_data if row_diff.diff_type == 'added' else row_diff.old_data
                    if data:
                        data_text = "  Data: " + str(data)
                        data_lbl = tk.Label(diff_frame, text=data_text, bg=bg_color, wraplength=600, justify=tk.LEFT)
                        data_lbl.pack(fill=tk.X, padx=5, pady=2)

    def _get_diff_color(self, diff_type: str) -> str:
        """Get background color for difference type."""
        colors = {
            'added': '#90EE90',      # Light green
            'deleted': '#FFB6C1',    # Light red
            'modified': '#FFD700'    # Gold
        }
        return colors.get(diff_type, '#FFFFFF')

    def _show_no_differences(self) -> None:
        """Show message when no differences found."""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Result")

        label = ttk.Label(
            frame,
            text="No differences found between the two files!",
            font=('Arial', 12),
            foreground='green'
        )
        label.pack(pady=20)
