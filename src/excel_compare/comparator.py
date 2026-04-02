"""Excel comparator core logic."""

from dataclasses import dataclass
from typing import List, Dict, Optional, Any
import pandas as pd
from openpyxl import load_workbook


@dataclass
class CellDiff:
    """Represents a difference in a single cell."""
    sheet_name: str
    row: int
    column: str
    old_value: Any
    new_value: Any


@dataclass
class RowDiff:
    """Represents a row difference."""
    sheet_name: str
    row_index: int
    diff_type: str  # 'added', 'deleted', 'modified'
    old_data: Optional[Dict[str, Any]] = None
    new_data: Optional[Dict[str, Any]] = None
    cell_diffs: List[CellDiff] = None


@dataclass
class ColumnDiff:
    """Represents a column structure difference."""
    sheet_name: str
    column_name: str
    diff_type: str  # 'added', 'deleted'


@dataclass
class SheetDiff:
    """Represents differences in a sheet."""
    sheet_name: str
    diff_type: str  # 'added', 'deleted', 'modified'
    column_diffs: List[ColumnDiff] = None
    row_diffs: List[RowDiff] = None


@dataclass
class CompareResult:
    """Result of comparing two Excel files."""
    file1: str
    file2: str
    sheet_diffs: List[SheetDiff]
    summary: Dict[str, int]


class ExcelComparator:
    """Compares two Excel files and identifies differences."""

    def __init__(self, file1_path: str, file2_path: str):
        """
        Initialize the comparator with two file paths.

        Args:
            file1_path: Path to the first Excel file
            file2_path: Path to the second Excel file
        """
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.df1_dict = {}
        self.df2_dict = {}
        self.sheet_names_1 = []
        self.sheet_names_2 = []

    def load_files(self) -> None:
        """Load both Excel files."""
        # Load file1
        wb1 = load_workbook(self.file1_path, read_only=True)
        self.sheet_names_1 = wb1.sheetnames
        for sheet in self.sheet_names_1:
            # Read with header=0 to use first row as column names
            # Don't skip rows to preserve original row numbers
            df = pd.read_excel(self.file1_path, sheet_name=sheet, header=0)
            # Filter out unnamed columns
            df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]
            self.df1_dict[sheet] = df

        # Load file2
        wb2 = load_workbook(self.file2_path, read_only=True)
        self.sheet_names_2 = wb2.sheetnames
        for sheet in self.sheet_names_2:
            df = pd.read_excel(self.file2_path, sheet_name=sheet, header=0)
            # Filter out unnamed columns
            df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]
            self.df2_dict[sheet] = df

    def compare(self) -> CompareResult:
        """
        Compare the two Excel files.

        Returns:
            CompareResult object containing all differences
        """
        self.load_files()

        sheet_diffs = []

        # Get all unique sheet names
        all_sheets = set(self.sheet_names_1) | set(self.sheet_names_2)

        for sheet_name in all_sheets:
            if sheet_name in self.sheet_names_1 and sheet_name not in self.sheet_names_2:
                sheet_diffs.append(SheetDiff(
                    sheet_name=sheet_name,
                    diff_type='deleted'
                ))
            elif sheet_name not in self.sheet_names_1 and sheet_name in self.sheet_names_2:
                sheet_diffs.append(SheetDiff(
                    sheet_name=sheet_name,
                    diff_type='added'
                ))
            else:
                # Both files have this sheet, compare content
                diff = self._compare_sheet(sheet_name)
                if diff:
                    sheet_diffs.append(diff)

        # Generate summary
        summary = self._generate_summary(sheet_diffs)

        return CompareResult(
            file1=self.file1_path,
            file2=self.file2_path,
            sheet_diffs=sheet_diffs,
            summary=summary
        )

    def _compare_sheet(self, sheet_name: str) -> Optional[SheetDiff]:
        """
        Compare a single sheet between both files.

        Args:
            sheet_name: Name of the sheet to compare

        Returns:
            SheetDiff object if differences found, None otherwise
        """
        df1 = self.df1_dict[sheet_name]
        df2 = self.df2_dict[sheet_name]

        column_diffs = []
        row_diffs = []

        # Compare columns
        cols1 = set(df1.columns)
        cols2 = set(df2.columns)

        for col in cols1 - cols2:
            column_diffs.append(ColumnDiff(
                sheet_name=sheet_name,
                column_name=str(col),
                diff_type='deleted'
            ))

        for col in cols2 - cols1:
            column_diffs.append(ColumnDiff(
                sheet_name=sheet_name,
                column_name=str(col),
                diff_type='added'
            ))

        # Only compare rows if there are common columns
        common_cols = cols1 & cols2
        if common_cols:
            # Reset index for comparison
            df1_common = df1[list(common_cols)].reset_index(drop=True)
            df2_common = df2[list(common_cols)].reset_index(drop=True)

            # Compare row counts
            len1, len2 = len(df1_common), len(df2_common)

            # Find added and deleted rows
            added_rows = max(0, len2 - len1)
            deleted_rows = max(0, len1 - len2)

            # Compare common rows
            common_length = min(len1, len2)
            for i in range(common_length):
                cell_diffs = []
                row1 = df1_common.iloc[i]
                row2 = df2_common.iloc[i]

                # Excel行号：数据从第2行开始（第1行是表头）
                excel_row_num = i + 2

                for col in common_cols:
                    val1 = row1[col]
                    val2 = row2[col]

                    # Handle NaN values
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    elif pd.isna(val1) and not pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,  # Excel行号
                            column=str(col),
                            old_value=None,
                            new_value=val2
                        ))
                    elif not pd.isna(val1) and pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,  # Excel行号
                            column=str(col),
                            old_value=val1,
                            new_value=None
                        ))
                    elif val1 != val2:
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,  # Excel行号
                            column=str(col),
                            old_value=val1,
                            new_value=val2
                        ))

                if cell_diffs:
                    row_diffs.append(RowDiff(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,  # Excel行号
                        diff_type='modified',
                        old_data=row1.to_dict(),
                        new_data=row2.to_dict(),
                        cell_diffs=cell_diffs
                    ))

            # Add row count differences
            if added_rows > 0:
                for i in range(added_rows):
                    # Excel行号：从common_length+2开始（+1是因为从0索引，+1是因为跳过表头行）
                    excel_row_num = common_length + i + 2
                    row_diffs.append(RowDiff(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        diff_type='added',
                        new_data=df2_common.iloc[common_length + i].to_dict()
                    ))

            if deleted_rows > 0:
                for i in range(deleted_rows):
                    excel_row_num = common_length + i + 2
                    row_diffs.append(RowDiff(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        diff_type='deleted',
                        old_data=df1_common.iloc[common_length + i].to_dict()
                    ))

        # Only return SheetDiff if there are differences
        if column_diffs or row_diffs:
            return SheetDiff(
                sheet_name=sheet_name,
                diff_type='modified',
                column_diffs=column_diffs if column_diffs else None,
                row_diffs=row_diffs if row_diffs else None
            )

        return None

    def _generate_summary(self, sheet_diffs: List[SheetDiff]) -> Dict[str, int]:
        """
        Generate a summary of all differences.

        Args:
            sheet_diffs: List of sheet differences

        Returns:
            Dictionary containing summary statistics
        """
        summary = {
            'total_sheets': 0,
            'added_sheets': 0,
            'deleted_sheets': 0,
            'modified_sheets': 0,
            'added_columns': 0,
            'deleted_columns': 0,
            'added_rows': 0,
            'deleted_rows': 0,
            'modified_cells': 0,
            'total_differences': 0
        }

        for diff in sheet_diffs:
            summary['total_sheets'] += 1
            summary['total_differences'] += 1

            if diff.diff_type == 'added':
                summary['added_sheets'] += 1
            elif diff.diff_type == 'deleted':
                summary['deleted_sheets'] += 1
            elif diff.diff_type == 'modified':
                summary['modified_sheets'] += 1

                if diff.column_diffs:
                    for col_diff in diff.column_diffs:
                        if col_diff.diff_type == 'added':
                            summary['added_columns'] += 1
                        elif col_diff.diff_type == 'deleted':
                            summary['deleted_columns'] += 1
                        summary['total_differences'] += 1

                if diff.row_diffs:
                    for row_diff in diff.row_diffs:
                        if row_diff.diff_type == 'added':
                            summary['added_rows'] += 1
                        elif row_diff.diff_type == 'deleted':
                            summary['deleted_rows'] += 1
                        elif row_diff.diff_type == 'modified':
                            if row_diff.cell_diffs:
                                summary['modified_cells'] += len(row_diff.cell_diffs)
                        summary['total_differences'] += 1

        return summary
