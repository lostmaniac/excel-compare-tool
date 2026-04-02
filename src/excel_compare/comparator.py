"""Excel comparator core logic."""

from dataclasses import dataclass
from typing import List, Dict, Optional, Any
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


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
        """Load both Excel files with original row and column positions."""
        # Load file1
        wb1 = load_workbook(self.file1_path, read_only=True, data_only=True)
        self.sheet_names_1 = wb1.sheetnames
        for sheet in self.sheet_names_1:
            # Read all data without header
            df = pd.read_excel(self.file1_path, sheet_name=sheet, header=None)

            # Generate column names A, B, C, ..., Z, AA, AB, etc.
            column_names = [get_column_letter(i+1) for i in range(df.shape[1])]
            df.columns = column_names

            # Set row numbers as index (Excel row numbers, starting from 1)
            df.index = range(1, len(df) + 1)

            self.df1_dict[sheet] = df

        # Load file2
        wb2 = load_workbook(self.file2_path, read_only=True, data_only=True)
        self.sheet_names_2 = wb2.sheetnames
        for sheet in self.sheet_names_2:
            # Read all data without header
            df = pd.read_excel(self.file2_path, sheet_name=sheet, header=None)

            # Generate column names A, B, C, ..., Z, AA, AB, etc.
            column_names = [get_column_letter(i+1) for i in range(df.shape[1])]
            df.columns = column_names

            # Set row numbers as index (Excel row numbers, starting from 1)
            df.index = range(1, len(df) + 1)

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
        common_cols = list(cols1 & cols2)
        if common_cols:
            # Reset index for comparison
            df1_common = df1[common_cols].reset_index(drop=True)
            df2_common = df2[common_cols].reset_index(drop=True)

            len1, len2 = len(df1_common), len(df2_common)

            # 滑动窗口算法对比行
            i = 0  # 文件1索引
            j = 0  # 文件2索引
            max_lookahead = 10  # 前瞻行数

            def rows_equal(row1, row2, cols):
                """判断两行是否所有公共列都相等"""
                for col in cols:
                    val1 = row1[col]
                    val2 = row2[col]
                    # 处理NaN值：两个NaN视为相等
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    # 一个是NaN另一个不是 → 不相等
                    elif pd.isna(val1) or pd.isna(val2):
                        return False
                    # 值不相等
                    elif val1 != val2:
                        return False
                return True

            def compare_cells(row1, row2, cols):
                """对比两行的单元格差异"""
                cell_diffs = []
                for col in cols:
                    val1 = row1[col]
                    val2 = row2[col]
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    elif pd.isna(val1) and not pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=row1.name + 1,  # Excel实际行号（index+1）
                            column=str(col),
                            old_value=None,
                            new_value=val2
                        ))
                    elif not pd.isna(val1) and pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=row1.name + 1,
                            column=str(col),
                            old_value=val1,
                            new_value=None
                        ))
                    elif val1 != val2:
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=row1.name + 1,
                            column=str(col),
                            old_value=val1,
                            new_value=val2
                        ))
                return cell_diffs

            while i < len1 and j < len2:
                row1 = df1_common.iloc[i]
                row2 = df2_common.iloc[j]

                # 判断两行是否相等
                if rows_equal(row1, row2, common_cols):
                    # 完全匹配，继续下一行
                    i += 1
                    j += 1
                else:
                    # 不匹配，向前查找对齐点
                    found_align = False

                    # 检查文件2中是否新增了行（file2[j+offset] 与 file1[i] 匹配）
                    for offset in range(1, max_lookahead + 1):
                        if j + offset < len2:
                            if rows_equal(row1, df2_common.iloc[j + offset], common_cols):
                                # 发现新增行
                                for k in range(offset):
                                    row_diffs.append(RowDiff(
                                        sheet_name=sheet_name,
                                        row_index=j + k + 1,  # Excel实际行号
                                        diff_type='added',
                                        new_data=df2_common.iloc[j + k].to_dict()
                                    ))
                                j += offset
                                found_align = True
                                break

                    # 检查文件1中是否删除了行（file1[i+offset] 与 file2[j] 匹配）
                    if not found_align:
                        for offset in range(1, max_lookahead + 1):
                            if i + offset < len1:
                                if rows_equal(df1_common.iloc[i + offset], row2, common_cols):
                                    # 发现删除行
                                    for k in range(offset):
                                        row_diffs.append(RowDiff(
                                            sheet_name=sheet_name,
                                            row_index=i + k + 1,  # Excel实际行号
                                            diff_type='deleted',
                                            old_data=df1_common.iloc[i + k].to_dict()
                                        ))
                                    i += offset
                                    found_align = True
                                    break

                    # 找不到对齐点，标记为修改
                    if not found_align:
                        cell_diffs = compare_cells(row1, row2, common_cols)
                        if cell_diffs:
                            row_diffs.append(RowDiff(
                                sheet_name=sheet_name,
                                row_index=i + 1,  # Excel实际行号
                                diff_type='modified',
                                old_data=row1.to_dict(),
                                new_data=row2.to_dict(),
                                cell_diffs=cell_diffs
                            ))
                        i += 1
                        j += 1

            # 处理剩余行
            while i < len1:
                row_diffs.append(RowDiff(
                    sheet_name=sheet_name,
                    row_index=i + 1,
                    diff_type='deleted',
                    old_data=df1_common.iloc[i].to_dict()
                ))
                i += 1

            while j < len2:
                row_diffs.append(RowDiff(
                    sheet_name=sheet_name,
                    row_index=j + 1,
                    diff_type='added',
                    new_data=df2_common.iloc[j].to_dict()
                ))
                j += 1

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
