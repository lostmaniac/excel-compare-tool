"""Excel comparator core logic."""

from dataclasses import dataclass
from typing import List, Dict, Optional, Any
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import difflib


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

    def _is_empty_row(self, row_data):
        """判断是否为空行（所有值都是NaN、None或空字符串）"""
        for val in row_data:
            if pd.notna(val) and val != '' and str(val).strip() != '':
                return False
        return True

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
        if common_cols and 'B' in common_cols:  # 需要B列（场所名称）作为唯一标识符进行匹配
            # 使用原始索引（Excel行号）进行比较
            df1_common = df1[common_cols]
            df2_common = df2[common_cols]

            # 基于场所名称（B列）进行匹配，而不是按行顺序比较
            # 收集file1中所有有效行（有场所名称的行）
            rows_by_name1 = {}  # {场所名称: (行号, 行数据)}
            for excel_row_num in df1_common.index:
                row_data = df1_common.loc[excel_row_num]
                site_name = row_data['B']
                if pd.notna(site_name) and site_name != '' and str(site_name).strip() != '':
                    rows_by_name1[str(site_name).strip()] = (excel_row_num, row_data)

            # 收集file2中所有有效行（有场所名称的行）
            rows_by_name2 = {}  # {场所名称: (行号, 行数据)}
            for excel_row_num in df2_common.index:
                row_data = df2_common.loc[excel_row_num]
                site_name = row_data['B']
                if pd.notna(site_name) and site_name != '' and str(site_name).strip() != '':
                    rows_by_name2[str(site_name).strip()] = (excel_row_num, row_data)

            # 获取所有场所名称
            all_names = set(rows_by_name1.keys()) | set(rows_by_name2.keys())

            # 对比单元格的函数
            def compare_cells(row1, row2, excel_row_num2):
                """对比两行的单元格差异"""
                cell_diffs = []
                for col in common_cols:
                    # 跳过序号列（A列）和场所名称列（B列）
                    if col == 'A' or col == 'B':
                        continue
                    val1 = row1[col]
                    val2 = row2[col]
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    elif pd.isna(val1) and not pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num2,
                            column=str(col),
                            old_value=None,
                            new_value=val2
                        ))
                    elif not pd.isna(val1) and pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num2,
                            column=str(col),
                            old_value=val1,
                            new_value=None
                        ))
                    elif val1 != val2:
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num2,
                            column=str(col),
                            old_value=val1,
                            new_value=val2
                        ))
                return cell_diffs

            # 按场所名称排序遍历
            for site_name in sorted(all_names):
                in_file1 = site_name in rows_by_name1
                in_file2 = site_name in rows_by_name2

                if not in_file1 and in_file2:
                    # 新增的行
                    excel_row_num, row_data = rows_by_name2[site_name]
                    row_diffs.append(RowDiff(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        diff_type='added',
                        new_data=row_data.to_dict()
                    ))

                elif in_file1 and not in_file2:
                    # 删除的行
                    excel_row_num, row_data = rows_by_name1[site_name]
                    row_diffs.append(RowDiff(
                        sheet_name=sheet_name,
                        row_index=excel_row_num,
                        diff_type='deleted',
                        old_data=row_data.to_dict()
                    ))

                elif in_file1 and in_file2:
                    # 两个文件都有此场所，比较内容
                    excel_row_num1, row_data1 = rows_by_name1[site_name]
                    excel_row_num2, row_data2 = rows_by_name2[site_name]

                    cell_diffs = compare_cells(row_data1, row_data2, excel_row_num2)
                    if cell_diffs:
                        row_diffs.append(RowDiff(
                            sheet_name=sheet_name,
                            row_index=excel_row_num2,  # 使用file2的行号
                            diff_type='modified',
                            old_data=row_data1.to_dict(),
                            new_data=row_data2.to_dict(),
                            cell_diffs=cell_diffs
                        ))

            # 按行号排序，使显示更有序
            row_diffs.sort(key=lambda x: x.row_index)

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
