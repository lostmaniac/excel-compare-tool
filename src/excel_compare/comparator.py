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
        if common_cols:
            # 使用原始索引（Excel行号）进行比较
            df1_common = df1[common_cols]
            df2_common = df2[common_cols]

            # 保存原始行号（Excel行号）
            df1_row_nums = df1_common.index.tolist()
            df2_row_nums = df2_common.index.tolist()

            # 重置索引以便于difflib比较
            df1_common_reset = df1_common.reset_index(drop=True)
            df2_common_reset = df2_common.reset_index(drop=True)

            # 过滤掉空行
            def filter_empty_rows(df_reset, row_nums):
                """过滤掉空行，返回(过滤后的df, 原始行号列表)"""
                non_empty_indices = []
                non_empty_row_nums = []
                for idx, row_num in zip(df_reset.index, row_nums):
                    row_data = df_reset.iloc[idx].values
                    if not self._is_empty_row(row_data):
                        non_empty_indices.append(idx)
                        non_empty_row_nums.append(row_num)
                return df_reset.iloc[non_empty_indices].reset_index(drop=True), non_empty_row_nums

            df1_filtered, df1_filtered_row_nums = filter_empty_rows(df1_common_reset, df1_row_nums)
            df2_filtered, df2_filtered_row_nums = filter_empty_rows(df2_common_reset, df2_row_nums)

            # 将DataFrame的每行转换为可哈希的元组，用于difflib比较
            def row_to_tuple(row):
                """将Series转换为可哈希的元组，处理NaN值"""
                return tuple(
                    (val if not pd.isna(val) else None) for val in row
                )

            # 生成行元组列表
            rows1 = [row_to_tuple(row) for _, row in df1_filtered.iterrows()]
            rows2 = [row_to_tuple(row) for _, row in df2_filtered.iterrows()]

            # 使用difflib.SequenceMatcher进行差异分析
            matcher = difflib.SequenceMatcher(None, rows1, rows2)

            def compare_cells(row1, row2, excel_row_num):
                """对比两行的单元格差异"""
                cell_diffs = []
                for col in common_cols:
                    val1 = row1[col]
                    val2 = row2[col]
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    elif pd.isna(val1) and not pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,
                            column=str(col),
                            old_value=None,
                            new_value=val2
                        ))
                    elif not pd.isna(val1) and pd.isna(val2):
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,
                            column=str(col),
                            old_value=val1,
                            new_value=None
                        ))
                    elif val1 != val2:
                        cell_diffs.append(CellDiff(
                            sheet_name=sheet_name,
                            row=excel_row_num,
                            column=str(col),
                            old_value=val1,
                            new_value=val2
                        ))
                return cell_diffs

            # 遍历所有操作序列
            for op, i1, i2, j1, j2 in matcher.get_opcodes():
                # op: 操作类型
                # i1, i2: file1的行索引范围 [i1, i2)
                # j1, j2: file2的行索引范围 [j1, j2)

                if op == 'equal':
                    # 完全相同的行，无需记录
                    continue

                elif op == 'delete':
                    # file1中删除了 [i1, i2) 的行
                    for idx in range(i1, i2):
                        row_diffs.append(RowDiff(
                            sheet_name=sheet_name,
                            row_index=df1_filtered_row_nums[idx],  # 使用原始Excel行号
                            diff_type='deleted',
                            old_data=df1_filtered.iloc[idx].to_dict()
                        ))

                elif op == 'insert':
                    # file2中新增了 [j1, j2) 的行
                    for idx in range(j1, j2):
                        row_diffs.append(RowDiff(
                            sheet_name=sheet_name,
                            row_index=df2_filtered_row_nums[idx],  # 使用原始Excel行号
                            diff_type='added',
                            new_data=df2_filtered.iloc[idx].to_dict()
                        ))

                elif op == 'replace':
                    # [i1, i2) 的行被替换为 [j1, j2) 的行
                    # 取较小的长度进行对比
                    min_len = min(i2 - i1, j2 - j1)
                    for k in range(min_len):
                        row1 = df1_filtered.iloc[i1 + k]
                        row2 = df2_filtered.iloc[j1 + k]
                        excel_row_num = df1_filtered_row_nums[i1 + k]  # 使用原始Excel行号

                        cell_diffs = compare_cells(row1, row2, excel_row_num)
                        if cell_diffs:
                            row_diffs.append(RowDiff(
                                sheet_name=sheet_name,
                                row_index=excel_row_num,
                                diff_type='modified',
                                old_data=row1.to_dict(),
                                new_data=row2.to_dict(),
                                cell_diffs=cell_diffs
                            ))

                    # 处理file1多出的行（删除）
                    for idx in range(i1 + min_len, i2):
                        row_diffs.append(RowDiff(
                            sheet_name=sheet_name,
                            row_index=df1_filtered_row_nums[idx],  # 使用原始Excel行号
                            diff_type='deleted',
                            old_data=df1_filtered.iloc[idx].to_dict()
                        ))

                    # 处理file2多出的行（新增）
                    for idx in range(j1 + min_len, j2):
                        row_diffs.append(RowDiff(
                            sheet_name=sheet_name,
                            row_index=df2_filtered_row_nums[idx],  # 使用原始Excel行号
                            diff_type='added',
                            new_data=df2_filtered.iloc[idx].to_dict()
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
