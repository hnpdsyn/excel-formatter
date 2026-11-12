"""
Excel 格式整理工具 - 核心逻辑
"""

import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import re
from typing import List, Optional, Tuple
import warnings
warnings.filterwarnings('ignore')


class ExcelFormatter:
    """Excel 格式整理器"""

    def __init__(self):
        self.logs = []
        self.input_path = None
        self.wb = None
        self.ws = None

    def _log(self, msg: str):
        self.logs.append(msg)

    def get_logs(self) -> List[str]:
        return self.logs

    def load(self, file_path: str) -> bool:
        """加载Excel文件"""
        self.logs = []
        self.input_path = file_path
        try:
            self.wb = openpyxl.load_workbook(file_path)
            self.ws = self.wb.active
            self._log(f"✅ 已加载：{Path(file_path).name}（{self.wb.sheetnames}）")
            return True
        except Exception as e:
            self._log(f"❌ 加载失败：{str(e)}")
            return False

    def get_sheet_names(self) -> List[str]:
        if self.wb:
            return self.wb.sheetnames
        return []

    def set_active_sheet(self, name: str):
        if name in self.wb.sheetnames:
            self.ws = self.wb[name]

    # ===================== 清理功能 =====================

    def clean_trim_spaces(self):
        """去除所有单元格前后空格"""
        count = 0
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    stripped = cell.value.strip()
                    if stripped != cell.value:
                        cell.value = stripped
                        count += 1
        self._log(f"✅ 已清理 {count} 个单元格的前后空格")
        return count

    def clean_empty_rows(self):
        """删除整行空行"""
        rows_to_delete = []
        for row_idx in range(self.ws.max_row, 0, -1):
            is_empty = True
            for col in range(1, self.ws.max_column + 1):
                if self.ws.cell(row=row_idx, column=col).value is not None:
                    is_empty = False
                    break
            if is_empty:
                rows_to_delete.append(row_idx)

        if rows_to_delete:
            for r in rows_to_delete:
                self.ws.delete_rows(r)
            self._log(f"✅ 已删除 {len(rows_to_delete)} 行空行")
        else:
            self._log("ℹ️ 未发现空行")
        return len(rows_to_delete)

    def clean_empty_columns(self):
        """删除整列空列"""
        cols_to_delete = []
        for col in range(self.ws.max_column, 0, -1):
            is_empty = True
            for row in range(1, self.ws.max_row + 1):
                if self.ws.cell(row=row, column=col).value is not None:
                    is_empty = False
                    break
            if is_empty:
                cols_to_delete.append(col)

        if cols_to_delete:
            for c in cols_to_delete:
                self.ws.delete_cols(c)
            self._log(f"✅ 已删除 {len(cols_to_delete)} 列空列")
        else:
            self._log("ℹ️ 未发现空列")
        return len(cols_to_delete)

    def clean_duplicates(self, has_header: bool = True):
        """删除重复行"""
        rows_data = []
        for row_idx in range(1, self.ws.max_row + 1):
            row_values = tuple(
                self.ws.cell(row=row_idx, column=col).value
                for col in range(1, self.ws.max_column + 1)
            )
            rows_data.append((row_idx, row_values))

        # 保留表头
        seen = set()
        to_delete = []
        start_idx = 1 if has_header else 0

        for row_idx, values in rows_data[start_idx:]:
            if values in seen:
                to_delete.append(row_idx)
            else:
                seen.add(values)

        if to_delete:
            for r in reversed(to_delete):
                self.ws.delete_rows(r)
            self._log(f"✅ 已删除 {len(to_delete)} 行重复数据")
        else:
            self._log("ℹ️ 未发现重复行")
        return len(to_delete)

    def clean_special_chars(self, keep_chinese: bool = False):
        """清除特殊字符（除了中文、字母、数字、常用标点）"""
        count = 0
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if keep_chinese:
                        cleaned = re.sub(r'[^\u4e00-\u9fff\uFF00-\uFFEFa-zA-Z0-9\s\.\,\-\+\%\#\@\(\)]', '', cell.value)
                    else:
                        cleaned = re.sub(r'[^\x20-\x7E\u4e00-\u9fff\uFF00-\uFFEF]', '', cell.value)
                    if cleaned != cell.value:
                        cell.value = cleaned
                        count += 1
        self._log(f"✅ 已清理 {count} 个单元格的特殊字符")
        return count

    def clean_all(self, has_header: bool = True, keep_chinese: bool = False):
        """一键全清理"""
        self._log("🔧 开始一键清理...")
        self.clean_trim_spaces()
        self.clean_empty_rows()
        self.clean_empty_columns()
        self.clean_duplicates(has_header)
        self.clean_special_chars(keep_chinese)
        self._log("🎉 一键清理完成！")

    # ===================== 格式功能 =====================

    def format_font(self, font_name: str = '微软雅黑', font_size: int = 11, bold_header: bool = True):
        """统一字体"""
        font = Font(name=font_name, size=font_size)
        header_font = Font(name=font_name, size=font_size, bold=True)

        for row_idx, row in enumerate(self.ws.iter_rows(), 1):
            for cell in row:
                if row_idx == 1 and bold_header:
                    cell.font = header_font
                else:
                    cell.font = font
        self._log(f"✅ 已统一字体：{font_name} {font_size}号")

    def format_alignment(self, horizontal: str = 'center', vertical: str = 'center',
                         wrap_text: bool = False):
        """统一对齐方式"""
        align = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)
        for row in self.ws.iter_rows():
            for cell in row:
                cell.alignment = align
        self._log(f"✅ 已统一对齐：{horizontal} / {vertical}")

    def format_auto_width(self, max_width: int = 50):
        """列宽自适应"""
        for col_idx in range(1, self.ws.max_column + 1):
            max_len = 0
            col_letter = get_column_letter(col_idx)
            for row_idx in range(1, self.ws.max_row + 1):
                cell = self.ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    # 中文字符算2个宽度
                    val = str(cell.value)
                    length = sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in val)
                    max_len = max(max_len, length)
            width = min(max_len + 2, max_width)
            self.ws.column_dimensions[col_letter].width = max(width, 8)
        self._log(f"✅ 列宽已自适应（最大{max_width}）")

    def format_border(self, style: str = 'thin'):
        """设置边框"""
        border_style = {
            'thin': Side(style='thin'),
            'medium': Side(style='medium'),
        }
        side = border_style.get(style, Side(style='thin'))
        border = Border(left=side, right=side, top=side, bottom=side)

        for row in self.ws.iter_rows():
            for cell in row:
                cell.border = border
        self._log(f"✅ 已设置边框（{style}）")

    # ===================== 格式转换 =====================

    def convert_to_csv(self, output_path: str, encoding: str = 'utf-8-sig'):
        """xlsx → csv"""
        df = pd.read_excel(self.input_path, dtype=str)
        df.to_csv(output_path, index=False, encoding=encoding)
        self._log(f"✅ 已导出CSV：{Path(output_path).name}")

    def convert_from_csv(self, csv_path: str, output_path: str):
        """csv → xlsx"""
        df = pd.read_csv(csv_path, dtype=str)
        df.to_excel(output_path, index=False, engine='openpyxl')
        self._log(f"✅ 已导出Excel：{Path(output_path).name}")

    # ===================== 保存 =====================

    def save(self, output_path: str = None) -> str:
        """保存文件"""
        save_path = output_path or self.input_path
        self.wb.save(save_path)
        self._log(f"💾 已保存：{Path(save_path).name}")
        return str(save_path)