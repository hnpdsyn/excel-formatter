"""
Excel 格式整理工具 - 图形界面
"""

import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QFileDialog, QTextEdit,
    QComboBox, QSpinBox, QLineEdit, QGroupBox, QCheckBox,
    QMessageBox, QGridLayout, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, QThread, Signal

from core import ExcelFormatter


class FormatterWorker(QThread):
    log_signal = Signal(str)
    done = Signal(bool, str)

    def __init__(self, work_func):
        super().__init__()
        self.work_func = work_func

    def run(self):
        try:
            self.work_func()
            self.done.emit(True, "✅ 操作完成！")
        except Exception as e:
            self.done.emit(False, f"❌ 操作失败：{str(e)}")


class ExcelFormatterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 格式整理工具")
        self.setMinimumSize(750, 580)
        self.resize(850, 640)

        self.formatter = ExcelFormatter()
        self.current_file = None
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 12)

        # 标题
        title = QLabel("✨ Excel 格式整理工具")
        title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 8px 0;")
        layout.addWidget(title)

        # 文件选择栏
        file_bar = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("选择Excel文件...")
        self.btn_open = QPushButton("📂 打开文件")
        self.btn_open.clicked.connect(self._open_file)
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: #888;")
        file_bar.addWidget(self.file_path, 1)
        file_bar.addWidget(self.btn_open)
        file_bar.addWidget(self.file_label)
        layout.addLayout(file_bar)

        # 主Tab
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #ccc; border-radius: 4px; padding: 8px; }
            QTabBar::tab { padding: 8px 20px; font-size: 13px; }
            QTabBar::tab:selected { font-weight: bold; border-bottom: 2px solid #0078d4; }
        """)
        tabs.addTab(self._build_clean_tab(), "  一键清理  ")
        tabs.addTab(self._build_format_tab(), "  格式统一  ")
        tabs.addTab(self._build_convert_tab(), "  格式转换  ")
        layout.addWidget(tabs, 1)

        # 底部提示
        tip = QLabel("💡 支持 .xlsx / .xls / .xlsm 格式 | 建议先备份原文件")
        tip.setStyleSheet("color: #888; font-size: 12px; padding: 4px 0;")
        layout.addWidget(tip)

    # ===================== 文件操作 =====================

    def _open_file(self, path=None):
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择Excel文件",
                "", "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
            )
        if path:
            self.file_path.setText(path)
            self.current_file = path
            fname = Path(path).name
            self.file_label.setText(f"📄 {fname}")
            self.file_label.setStyleSheet("color: #333; font-weight: bold;")
            self.formatter.load(path)
            self._load_sheets()

    def _load_sheets(self):
        sheets = self.formatter.get_sheet_names()
        if hasattr(self, 'sheet_combo') and sheets:
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)

    def _get_save_path(self, suffix: str = "_整理"):
        if not self.current_file:
            return None
        p = Path(self.current_file)
        return str(p.parent / f"{p.stem}{suffix}{p.suffix}")

    # ===================== 清理Tab =====================

    def _build_clean_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(10)

        # 清理选项
        opt_group = QGroupBox("清理选项")
        opt_layout = QGridLayout(opt_group)

        self.cb_trim = QCheckBox("去除前后空格")
        self.cb_trim.setChecked(True)
        self.cb_empty_row = QCheckBox("删除空行")
        self.cb_empty_row.setChecked(True)
        self.cb_empty_col = QCheckBox("删除空列")
        self.cb_duplicate = QCheckBox("删除重复行")
        self.cb_duplicate.setChecked(True)
        self.cb_special = QCheckBox("清除特殊字符")
        self.cb_special.setChecked(True)
        self.cb_chinese = QCheckBox("仅保留中文（配合特殊字符使用）")

        opt_layout.addWidget(self.cb_trim, 0, 0)
        opt_layout.addWidget(self.cb_empty_row, 0, 1)
        opt_layout.addWidget(self.cb_empty_col, 0, 2)
        opt_layout.addWidget(self.cb_duplicate, 1, 0)
        opt_layout.addWidget(self.cb_special, 1, 1)
        opt_layout.addWidget(self.cb_chinese, 1, 2)
        layout.addWidget(opt_group)

        # 表头设置
        header_group = QGroupBox("设置")
        header_layout = QHBoxLayout(header_group)
        self.cb_has_header = QCheckBox("第一行为表头（删除重复行时保留）")
        self.cb_has_header.setChecked(True)
        header_layout.addWidget(self.cb_has_header)
        header_layout.addStretch()
        layout.addWidget(header_group)

        # 输出设置
        output_group = QGroupBox("输出")
        output_layout = QHBoxLayout(output_group)
        self.radio_clean_overwrite = QRadioButton("覆盖原文件")
        self.radio_clean_new = QRadioButton("另存为新文件")
        self.radio_clean_new.setChecked(True)
        self.clean_new_path = QLineEdit()
        self.clean_new_path.setPlaceholderText("自动生成：原文件名_整理.xlsx")
        self.btn_clean_browse = QPushButton("选择")
        self.btn_clean_browse.clicked.connect(lambda: self._select_save(self.clean_new_path))
        output_layout.addWidget(self.radio_clean_overwrite)
        output_layout.addWidget(self.radio_clean_new)
        output_layout.addWidget(self.clean_new_path, 1)
        output_layout.addWidget(self.btn_clean_browse)
        layout.addWidget(output_group)

        # 执行按钮
        self.btn_clean = QPushButton("🚀 开始清理")
        self.btn_clean.setStyleSheet("""
            QPushButton {
                background: #0078d4; color: white; font-size: 14px;
                padding: 10px 30px; border-radius: 6px; font-weight: bold;
            }
            QPushButton:hover { background: #106ebe; }
            QPushButton:disabled { background: #ccc; }
        """)
        self.btn_clean.clicked.connect(self._run_clean)
        layout.addWidget(self.btn_clean)

        # 日志
        log_group = QGroupBox("执行日志")
        log_layout = QVBoxLayout(log_group)
        self.clean_log = QTextEdit()
        self.clean_log.setReadOnly(True)
        self.clean_log.setStyleSheet("background: #1e1e1e; color: #d4d4d4; font-family: Consolas; font-size: 12px;")
        log_layout.addWidget(self.clean_log)
        layout.addWidget(log_group, 1)

        return widget

    # ===================== 格式Tab =====================

    def _build_format_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(10)

        # 字体设置
        font_group = QGroupBox("字体设置")
        font_layout = QGridLayout(font_group)
        font_layout.addWidget(QLabel("字体："), 0, 0)
        self.font_combo = QComboBox()
        self.font_combo.addItems(['微软雅黑', '宋体', '黑体', 'Arial', 'Calibri', 'Times New Roman'])
        self.font_combo.setCurrentText('微软雅黑')
        font_layout.addWidget(self.font_combo, 0, 1)
        font_layout.addWidget(QLabel("字号："), 0, 2)
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 48)
        self.font_size.setValue(11)
        font_layout.addWidget(self.font_size, 0, 3)
        self.cb_bold_header = QCheckBox("表头加粗")
        self.cb_bold_header.setChecked(True)
        font_layout.addWidget(self.cb_bold_header, 0, 4)
        layout.addWidget(font_group)

        # 对齐设置
        align_group = QGroupBox("对齐方式")
        align_layout = QHBoxLayout(align_group)
        align_layout.addWidget(QLabel("水平："))
        self.align_h = QComboBox()
        self.align_h.addItems(['center', 'left', 'right', 'general'])
        self.align_h.setCurrentText('center')
        align_layout.addWidget(self.align_h)
        align_layout.addWidget(QLabel("垂直："))
        self.align_v = QComboBox()
        self.align_v.addItems(['center', 'top', 'bottom'])
        self.align_v.setCurrentText('center')
        align_layout.addWidget(self.align_v)
        self.cb_wrap = QCheckBox("自动换行")
        align_layout.addWidget(self.cb_wrap)
        align_layout.addStretch()
        layout.addWidget(align_group)

        # 其他格式
        other_group = QGroupBox("其他设置")
        other_layout = QHBoxLayout(other_group)
        self.cb_auto_width = QCheckBox("列宽自适应")
        self.cb_auto_width.setChecked(True)
        other_layout.addWidget(self.cb_auto_width)
        self.cb_border = QCheckBox("添加边框")
        self.cb_border.setChecked(True)
        other_layout.addWidget(self.cb_border)
        other_layout.addStretch()
        layout.addWidget(other_group)

        # 输出
        output_group = QGroupBox("输出")
        output_layout = QHBoxLayout(output_group)
        self.radio_fmt_overwrite = QRadioButton("覆盖原文件")
        self.radio_fmt_new = QRadioButton("另存为新文件")
        self.radio_fmt_new.setChecked(True)
        self.fmt_new_path = QLineEdit()
        self.fmt_new_path.setPlaceholderText("自动生成：原文件名_格式化.xlsx")
        self.btn_fmt_browse = QPushButton("选择")
        self.btn_fmt_browse.clicked.connect(lambda: self._select_save(self.fmt_new_path))
        output_layout.addWidget(self.radio_fmt_overwrite)
        output_layout.addWidget(self.radio_fmt_new)
        output_layout.addWidget(self.fmt_new_path, 1)
        output_layout.addWidget(self.btn_fmt_browse)
        layout.addWidget(output_group)

        # 执行按钮
        self.btn_format = QPushButton("🎨 开始格式化")
        self.btn_format.setStyleSheet("""
            QPushButton {
                background: #0078d4; color: white; font-size: 14px;
                padding: 10px 30px; border-radius: 6px; font-weight: bold;
            }
            QPushButton:hover { background: #106ebe; }
            QPushButton:disabled { background: #ccc; }
        """)
        self.btn_format.clicked.connect(self._run_format)
        layout.addWidget(self.btn_format)

        # 日志
        log_group = QGroupBox("执行日志")
        log_layout = QVBoxLayout(log_group)
        self.fmt_log = QTextEdit()
        self.fmt_log.setReadOnly(True)
        self.fmt_log.setStyleSheet("background: #1e1e1e; color: #d4d4d4; font-family: Consolas; font-size: 12px;")
        log_layout.addWidget(self.fmt_log)
        layout.addWidget(log_group, 1)

        return widget

    # ===================== 转换Tab =====================

    def _build_convert_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(10)

        # 转换模式
        mode_group = QGroupBox("转换模式")
        mode_layout = QVBoxLayout(mode_group)

        row1 = QHBoxLayout()
        self.radio_xlsx2csv = QRadioButton("Excel → CSV")
        self.radio_xlsx2csv.setChecked(True)
        self.radio_xlsx2csv.toggled.connect(self._on_conv_mode_change)
        self.radio_csv2xlsx = QRadioButton("CSV → Excel")
        row1.addWidget(self.radio_xlsx2csv)
        row1.addWidget(self.radio_csv2xlsx)
        row1.addStretch()
        mode_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("编码："))
        self.csv_encoding = QComboBox()
        self.csv_encoding.addItems(['utf-8-sig', 'utf-8', 'gbk', 'gb2312'])
        self.csv_encoding.setCurrentText('utf-8-sig')
        row2.addWidget(self.csv_encoding)
        row2.addStretch()
        mode_layout.addLayout(row2)
        layout.addWidget(mode_group)

        # 输入
        input_group = QGroupBox("输入文件")
        input_layout = QHBoxLayout(input_group)
        self.conv_input = QLineEdit()
        self.conv_input.setPlaceholderText("选择要转换的文件...")
        self.btn_conv_input = QPushButton("📂 选择")
        self.btn_conv_input.clicked.connect(self._select_conv_input)
        input_layout.addWidget(self.conv_input, 1)
        input_layout.addWidget(self.btn_conv_input)
        layout.addWidget(input_group)

        # 输出
        output_group = QGroupBox("输出文件")
        output_layout = QHBoxLayout(output_group)
        self.conv_output = QLineEdit()
        self.conv_output.setPlaceholderText("自动生成...")
        self.btn_conv_output = QPushButton("选择路径")
        self.btn_conv_output.clicked.connect(lambda: self._select_save(self.conv_output))
        output_layout.addWidget(self.conv_output, 1)
        output_layout.addWidget(self.btn_conv_output)
        layout.addWidget(output_group)

        # 执行按钮
        self.btn_convert = QPushButton("🔄 开始转换")
        self.btn_convert.setStyleSheet("""
            QPushButton {
                background: #0078d4; color: white; font-size: 14px;
                padding: 10px 30px; border-radius: 6px; font-weight: bold;
            }
            QPushButton:hover { background: #106ebe; }
            QPushButton:disabled { background: #ccc; }
        """)
        self.btn_convert.clicked.connect(self._run_convert)
        layout.addWidget(self.btn_convert)

        # 日志
        log_group = QGroupBox("执行日志")
        log_layout = QVBoxLayout(log_group)
        self.conv_log = QTextEdit()
        self.conv_log.setReadOnly(True)
        self.conv_log.setStyleSheet("background: #1e1e1e; color: #d4d4d4; font-family: Consolas; font-size: 12px;")
        log_layout.addWidget(self.conv_log)
        layout.addWidget(log_group, 1)

        return widget

    def _on_conv_mode_change(self):
        is_xlsx2csv = self.radio_xlsx2csv.isChecked()
        self.csv_encoding.setEnabled(is_xlsx2csv)

    def _select_conv_input(self):
        if self.radio_xlsx2csv.isChecked():
            path, _ = QFileDialog.getOpenFileName(
                self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
            )
            if path:
                self.conv_input.setText(path)
                p = Path(path)
                self.conv_output.setText(str(p.parent / f"{p.stem}.csv"))
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择CSV文件", "", "CSV文件 (*.csv);;所有文件 (*.*)"
            )
            if path:
                self.conv_input.setText(path)
                p = Path(path)
                self.conv_output.setText(str(p.parent / f"{p.stem}.xlsx"))

    # ===================== 通用方法 =====================

    def _select_save(self, line_edit):
        path, _ = QFileDialog.getSaveFileName(self, "选择保存路径", line_edit.text() or "output.xlsx")
        if path:
            line_edit.setText(path)

    def _log(self, widget, msg):
        widget.append(msg)
        widget.verticalScrollBar().setValue(widget.verticalScrollBar().maximum())

    def _run_in_thread(self, work_func, log_widget, button):
        def wrapper():
            try:
                work_func()
            except Exception as e:
                raise e

        class Worker(QThread):
            log_signal = Signal(str)
            done = Signal(bool, str)

            def run(self):
                try:
                    work_func()
                    self.done.emit(True, "✅ 操作完成！")
                except Exception as e:
                    self.done.emit(False, f"❌ {str(e)}")

        self.thread = Worker()
        self.thread.log_signal.connect(lambda msg: self._log(log_widget, msg))
        self.thread.done.connect(lambda ok, msg: (
            self._log(log_widget, msg),
            button.setEnabled(True)
        ))
        button.setEnabled(False)
        self.thread.start()

    # ===================== 执行清理 =====================

    def _run_clean(self):
        if not self.current_file or not os.path.exists(self.current_file):
            QMessageBox.warning(self, "提示", "请先打开Excel文件")
            return

        self.formatter.load(self.current_file)

        def work():
            self._log(self.clean_log, f"📂 处理文件：{Path(self.current_file).name}")

            if self.cb_trim.isChecked():
                self.formatter.clean_trim_spaces()
            if self.cb_empty_row.isChecked():
                self.formatter.clean_empty_rows()
            if self.cb_empty_col.isChecked():
                self.formatter.clean_empty_columns()
            if self.cb_duplicate.isChecked():
                self.formatter.clean_duplicates(self.cb_has_header.isChecked())
            if self.cb_special.isChecked():
                self.formatter.clean_special_chars(self.cb_chinese.isChecked())

            if self.radio_clean_overwrite.isChecked():
                self.formatter.save(self.current_file)
            else:
                out_path = self.clean_new_path.text().strip() or self._get_save_path("_整理")
                if out_path:
                    self.formatter.save(out_path)

        self._run_in_thread(work, self.clean_log, self.btn_clean)

    # ===================== 执行格式化 =====================

    def _run_format(self):
        if not self.current_file or not os.path.exists(self.current_file):
            QMessageBox.warning(self, "提示", "请先打开Excel文件")
            return

        self.formatter.load(self.current_file)

        def work():
            self._log(self.fmt_log, f"📂 处理文件：{Path(self.current_file).name}")

            self.formatter.format_font(
                self.font_combo.currentText(),
                self.font_size.value(),
                self.cb_bold_header.isChecked()
            )
            self.formatter.format_alignment(
                self.align_h.currentText(),
                self.align_v.currentText(),
                self.cb_wrap.isChecked()
            )
            if self.cb_auto_width.isChecked():
                self.formatter.format_auto_width()
            if self.cb_border.isChecked():
                self.formatter.format_border()

            if self.radio_fmt_overwrite.isChecked():
                self.formatter.save(self.current_file)
            else:
                out_path = self.fmt_new_path.text().strip() or self._get_save_path("_格式化")
                if out_path:
                    self.formatter.save(out_path)

        self._run_in_thread(work, self.fmt_log, self.btn_format)

    # ===================== 执行转换 =====================

    def _run_convert(self):
        in_path = self.conv_input.text().strip()
        out_path = self.conv_output.text().strip()

        if not in_path or not os.path.exists(in_path):
            QMessageBox.warning(self, "提示", "请选择输入文件")
            return
        if not out_path:
            QMessageBox.warning(self, "提示", "请设置输出路径")
            return

        def work():
            self._log(self.conv_log, f"📂 输入：{Path(in_path).name}")

            if self.radio_xlsx2csv.isChecked():
                self.formatter.load(in_path)
                self.formatter.convert_to_csv(out_path, self.csv_encoding.currentText())
            else:
                self.formatter.convert_from_csv(in_path, out_path)

            self._log(self.conv_log, f"📁 输出：{Path(out_path).name}")

        self._run_in_thread(work, self.conv_log, self.btn_convert)


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = ExcelFormatterApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()