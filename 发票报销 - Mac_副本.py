import sys
import os
import re
import fitz
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QFileDialog,
    QMessageBox, QProgressBar, QDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent


# ================== PDF 提取逻辑 ==================
def extract_invoice_data(path):
    """从单张 + 循环"""
    try:
        doc = fitz.open(path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"[错误] 无法读取 {path}: {e}")
        return None

    data = {"文件名": os.path.basename(path)}


    # ==================== 付款人 ====================
    filename = os.path.basename(path)  
    name_without_ext = os.path.splitext(filename)[0] 

    parts = name_without_ext.split('-')
    candidate = parts[-1]  
    data["付款人"] = candidate if candidate else ""


    # ======== 发票号码：20位数字 ========
    m = re.search(r"(\d{20})", full_text)
    data["发票号码"] = m.group(1) if m else "⛔️"


    # ======== 开票日期 ========
    # 搜索年月日
    m = re.search(r"(\d{4}年\d{2}月\d{2}日)", full_text)
    data["开票日期"] = m.group(1) if m else "⛔️"


    # ======== 供应商 ========
    # 寻找特定规律，供应商在发票号前面
    """ ======== 优化空间 ======== """
    # 搜索方式

    try:
        pairs = re.findall(r"([^\n]{2,30}?)\s*\n\s*([A-Z0-9]{15,20})", full_text)

        valid_pairs = [(name, tax) for name, tax in pairs 
                if re.search(r'[A-Z]', tax) and re.search(r'\d', tax) or re.search(r'\d{18}', tax) ]

        data["供应商"] = valid_pairs[2][0] if len(valid_pairs) >= 3 else valid_pairs[1][0]
    except Exception as e:
        print(f'   ⛔️ {path} 无供应商')
        data["供应商"] = "⛔️"


    # ======== 数量 + 单位 ======== #
    """ ========= 注释 ========== """
    # 代码识别不同发票的差异大 - 数量一般集中在单位后四行内
    # 先搜索单位后，遍历单位后四行的数值，只搜寻整数
    # 输出整数的同时输出相应的单位
    """ ======== 优化空间 ======== """
    # 单位的识别
    # 搜索方式

    lines = [t.strip() for t in re.split(r'\s+', full_text) if t.strip()]
    units = ["条", "套", "个", "盒", "卷", "支", "包", "批", "台", "件",
             "根", "只", "把", "米", "次", "份", "粒", "克", "斤", "张"]

    quantity = 0
    for i, line in enumerate(lines):
        if line in units:
            next_line_i = i+1
            for i in range (next_line_i,next_line_i+4):
                next_line = lines[i]
                if re.match(r"^\d+$", next_line):  # 纯整数
                    quantity += int(next_line)
                    data["数量"] = quantity
                    data["单位"] = line if line not in r"套包盒批份次" else line + "❌"
 

    # ======== 金额：提取所有 ¥开头的金额 ========
    #查找以¥开头的数字，取最大值

    amounts = re.findall(r"¥\s*(\d+\.\d{2})", full_text)
    data["金额"] = max(amounts) if amounts else ""


    # ======== 单价 ======== 
    # 金额 / 数量

    try:
        ave = round(float(max(amounts)) / int(quantity),2)
    except Exception as e:
        print(f'   ⛔️ {path} 无单价')
        ave = 0
    data["单价"] = ave


    # ======== 类型：易耗品 + 低值判断 ======== 
    # 通过单价判断
    data["类型"] = "易耗品" if ave < 200 else "❗️低值"


    # ======== 项目名称 ======== 
    # 文件名用 "-" 分割，取倒一

    try:
        m = re.findall(r"\*([\u4e00-\u9fa5A-Z0-9]+)\*", full_text)
        data["项目名称"] = m[0]
    except Exception as e:
        print(f'   ⛔️ {path} 无项目名称')
        data["项目名称"] = "⛔️"


    # ======== 规格型号 ========
    # 文件名用 "-" 分割，取倒一

    try:
        a = re.findall(r"\*[\u4e00-\u9fa5A-Z0-9]+\*[\w\u4e00-\u9fa5A-Z0-9]+\n([\w\u4e00-\u9fa5A-Z0-9]+)", full_text)
        candidate = ""
        for can in a:
            candidate += "\n" + can
        if len(candidate) > 4:
            data["规格型号"] = candidate
        else:
            data["规格型号"] = ""
        # ["1","2","3","4","5","6","7","8","9","0"]
    except Exception as e:
        print(f'   ⛔️ {path} 规格型号')
        data["项目名称"] = "⛔️"


    return data



# ========== batch_invoices_to_excel ========== #
#遍历文件 + 提取信息

def process_pdfs_to_excel(file_paths, save_path, columns_order=None):

    records = []
    for path in file_paths:
        record = extract_invoice_data(path)
        if record:
            records.append(record)

    if not records:
        return None

    # 定义默认列顺序
    default_columns = [
        "文件名", "类型", "项目名称", "单价", "数量", 
        "单位", "金额", "发票号码", 
        "开票日期", "供应商", "规格型号","付款人"
    ]

    columns = columns_order if columns_order else default_columns

    df = pd.DataFrame(records)
    # 确保所有默认列存在，缺失列补空字符串
    for col in default_columns:
        if col not in df.columns:
            df[col] = ""
    
    # 先计算汇总（使用完整数据，不受用户删除列的影响）
    df["金额"] = pd.to_numeric(df["金额"], errors="coerce")
    summary = df.groupby("付款人", as_index=False)["金额"].sum()
    summary["金额合计"] = summary["金额"].apply(lambda x: f"¥{x:.2f}")
    summary = summary[["付款人", "金额合计"]]

    # 再按用户指定顺序和选择输出列
    output_columns = [col for col in columns if col in df.columns]
    df_output = df[output_columns]

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
    
        # Sheet 1：按用户指定列顺序和筛选输出
        df_output.to_excel(writer, sheet_name="发票明细", index=False)
        
        # Sheet 2：每人合计
        summary.to_excel(writer, sheet_name="按人汇总", index=False)

    # 生成Excel
    return save_path


# ================== 列顺序确认对话框 ==================
class ColumnOrderDialog(QDialog):
    def __init__(self, default_columns, parent=None):
        super().__init__(parent)
        self.default_columns = default_columns[:]
        self.setWindowTitle("确认输出列顺序")
        self.setFixedSize(460, 600)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        # === 顶部标题栏 ===
        header = QWidget()
        header.setStyleSheet("background-color: #B91C1C; border-top-left-radius: 10px; border-top-right-radius: 10px;")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(24, 24, 24, 20)
        
        title = QLabel("确认输出列顺序")
        title_font = QFont("Yuanti SC", 18, QFont.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #ffffff;")
        header_layout.addWidget(title)
        
        subtitle = QLabel("拖拽调整顺序 · 点击 ✕ 删除不需要的列")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #FECACA; font-size: 13px; margin-top: 4px;")
        header_layout.addWidget(subtitle)
        
        layout.addWidget(header)

        # === 内容区 ===
        content = QWidget()
        content.setStyleSheet("background-color: #ffffff;")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(24, 20, 24, 16)
        content_layout.setSpacing(12)
        
        hint = QLabel("提示：长按 ☰ 上下拖拽调整顺序，点击右侧 ✕ 移除列")
        hint.setStyleSheet("color: #9CA3AF; font-size: 12px; padding: 8px 12px; background: #F9FAFB; border-radius: 6px;")
        hint.setWordWrap(True)
        content_layout.addWidget(hint)

        # 可拖拽列表
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QListWidget.InternalMove)
        self.list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #E5E7EB;
                border-radius: 10px;
                padding: 10px;
                font-size: 14px;
                background-color: #FAFAFA;
            }
            QListWidget::item {
                padding: 0px;
                margin: 5px 2px;
                border-radius: 8px;
                border: 1px solid #F3F4F6;
                background: #ffffff;
            }
            QListWidget::item:selected {
                background: #FEE2E2;
                border: 1px solid #FCA5A5;
            }
        """)
        
        for col in default_columns:
            self._add_column_item(col)
            
        content_layout.addWidget(self.list_widget)
        layout.addWidget(content, stretch=1)

        # === 底部按钮栏 ===
        footer = QWidget()
        footer.setStyleSheet("background-color: #F9FAFB; border-bottom-left-radius: 10px; border-bottom-right-radius: 10px;")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(24, 16, 24, 20)
        footer_layout.setSpacing(12)
        
        self.btn_reset = QPushButton("恢复默认")
        self.btn_reset.setFixedHeight(40)
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.setStyleSheet("""
            QPushButton {
                background-color: #E5E7EB;
                color: #4B5563;
                border: none;
                border-radius: 8px;
                padding: 0 18px;
                font-size: 13px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #D1D5DB;
            }
            QPushButton:pressed {
                background-color: #9CA3AF;
            }
        """)
        self.btn_reset.clicked.connect(self.reset_columns)
        
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedHeight(40)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #6B7280;
                border: 1px solid #E5E7EB;
                border-radius: 8px;
                padding: 0 24px;
                font-size: 14px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #F9FAFB;
                border-color: #D1D5DB;
            }
            QPushButton:pressed {
                background-color: #F3F4F6;
            }
        """)
        self.btn_cancel.clicked.connect(self.reject)
        
        self.btn_ok = QPushButton("确认输出")
        self.btn_ok.setFixedHeight(40)
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #B91C1C;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 0 28px;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #991B1B;
            }
            QPushButton:pressed {
                background-color: #7F1D1D;
            }
        """)
        self.btn_ok.clicked.connect(self.on_confirm)
        
        footer_layout.addWidget(self.btn_reset)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_cancel)
        footer_layout.addWidget(self.btn_ok)
        layout.addWidget(footer)
        
        # 对话框整体圆角和阴影效果
        self.setStyleSheet("""
            QDialog {
                background-color: transparent;
                border-radius: 10px;
            }
        """)

    def _add_column_item(self, col_name):
        item = QListWidgetItem()
        item.setSizeHint(QSize(0, 52))
        item.setFlags(item.flags() | Qt.ItemIsEnabled)
        self.list_widget.addItem(item)
        
        widget = QWidget()
        hlayout = QHBoxLayout(widget)
        hlayout.setContentsMargins(16, 0, 12, 0)
        hlayout.setSpacing(12)
        
        drag_label = QLabel("☰")
        drag_label.setStyleSheet("font-size: 15px; color: #D1D5DB; border: none; background: transparent;")
        hlayout.addWidget(drag_label)
        
        name_label = QLabel(col_name)
        name_label.setStyleSheet("font-size: 15px; color: #374151; font-weight: 500; border: none; background: transparent;")
        hlayout.addWidget(name_label, stretch=1)
        
        btn_delete = QPushButton("✕")
        btn_delete.setFixedSize(28, 28)
        btn_delete.setCursor(Qt.PointingHandCursor)
        btn_delete.setStyleSheet("""
            QPushButton {
                background-color: #F3F4F6;
                color: #9CA3AF;
                border: none;
                border-radius: 14px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                color: #B91C1C;
            }
            QPushButton:pressed {
                background-color: #FECACA;
            }
        """)
        btn_delete.clicked.connect(lambda checked, i=item: self.delete_item(i))
        hlayout.addWidget(btn_delete)
        
        self.list_widget.setItemWidget(item, widget)
    
    def delete_item(self, item):
        row = self.list_widget.row(item)
        self.list_widget.takeItem(row)
    
    def reset_columns(self):
        self.list_widget.clear()
        for col in self.default_columns:
            self._add_column_item(col)
    
    def on_confirm(self):
        if self.list_widget.count() == 0:
            QMessageBox.warning(self, "提示", "至少需要保留一列才能输出！")
            return
        self.accept()

    def get_order(self):
        columns = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            widget = self.list_widget.itemWidget(item)
            if widget:
                labels = widget.findChildren(QLabel)
                for label in labels:
                    text = label.text().strip()
                    if text and text != "☰":
                        columns.append(text)
                        break
        return columns


# ================== 后台线程 ==================
class Worker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, file_paths, save_path, columns_order=None):
        super().__init__()
        self.file_paths = file_paths
        self.save_path = save_path
        self.columns_order = columns_order

    def run(self):
        try:
            total = len(self.file_paths)
            for i, _ in enumerate(self.file_paths, 1):
                self.progress.emit(int(i / total * 100))

            result = process_pdfs_to_excel(self.file_paths, self.save_path, self.columns_order)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit("未能从PDF中提取到有效数据")
        except Exception as e:
            self.error.emit(str(e))


# ================== 主窗口 ==================
class InvoiceWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("机器人实验室发票报销小程序 · 3.5")
        self.setGeometry(543, 124, 560, 780)
        self.setAcceptDrops(True)

        self.pdf_files = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        # === 顶部标题栏 ===
        header = QWidget()
        header.setStyleSheet("background-color: #B91C1C;")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(28, 32, 28, 24)
        header_layout.setSpacing(6)

        title = QLabel("低值及耗材类采购申请明细")
        title_font = QFont("Yuanti SC", 22, QFont.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #ffffff;")
        header_layout.addWidget(title)

        version = QLabel("机器人实验室 · 发票报销小程序 v3.5")
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet("color: #FECACA; font-size: 13px;")
        header_layout.addWidget(version)

        layout.addWidget(header)

        # === 主体内容区 ===
        body = QWidget()
        body.setStyleSheet("background-color: #F9FAFB;")
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(28, 24, 28, 20)
        body_layout.setSpacing(16)

        # 拖拽区
        self.drop_label = QLabel("拖拽 PDF 发票文件到此处\n或点击下方按钮选择文件")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setMinimumHeight(110)
        self.drop_label.setStyleSheet("""
            QLabel {
                background-color: #ffffff;
                border: 1px solid #E5E7EB;
                border-radius: 14px;
                color: #9CA3AF;
                font-size: 14px;
                padding: 20px;
                line-height: 1.6;
            }
        """)
        body_layout.addWidget(self.drop_label)

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)

        self.btn_add = QPushButton("选择 PDF 文件")
        self.btn_add.setFixedHeight(42)
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setStyleSheet("""
            QPushButton {
                background-color: #B91C1C;
                color: white;
                border: none;
                border-radius: 10px;
                padding: 0 20px;
                font-size: 14px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #991B1B;
            }
            QPushButton:pressed {
                background-color: #7F1D1D;
            }
            QPushButton:disabled {
                background-color: #E5E7EB;
                color: #9CA3AF;
            }
        """)
        self.btn_add.clicked.connect(self.select_files)

        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.setFixedHeight(42)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #6B7280;
                border: 1px solid #E5E7EB;
                border-radius: 10px;
                padding: 0 20px;
                font-size: 14px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #F3F4F6;
                border-color: #D1D5DB;
            }
            QPushButton:pressed {
                background-color: #E5E7EB;
            }
            QPushButton:disabled {
                background-color: #F9FAFB;
                color: #D1D5DB;
            }
        """)
        self.btn_clear.clicked.connect(self.clear_files)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_clear)
        body_layout.addLayout(btn_layout)

        # 文件列表标题
        list_header_layout = QHBoxLayout()
        list_title = QLabel("待处理文件")
        list_title.setStyleSheet("color: #374151; font-size: 14px; font-weight: 600;")
        list_header_layout.addWidget(list_title)
        
        self.file_count_label = QLabel("0 个文件")
        self.file_count_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.file_count_label.setStyleSheet("color: #9CA3AF; font-size: 12px;")
        list_header_layout.addWidget(self.file_count_label)
        body_layout.addLayout(list_header_layout)

        # 文件列表
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #E5E7EB;
                border-radius: 10px;
                padding: 8px;
                font-size: 13px;
                background-color: #ffffff;
            }
            QListWidget::item {
                padding: 10px 12px;
                margin: 3px 0px;
                border-radius: 6px;
                border: none;
                color: #374151;
            }
            QListWidget::item:selected {
                background-color: #FEE2E2;
                color: #B91C1C;
            }
            QListWidget::item:hover {
                background-color: #F9FAFB;
            }
        """)
        body_layout.addWidget(self.list_widget, stretch=1)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFixedHeight(8)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                text-align: center;
                background-color: #E5E7EB;
            }
            QProgressBar::chunk {
                background-color: #B91C1C;
                border-radius: 4px;
            }
        """)
        body_layout.addWidget(self.progress)

        # 生成按钮
        self.btn_run = QPushButton("生成 Excel 明细")
        self.btn_run.setFixedHeight(50)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        run_font = QFont("Yuanti SC", 14, QFont.Bold)
        self.btn_run.setFont(run_font)
        self.btn_run.setStyleSheet("""
            QPushButton {
                background-color: #B91C1C;
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #991B1B;
            }
            QPushButton:pressed {
                background-color: #7F1D1D;
            }
            QPushButton:disabled {
                background-color: #E5E7EB;
                color: #9CA3AF;
            }
        """)
        self.btn_run.clicked.connect(self.run_extraction)
        body_layout.addWidget(self.btn_run)

        layout.addWidget(body, stretch=1)

        # === 底部状态栏 ===
        footer = QWidget()
        footer.setStyleSheet("background-color: #ffffff;")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(28, 12, 28, 12)
        
        self.status = QLabel("KAISEN ©")
        self.status.setStyleSheet("color: #D1D5DB; font-size: 11px;")
        footer_layout.addWidget(self.status)
        
        layout.addWidget(footer)

    # ================== 拖拽 ==================
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith('.pdf') and path not in self.pdf_files:
                self.pdf_files.append(path)
                item = QListWidgetItem(f"  {os.path.basename(path)}")
                item.setToolTip(path)
                self.list_widget.addItem(item)
        self._update_status()

    # ================== 文件操作 ==================
    def select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择 PDF 发票文件", "", "PDF Files (*.pdf)"
        )
        for path in paths:
            if path not in self.pdf_files:
                self.pdf_files.append(path)
                item = QListWidgetItem(f"  {os.path.basename(path)}")
                item.setToolTip(path)
                self.list_widget.addItem(item)
        self._update_status()

    def clear_files(self):
        self.pdf_files.clear()
        self.list_widget.clear()
        self._update_status()

    def _update_status(self):
        count = len(self.pdf_files)
        self.file_count_label.setText(f"{count} 个文件")
        if count:
            self.drop_label.setText(f"已加载 {count} 个 PDF 文件\n可继续拖拽追加更多文件")
            self.drop_label.setStyleSheet("""
                QLabel {
                    background-color: #FEF2F2;
                    border: 1px solid #E5E7EB;
                    border-radius: 14px;
                    color: #B91C1C;
                    font-size: 14px;
                    padding: 20px;
                    line-height: 1.6;
                }
            """)
        else:
            self.drop_label.setText("拖拽 PDF 发票文件到此处\n或点击下方按钮选择文件")
            self.drop_label.setStyleSheet("""
                QLabel {
                    background-color: #ffffff;
                    border: 1px solid #E5E7EB;
                    border-radius: 14px;
                    color: #9CA3AF;
                    font-size: 14px;
                    padding: 20px;
                    line-height: 1.6;
                }
            """)

    # ================== 核心处理 ==================
    def run_extraction(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "提示", "请先拖拽或选择 PDF 文件")
            return

        # 弹出列顺序确认对话框
        default_columns = [
            "文件名", "类型", "项目名称", "单价", "数量", 
            "单位", "金额", "发票号码", 
            "开票日期", "供应商", "规格型号", "付款人"
        ]
        dialog = ColumnOrderDialog(default_columns, self)
        if dialog.exec_() != QDialog.Accepted:
            return
        
        columns_order = dialog.get_order()

        default_name = f"{datetime.now().strftime('%Y%m')}低值及耗材类采购申请明细.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存 Excel 报表", default_name, "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        self.btn_run.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.status.setText("正在提取数据，请稍候...")
        self.progress.setValue(0)

        self.worker = Worker(self.pdf_files, save_path, columns_order)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_success(self, path):
        self.btn_run.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.status.setText("KAISEN ©")
        self.progress.setValue(100)

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("处理完成")
        msg_box.setText(f"<p style='font-size:14px;'>Excel 已保存到：</p><p style='color:#B91C1C;font-weight:600;'>{path}</p>")
        msg_box.setInformativeText("是否立即打开所在文件夹？")
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.Yes)
        msg_box.button(QMessageBox.Yes).setText("打开文件夹")
        msg_box.button(QMessageBox.No).setText("关闭")
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #B91C1C;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #991B1B;
            }
        """)
        
        reply = msg_box.exec_()
        if reply == QMessageBox.Yes:
            os.system(f'open "{os.path.dirname(path)}"')

    def on_error(self, msg):
        self.btn_run.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.progress.setValue(0)
        self.status.setText("KAISEN ©")
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("错误")
        msg_box.setText(f"处理失败")
        msg_box.setInformativeText(msg)
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #B91C1C;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #991B1B;
            }
        """)
        msg_box.exec_()


# ================== MAIN ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 使用圆润中文字体
    font = QFont("Yuanti SC", 10)
    if not QFont(font).exactMatch():
        font = QFont("PingFang SC", 10)
    app.setFont(font)

    window = InvoiceWindow()
    window.show()
    sys.exit(app.exec_())
