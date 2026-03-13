"""
导出对话框
"""

from PySide6.QtWidgets import (
    QDialog, QLineEdit, QCheckBox, QRadioButton, 
    QButtonGroup, QDialogButtonBox, QFormLayout
)

from src.core import ExportOptions


class ExportDialog(QDialog):
    """导出设置对话框"""
    
    def __init__(self, parent=None, default_path="./output/"):
        super().__init__(parent)
        self.setWindowTitle("导出文件")
        self.setModal(True)
        
        # 创建控件
        self.filename_edit = QLineEdit()
        self.filename_edit.setPlaceholderText("请输入文件名，默认为New")
        self.filename_edit.setMinimumWidth(300)
        
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("请输入文档标题，可选")
        self.title_edit.setMinimumWidth(300)
        
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("请指定生成目录，默认为./output/")
        self.path_edit.setMinimumWidth(300)
        self.path_edit.setText(default_path)
        
        self.answer_checkbox = QCheckBox("将答案写入文件末尾（答案单独一页）")
        self.answer_card_checkbox = QCheckBox("创建对应题库并附带答题卡")
        self.source_checkbox = QCheckBox("在题目结尾添加出处")
        
        self.docx_radio = QRadioButton(".docx")
        self.pdf_radio = QRadioButton(".pdf")
        self.csv_radio = QRadioButton(".csv")
        
        self.format_group = QButtonGroup(self)
        self.format_group.addButton(self.docx_radio, 0)
        self.format_group.addButton(self.pdf_radio, 1)
        self.format_group.addButton(self.csv_radio, 2)
        
        # 布局
        buttons = QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        button_box = QDialogButtonBox(buttons, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        layout = QFormLayout(self)
        layout.addRow("文件名称", self.filename_edit)
        layout.addRow("文档标题", self.title_edit)
        layout.addRow("文件路径", self.path_edit)
        layout.addWidget(self.docx_radio)
        layout.addWidget(self.pdf_radio)
        layout.addWidget(self.csv_radio)
        layout.addWidget(self.answer_checkbox)
        layout.addWidget(self.answer_card_checkbox)
        layout.addWidget(self.source_checkbox)
        layout.addWidget(button_box)
    
    def get_filename(self) -> str:
        """获取文件名"""
        name = self.filename_edit.text()
        return name if name else "new"
    
    def get_title(self) -> str:
        """获取文档标题"""
        return self.title_edit.text()
    
    def get_path(self) -> str:
        """获取文件路径"""
        path = self.path_edit.text()
        if path and not (path.endswith("/") or path.endswith("\\")):
            path += "/"
        return path if path else "./output/"
    
    def get_export_options(self) -> ExportOptions:
        """获取导出选项"""
        return ExportOptions(
            include_answer=self.answer_checkbox.isChecked(),
            include_answer_card=self.answer_card_checkbox.isChecked(),
            include_source=self.source_checkbox.isChecked()
        )
    
    def get_format(self) -> int:
        """获取选择的格式 0=docx, 1=pdf, 2=csv"""
        return self.format_group.checkedId()
