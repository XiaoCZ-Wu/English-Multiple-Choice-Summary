"""
设置来源对话框
用于批量设置题目来源
"""

from PySide6.QtWidgets import (
    QDialog, QLineEdit, QDialogButtonBox,
    QFormLayout, QVBoxLayout, QLabel
)
from PySide6.QtCore import Qt


class SourceDialog(QDialog):
    """设置来源输入对话框"""

    def __init__(self, parent=None, current_source="", title="设置来源"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(400)

        # 创建控件
        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("请输入题目来源，如：2024年高考真题")
        self.source_edit.setMinimumWidth(350)
        self.source_edit.setText(current_source)

        # 布局
        layout = QFormLayout()
        layout.setSpacing(15)
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

        layout.addRow("题目来源:", self.source_edit)

        # 按钮
        buttons = QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        button_box = QDialogButtonBox(buttons, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.addLayout(layout)
        main_layout.addWidget(button_box)

    def get_source(self):
        """获取输入的来源"""
        return self.source_edit.text().strip()
