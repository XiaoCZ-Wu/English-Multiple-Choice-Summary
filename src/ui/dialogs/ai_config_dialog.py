"""
AI配置对话框
用于添加/编辑AI配置
"""

from PySide6.QtWidgets import (
    QDialog, QLineEdit, QDialogButtonBox,
    QFormLayout, QVBoxLayout, QLabel
)
from PySide6.QtCore import Qt


class AIConfigDialog(QDialog):
    """AI配置输入对话框"""

    def __init__(self, parent=None, name="", base_url="", api_key="", model="", title="添加AI配置", is_edit=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(700)  # 窗口加宽
        self.is_edit = is_edit
        self.original_api_key = api_key if is_edit else ""

        # 创建控件
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("请输入AI名称，如：智谱AI、OpenAI等")
        self.name_edit.setMinimumWidth(600)
        self.name_edit.setText(name)

        self.base_url_edit = QLineEdit()
        self.base_url_edit.setPlaceholderText("请输入Base URL，如：https://api.openai.com/v1")
        self.base_url_edit.setMinimumWidth(600)
        self.base_url_edit.setText(base_url)

        self.api_key_edit = QLineEdit()
        if is_edit:
            # 编辑模式：显示提示，留空表示不修改
            self.api_key_edit.setPlaceholderText("留空表示不修改API Key，输入新值则替换")
            self.api_key_edit.setText("")  # 编辑时默认空
        else:
            # 添加模式
            self.api_key_edit.setPlaceholderText("请输入API Key")
        self.api_key_edit.setMinimumWidth(600)
        self.api_key_edit.setEchoMode(QLineEdit.Password)  # 密码模式

        # 模型输入（改为文本输入）
        self.model_edit = QLineEdit()
        self.model_edit.setPlaceholderText("请输入模型名称，如：gpt-4o、glm-4等")
        self.model_edit.setMinimumWidth(600)
        self.model_edit.setText(model)

        # 布局 - 顺序：名称、baseurl、模型、key
        layout = QFormLayout()
        layout.setSpacing(15)  # 增加行间距
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)

        layout.addRow("AI名称:", self.name_edit)
        layout.addRow("Base URL:", self.base_url_edit)
        layout.addRow("模型:", self.model_edit)
        layout.addRow("API Key:", self.api_key_edit)

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

    def get_values(self):
        """获取输入的值"""
        api_key = self.api_key_edit.text().strip()
        # 编辑模式下，如果API Key为空，则保留原值
        if self.is_edit and not api_key:
            api_key = self.original_api_key
        return {
            "name": self.name_edit.text().strip(),
            "base_url": self.base_url_edit.text().strip(),
            "api_key": api_key,
            "model": self.model_edit.text().strip()
        }
