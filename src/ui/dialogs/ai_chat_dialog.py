"""
AI对话窗口
用于与AI进行交互式对话，可以询问题目相关问题
"""

from typing import Optional
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QTextCursor
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QPushButton, QMessageBox,
    QMenu, QSplitter, QWidget, QScrollArea, QFrame, QSpacerItem, QSizePolicy
)

from src.models import Question, AIConfig


class AIChatThread(QThread):
    """AI对话线程"""
    response_received = Signal(str)  # 收到回复
    error_occurred = Signal(str)  # 发生错误
    finished_signal = Signal()  # 完成信号

    def __init__(self, ai_config: AIConfig, messages: list):
        super().__init__()
        self.ai_config = ai_config
        self.messages = messages

    def run(self):
        """执行AI请求"""
        try:
            import requests

            headers = {
                "Authorization": f"Bearer {self.ai_config.api_key}",
                "Content-Type": "application/json"
            }

            data = {
                "model": self.ai_config.model if self.ai_config.model else "glm-4v-flash",
                "messages": self.messages
            }

            base_url = self.ai_config.base_url.rstrip('/')
            if not base_url.endswith('/chat/completions'):
                base_url += '/chat/completions'

            response = requests.post(
                base_url,
                headers=headers,
                json=data,
                timeout=60
            )
            response.raise_for_status()

            result = response.json()
            if 'choices' in result and len(result['choices']) > 0:
                content = result['choices'][0]['message']['content']
                self.response_received.emit(content)
            else:
                self.error_occurred.emit("AI返回格式错误")

        except Exception as e:
            self.error_occurred.emit(f"请求失败: {str(e)}")
        finally:
            self.finished_signal.emit()


class MessageWidget(QFrame):
    """单条消息组件"""
    
    def __init__(self, sender: str, content: str, parent=None):
        super().__init__(parent)
        self.sender = sender
        self.content = content
        self._setup_ui()
        
    def _setup_ui(self):
        """设置UI - 使用普通主题色"""
        layout = QVBoxLayout(self)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 8, 10, 8)
        
        # 使用统一的普通样式
        self.setStyleSheet("""
            MessageWidget {
                background-color: palette(base);
                border: 1px solid palette(mid);
                border-radius: 4px;
            }
        """)
        
        # 发送者标签
        sender_label = QLabel(f"<b>{self.sender}</b>")
        layout.addWidget(sender_label)
        
        # 内容文本框（只读，但可选择复制）
        content_edit = QTextEdit()
        content_edit.setPlainText(self.content)
        content_edit.setReadOnly(True)
        content_edit.setFrameStyle(QFrame.NoFrame)
        content_edit.setStyleSheet("""
            QTextEdit {
                background-color: transparent;
                border: none;
            }
        """)
        # 根据内容自动调整高度
        doc = content_edit.document()
        doc.setTextWidth(content_edit.viewport().width())
        height = doc.size().height() + 10
        content_edit.setMaximumHeight(int(min(height, 400)))
        content_edit.setMinimumHeight(int(min(max(height, 60), 400)))
        
        layout.addWidget(content_edit)
        
        # 设置右键菜单
        content_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        content_edit.customContextMenuRequested.connect(self._show_context_menu)
        
        self.content_edit = content_edit
        
    def _show_context_menu(self, position):
        """显示右键菜单"""
        menu = QMenu(self)
        
        # 复制选中内容
        copy_action = QAction("复制选中内容", self)
        copy_action.triggered.connect(self.content_edit.copy)
        menu.addAction(copy_action)
        
        # 复制整条消息
        copy_all_action = QAction("复制整条消息", self)
        copy_all_action.triggered.connect(lambda: self._copy_all())
        menu.addAction(copy_all_action)
        
        menu.addSeparator()
        
        # 添加到解析
        add_action = QAction("添加到题目解析", self)
        add_action.triggered.connect(self._add_to_analysis)
        menu.addAction(add_action)
        
        menu.exec(self.content_edit.mapToGlobal(position))
        
    def _copy_all(self):
        """复制整条消息"""
        text = f"{self.sender}: {self.content}"
        from PySide6.QtWidgets import QApplication
        QApplication.clipboard().setText(text)
        
    def _add_to_analysis(self):
        """添加到题目解析"""
        parent = self.parent()
        while parent and not isinstance(parent, AIChatDialog):
            parent = parent.parent()
        
        if parent and isinstance(parent, AIChatDialog):
            current_text = parent.analysis_edit.toPlainText()
            if current_text:
                current_text += "\n\n"
            current_text += self.content
            parent.analysis_edit.setText(current_text)


class AIChatDialog(QDialog):
    """AI对话窗口"""
    
    analysis_saved = Signal(str)  # 解析保存信号，传递新的解析内容

    def __init__(self, parent=None, question: Question = None,
                 ai_config: AIConfig = None, current_analysis: str = ""):
        super().__init__(parent)
        self.question = question
        self.ai_config = ai_config
        self.current_analysis = current_analysis
        self.chat_thread: Optional[AIChatThread] = None
        self.messages = []  # 对话历史

        self._setup_ui()
        self._init_content()

    def _setup_ui(self):
        """设置UI"""
        ai_name = self.ai_config.name if self.ai_config else "AI"
        self.setWindowTitle(f"与 {ai_name} 对话")
        self.setMinimumSize(1000, 700)
        self.setModal(True)

        layout = QHBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        # 左侧区域：聊天记录 + 输入框
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(10)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # 聊天记录标签
        chat_label = QLabel("<b>对话记录：</b>")
        left_layout.addWidget(chat_label)

        # 聊天记录滚动区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameStyle(QFrame.StyledPanel)
        
        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setSpacing(10)
        self.chat_layout.setContentsMargins(10, 10, 10, 10)
        self.chat_layout.addStretch()
        
        self.scroll_area.setWidget(self.chat_container)
        left_layout.addWidget(self.scroll_area)

        # 输入区域
        input_label = QLabel("<b>输入你的问题：</b>")
        left_layout.addWidget(input_label)

        self.input_edit = QTextEdit()
        self.input_edit.setPlaceholderText("在此输入你的问题...")
        self.input_edit.setMaximumHeight(120)
        left_layout.addWidget(self.input_edit)

        # 发送按钮
        send_layout = QHBoxLayout()
        send_layout.addStretch()
        
        self.send_btn = QPushButton("发送")
        self.send_btn.setFixedWidth(100)
        self.send_btn.clicked.connect(self._send_message)
        send_layout.addWidget(self.send_btn)
        
        left_layout.addLayout(send_layout)

        # 右侧区域：题目解析
        right_widget = QWidget()
        right_widget.setMinimumWidth(350)
        right_widget.setMaximumWidth(450)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(10)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # 题目信息
        question_label = QLabel("<b>当前题目：</b>")
        right_layout.addWidget(question_label)

        self.question_text = QTextEdit()
        self.question_text.setReadOnly(True)
        self.question_text.setMaximumHeight(150)
        right_layout.addWidget(self.question_text)

        # 题目解析编辑区域
        analysis_label = QLabel("<b>题目解析（可编辑）：</b>")
        right_layout.addWidget(analysis_label)

        self.analysis_edit = QTextEdit()
        self.analysis_edit.setPlaceholderText("AI回复可以右键添加到此处...")
        right_layout.addWidget(self.analysis_edit)

        # 右侧按钮区域
        btn_layout = QHBoxLayout()

        clear_btn = QPushButton("清空对话")
        clear_btn.clicked.connect(self._clear_chat)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()

        save_btn = QPushButton("保存解析")
        save_btn.clicked.connect(self._save_analysis)
        btn_layout.addWidget(save_btn)

        right_layout.addLayout(btn_layout)

        # 添加分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([600, 350])
        
        layout.addWidget(splitter)

    def _init_content(self):
        """初始化内容"""
        # 设置题目文本
        if self.question:
            question_text = f"【题目】{self.question.question}\n\n"
            question_text += f"A. {self.question.A}\n"
            question_text += f"B. {self.question.B}\n"
            question_text += f"C. {self.question.C}\n"
            question_text += f"D. {self.question.D}\n\n"
            question_text += f"【答案】{self.question.answer}"
            self.question_text.setText(question_text)
            
            # 准备输入框内容
            input_text = f"题目：{self.question.question}\n"
            input_text += f"选项：A.{self.question.A} B.{self.question.B} C.{self.question.C} D.{self.question.D}\n"
            input_text += f"正确答案：{self.question.answer}"
            if self.current_analysis:
                input_text += f"\n现有解析：{self.current_analysis}"
            input_text += "\n\n我的问题是："
            self.input_edit.setText(input_text)
        else:
            self.question_text.setText("（无题目信息）")

        # 设置当前解析
        self.analysis_edit.setText(self.current_analysis)

        # 初始化系统消息
        if self.question:
            system_msg = "你是一个英语选择题辅导老师。请帮助学生理解题目，但不要直接给出答案，而是引导学生思考。回答时请使用纯文本，不要使用Markdown格式（如**粗体**、*斜体*、代码块等）"
            self.messages.append({"role": "system", "content": system_msg})

            # 添加系统提示消息
            self._add_message_widget("系统", "题目信息已加载。在左侧输入框中输入问题后点击发送，AI将为你解答。")

    def _add_message_widget(self, sender: str, content: str):
        """添加消息组件到聊天记录"""
        msg_widget = MessageWidget(sender, content)
        # 插入到 stretch 之前
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, msg_widget)
        
        # 滚动到底部
        from PySide6.QtCore import QTimer
        QTimer.singleShot(100, self._scroll_to_bottom)
        
    def _scroll_to_bottom(self):
        """滚动到底部"""
        scrollbar = self.scroll_area.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def _send_message(self):
        """发送消息"""
        text = self.input_edit.toPlainText().strip()
        if not text:
            return

        if not self.ai_config or not self.ai_config.api_key:
            QMessageBox.warning(self, "警告", "请先配置AI！")
            return

        # 显示用户消息
        self._add_message_widget("你", text)
        self.input_edit.clear()

        # 添加到消息历史
        self.messages.append({"role": "user", "content": text})

        # 禁用发送按钮
        self.send_btn.setEnabled(False)

        # 启动AI线程
        self.chat_thread = AIChatThread(self.ai_config, self.messages)
        self.chat_thread.response_received.connect(self._on_response)
        self.chat_thread.error_occurred.connect(self._on_error)
        self.chat_thread.finished_signal.connect(self._on_finished)
        self.chat_thread.start()

    def _on_response(self, content: str):
        """收到AI回复"""
        self._add_message_widget("AI", content)
        self.messages.append({"role": "assistant", "content": content})

    def _on_error(self, error: str):
        """处理错误"""
        QMessageBox.critical(self, "错误", error)

    def _on_finished(self):
        """线程完成"""
        self.send_btn.setEnabled(True)

    def _clear_chat(self):
        """清空对话"""
        # 清除所有消息组件
        while self.chat_layout.count() > 1:  # 保留 stretch
            item = self.chat_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # 保留系统消息，清空其他对话
        if self.messages:
            self.messages = [self.messages[0]] if self.messages[0].get("role") == "system" else []

    def _save_analysis(self):
        """保存解析但不关闭窗口"""
        new_analysis = self.analysis_edit.toPlainText()
        if new_analysis != self.current_analysis:
            self.current_analysis = new_analysis
            # 发送信号通知主窗口保存
            self.analysis_saved.emit(new_analysis)
            QMessageBox.information(self, "提示", "题目解析已保存到题库！")
        else:
            QMessageBox.information(self, "提示", "解析内容未变化")

    def get_analysis(self) -> str:
        """获取编辑后的解析文本"""
        return self.analysis_edit.toPlainText()
