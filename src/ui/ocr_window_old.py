"""
OCR录入窗口
使用浏览器自动化 + 豆包AI识别题目
"""

import os
import re
import shutil
import sys
import json
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from PySide6.QtCore import Qt, QThread, Signal, QPoint
from PySide6.QtGui import QPixmap, QImage, QWheelEvent, QMouseEvent, QAction, QPainter
from PySide6.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox, QFileDialog, QProgressDialog,
    QMenu, QAbstractItemView
)
from PySide6.QtUiTools import QUiLoader

# Playwright导入
try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

from src.models import Question
from src.utils import UI_DIR, CLASSIFICATIONS
from .screenshot_tool import take_screenshot


class DoubaoAIThread(QThread):
    """豆包AI识别线程"""
    progress_signal = Signal(int, int)  # 当前进度, 总数
    result_signal = Signal(str, object)  # 图片路径, 识别结果列表(List[Dict])
    error_signal = Signal(str)  # 错误图片路径
    log_signal = Signal(str)  # 日志信息

    def __init__(self, image_tasks: List[Tuple[str, str]], generate_analysis: bool = False):
        """
        Args:
            image_tasks: [(图片路径, 题号范围), ...]
            generate_analysis: 是否生成解析
        """
        super().__init__()
        self.image_tasks = image_tasks
        self.generate_analysis = generate_analysis
        self._is_running = True
        self.browser = None
        self.page = None

    def run(self):
        """执行AI识别"""
        total = len(self.image_tasks)

        try:
            # 启动浏览器（使用本地Edge）
            self.log_signal.emit("正在启动Edge浏览器...")
            with sync_playwright() as p:
                # 使用本地Edge浏览器
                edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                if not os.path.exists(edge_path):
                    # 尝试另一个路径
                    edge_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"

                # 浏览器启动参数，模拟真实用户
                browser_args = [
                    '--disable-blink-features=AutomationControlled',
                    '--disable-web-security',
                    '--disable-features=IsolateOrigins,site-per-process',
                    '--disable-site-isolation-trials',
                    '--disable-dev-shm-usage',
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-gpu',
                    '--disable-webgl',
                    '--disable-software-rasterizer',
                ]

                if os.path.exists(edge_path):
                    self.browser = p.chromium.launch(
                        headless=False,
                        executable_path=edge_path,
                        args=browser_args
                    )
                else:
                    # 如果找不到Edge，使用默认的chromium
                    self.log_signal.emit("未找到Edge，使用默认浏览器...")
                    self.browser = p.chromium.launch(
                        headless=False,
                        args=browser_args
                    )

                # 创建页面并设置用户代理
                context = self.browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.0 Edg/120.0.0.0',
                    locale='zh-CN',
                    timezone_id='Asia/Shanghai',
                )
                self.page = context.new_page()

                # 注入脚本隐藏自动化特征
                self.page.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    });
                    Object.defineProperty(navigator, 'plugins', {
                        get: () => [1, 2, 3, 4, 5]
                    });
                    window.chrome = { runtime: {} };
                """)

                # 打开豆包
                self.log_signal.emit("正在打开豆包...")
                self.page.goto("https://www.doubao.com")
                self.page.wait_for_timeout(3000)  # 等待页面加载

                for i, (image_path, question_range) in enumerate(self.image_tasks):
                    if not self._is_running:
                        break

                    self.log_signal.emit(f"正在处理第 {i+1}/{total} 张图片...")

                    try:
                        questions = self._recognize_image(image_path, question_range)
                        if questions:
                            self.result_signal.emit(image_path, questions)
                        else:
                            self.error_signal.emit(image_path)
                    except Exception as e:
                        self.log_signal.emit(f"识别失败: {e}")
                        self.error_signal.emit(image_path)

                    self.progress_signal.emit(i + 1, total)

                # 关闭浏览器
                if self.browser:
                    self.browser.close()

        except Exception as e:
            self.log_signal.emit(f"浏览器错误: {e}")
            # 标记所有任务失败
            for image_path, _ in self.image_tasks:
                self.error_signal.emit(image_path)

    def _recognize_image(self, image_path: str, question_range: str) -> List[Dict]:
        """识别单张图片"""
        # 构建提示词 - 简化版，只获取题号和选项
        if self.generate_analysis:
            prompt = f"""请识别图片中的英语选择题，图片包含以下题号：{question_range}

请按以下JSON格式返回纯文本，不要包含其他内容，不要使用Markdown格式：
{{
  "questions": [
    {{
      "number": "原始题号",
      "question": "题目内容",
      "optionA": "选项A",
      "optionB": "选项B",
      "optionC": "选项C",
      "optionD": "选项D",
      "analysis": "解析内容"
    }}
  ]
}}

请为每道题提供详细的解析。直接返回JSON文本，不要加```json标记。"""
        else:
            prompt = f"""请识别图片中的英语选择题，图片包含以下题号：{question_range}

请按以下JSON格式返回纯文本，不要包含其他内容，不要使用Markdown格式：
{{
  "questions": [
    {{
      "number": "原始题号",
      "question": "题目内容",
      "optionA": "选项A",
      "optionB": "选项B",
      "optionC": "选项C",
      "optionD": "选项D"
    }}
  ]
}}

只需要识别题号和选项内容即可，不需要答案和解析。直接返回JSON文本，不要加```json标记。"""

        # 上传图片
        self.log_signal.emit(f"  上传图片: {os.path.basename(image_path)}")

        # 找到文件输入框并上传
        file_input = self.page.locator('input[type="file"]').first
        if file_input.count() == 0:
            # 可能需要点击上传按钮
            upload_btn = self.page.locator('button:has-text("上传")').first
            if upload_btn.count() > 0:
                upload_btn.click()
                self.page.wait_for_timeout(1000)
                file_input = self.page.locator('input[type="file"]').first

        if file_input.count() > 0:
            file_input.set_input_files(image_path)
            self.page.wait_for_timeout(3000)  # 等待上传完成

        # 输入提示词
        self.log_signal.emit("  发送提示词...")
        textarea = self.page.locator('textarea').first
        if textarea.count() > 0:
            # 模拟人工输入，逐字输入
            textarea.click()
            self.page.wait_for_timeout(500)
            textarea.fill(prompt)
            self.page.wait_for_timeout(2000)  # 输入后等待
            textarea.press('Enter')
            self.page.wait_for_timeout(1000)  # 发送后等待

        # 等待回复
        self.log_signal.emit("  等待AI回复...")
        self.page.wait_for_timeout(20000)  # 等待20秒

        # 获取回复内容
        self.log_signal.emit("  获取AI回复内容...")

        # 等待AI回复完成（通过检查是否有新的消息出现）
        max_wait = 30  # 最多等待30秒
        wait_count = 0
        last_text = ""

        while wait_count < max_wait:
            self.page.wait_for_timeout(1000)
            # 获取页面可见文本
            current_text = self.page.inner_text('body')

            # 如果文本不再变化，说明回复完成
            if current_text == last_text and len(current_text) > len(prompt):
                self.log_signal.emit(f"  检测到回复完成")
                break

            last_text = current_text
            wait_count += 1
            self.log_signal.emit(f"  等待回复中... {wait_count}s")

        # 尝试获取AI回复的文本
        # 豆包的回复通常在最后一个气泡中
        try:
            # 获取所有消息气泡
            message_bubbles = self.page.locator('div[class*="bubble"], div[class*="message"], .chat-item').all()
            if message_bubbles:
                # 获取最后一个消息
                last_message = message_bubbles[-1].inner_text()
                self.log_signal.emit(f"  获取到消息，长度: {len(last_message)}")
                return self._parse_ai_response(last_message)
        except Exception as e:
            self.log_signal.emit(f"  获取消息失败: {e}")

        # 如果上面的方法失败，尝试直接获取页面文本
        try:
            page_text = self.page.inner_text('body')
            self.log_signal.emit(f"  获取页面文本，长度: {len(page_text)}")
            # 只保留prompt之后的内容
            if prompt[:50] in page_text:
                response_text = page_text.split(prompt[:50])[-1]
                return self._parse_ai_response(response_text)
            else:
                return self._parse_ai_response(page_text)
        except Exception as e:
            self.log_signal.emit(f"  获取页面文本失败: {e}")

        return []

    def _parse_ai_response(self, text: str) -> List[Dict]:
        """解析AI返回的JSON"""
        questions = []

        try:
            # 尝试提取JSON
            json_match = re.search(r'\{[\s\S]*\}', text)
            if json_match:
                json_str = json_match.group()
                self.log_signal.emit(f"  提取到JSON，长度: {len(json_str)}")
                data = json.loads(json_str)
                question_list = data.get('questions', [])
                self.log_signal.emit(f"  解析到{len(question_list)}道题目")
                for q in question_list:
                    questions.append({
                        'question': q.get('question', ''),
                        'A': q.get('optionA', ''),
                        'B': q.get('optionB', ''),
                        'C': q.get('optionC', ''),
                        'D': q.get('optionD', ''),
                        'answer': q.get('answer', ''),
                        'classification': '',
                        'source': '',
                        'analysis': q.get('analysis', '')
                    })
            else:
                self.log_signal.emit("  未在回复中找到JSON格式内容")
        except Exception as e:
            self.log_signal.emit(f"  解析失败: {e}")

        return questions

    def stop(self):
        """停止识别"""
        self._is_running = False
        if self.browser:
            try:
                self.browser.close()
            except:
                pass


class ImageLabel(QLabel):
    """支持缩放和拖拽的图片标签"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("border: 1px solid gray; background-color: #f0f0f0;")

        self._pixmap: Optional[QPixmap] = None
        self._scale_factor = 1.0
        self._min_scale = 0.1
        self._max_scale = 5.0
        self._dragging = False
        self._drag_start: Optional[QPoint] = None
        self._offset = QPoint(0, 0)

        self.setMinimumSize(400, 300)
        self.setText("点击浏览或截图添加图片")

    def setPixmap(self, pixmap: QPixmap):
        """设置图片"""
        self._pixmap = pixmap
        self._scale_factor = 1.0
        self._offset = QPoint(0, 0)
        self._update_display()

    def _update_display(self):
        """更新显示（支持拖拽偏移）"""
        if self._pixmap is None:
            return

        # 创建一个与标签大小相同的空白画布
        canvas = QPixmap(self.size())
        canvas.fill(Qt.transparent)

        # 缩放图片
        scaled_size = self._pixmap.size() * self._scale_factor
        scaled_pixmap = self._pixmap.scaled(
            scaled_size.toSizeF().toSize(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )

        # 计算绘制位置（居中 + 偏移）
        x = (self.width() - scaled_pixmap.width()) // 2 + self._offset.x()
        y = (self.height() - scaled_pixmap.height()) // 2 + self._offset.y()

        # 在画布上绘制图片
        painter = QPainter(canvas)
        painter.drawPixmap(x, y, scaled_pixmap)
        painter.end()

        super().setPixmap(canvas)

    def wheelEvent(self, event: QWheelEvent):
        """鼠标滚轮缩放"""
        if self._pixmap is None:
            return

        # 向上放大，向下缩小
        delta = event.angleDelta().y()
        if delta > 0:
            self._scale_factor *= 1.1
        else:
            self._scale_factor /= 1.1

        # 限制缩放范围
        self._scale_factor = max(self._min_scale, min(self._max_scale, self._scale_factor))

        self._update_display()
        event.accept()

    def mousePressEvent(self, event: QMouseEvent):
        """鼠标按下开始拖拽"""
        if event.button() == Qt.LeftButton and self._pixmap:
            self._dragging = True
            self._drag_start = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()

    def mouseMoveEvent(self, event: QMouseEvent):
        """鼠标移动拖拽"""
        if self._dragging and self._pixmap:
            delta = event.pos() - self._drag_start
            self._offset += delta
            self._drag_start = event.pos()
            self._update_display()
            event.accept()

    def mouseReleaseEvent(self, event: QMouseEvent):
        """鼠标释放结束拖拽"""
        if event.button() == Qt.LeftButton:
            self._dragging = False
            self.setCursor(Qt.ArrowCursor)
            event.accept()

    def resizeEvent(self, event):
        """窗口大小改变时重绘"""
        super().resizeEvent(event)
        if self._pixmap:
            self._update_display()


class OCRWindow(QWidget):
    """OCR录入窗口 - 使用豆包AI识别"""

    def __init__(self):
        super().__init__()

        # 设置窗口属性（正常窗口，不置顶，不模态）
        self.setWindowFlags(Qt.Window)

        # 初始化变量
        self.temp_dir = os.path.join(os.path.dirname(__file__), '..', 'temp', 'ocr')
        self.screenshot_counter = 1
        self.image_paths: List[str] = []  # listWidget中的图片路径
        self.image_to_questions: Dict[str, str] = {}  # 图片路径 -> 题号范围
        self.current_image_path: Optional[str] = None  # 当前选中的图片
        self.table_to_image: Dict[int, str] = {}  # 表格行到图片路径的映射
        self.ai_thread: Optional[DoubaoAIThread] = None

        # 清理临时目录
        self._clear_temp_dir()

        # 加载UI
        self._setup_ui()

    def _center_window(self):
        """窗口居中显示"""
        from PySide6.QtGui import QScreen
        from PySide6.QtWidgets import QApplication

        # 获取屏幕几何信息
        screen = QApplication.primaryScreen().geometry()
        # 获取窗口几何信息
        size = self.geometry()
        # 计算居中位置
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        self.move(x, y)

    def _clear_temp_dir(self):
        """清空临时目录"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir, exist_ok=True)
    
    def _setup_ui(self):
        """设置UI"""
        # 加载UI文件 - 使用绝对路径
        from PySide6.QtUiTools import QUiLoader
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        ui_file = os.path.join(base_dir, 'ui_dir', 'ocr.ui')

        loader = QUiLoader()
        self.ui = loader.load(ui_file, self)
        if not self.ui:
            raise RuntimeError(f"无法加载UI文件: {ui_file}")

        # 设置窗口标题
        self.setWindowTitle(self.ui.windowTitle() or "OCR录入")

        # 设置布局
        layout = QVBoxLayout(self)
        layout.addWidget(self.ui)
        layout.setContentsMargins(0, 0, 0, 0)

        # 窗口居中显示
        self._center_window()
        
        # 替换图片预览标签为自定义的ImageLabel
        self.image_label = ImageLabel()
        self.ui.verticalLayout_2.replaceWidget(self.ui.label_2, self.image_label)
        self.ui.label_2.deleteLater()
        
        # 设置表格为可编辑
        self.ui.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        
        # 连接信号
        self._connect_signals()
        
        # 设置表格右键菜单
        self._setup_table_context_menu()
        
        # 设置listWidget右键菜单
        self._setup_list_widget_context_menu()
    
    def _connect_signals(self):
        """连接信号槽"""
        # 按钮信号 - 使用列表和循环
        button_connections = [
            (self.ui.pushButton, self._on_browse_files),
            (self.ui.pushButton_2, self._on_screenshot),
            (self.ui.pushButton_3, self._on_clear),
            (self.ui.pushButton_4, self._on_start_ocr),
            (self.ui.pushButton_5, self._on_confirm_import),
        ]
        for button, handler in button_connections:
            button.clicked.connect(handler)

        # listWidget信号
        self.ui.listWidget.currentItemChanged.connect(self._on_list_item_changed)

        # textEdit信号 - 记录题号
        self.ui.textEdit.textChanged.connect(self._on_question_range_changed)

        # tableWidget信号
        self.ui.tableWidget.currentCellChanged.connect(self._on_table_cell_changed)
        self.ui.tableWidget.cellDoubleClicked.connect(self._on_table_cell_double_clicked)
    
    def _setup_table_context_menu(self):
        """设置表格右键菜单"""
        self.ui.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.ui.tableWidget.customContextMenuRequested.connect(self._show_table_context_menu)
    
    def _setup_list_widget_context_menu(self):
        """设置listWidget右键菜单"""
        self.ui.listWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.ui.listWidget.customContextMenuRequested.connect(self._show_list_context_menu)
    
    def _show_table_context_menu(self, position):
        """显示表格右键菜单"""
        menu = QMenu()
        delete_action = QAction("删除", self)
        delete_action.triggered.connect(self._on_delete_table_row)
        menu.addAction(delete_action)
        menu.exec(self.ui.tableWidget.viewport().mapToGlobal(position))
    
    def _show_list_context_menu(self, position):
        """显示listWidget右键菜单"""
        item = self.ui.listWidget.itemAt(position)
        if item is None:
            return

        menu = QMenu()

        open_dir_action = QAction("在目录中打开", self)
        open_dir_action.triggered.connect(lambda: self._on_open_image_dir(item))
        menu.addAction(open_dir_action)

        delete_action = QAction("删除", self)
        delete_action.triggered.connect(lambda: self._on_delete_list_item(item))
        menu.addAction(delete_action)

        menu.exec(self.ui.listWidget.viewport().mapToGlobal(position))

    def _on_delete_list_item(self, item: QListWidgetItem):
        """删除listWidget中的项"""
        reply = QMessageBox.question(
            self, "确认删除", f"确定要删除 {item.text()} 吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            image_path = item.data(Qt.UserRole)
            if image_path in self.image_to_questions:
                del self.image_to_questions[image_path]
            if image_path in self.image_paths:
                self.image_paths.remove(image_path)
            self.ui.listWidget.takeItem(self.ui.listWidget.row(item))
    
    def _on_browse_files(self):
        """浏览本地文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择图片文件", "",
            "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        
        for file_path in file_paths:
            # 复制到临时目录
            filename = os.path.basename(file_path)
            dest_path = os.path.join(self.temp_dir, filename)
            shutil.copy2(file_path, dest_path)
            
            # 添加到listWidget
            item = QListWidgetItem(filename)
            item.setData(Qt.UserRole, dest_path)
            self.ui.listWidget.addItem(item)
            self.image_paths.append(dest_path)
    
    def _on_screenshot(self):
        """截取屏幕"""
        def on_screenshot_taken(pixmap: QPixmap):
            """截图完成回调"""
            # 生成文件名
            filename = f"屏幕捕获_{self.screenshot_counter}.png"
            self.screenshot_counter += 1

            # 保存截图
            save_path = os.path.join(self.temp_dir, filename)
            pixmap.save(save_path)

            # 添加到listWidget
            item = QListWidgetItem(filename)
            item.setData(Qt.UserRole, save_path)
            self.ui.listWidget.addItem(item)
            self.image_paths.append(save_path)

            # 显示预览
            self.image_label.setPixmap(pixmap)
            self.ui.listWidget.setCurrentItem(item)

            # 截图完成后显示OCR窗口
            self.showNormal()
            self.raise_()
            self.activateWindow()

        # 隐藏OCR窗口，开始截图
        self.hide()

        # 延迟一点再启动截图工具，确保窗口隐藏完成
        from PySide6.QtCore import QTimer
        def start_screenshot():
            self.screenshot_widget = take_screenshot(on_screenshot_taken)
        QTimer.singleShot(300, start_screenshot)
    
    def _on_clear(self):
        """清空所有数据"""
        reply = QMessageBox.question(
            self, "确认清空", "确定要清空所有图片和识别结果吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        
        # 清空listWidget
        self.ui.listWidget.clear()
        self.image_paths.clear()
        
        # 清空表格
        self.ui.tableWidget.setRowCount(0)
        self.table_to_image.clear()
        
        # 清空预览
        self.image_label.setText("点击浏览或截图添加图片")
        
        # 清空临时目录
        self._clear_temp_dir()
        
        # 重置截图计数器
        self.screenshot_counter = 1
    
    def _on_start_ocr(self):
        """开始AI识别"""
        if not PLAYWRIGHT_AVAILABLE:
            QMessageBox.warning(
                self, "警告",
                "Playwright未安装，无法使用AI识别。\n"
                "请运行: pip install playwright\n"
                "然后运行: python -m playwright install chromium"
            )
            return

        if self.ui.listWidget.count() == 0:
            QMessageBox.information(self, "提示", "请先添加图片")
            return

        # 保存当前图片的题号
        if self.current_image_path:
            self.image_to_questions[self.current_image_path] = self.ui.textEdit.toPlainText().strip()

        # 验证所有图片的题号
        image_tasks = []
        for i in range(self.ui.listWidget.count()):
            item = self.ui.listWidget.item(i)
            path = item.data(Qt.UserRole)
            if path:
                question_range = self.image_to_questions.get(path, "")

                # 验证题号格式
                if not question_range:
                    QMessageBox.warning(self, "题号为空", f"第 {i+1} 张图片未填写题号")
                    self.ui.listWidget.setCurrentItem(item)
                    return

                if not self._validate_question_range(question_range):
                    QMessageBox.warning(
                        self, "题号格式错误",
                        f"第 {i+1} 张图片的题号格式不正确\n"
                        f"仅允许：数字、英文逗号(,)、连字符(-)\n"
                        f"例如：1-9 或 1,3,10 或 1-5,8,10-12"
                    )
                    self.ui.listWidget.setCurrentItem(item)
                    return

                image_tasks.append((path, question_range))

        if not image_tasks:
            return

        # 获取是否生成解析
        generate_analysis = self.ui.checkBox.isChecked()

        # 创建进度对话框
        self.progress_dialog = QProgressDialog("正在启动浏览器...", "取消", 0, len(image_tasks), self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setMinimumDuration(0)

        # 创建并启动AI识别线程
        self.ai_thread = DoubaoAIThread(image_tasks, generate_analysis)
        self.ai_thread.progress_signal.connect(self._on_ai_progress)
        self.ai_thread.result_signal.connect(self._on_ai_result)
        self.ai_thread.error_signal.connect(self._on_ai_error)
        self.ai_thread.log_signal.connect(self._on_ai_log)
        self.ai_thread.finished.connect(self._on_ai_finished)

        self.progress_dialog.canceled.connect(self.ai_thread.stop)

        self.ai_thread.start()

    def _on_ai_progress(self, current: int, total: int):
        """AI识别进度更新"""
        if self.progress_dialog:
            self.progress_dialog.setValue(current)
            percentage = int(current / total * 100)
            self.progress_dialog.setLabelText(f"正在识别... {percentage}%")

    def _on_ai_result(self, image_path: str, questions: List[Dict]):
        """AI识别成功"""
        # 添加到表格
        for question in questions:
            self._add_question_to_table(question, image_path)

    def _on_ai_error(self, image_path: str):
        """AI识别失败"""
        QMessageBox.warning(self, "识别失败", f"图片识别失败: {os.path.basename(image_path)}")

    def _on_ai_log(self, message: str):
        """AI识别日志"""
        print(f"[AI] {message}")

    def _on_ai_finished(self):
        """AI识别完成"""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None

        QMessageBox.information(self, "完成", "AI识别完成！")

    def _add_question_to_table(self, question: Dict, image_path: str):
        """添加题目到表格"""
        row = self.ui.tableWidget.rowCount()
        self.ui.tableWidget.insertRow(row)
        
        # 设置单元格内容
        self.ui.tableWidget.setItem(row, 0, QTableWidgetItem(question.get('question', '')))
        self.ui.tableWidget.setItem(row, 1, QTableWidgetItem(question.get('A', '')))
        self.ui.tableWidget.setItem(row, 2, QTableWidgetItem(question.get('B', '')))
        self.ui.tableWidget.setItem(row, 3, QTableWidgetItem(question.get('C', '')))
        self.ui.tableWidget.setItem(row, 4, QTableWidgetItem(question.get('D', '')))
        self.ui.tableWidget.setItem(row, 5, QTableWidgetItem(question.get('answer', '')))
        self.ui.tableWidget.setItem(row, 6, QTableWidgetItem(question.get('classification', '')))
        self.ui.tableWidget.setItem(row, 7, QTableWidgetItem(question.get('source', '')))
        self.ui.tableWidget.setItem(row, 8, QTableWidgetItem(question.get('analysis', '')))
        
        # 记录映射关系
        self.table_to_image[row] = image_path
    
    def _on_delete_table_row(self):
        """删除表格行"""
        current_row = self.ui.tableWidget.currentRow()
        if current_row >= 0:
            self.ui.tableWidget.removeRow(current_row)
            # 更新映射关系
            if current_row in self.table_to_image:
                del self.table_to_image[current_row]
    
    def _on_list_item_changed(self, current: QListWidgetItem, previous: QListWidgetItem):
        """listWidget项切换 - 保存上一张图片的题号，加载新图片的题号"""
        # 保存上一张图片的题号
        if previous:
            prev_path = previous.data(Qt.UserRole)
            if prev_path:
                self.image_to_questions[prev_path] = self.ui.textEdit.toPlainText().strip()

        # 加载新图片的题号
        if current:
            image_path = current.data(Qt.UserRole)
            self.current_image_path = image_path

            if image_path and os.path.exists(image_path):
                pixmap = QPixmap(image_path)
                self.image_label.setPixmap(pixmap)

            # 加载该图片对应的题号
            question_range = self.image_to_questions.get(image_path, "")
            self.ui.textEdit.setPlainText(question_range)

    def _on_question_range_changed(self):
        """题号范围文本改变时实时保存"""
        if self.current_image_path:
            self.image_to_questions[self.current_image_path] = self.ui.textEdit.toPlainText().strip()

    def _validate_question_range(self, text: str) -> bool:
        """验证题号格式（仅允许数字、英文逗号、连字符）"""
        if not text:
            return False
        # 允许的字符：数字0-9、逗号、连字符、空格
        import re
        pattern = r'^[\d,\-\s]+$'
        return bool(re.match(pattern, text.strip()))

    def _parse_question_range(self, text: str) -> List[int]:
        """解析题号范围，返回题号列表"""
        questions = []
        parts = text.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                # 范围格式：1-5
                try:
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    questions.extend(range(start, end + 1))
                except ValueError:
                    continue
            else:
                # 单个数字
                try:
                    questions.append(int(part))
                except ValueError:
                    continue
        return questions
    
    def _on_table_cell_changed(self, current_row: int, current_column: int, previous_row: int, previous_column: int):
        """表格单元格切换"""
        if current_row >= 0 and current_row in self.table_to_image:
            image_path = self.table_to_image[current_row]
            # 在listWidget中选中对应的项
            for i in range(self.ui.listWidget.count()):
                item = self.ui.listWidget.item(i)
                if item.data(Qt.UserRole) == image_path:
                    self.ui.listWidget.setCurrentItem(item)
                    break
    
    def _on_table_cell_double_clicked(self, row: int, column: int):
        """双击表格单元格，进入编辑模式"""
        # 设置该单元格为可编辑
        item = self.ui.tableWidget.item(row, column)
        if item:
            item.setFlags(item.flags() | Qt.ItemIsEditable)
            self.ui.tableWidget.editItem(item)
    
    def _on_open_image_dir(self, item: QListWidgetItem):
        """在目录中打开图片"""
        image_path = item.data(Qt.UserRole)
        if image_path:
            import subprocess
            subprocess.run(['explorer', '/select,', os.path.normpath(image_path)])
    
    def _on_confirm_import(self):
        """确认导入"""
        # 验证数据
        invalid_rows = []
        for row in range(self.ui.tableWidget.rowCount()):
            # 检查选项
            has_options = all(
                self.ui.tableWidget.item(row, i) and 
                self.ui.tableWidget.item(row, i).text().strip()
                for i in range(1, 5)  # A, B, C, D
            )
            
            # 检查分类
            classification_item = self.ui.tableWidget.item(row, 6)
            classification = classification_item.text().strip() if classification_item else ""
            valid_classification = classification in CLASSIFICATIONS
            
            # 检查来源
            source_item = self.ui.tableWidget.item(row, 7)
            source = source_item.text().strip() if source_item else ""
            
            if not has_options or not valid_classification:
                invalid_rows.append(row + 1)
        
        if invalid_rows:
            QMessageBox.warning(
                self, "数据不完整",
                f"以下行数据不完整，请检查：\n行号: {', '.join(map(str, invalid_rows))}"
            )
            return
        
        # 检查来源为空的情况
        empty_source_rows = []
        for row in range(self.ui.tableWidget.rowCount()):
            source_item = self.ui.tableWidget.item(row, 7)
            if not source_item or not source_item.text().strip():
                empty_source_rows.append(row + 1)
        
        if empty_source_rows:
            reply = QMessageBox.question(
                self, "来源为空",
                f"以下行题目来源为空，将默认填充为\"无\"：\n行号: {', '.join(map(str, empty_source_rows))}\n\n是否继续？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
            
            # 填充"无"
            for row in empty_source_rows:
                self.ui.tableWidget.setItem(row - 1, 7, QTableWidgetItem("无"))
        
        # 收集数据
        questions_data = []
        for row in range(self.ui.tableWidget.rowCount()):
            question_data = {
                'question': self.ui.tableWidget.item(row, 0).text() if self.ui.tableWidget.item(row, 0) else '',
                'A': self.ui.tableWidget.item(row, 1).text() if self.ui.tableWidget.item(row, 1) else '',
                'B': self.ui.tableWidget.item(row, 2).text() if self.ui.tableWidget.item(row, 2) else '',
                'C': self.ui.tableWidget.item(row, 3).text() if self.ui.tableWidget.item(row, 3) else '',
                'D': self.ui.tableWidget.item(row, 4).text() if self.ui.tableWidget.item(row, 4) else '',
                'answer': self.ui.tableWidget.item(row, 5).text() if self.ui.tableWidget.item(row, 5) else '',
                'classification': self.ui.tableWidget.item(row, 6).text() if self.ui.tableWidget.item(row, 6) else '',
                'source': self.ui.tableWidget.item(row, 7).text() if self.ui.tableWidget.item(row, 7) else '',
                'analysis': self.ui.tableWidget.item(row, 8).text() if self.ui.tableWidget.item(row, 8) else '',
            }
            questions_data.append(question_data)
        
        # 先在控制台输出数据供检查
        print("=" * 80)
        print("OCR识别结果（供检查）：")
        print("=" * 80)
        for i, q in enumerate(questions_data, 1):
            print(f"\n题目 {i}:")
            for key, value in q.items():
                print(f"  {key}: {value}")
        print("=" * 80)
        
        # 选择导入方式
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("选择导入方式")
        msg_box.setText(f"共 {len(questions_data)} 道题目，数据已在控制台输出，请选择导入方式：")
        
        import_to_input_btn = msg_box.addButton("导入到录入页面", QMessageBox.ActionRole)
        # save_directly_btn = msg_box.addButton("直接保存到题库", QMessageBox.ActionRole)
        cancel_btn = msg_box.addButton("取消", QMessageBox.RejectRole)
        
        msg_box.exec()
        
        clicked_btn = msg_box.clickedButton()
        
        if clicked_btn == import_to_input_btn:
            # 导入到录入页面
            self._import_to_main(questions_data)
        # elif clicked_btn == save_directly_btn:
        #     # 直接保存到题库
        #     self._save_to_database(questions_data)
        # 取消则不执行任何操作
    
    def _import_to_main(self, questions_data: List[Dict]):
        """将识别的题目导入到主程序录入页面"""
        if not questions_data:
            return
        
        # 获取父窗口（主窗口）
        parent = self.parent()
        if not parent:
            QMessageBox.warning(self, "错误", "无法获取主窗口")
            return
        
        # 获取主窗口控制器
        main_window = None
        if hasattr(parent, 'main_window'):
            main_window = parent
        
        if not main_window:
            QMessageBox.warning(self, "错误", "无法获取主窗口控制器")
            return
        
        # 导入第一道题到录入页面
        first_question = questions_data[0]
        main_window.main_window.textEdit.setPlainText(first_question.get('question', ''))
        main_window.main_window.textEdit_2.setPlainText(first_question.get('A', ''))
        main_window.main_window.textEdit_3.setPlainText(first_question.get('B', ''))
        main_window.main_window.textEdit_4.setPlainText(first_question.get('C', ''))
        main_window.main_window.textEdit_5.setPlainText(first_question.get('D', ''))
        main_window.main_window.textEdit_6.setPlainText(first_question.get('source', ''))
        main_window.main_window.textEdit_7.setPlainText(first_question.get('analysis', ''))
        
        # 设置答案
        answer = first_question.get('answer', '')
        if answer in ['A', 'B', 'C', 'D']:
            answer_index = ['A', 'B', 'C', 'D'].index(answer)
            # 设置答案按钮组
            if hasattr(main_window, 'btn_group_answer'):
                buttons = main_window.btn_group_answer.buttons()
                if answer_index < len(buttons):
                    buttons[answer_index].setChecked(True)
        
        # 设置分类（分类是索引值）
        classification = first_question.get('classification', '')
        try:
            class_index = int(classification)
            if 0 <= class_index < len(CLASSIFICATIONS):
                if hasattr(main_window, 'btn_group_classification'):
                    buttons = main_window.btn_group_classification.buttons()
                    if class_index < len(buttons):
                        buttons[class_index].setChecked(True)
        except (ValueError, TypeError):
            # 如果分类不是数字，尝试从字符串匹配
            if classification in CLASSIFICATIONS:
                class_index = CLASSIFICATIONS.index(classification)
                if hasattr(main_window, 'btn_group_classification'):
                    buttons = main_window.btn_group_classification.buttons()
                    if class_index < len(buttons):
                        buttons[class_index].setChecked(True)
        
        # 如果有更多题目，提示用户
        if len(questions_data) > 1:
            reply = QMessageBox.question(
                self, "多道题目",
                f"共识别出 {len(questions_data)} 道题目。\n"
                f"第一道题已填入录入页面。\n"
                f"是否继续导入下一道？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                # 保存剩余题目到临时列表
                self._remaining_questions = questions_data[1:]
                # 设置标志表示正在批量导入
                self._batch_importing = True
                self._main_window_ref = main_window
                # 切换到录入页面
                main_window._goto_page(1)  # PAGE_CREATE = 1
            else:
                # 用户选择不继续，切换到录入页面让用户手动保存
                main_window._goto_page(1)  # PAGE_CREATE = 1
        else:
            # 只有一道题，切换到录入页面
            main_window._goto_page(1)  # PAGE_CREATE = 1
            QMessageBox.information(
                self, "导入成功",
                f"已导入题目到录入页面，请检查并保存。"
            )
    
    def check_and_import_next(self):
        """检查并导入下一道题目（由主窗口调用）"""
        if not hasattr(self, '_batch_importing') or not self._batch_importing:
            return False
        
        if not hasattr(self, '_remaining_questions') or not self._remaining_questions:
            # 所有题目已导入完成
            self._batch_importing = False
            self._main_window_ref = None
            QMessageBox.information(self, "完成", "所有题目已导入完成！")
            return False
        
        # 导入下一道
        next_question = self._remaining_questions.pop(0)
        main_window = self._main_window_ref
        
        main_window.main_window.textEdit.setPlainText(next_question.get('question', ''))
        main_window.main_window.textEdit_2.setPlainText(next_question.get('A', ''))
        main_window.main_window.textEdit_3.setPlainText(next_question.get('B', ''))
        main_window.main_window.textEdit_4.setPlainText(next_question.get('C', ''))
        main_window.main_window.textEdit_5.setPlainText(next_question.get('D', ''))
        main_window.main_window.textEdit_6.setPlainText(next_question.get('source', ''))
        main_window.main_window.textEdit_7.setPlainText(next_question.get('analysis', ''))
        
        # 设置答案
        answer = next_question.get('answer', '')
        if answer in ['A', 'B', 'C', 'D']:
            answer_index = ['A', 'B', 'C', 'D'].index(answer)
            if hasattr(main_window, 'btn_group_answer'):
                buttons = main_window.btn_group_answer.buttons()
                if answer_index < len(buttons):
                    buttons[answer_index].setChecked(True)
        
        # 设置分类（分类是索引值）
        classification = next_question.get('classification', '')
        try:
            class_index = int(classification)
            if 0 <= class_index < len(CLASSIFICATIONS):
                if hasattr(main_window, 'btn_group_classification'):
                    buttons = main_window.btn_group_classification.buttons()
                    if class_index < len(buttons):
                        buttons[class_index].setChecked(True)
        except (ValueError, TypeError):
            # 如果分类不是数字，尝试从字符串匹配
            if classification in CLASSIFICATIONS:
                class_index = CLASSIFICATIONS.index(classification)
                if hasattr(main_window, 'btn_group_classification'):
                    buttons = main_window.btn_group_classification.buttons()
                    if class_index < len(buttons):
                        buttons[class_index].setChecked(True)
        
        QMessageBox.information(
            self, "下一道题",
            f"已导入下一道题，还有 {len(self._remaining_questions)} 道待导入。"
        )
        return True
    
    def _save_to_database(self, questions_data: List[Dict]):
        """直接保存到题库"""
        # 获取父窗口（主窗口）
        parent = self.parent()
        if not parent:
            QMessageBox.warning(self, "错误", "无法获取主窗口")
            return
        
        # 获取主窗口控制器
        main_window = None
        if hasattr(parent, 'main_window'):
            main_window = parent
        
        if not main_window:
            QMessageBox.warning(self, "错误", "无法获取主窗口控制器")
            return
        
        # 转换数据格式
        saved_count = 0
        for question_data in questions_data:
            try:
                # 处理分类（如果是索引值）
                classification = question_data.get('classification', '')
                try:
                    class_index = int(classification)
                    if 0 <= class_index < len(CLASSIFICATIONS):
                        classification = CLASSIFICATIONS[class_index]
                except (ValueError, TypeError):
                    pass  # 保持原值
                
                # 创建Question对象
                question = Question(
                    question=question_data.get('question', ''),
                    A=question_data.get('A', ''),
                    B=question_data.get('B', ''),
                    C=question_data.get('C', ''),
                    D=question_data.get('D', ''),
                    answer=question_data.get('answer', ''),
                    classification=classification,
                    source=question_data.get('source', '无'),
                    analysis=question_data.get('analysis', ''),
                    total=0,
                    correct=0
                )
                
                # 添加到题库
                main_window.data_manager.add_question(question)
                saved_count += 1
                
            except Exception as e:
                print(f"保存题目失败: {e}")
                continue
        
        # 保存到文件
        main_window.data_manager.save_questions()
        
        QMessageBox.information(
            self, "保存成功",
            f"成功保存 {saved_count}/{len(questions_data)} 道题目到题库！"
        )
        
        # 清空表格
        self.ui.tableWidget.setRowCount(0)
        self.table_to_image.clear()
    
    def keyPressEvent(self, event):
        """按键事件"""
        # 在表格中按Enter键移动到下一行
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            if self.ui.tableWidget.hasFocus():
                current_row = self.ui.tableWidget.currentRow()
                current_col = self.ui.tableWidget.currentColumn()
                next_row = current_row + 1
                
                if next_row < self.ui.tableWidget.rowCount():
                    self.ui.tableWidget.setCurrentCell(next_row, current_col)
                    self.ui.tableWidget.editItem(self.ui.tableWidget.item(next_row, current_col))
                return
        
        super().keyPressEvent(event)
    
    def closeEvent(self, event):
        """关闭事件"""
        # 检查是否有未导入的数据
        if self.ui.tableWidget.rowCount() > 0:
            reply = QMessageBox.question(
                self, "确认关闭",
                f"表格中还有 {self.ui.tableWidget.rowCount()} 道题目未导入，确定要关闭窗口吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
        
        # 停止AI识别线程
        if self.ai_thread and self.ai_thread.isRunning():
            self.ai_thread.stop()
            self.ai_thread.wait()
        
        event.accept()


def show_ocr_window(parent=None):
    """显示OCR窗口"""
    window = OCRWindow()
    window.show()
    # 窗口显示后再居中
    window._center_window()
    return window
