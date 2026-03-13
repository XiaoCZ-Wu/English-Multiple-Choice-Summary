"""
OCR录入窗口
使用智谱AI服务识别题目
"""

import os
import re
import shutil
import json
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from PySide6.QtCore import Qt, QThread, Signal, QPoint
from PySide6.QtGui import QPixmap, QImage, QWheelEvent, QMouseEvent, QAction, QPainter, QPaintEvent
from PySide6.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QPushButton, QMessageBox, QFileDialog, QProgressDialog,
    QMenu, QAbstractItemView
)

from src.utils.constants import CLASSIFICATIONS, OCR_TEMP_DIR
from src.utils import app_logger
from PySide6.QtUiTools import QUiLoader

from src.models import Question
from src.utils import UI_DIR, CLASSIFICATIONS
from .screenshot_tool import take_screenshot
from .dialogs import SourceDialog

# 配置日志
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)


class ZhipuAIThread(QThread):
    """使用AI的识别线程（支持自定义AI配置）"""
    progress_signal = Signal(int, int)  # 当前进度, 总数
    result_signal = Signal(str, object)  # 图片路径, 识别结果列表(List[Dict])
    error_signal = Signal(str, str)  # 错误图片路径, 错误信息
    log_signal = Signal(str)  # 日志信息

    def __init__(self, image_tasks: List[Tuple[str, str]], generate_analysis: bool = False, ai_config: Optional[Dict] = None):
        """
        Args:
            image_tasks: [(图片路径, 题号范围), ...]
            generate_analysis: 是否生成解析
            ai_config: AI配置字典，包含 name, base_url, api_key
        """
        super().__init__()
        self.image_tasks = image_tasks
        self.generate_analysis = generate_analysis
        self.ai_config = ai_config or {}
        self._is_running = True

    def _parse_question_range(self, range_str: str) -> List[int]:
        """解析题号范围字符串，返回具体的题号列表
        
        支持格式：
        - 单个数字：1
        - 范围：1-5
        - 多个范围：1-5,7,9-11
        """
        result = []
        if not range_str or not range_str.strip():
            return result
        
        # 移除所有空格
        range_str = range_str.replace(' ', '').replace('，', ',')
        
        parts = range_str.split(',')
        for part in parts:
            if not part:
                continue
            if '-' in part:
                # 处理范围，如 9-11
                try:
                    start, end = part.split('-', 1)
                    start_num = int(start)
                    end_num = int(end)
                    result.extend(range(start_num, end_num + 1))
                except ValueError:
                    continue
            else:
                # 处理单个数字
                try:
                    result.append(int(part))
                except ValueError:
                    continue
        
        return sorted(list(set(result)))  # 去重并排序

    def run(self):
        """执行AI识别"""
        import requests
        import base64

        total = len(self.image_tasks)

        # 获取 API Key、Base URL 和模型名称从 AI 配置
        api_key = self.ai_config.get('api_key', '')
        base_url = self.ai_config.get('base_url', '')
        ai_name = self.ai_config.get('name', 'AI')
        model_name = self.ai_config.get('model', '')
        
        # 检查必要配置
        if not model_name:
            error_msg = f"错误：AI模型 \"{ai_name}\" 未配置模型名称"
            app_logger.error(f"[OCR错误] {error_msg}")
            self.log_signal.emit(error_msg)
            for image_path, _ in self.image_tasks:
                self.error_signal.emit(image_path, error_msg)
            return

        # 检查是否为视觉模型（OCR需要多模态能力）
        vision_keywords = ['vision', 'v-', '4v', 'gpt-4o', 'glm-4v', 'qwen-vl', 'claude-3']
        model_lower = model_name.lower()
        if not any(keyword in model_lower for keyword in vision_keywords):
            error_msg = f"错误：OCR识别需要使用支持视觉的多模态模型\n当前模型 \"{model_name}\" 可能不支持图像识别\n请使用如 glm-4v-flash、gpt-4o 等视觉模型\n\n提示：图片过大也可能导致API错误，建议图片大小不超过5MB"
            app_logger.error(f"[OCR错误] {error_msg}")
            self.log_signal.emit(error_msg)
            for image_path, _ in self.image_tasks:
                self.error_signal.emit(image_path, error_msg)
            return

        if not base_url:
            error_msg = f"错误：AI模型 \"{ai_name}\" 未配置 Base URL"
            app_logger.error(f"[OCR错误] {error_msg}")
            self.log_signal.emit(error_msg)
            for image_path, _ in self.image_tasks:
                self.error_signal.emit(image_path, error_msg)
            return

        if not api_key:
            error_msg = f"错误：AI模型 \"{ai_name}\" 未设置 API Key"
            app_logger.error(f"[OCR错误] {error_msg}")
            self.log_signal.emit(error_msg)
            for image_path, _ in self.image_tasks:
                self.error_signal.emit(image_path, error_msg)
            return

        # 确保 base_url 以 / 结尾
        if not base_url.endswith('/'):
            base_url += '/'

        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }

        self.log_signal.emit(f"使用 AI 模型: {ai_name}")
        self.log_signal.emit(f"Base URL: {base_url}")
        self.log_signal.emit(f"模型名称: {model_name}")

        for i, (image_path, question_range) in enumerate(self.image_tasks):
            if not self._is_running:
                break

            self.log_signal.emit(f"正在处理第 {i+1}/{total} 张图片...")

            try:
                # 读取图片并转为 base64
                with open(image_path, 'rb') as f:
                    image_base64 = base64.b64encode(f.read()).decode('utf-8')

                # 解析题号范围，生成具体的题号列表
                question_numbers = self._parse_question_range(question_range)
                if question_numbers:
                    question_list_str = ', '.join(map(str, question_numbers))
                    question_range_desc = f"图片中只包含以下题号的题目：{question_list_str}"
                else:
                    question_list_str = "全部"
                    question_range_desc = "请识别图片中的所有题目"

                # 构建分类选项字符串
                classifications_str = '、'.join(CLASSIFICATIONS)

                # 构建提示词
                if self.generate_analysis:
                    prompt = f"""请识别图片中的英语选择题。

【重要说明】
- {question_range_desc}
- 【关键】必须识别图片中所有可见的题目，一道都不能遗漏
- 严格按照图片中实际显示的题号来填写
- 【关键】题目内容必须包含原始题号，格式如："1. What is..." 或 "(1) What is..."
- 请根据题目内容判断正确答案（A/B/C/D）
- 请从以下分类中选择最符合的一个：{classifications_str}
- 【关键】题目解析请使用中文回答

【输出格式要求】
请使用HTML标签格式输出，每道题用<question>标签包裹，每个字段用对应的标签：

```
<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
<ana>解析内容（中文）</ana>
</question>

<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
<ana>解析内容（中文）</ana>
</question>
```

如果有多个题目，就重复<question>标签块。确保所有内容都在代码块内。"""
                else:
                    prompt = f"""请识别图片中的英语选择题。

【重要说明】
- {question_range_desc}
- 【关键】必须识别图片中所有可见的题目，一道都不能遗漏
- 严格按照图片中实际显示的题号来填写
- 【关键】题目内容必须包含原始题号，格式如："1. What is..." 或 "(1) What is..."
- 请根据题目内容判断正确答案（A/B/C/D）
- 请从以下分类中选择最符合的一个：{classifications_str}

【输出格式要求】
请使用HTML标签格式输出，每道题用<question>标签包裹，每个字段用对应的标签：

```
<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
</question>

<question>
<q>题号. 题目内容</q>
<a>选项A</a>
<b>选项B</b>
<c>选项C</c>
<d>选项D</d>
<ans>正确选项(A/B/C/D)</ans>
<cat>分类</cat>
</question>
```

如果有多个题目，就重复<question>标签块。确保所有内容都在代码块内。"""

                # 输出提示词到日志
                app_logger.info(f"{'='*60}")
                app_logger.info(f"[提示词] 图片: {os.path.basename(image_path)}")
                app_logger.info(f"[提示词] 题号范围: {question_range}")
                app_logger.info(f"[提示词] 内容:\n{prompt}")
                app_logger.info(f"{'='*60}")

                # 构建请求体（智谱AI要求：先文本后图片）
                payload = {
                    "model": model_name,
                    "messages": [
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": prompt
                                },
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:image/jpeg;base64,{image_base64}"
                                    }
                                }
                            ]
                        }
                    ]
                }

                # 发送请求 - 使用配置的base_url
                api_endpoint = f"{base_url}chat/completions"
                app_logger.info(f"[OCR] API端点: {api_endpoint}")
                # 输出payload但不包含图片base64数据（太长）
                payload_copy = json.loads(json.dumps(payload))
                payload_copy['messages'][0]['content'][1]['image_url']['url'] = '[图片base64数据已省略]'
                app_logger.info(f"[OCR] 请求体: {json.dumps(payload_copy, ensure_ascii=False, indent=2)}")
                response = requests.post(
                    api_endpoint,
                    headers=headers,
                    json=payload,
                    timeout=60
                )

                if response.status_code == 200:
                    result = response.json()
                    content = result['choices'][0]['message']['content']

                    # 输出AI回复
                    app_logger.info(f"{'='*60}")
                    app_logger.info(f"[AI回复] 图片: {os.path.basename(image_path)}")
                    app_logger.info(f"[AI回复] 内容:\n{content}")
                    app_logger.info(f"{'='*60}")

                    questions = self._parse_result(content)

                    if questions:
                        self.result_signal.emit(image_path, questions)
                    else:
                        error_msg = "未能从AI回复中解析出题目，请检查图片内容或AI模型是否正常"
                        self.log_signal.emit(error_msg)
                        self.error_signal.emit(image_path, error_msg)
                else:
                    app_logger.error(f"[OCR错误] API请求失败: {response.status_code}")
                    app_logger.error(f"[OCR错误] 响应内容: {response.text}")
                    error_msg = f"API请求失败: {response.status_code} - {response.text}"
                    self.log_signal.emit(error_msg)
                    self.error_signal.emit(image_path, error_msg)

            except Exception as e:
                error_msg = f"识别失败: {str(e)}"
                app_logger.error(f"[OCR错误] {error_msg}")
                import traceback
                app_logger.error(traceback.format_exc())
                self.log_signal.emit(error_msg)
                self.error_signal.emit(image_path, error_msg)

            self.progress_signal.emit(i + 1, total)

    def _parse_result(self, result) -> List[Dict]:
        """解析AI返回的结果（HTML格式）"""
        questions = []

        try:
            # 从结果中提取文本
            if hasattr(result, 'content'):
                text = result.content
            elif hasattr(result, 'text'):
                text = result.text
            else:
                text = str(result)

            logger.info(f"[解析] 原始文本长度: {len(text)}")
            logger.info(f"[解析] 原始文本内容:\n{text}")
            logger.info("=" * 80)

            # 首先尝试从代码块中提取内容
            code_pattern = r'```[\s\S]*?\n([\s\S]*?)```'
            code_matches = re.findall(code_pattern, text, re.DOTALL)

            if code_matches:
                logger.info(f"[解析] 找到 {len(code_matches)} 个代码块")
                for i, match in enumerate(code_matches):
                    logger.info(f"[解析] 代码块 {i+1} 内容:\n{match}")
                # 合并所有代码块的内容
                html_content = '\n'.join(code_matches)
            else:
                logger.info("[解析] 未找到代码块，使用完整文本")
                html_content = text

            logger.info(f"[解析] 用于解析的HTML内容:\n{html_content}")
            logger.info("=" * 80)

            # 使用HTML解析器解析内容
            from html.parser import HTMLParser

            class QuestionParser(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.questions = []
                    self.current_question = {}
                    self.current_tag = None
                    self.current_data = []
                    self.tag_stack = []

                def handle_starttag(self, tag, attrs):
                    logger.debug(f"[HTML解析] 开始标签: <{tag}>")
                    self.current_tag = tag
                    self.tag_stack.append(tag)
                    self.current_data = []

                def handle_endtag(self, tag):
                    logger.debug(f"[HTML解析] 结束标签: </{tag}>")
                    content = ''.join(self.current_data).strip()
                    logger.info(f"[HTML解析] 标签 <{tag}> 内容: '{content}'")

                    if tag == 'question':
                        # 一道题结束，保存当前题目
                        if self.current_question:
                            logger.info(f"[HTML解析] 完成一道题解析: {self.current_question}")
                            self.questions.append(self.current_question)
                            self.current_question = {}
                        else:
                            logger.warning("[HTML解析] 空题目标签，无内容")
                    elif tag in ('q', 'a', 'b', 'c', 'd', 'ans', 'cat', 'ana'):
                        # 字段结束，保存字段内容
                        if not content:
                            logger.warning(f"[HTML解析] 标签 <{tag}> 内容为空")
                        if tag == 'q':
                            self.current_question['question'] = content
                        elif tag == 'a':
                            self.current_question['A'] = content
                        elif tag == 'b':
                            self.current_question['B'] = content
                        elif tag == 'c':
                            self.current_question['C'] = content
                        elif tag == 'd':
                            self.current_question['D'] = content
                        elif tag == 'ans':
                            self.current_question['answer'] = content.upper()
                        elif tag == 'cat':
                            self.current_question['classification'] = content
                        elif tag == 'ana':
                            self.current_question['analysis'] = content
                    self.current_tag = None
                    self.current_data = []
                    if self.tag_stack and self.tag_stack[-1] == tag:
                        self.tag_stack.pop()

                def handle_data(self, data):
                    if self.current_tag:
                        self.current_data.append(data)

            parser = QuestionParser()
            parser.feed(html_content)

            logger.info(f"[解析] HTML解析器找到 {len(parser.questions)} 道题目")

            # 处理解析结果
            for idx, q in enumerate(parser.questions):
                try:
                    logger.info(f"[解析] 处理第 {idx+1} 道题目原始数据: {q}")

                    # 验证分类
                    classification = q.get('classification', '')
                    if classification not in CLASSIFICATIONS:
                        logger.warning(f"[解析] 题目 {idx+1} 分类 '{classification}' 不在有效分类列表中，置为空")
                        classification = ''

                    question_data = {
                        'question': q.get('question', ''),
                        'A': q.get('A', ''),
                        'B': q.get('B', ''),
                        'C': q.get('C', ''),
                        'D': q.get('D', ''),
                        'answer': q.get('answer', ''),
                        'classification': classification,
                        'source': '',
                        'analysis': q.get('analysis', '')
                    }

                    # 从题目中提取题号
                    question_text = question_data['question']
                    number_match = re.match(r'^(\d+)[\.\)\s]', question_text)
                    question_number = number_match.group(1) if number_match else ''

                    logger.info(f"[解析] 题目 {idx+1} 最终数据:")
                    logger.info(f"  题号: {question_number}")
                    logger.info(f"  题目: '{question_text}'")
                    logger.info(f"  A: '{question_data['A']}'")
                    logger.info(f"  B: '{question_data['B']}'")
                    logger.info(f"  C: '{question_data['C']}'")
                    logger.info(f"  D: '{question_data['D']}'")
                    logger.info(f"  答案: '{question_data['answer']}'")
                    logger.info(f"  分类: '{question_data['classification']}'")
                    logger.info(f"  解析: '{question_data['analysis'][:50]}...'" if question_data['analysis'] else "  解析: (空)")

                    if question_data['question']:
                        questions.append(question_data)
                        logger.info(f"[解析] ✓ 成功添加题目 {idx+1} 到列表")
                    else:
                        logger.error(f"[解析] ✗ 题目 {idx+1} 缺少题目内容，跳过")

                except Exception as e:
                    logger.error(f"[解析] ✗ 处理题目 {idx+1} 时发生异常: {e}")
                    import traceback
                    logger.error(f"[解析] 异常详情: {traceback.format_exc()}")
                    continue

            logger.info(f"[解析] 总共成功解析并添加 {len(questions)} 道题目")

            if not questions:
                logger.error("[解析] ✗ 未解析到任何有效题目")
                logger.error(f"[解析] 原始文本:\n{text}")

        except Exception as e:
            logger.error(f"[解析] ✗ 整体解析失败: {e}")
            import traceback
            logger.error(f"[解析] 错误详情: {traceback.format_exc()}")

        return questions

    def stop(self):
        """停止识别"""
        self._is_running = False


class ImageLabel(QLabel):
    """支持缩放和拖拽的图片标签"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pixmap: Optional[QPixmap] = None
        self._scale = 1.0
        self._offset = QPoint(0, 0)
        self._dragging = False
        self._last_pos = QPoint()
        self.setMinimumSize(400, 300)
        self.setStyleSheet("background-color: #f0f0f0; border: 1px solid #ccc;")

    def setPixmap(self, pixmap: QPixmap):
        """设置图片"""
        self._pixmap = pixmap
        self._scale = 1.0
        self._offset = QPoint(0, 0)
        self.update()

    def wheelEvent(self, event: QWheelEvent):
        """鼠标滚轮缩放"""
        if not self._pixmap:
            return

        delta = event.angleDelta().y()
        if delta > 0:
            self._scale = min(self._scale * 1.1, 5.0)  # 最大放大5倍
        else:
            self._scale = max(self._scale / 1.1, 0.1)  # 最小缩小到0.1倍

        self.update()
        event.accept()

    def mousePressEvent(self, event: QMouseEvent):
        """鼠标按下开始拖拽"""
        if event.button() == Qt.LeftButton and self._pixmap:
            self._dragging = True
            self._last_pos = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()

    def mouseMoveEvent(self, event: QMouseEvent):
        """鼠标移动拖拽"""
        if self._dragging and self._pixmap:
            delta = event.pos() - self._last_pos
            self._offset += delta
            self._last_pos = event.pos()
            self.update()
            event.accept()

    def mouseReleaseEvent(self, event: QMouseEvent):
        """鼠标释放结束拖拽"""
        if event.button() == Qt.LeftButton:
            self._dragging = False
            self.setCursor(Qt.ArrowCursor)
            event.accept()

    def paintEvent(self, event: QPaintEvent):
        """重写绘制事件，支持拖拽和缩放"""
        super().paintEvent(event)
        if not self._pixmap:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        # 计算缩放后的图片大小
        scaled_width = int(self._pixmap.width() * self._scale)
        scaled_height = int(self._pixmap.height() * self._scale)

        # 计算绘制位置（考虑偏移量）
        x = (self.width() - scaled_width) // 2 + self._offset.x()
        y = (self.height() - scaled_height) // 2 + self._offset.y()

        # 绘制图片
        painter.drawPixmap(x, y, scaled_width, scaled_height, self._pixmap)

    def resizeEvent(self, event):
        """窗口大小改变时重绘"""
        super().resizeEvent(event)
        if self._pixmap:
            self.update()


class OCRWindow(QWidget):
    """OCR录入窗口 - 使用 browser-use + AI识别"""

    # 导入信号 - 当用户确认导入时触发，参数为题目列表
    import_signal = Signal(list)  # List[Dict] 题目列表

    def __init__(self, ai_config: Optional[Dict] = None):
        """
        Args:
            ai_config: AI配置字典，包含 name, base_url, api_key
        """
        super().__init__()

        # 设置窗口属性（正常窗口，不置顶，不模态）
        self.setWindowFlags(Qt.Window)

        # 初始化变量
        self.temp_dir = OCR_TEMP_DIR
        self.screenshot_counter = 1
        self.image_paths: List[str] = []  # listWidget中的图片路径
        self.image_to_questions: Dict[str, str] = {}  # 图片路径 -> 题号范围
        self.current_image_path: Optional[str] = None  # 当前选中的图片
        self.table_to_image: Dict[int, str] = {}  # 表格行到图片路径的映射
        self.ai_thread: Optional[ZhipuAIThread] = None
        self.ai_config = ai_config  # AI配置

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
        """清空临时目录（保留网页端上传的文件）"""
        if os.path.exists(self.temp_dir):
            for filename in os.listdir(self.temp_dir):
                # 跳过网页端上传的文件（以 web_ 开头）
                if filename.startswith('web_'):
                    continue
                file_path = os.path.join(self.temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    app_logger.warning(f"删除文件失败 {file_path}: {e}")
        os.makedirs(self.temp_dir, exist_ok=True)

    def _setup_ui(self):
        """设置UI"""
        # 加载UI文件 - 使用 constants 中的路径
        from PySide6.QtUiTools import QUiLoader
        from src.utils.constants import get_resource_path
        ui_file = get_resource_path('src/ui_dir/ocr.ui')

        loader = QUiLoader()
        self.ui = loader.load(ui_file, self)
        if not self.ui:
            raise RuntimeError(f"无法加载UI文件: {ui_file}")

        # 设置窗口标题，显示当前使用的AI
        ai_name = self.ai_config.get('name', '未配置') if self.ai_config else '未配置'
        self.setWindowTitle(f"OCR录入 - 使用AI: {ai_name}")

        # 设置布局
        layout = QVBoxLayout(self)
        layout.addWidget(self.ui)
        layout.setContentsMargins(0, 0, 0, 0)

        # 替换图片预览标签为自定义的ImageLabel
        self.image_label = ImageLabel()
        self.ui.verticalLayout_2.replaceWidget(self.ui.label_2, self.image_label)
        self.ui.label_2.deleteLater()

        # 设置表格为可编辑
        self.ui.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)

        # 设置表格支持多选
        self.ui.tableWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.ui.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)

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

        # tableWidget信号 - 点击行显示对应图片
        self.ui.tableWidget.itemSelectionChanged.connect(self._on_table_selection_changed)

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

        # 获取选中的行
        selected_rows = set()
        for item in self.ui.tableWidget.selectedItems():
            selected_rows.add(item.row())

        # 如果有选中的行，显示"统一设置来源"
        if len(selected_rows) > 0:
            set_source_action = QAction(f"统一设置来源 ({len(selected_rows)} 行)", self)
            set_source_action.triggered.connect(self._on_set_source_for_selected)
            menu.addAction(set_source_action)
            menu.addSeparator()

        delete_action = QAction("删除", self)
        delete_action.triggered.connect(self._on_delete_table_row)
        menu.addAction(delete_action)
        menu.exec(self.ui.tableWidget.viewport().mapToGlobal(position))

    def _on_set_source_for_selected(self):
        """为选中的行统一设置来源"""
        # 获取选中的行
        selected_rows = set()
        for item in self.ui.tableWidget.selectedItems():
            selected_rows.add(item.row())

        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择要设置的行!")
            return

        # 获取当前第一行的来源作为默认值
        default_source = ""
        first_row = min(selected_rows)
        source_item = self.ui.tableWidget.item(first_row, 7)  # 来源在第7列
        if source_item:
            default_source = source_item.text()

        # 弹出对话框
        dialog = SourceDialog(self, current_source=default_source, title=f"统一设置来源 ({len(selected_rows)} 行)")
        if dialog.exec() != QDialog.Accepted:
            return

        new_source = dialog.get_source()

        # 为所有选中的行设置来源
        for row in selected_rows:
            self.ui.tableWidget.setItem(row, 7, QTableWidgetItem(new_source))

        QMessageBox.information(self, "完成", f"已为 {len(selected_rows)} 行设置来源!")

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

        # 启动截图工具，传入self作为父窗口
        take_screenshot(on_screenshot_taken, self)

    def _on_clear(self):
        """清空所有"""
        reply = QMessageBox.question(
            self, "确认清空", "确定要清空所有图片和识别结果吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.ui.listWidget.clear()
            self.ui.tableWidget.setRowCount(0)
            self.image_paths.clear()
            self.image_to_questions.clear()
            self.table_to_image.clear()
            self._clear_temp_dir()

    def _on_list_item_changed(self, current, previous):
        """listWidget选中项改变"""
        if current:
            image_path = current.data(Qt.UserRole)
            self.current_image_path = image_path

            # 加载并显示图片
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                self.image_label.setPixmap(pixmap)

            # 恢复题号记录
            if image_path in self.image_to_questions:
                self.ui.textEdit.setPlainText(self.image_to_questions[image_path])
            else:
                self.ui.textEdit.clear()

    def _on_question_range_changed(self):
        """题号范围改变"""
        if self.current_image_path:
            self.image_to_questions[self.current_image_path] = self.ui.textEdit.toPlainText()

    def _on_table_selection_changed(self):
        """表格选中项改变"""
        selected_items = self.ui.tableWidget.selectedItems()
        if selected_items:
            row = selected_items[0].row()
            if row in self.table_to_image:
                image_path = self.table_to_image[row]
                # 在listWidget中选中对应项
                for i in range(self.ui.listWidget.count()):
                    item = self.ui.listWidget.item(i)
                    if item.data(Qt.UserRole) == image_path:
                        self.ui.listWidget.setCurrentItem(item)
                        break

    def _validate_question_ranges(self) -> bool:
        """验证所有图片的题号范围"""
        invalid_items = []

        for i in range(self.ui.listWidget.count()):
            item = self.ui.listWidget.item(i)
            image_path = item.data(Qt.UserRole)
            question_range = self.image_to_questions.get(image_path, "")

            if not self._is_valid_question_range(question_range):
                invalid_items.append((i, item.text()))

        if invalid_items:
            msg = "以下图片的题号范围格式不正确（只能包含数字、英文逗号和连字符）：\n"
            for idx, name in invalid_items:
                msg += f"  - {name}\n"
            QMessageBox.warning(self, "题号格式错误", msg)
            # 选中第一个有问题的项
            self.ui.listWidget.setCurrentRow(invalid_items[0][0])
            return False

        return True

    def _is_valid_question_range(self, text: str) -> bool:
        """检查题号范围格式是否有效"""
        text = text.strip()
        if not text:
            return True  # 空值表示识别全部题目
        # 只能包含数字、英文逗号、连字符
        import re
        return bool(re.match(r'^[\d,\-\s]+$', text))

    def _on_start_ocr(self):
        """开始AI识别"""
        # 检查是否有图片
        if self.ui.listWidget.count() == 0:
            QMessageBox.warning(self, "警告", "请先添加图片！")
            return

        # 验证题号范围
        if not self._validate_question_ranges():
            return

        # 检查是否配置了AI模型
        if not self.ai_config:
            QMessageBox.warning(
                self, "警告",
                "未配置AI模型！\n\n"
                "请在设置页面的\"AI模型管理\"标签中：\n"
                "1. 添加AI模型配置\n"
                "2. 选择\"自动识别AI模型配置\"\n"
                "3. 点击保存按钮"
            )
            return

        # 检查AI配置是否完整
        if not self.ai_config.get('api_key'):
            QMessageBox.warning(
                self, "警告",
                f"AI模型 \"{self.ai_config.get('name', '未知')}\" 未设置API Key！\n\n"
                "请在设置页面中修改该AI配置，填入正确的API Key。"
            )
            return

        # 准备任务列表
        image_tasks = []
        for i in range(self.ui.listWidget.count()):
            item = self.ui.listWidget.item(i)
            image_path = item.data(Qt.UserRole)
            question_range = self.image_to_questions.get(image_path, "")
            image_tasks.append((image_path, question_range))

        # 获取是否生成解析
        generate_analysis = self.ui.checkBox.isChecked() if hasattr(self.ui, 'checkBox') else False

        # 创建进度对话框
        self.progress_dialog = QProgressDialog(
            f"正在使用 {self.ai_config.get('name', 'AI')} 识别...", "取消", 0, len(image_tasks), self
        )
        self.progress_dialog.setWindowTitle("请稍候")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.show()

        # 创建并启动AI线程
        self.ai_thread = ZhipuAIThread(image_tasks, generate_analysis, self.ai_config)
        self.ai_thread.progress_signal.connect(self._on_ai_progress)
        self.ai_thread.result_signal.connect(self._on_ai_result)
        self.ai_thread.error_signal.connect(self._on_ai_error)
        self.ai_thread.log_signal.connect(self._on_ai_log)
        self.ai_thread.finished.connect(self._on_ai_finished)
        self.progress_dialog.canceled.connect(self.ai_thread.stop)
        self.ai_thread.start()

    def _on_ai_progress(self, current: int, total: int):
        """AI识别进度更新"""
        if self.progress_dialog and self.progress_dialog is not None:
            try:
                self.progress_dialog.setValue(current)
                self.progress_dialog.setLabelText(f"正在识别... {current}/{total}")
            except:
                pass

    def _on_ai_result(self, image_path: str, questions: List[Dict]):
        """AI识别成功"""
        # 添加到表格
        for question in questions:
            self._add_question_to_table(question, image_path)

    def _on_ai_error(self, image_path: str, error_msg: str):
        """AI识别失败"""
        # 日志输出
        app_logger.error(f"[OCR错误] 图片: {os.path.basename(image_path)}")
        app_logger.error(f"[OCR错误] 错误信息: {error_msg}")
        logger.error(f"[OCR错误] 图片: {os.path.basename(image_path)}, 错误: {error_msg}")

        # 显示提示框
        QMessageBox.warning(
            self,
            "识别失败",
            f"图片识别失败: {os.path.basename(image_path)}\n\n错误信息:\n{error_msg}"
        )

    def _on_ai_log(self, message: str):
        """AI识别日志"""
        logger.info(f"[AI] {message}")

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
        """删除表格行（支持多选）"""
        # 获取所有选中的行
        selected_rows = set()
        for item in self.ui.tableWidget.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            return
        
        # 确认删除
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除选中的 {len(selected_rows)} 行吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 按降序删除，避免行号变化影响
        for row in sorted(selected_rows, reverse=True):
            self.ui.tableWidget.removeRow(row)
            # 更新映射关系
            if row in self.table_to_image:
                del self.table_to_image[row]
        
        # 重新构建映射关系（因为行号变了）
        new_mapping = {}
        for row in range(self.ui.tableWidget.rowCount()):
            # 找到原来映射到这个行的图片
            for old_row, image_path in self.table_to_image.items():
                if old_row not in new_mapping.values():
                    new_mapping[row] = image_path
                    break
        self.table_to_image = new_mapping

    def _on_open_image_dir(self, item: QListWidgetItem):
        """在目录中打开图片"""
        image_path = item.data(Qt.UserRole)
        if image_path and os.path.exists(image_path):
            import subprocess
            subprocess.run(['explorer', '/select,', os.path.normpath(image_path)])

    def _on_confirm_import(self):
        """确认导入"""
        # 验证数据
        invalid_rows = []
        invalid_answer_rows = []
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

            # 检查答案（只能是A、B、C、D）
            answer_item = self.ui.tableWidget.item(row, 5)
            answer = answer_item.text().strip().upper() if answer_item else ""
            valid_answer = answer in ["A", "B", "C", "D"]
            if answer and not valid_answer:
                invalid_answer_rows.append(row + 1)

            if not has_options or not valid_classification:
                invalid_rows.append(row + 1)

        # 检查答案是否合法
        if invalid_answer_rows:
            QMessageBox.warning(
                self, "答案格式错误",
                f"以下行的答案格式不正确，只能是 A、B、C 或 D：\n行号: {', '.join(map(str, invalid_answer_rows))}"
            )
            return

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
                f'以下行的来源为空，是否自动填充为"无"？\n行号: {", ".join(map(str, empty_source_rows))}',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                for row in empty_source_rows:
                    self.ui.tableWidget.setItem(row - 1, 7, QTableWidgetItem("无"))
            else:
                return

        # 收集所有题目
        questions = []
        for row in range(self.ui.tableWidget.rowCount()):
            question = {
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
            questions.append(question)

        # 输出到日志
        logger.info("=" * 50)
        logger.info(f"准备导入 {len(questions)} 道题目：")
        logger.info("=" * 50)
        for i, q in enumerate(questions, 1):
            logger.info(f"题目 {i}:")
            logger.info(f"  问题: {q['question'][:50]}...")
            logger.info(f"  A: {q['A'][:30]}...")
            logger.info(f"  B: {q['B'][:30]}...")
            logger.info(f"  C: {q['C'][:30]}...")
            logger.info(f"  D: {q['D'][:30]}...")
            logger.info(f"  答案: {q['answer']}")
            logger.info(f"  分类: {q['classification']}")
            logger.info(f"  来源: {q['source']}")
        logger.info("=" * 50)

        # 发射导入信号（由主窗口处理导入逻辑并显示提示）
        self.import_signal.emit(questions)

        # 清空表格，避免重复导入
        self.ui.tableWidget.setRowCount(0)
        self.table_to_image.clear()

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


def show_ocr_window(parent=None, ai_config: Optional[Dict] = None):
    """显示OCR窗口（模态对话框，阻塞父窗口）

    Args:
        parent: 父窗口
        ai_config: AI配置字典，包含 name, base_url, api_key
    """
    window = OCRWindow(ai_config)
    window.setParent(parent, Qt.Dialog)
    window.setWindowModality(Qt.ApplicationModal)
    window.show()
    # 窗口显示后再居中
    window._center_window()
    return window
