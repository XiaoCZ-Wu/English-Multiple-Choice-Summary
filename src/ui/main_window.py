"""
主窗口类
整合所有页面和功能
"""

import sys
import os
from typing import List, Dict

from PySide6.QtCore import Qt, QEvent, QObject
from PySide6.QtGui import QFont, QAction, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QFileDialog,
    QTableWidgetItem, QButtonGroup, QMenu, QTableWidget, QDialog
)
from PySide6.QtUiTools import QUiLoader

from src.models import DataManager, AppConfig, Question, AIConfig
from src.core import PracticeManager, PracticeMode, ExportManager
from src.utils import (
    CLASSIFICATIONS, OPTIONS, format_time, format_accuracy, UI_FILE, BACKUP_DIR,
    OCR_TEMP_DIR, TEMP_DIR, app_logger,
    COL_QUESTION, COL_OPTION_A, COL_OPTION_B, COL_OPTION_C, COL_OPTION_D,
    COL_ANSWER, COL_CLASSIFICATION, COL_ACCURACY, COL_SOURCE, COL_ANALYSIS,
    PAGE_HOME, PAGE_CREATE, PAGE_MANAGE, PAGE_SETTINGS, PAGE_PRACTICE, PAGE_REPORT
)
from src.ui.dialogs import ExportDialog, AIConfigDialog, AIChatDialog


class MainWindow(QObject):
    """主窗口控制器"""
    
    def __init__(self):
        super().__init__()
        
        # 重启标志
        self._restart_required = False
        
        # 报告导出标志
        self._report_exported = False
        
        # 初始化数据
        self.config = AppConfig.load()
        self.data_manager = DataManager()
        self.data_manager.load()
        self.practice_manager = PracticeManager(self.data_manager)
        
        # 状态变量
        self.window_title = ""
        self.edit_mode = False
        self.editable_rows = []
        self.selected_rows = []
        self.deleted_questions = []
        self.result_list = []
        
        # 存储原始API Key（用于保存时恢复，避免保存脱敏后的值）
        self._ai_api_keys = {}
        
        # 按钮组
        self.btn_group_answer = None
        self.btn_group_classification = None
        self.btn_group_accuracy = None
        self.btn_group_mode = None
        self.btn_group_option = None
        
        # 右键菜单
        self.context_menu = None
        
        # 局域网服务器线程
        self.lan_server_thread = None
        
        # 清理temp/ocr目录
        self._clean_ocr_temp_dir()
        
        self._setup_ui()
        self._setup_connections()
        self._init_state()
    
    def _clean_ocr_temp_dir(self):
        """清理OCR临时目录"""
        try:
            import os
            import shutil
            
            # 使用常量中的OCR临时目录路径
            ocr_temp_dir = OCR_TEMP_DIR
            
            # 如果目录存在，删除软件端创建的文件（保留网页端上传的 web_ 开头文件）
            if os.path.exists(ocr_temp_dir):
                for filename in os.listdir(ocr_temp_dir):
                    # 跳过网页端上传的文件（以 web_ 开头）
                    if filename.startswith('web_'):
                        continue
                    file_path = os.path.join(ocr_temp_dir, filename)
                    try:
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        app_logger.warning(f"删除文件失败 {file_path}: {e}")
                app_logger.info(f"已清理OCR临时目录（保留网页端文件）: {ocr_temp_dir}")
        except Exception as e:
            app_logger.warning(f"清理OCR临时目录失败: {e}")
    
    def _setup_ui(self):
        """设置UI"""
        # 加载UI文件 - 使用 constants 中的路径
        from src.utils.constants import UI_FILE
        ui_file_path = UI_FILE

        loader = QUiLoader()
        self.main_window = loader.load(ui_file_path)
        if not self.main_window:
            raise RuntimeError(f"无法加载UI文件: {ui_file_path}")
        self.window_title = self.main_window.windowTitle()
        
        # 设置窗口图标
        from src.utils.constants import get_resource_path
        import os
        icon_path = get_resource_path('src/ico/ico.ico')
        if os.path.exists(icon_path):
            self.main_window.setWindowIcon(QIcon(icon_path))
        
        # 安装事件过滤器
        self.main_window.installEventFilter(self)
        
        # 设置字体
        font = QFont(self.config.font_name, self.config.font_size)
        QApplication.instance().setFont(font)
        
        # 设置文本编辑器字体
        text_edits = [
            self.main_window.textEdit, self.main_window.textEdit_2,
            self.main_window.textEdit_3, self.main_window.textEdit_4,
            self.main_window.textEdit_5, self.main_window.textEdit_6,
            self.main_window.textEdit_7, self.main_window.textEdit_8,
            self.main_window.textEdit_9, self.main_window.textEdit_10,
            self.main_window.textEdit_11
        ]
        for te in text_edits:
            te.setFont(font)
        
        # 初始化按钮组
        self._init_button_groups()
        
        # 初始化右键菜单
        self._init_context_menu()
        
        # 初始化设置页面
        self._init_settings_page()
        
        # 初始化题目列表
        self._init_questions_list()
    
    def _init_button_groups(self):
        """初始化按钮组"""
        # 答案按钮组
        self.btn_group_answer = QButtonGroup(self.main_window)
        for i, btn in enumerate([
            self.main_window.radioButton, self.main_window.radioButton_2,
            self.main_window.radioButton_3, self.main_window.radioButton_4
        ]):
            self.btn_group_answer.addButton(btn, i)
        
        # 分类按钮组
        self.btn_group_classification = QButtonGroup(self.main_window)
        for i, btn in enumerate([
            self.main_window.radioButton_5, self.main_window.radioButton_6,
            self.main_window.radioButton_7, self.main_window.radioButton_8,
            self.main_window.radioButton_9, self.main_window.radioButton_10,
            self.main_window.radioButton_11, self.main_window.radioButton_12
        ]):
            self.btn_group_classification.addButton(btn, i)
        
        # 正确率按钮组
        self.btn_group_accuracy = QButtonGroup(self.main_window)
        for i, btn in enumerate([
            self.main_window.radioButton_13, self.main_window.radioButton_14,
            self.main_window.radioButton_15
        ]):
            self.btn_group_accuracy.addButton(btn, i)
        
        # 模式按钮组
        self.btn_group_mode = QButtonGroup(self.main_window)
        for i, btn in enumerate([
            self.main_window.radioButton_16, self.main_window.radioButton_17
        ]):
            self.btn_group_mode.addButton(btn, i)
        
        # 选项按钮组（练习页面）
        self.btn_group_option = QButtonGroup(self.main_window)
        for i, btn in enumerate([
            self.main_window.radioButton_18, self.main_window.radioButton_19,
            self.main_window.radioButton_20, self.main_window.radioButton_21
        ]):
            self.btn_group_option.addButton(btn, i)
    
    def _init_context_menu(self):
        """初始化右键菜单"""
        from PySide6.QtWidgets import QMenu
        self.context_menu = QMenu()

        # 统一设置来源
        self.set_source_action = QAction("统一设置来源", self.main_window.tableWidget)
        self.set_source_action.triggered.connect(self._on_set_source_for_selected)
        self.context_menu.addAction(self.set_source_action)

        self.context_menu.addSeparator()

        edit_action = QAction("编辑", self.main_window.tableWidget)
        edit_action.triggered.connect(self._on_edit_question)

        delete_action = QAction("删除", self.main_window.tableWidget)
        delete_action.triggered.connect(self._on_delete_question)

        self.context_menu.addAction(edit_action)
        self.context_menu.addAction(delete_action)

        self.main_window.tableWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.main_window.tableWidget.customContextMenuRequested.connect(self._show_context_menu)
    
    def _init_settings_page(self):
        """初始化设置页面"""
        # 基本设置
        self.main_window.spinBox.setValue(self.config.font_size)
        self.main_window.fontComboBox_2.setCurrentText(self.config.font_name)
        self.main_window.lineEdit.setPlaceholderText(".\\output\\")
        self.main_window.lineEdit.setText(self.config.output_dir)

        # AI设置
        self._refresh_ai_combos()
        self._refresh_ai_table()

        # 设置AI表格右键菜单（只连接一次信号）
        if not hasattr(self, '_ai_table_context_menu_connected'):
            table = self.main_window.tableWidget_2
            table.setContextMenuPolicy(Qt.CustomContextMenu)
            table.customContextMenuRequested.connect(self._show_ai_table_context_menu)
            self._ai_table_context_menu_connected = True
        
        # 局域网端口设置
        self.main_window.lineEdit_2.setText(str(self.config.lan_port))
        self.main_window.lineEdit_2.textChanged.connect(self._on_lan_port_changed)
        
        # 设置OCR网页链接label
        self._update_ocr_link_label()
        self.main_window.label_26.setCursor(Qt.PointingHandCursor)
        # 安装事件过滤器来捕获点击事件
        self.main_window.label_26.installEventFilter(self)
        
        # 设置手机端网址label
        self._update_mobile_link_label()
        self.main_window.label_29.setCursor(Qt.PointingHandCursor)
        self.main_window.label_29.installEventFilter(self)

        # 设置软件版本信息
        self._set_version_info()

        # 启动局域网服务器（延迟启动，避免启动时崩溃）
        from PySide6.QtCore import QTimer
        QTimer.singleShot(3000, self._start_lan_server_safe)

    def _refresh_ai_combos(self):
        """刷新AI配置下拉框"""
        # 清空下拉框
        self.main_window.comboBox_3.clear()
        self.main_window.comboBox_4.clear()

        # 添加AI配置
        ai_names = [cfg.name for cfg in self.config.ai_configs if cfg.name]
        self.main_window.comboBox_3.addItems(ai_names)
        self.main_window.comboBox_4.addItems(ai_names)

        # 设置当前选中的AI
        if self.config.ocr_ai_name in ai_names:
            self.main_window.comboBox_3.setCurrentText(self.config.ocr_ai_name)
        if self.config.chat_ai_name in ai_names:
            self.main_window.comboBox_4.setCurrentText(self.config.chat_ai_name)

    def _set_version_info(self):
        """设置软件版本信息到软件信息页面"""
        from src.models.config import CONFIG_VERSION

        version_text = f"""<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
<html><head><meta name="qrichtext" content="1" /><meta charset="utf-8" /><style type="text/css">
p, li {{ white-space: pre-wrap; }}
hr {{ height: 1px; border-width: 0; }}
li.unchecked::marker {{ content: "\2610"; }}
li.checked::marker {{ content: "\2612"; }}
</style></head><body style=" font-family:'Microsoft YaHei UI'; font-size:11.25pt; font-weight:400; font-style:normal;">
<p style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">作者：小茶子XiaoCZ<br />版本：{CONFIG_VERSION}</p>
<p style=" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">仓库：<a href="https://github.com/XiaoCZ-Wu/English-Multiple-Choice-Summary"><span style=" font-size:11pt; text-decoration: underline; color:#008c67;">XiaoCZ-Wu/English-Multiple-Choice-Summary</span></a></p></body></html>"""

        self.main_window.textEdit_8.setHtml(version_text)

    def _refresh_ai_table(self):
        """刷新AI配置表格"""
        table = self.main_window.tableWidget_2
        table.setRowCount(len(self.config.ai_configs))

        # 设置表格为只读模式
        table.setEditTriggers(QTableWidget.NoEditTriggers)

        # 设置为单选模式
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setSelectionBehavior(QTableWidget.SelectRows)

        # 设置列宽 - 列顺序：名称(0)、baseurl(1)、模型(2)、key(3)
        table.setColumnWidth(0, 150)  # 名称列
        table.setColumnWidth(1, 350)  # Base URL列
        table.setColumnWidth(2, 150)  # 模型列
        table.setColumnWidth(3, 400)  # API Key列

        # 清空原始API Key映射（重新加载时）
        self._ai_api_keys = {}
        
        for row, cfg in enumerate(self.config.ai_configs):
            table.setItem(row, 0, QTableWidgetItem(cfg.name))
            table.setItem(row, 1, QTableWidgetItem(cfg.base_url))
            # 模型列
            table.setItem(row, 2, QTableWidgetItem(cfg.model if cfg.model else "glm-4v-flash"))
            # 存储原始API Key
            self._ai_api_keys[cfg.name] = cfg.api_key
            # API Key脱敏显示：保留前4位和后4位，中间用***代替
            api_key_display = self._mask_api_key(cfg.api_key)
            table.setItem(row, 3, QTableWidgetItem(api_key_display))

    def _show_ai_table_context_menu(self, position):
        """显示AI表格右键菜单"""
        menu = QMenu()

        add_action = QAction("添加", self)
        add_action.triggered.connect(self._on_add_ai_config)
        menu.addAction(add_action)

        # 获取当前选中的行
        table = self.main_window.tableWidget_2
        current_row = table.currentRow()

        if current_row >= 0:
            edit_action = QAction("修改", self)
            edit_action.triggered.connect(self._on_edit_ai_config)
            menu.addAction(edit_action)

            delete_action = QAction("删除", self)
            delete_action.triggered.connect(self._on_delete_ai_config)
            menu.addAction(delete_action)

        menu.exec(table.viewport().mapToGlobal(position))

    def _mask_api_key(self, api_key: str) -> str:
        """对API Key进行脱敏显示，保留前4位和后4位"""
        if not api_key:
            return ""
        if len(api_key) <= 8:
            return "*" * len(api_key)
        return api_key[:4] + "***" + api_key[-4:]

    def _on_add_ai_config(self):
        """添加AI配置（仅添加到界面，保存后才生效）"""
        # 使用自定义对话框
        dialog = AIConfigDialog(self.main_window, title="添加AI配置")
        if dialog.exec() != QDialog.Accepted:
            return

        values = dialog.get_values()
        name = values["name"]
        base_url = values["base_url"]
        api_key = values["api_key"]
        model = values["model"]

        if not name:
            QMessageBox.warning(self.main_window, "警告", "AI名称不能为空!")
            return

        if not model:
            QMessageBox.warning(self.main_window, "警告", "模型不能为空!")
            return

        # 检查名称是否已存在（在当前表格中）
        table = self.main_window.tableWidget_2
        for row in range(table.rowCount()):
            if table.item(row, 0) and table.item(row, 0).text() == name:
                QMessageBox.warning(self.main_window, "警告", "该名称已存在!")
                return

        # 添加到表格（仅界面显示，不保存到配置）
        # 列顺序：名称(0)、baseurl(1)、模型(2)、key(3)
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(name))
        table.setItem(row, 1, QTableWidgetItem(base_url))
        table.setItem(row, 2, QTableWidgetItem(model))
        table.setItem(row, 3, QTableWidgetItem(self._mask_api_key(api_key)))
        
        # 存储原始API Key
        self._ai_api_keys[name] = api_key

        # 添加到下拉框
        self.main_window.comboBox_3.addItem(name)
        self.main_window.comboBox_4.addItem(name)

        QMessageBox.information(self.main_window, "提示", "AI配置已添加，请点击保存按钮生效!")

    def _on_edit_ai_config(self):
        """修改AI配置（仅修改界面，保存后才生效）"""
        table = self.main_window.tableWidget_2
        current_row = table.currentRow()
        if current_row < 0:
            return

        old_name = table.item(current_row, 0).text()
        old_base_url = table.item(current_row, 1).text()
        old_model = table.item(current_row, 2).text() if table.item(current_row, 2) else ""

        # 从映射中获取原始API Key（表格中的是脱敏的）
        original_api_key = self._ai_api_keys.get(old_name, "")

        # 使用自定义对话框
        dialog = AIConfigDialog(
            self.main_window,
            name=old_name,
            base_url=old_base_url,
            api_key=original_api_key,
            model=old_model,
            title="修改AI配置",
            is_edit=True
        )
        if dialog.exec() != QDialog.Accepted:
            return

        values = dialog.get_values()
        name = values["name"]
        base_url = values["base_url"]
        api_key = values["api_key"]
        model = values["model"]

        if not name:
            QMessageBox.warning(self.main_window, "警告", "AI名称不能为空!")
            return

        if not model:
            QMessageBox.warning(self.main_window, "警告", "模型不能为空!")
            return

        # 检查新名称是否与其他行冲突
        for row in range(table.rowCount()):
            if row != current_row and table.item(row, 0) and table.item(row, 0).text() == name:
                QMessageBox.warning(self.main_window, "警告", "该名称已存在!")
                return

        # 更新表格中的数据（仅界面显示，不保存到配置）
        # 列顺序：名称(0)、baseurl(1)、模型(2)、key(3)
        table.item(current_row, 0).setText(name)
        table.item(current_row, 1).setText(base_url)
        table.item(current_row, 2).setText(model)
        table.item(current_row, 3).setText(self._mask_api_key(api_key))
        
        # 更新API Key映射
        if old_name in self._ai_api_keys:
            del self._ai_api_keys[old_name]
        self._ai_api_keys[name] = api_key

        # 更新下拉框中的名称
        combo3_index = self.main_window.comboBox_3.findText(old_name)
        combo4_index = self.main_window.comboBox_4.findText(old_name)
        if combo3_index >= 0:
            self.main_window.comboBox_3.setItemText(combo3_index, name)
        if combo4_index >= 0:
            self.main_window.comboBox_4.setItemText(combo4_index, name)

        QMessageBox.information(self.main_window, "提示", "AI配置已修改，请点击保存按钮生效!")

    def _on_delete_ai_config(self):
        """删除AI配置（仅标记删除，保存后才生效）"""
        table = self.main_window.tableWidget_2
        current_row = table.currentRow()
        if current_row < 0:
            return

        name = table.item(current_row, 0).text()

        reply = QMessageBox.question(
            self.main_window, "确认删除",
            f"确定要删除AI配置 '{name}' 吗?\n（点击保存后生效）",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 从表格中移除该行（仅界面显示，不实际删除配置）
            table.removeRow(current_row)
            # 从API Key映射中移除
            if name in self._ai_api_keys:
                del self._ai_api_keys[name]
            # 从下拉框中移除
            self._refresh_ai_combos_after_delete(name)
            QMessageBox.information(self.main_window, "提示", "已标记删除，请点击保存按钮生效!")

    def _refresh_ai_combos_after_delete(self, deleted_name: str):
        """删除后刷新下拉框（从界面中移除）"""
        # 获取表格中剩余的所有AI名称
        table = self.main_window.tableWidget_2
        remaining_names = []
        for row in range(table.rowCount()):
            name = table.item(row, 0).text()
            if name:
                remaining_names.append(name)

        # 清空并重新填充下拉框
        self.main_window.comboBox_3.clear()
        self.main_window.comboBox_4.clear()
        self.main_window.comboBox_3.addItems(remaining_names)
        self.main_window.comboBox_4.addItems(remaining_names)

        # 如果被删除的是当前选中的，则清空选择
        if self.main_window.comboBox_3.currentText() == deleted_name:
            self.main_window.comboBox_3.setCurrentIndex(-1)
        if self.main_window.comboBox_4.currentText() == deleted_name:
            self.main_window.comboBox_4.setCurrentIndex(-1)
    
    def _init_questions_list(self):
        """初始化题目列表"""
        self.result_list = self.data_manager.get_all_questions()
        self._refresh_table(self.result_list)
        
        # 初始化下拉菜单
        self.main_window.comboBox.clear()
        self.main_window.comboBox.addItem("Any")
        self.main_window.comboBox.addItems(self.data_manager.get_papers())
    
    def _on_open_ocr_window(self):
        """打开OCR窗口"""
        from .ocr_window import show_ocr_window

        # 获取当前选中的OCR AI配置
        ocr_ai_config = None
        if self.config.ocr_ai_name:
            ocr_ai_config = self.config.get_ai_config(self.config.ocr_ai_name)
            if ocr_ai_config:
                ocr_ai_config = ocr_ai_config.to_dict()

        self.ocr_window = show_ocr_window(self.main_window, ocr_ai_config)
        
        # 连接导入信号
        self.ocr_window.import_signal.connect(self._on_ocr_import_questions)
    
    def _on_ocr_import_questions(self, questions: List[Dict]):
        """处理OCR导入的题目
        
        Args:
            questions: 题目列表，每个题目是一个字典
        """
        if not questions:
            return
        
        # 导入题目到题库
        imported_count = 0
        for q in questions:
            # 获取分类索引
            classification = q.get('classification', '')
            if classification in CLASSIFICATIONS:
                classification_idx = CLASSIFICATIONS.index(classification)
            else:
                classification_idx = 0  # 默认使用第一个分类
            
            # 创建题目对象
            question = Question(
                question=q.get('question', ''),
                A=q.get('A', ''),
                B=q.get('B', ''),
                C=q.get('C', ''),
                D=q.get('D', ''),
                answer=q.get('answer', '').upper(),
                classification=classification_idx,
                source=q.get('source', '无'),
                analysis=q.get('analysis', '')
            )
            
            # 添加到数据管理器
            self.data_manager.questions.append(question)
            
            # 添加来源到papers（如果不存在）
            if question.source and question.source not in self.data_manager.papers:
                self.data_manager.papers.append(question.source)
            
            imported_count += 1
        
        # 保存到文件
        try:
            if self.data_manager.save():
                # 更新UI
                self._check_comboBox()  # 更新来源下拉框
                QMessageBox.information(
                    self.main_window, "导入成功",
                    f"成功导入 {imported_count} 道题目到题库！"
                )
            else:
                QMessageBox.critical(
                    self.main_window, "导入失败",
                    "保存题库失败，请检查文件权限。"
                )
        except Exception as e:
            QMessageBox.critical(
                self.main_window, "导入失败",
                f"保存题库时发生错误：{e}"
            )
    
    def _start_lan_server(self):
        """启动局域网服务器（只启动一次）"""
        # 如果服务器已经启动，不再重复启动
        if self.lan_server_thread is not None:
            return
            
        try:
            from src.core.lan_server import start_server
            import socket
            
            # 获取本机IP地址
            hostname = socket.gethostname()
            try:
                local_ip = socket.getaddrinfo(hostname, None, socket.AF_INET)[0][4][0]
            except:
                local_ip = "127.0.0.1"
            
            # 启动服务器 - 监听所有网络接口，允许局域网访问
            self.lan_server_thread = start_server(
                self.data_manager,
                self.config,
                port=self.config.lan_port,
                threaded=True,
                host='0.0.0.0',
                use_https=False
            )
            
            # 输出访问地址到日志
            app_logger.info("=" * 60)
            app_logger.info("📱 Web OCR服务已启动！")
            app_logger.info("-" * 60)
            app_logger.info(f"🌐 本地访问: http://127.0.0.1:{self.config.lan_port}")
            app_logger.info(f"🌐 局域网访问: http://{local_ip}:{self.config.lan_port}")
            app_logger.info("-" * 60)
            app_logger.info("💡 在同一WiFi下的手机浏览器中访问上述地址")
            app_logger.info("   即可使用OCR识别功能")
            app_logger.info("=" * 60)
            
        except Exception as e:
            app_logger.error(f"❌ 启动局域网服务器失败: {e}")

    def _start_lan_server_safe(self):
        """安全启动局域网服务器（使用QThread避免阻塞UI）"""
        from PySide6.QtCore import QThread

        class ServerThread(QThread):
            def __init__(self, main_window):
                super().__init__()
                self.main_window = main_window

            def run(self):
                self.main_window._start_lan_server()

        self._server_thread = ServerThread(self)
        self._server_thread.start()

    def _update_ocr_link_label(self):
        """更新OCR网页链接label的文本"""
        # 使用回环地址，仅供本机访问
        url = f"http://127.0.0.1:{self.config.lan_port}"
        self.main_window.label_26.setText(f"🌐 OCR网页：{url}（点击打开）")
        self.main_window.label_26.setToolTip(f"点击打开 {url}")
    
    def _on_open_ocr_web(self):
        """点击label打开OCR网页"""
        import webbrowser
        
        # 使用回环地址打开网页
        url = f"http://127.0.0.1:{self.config.lan_port}"
        webbrowser.open(url)
    
    def _update_mobile_link_label(self):
        """更新手机端网址label的文本"""
        import socket
        try:
            hostname = socket.gethostname()
            local_ip = socket.getaddrinfo(hostname, None, socket.AF_INET)[0][4][0]
        except:
            local_ip = "127.0.0.1"
        
        url = f"http://{local_ip}:{self.config.lan_port}"
        self.main_window.label_29.setText(f"📱 手机端：{url}（点击复制）")
        self.main_window.label_29.setToolTip(f"点击复制网址：{url}")
    
    def _on_copy_mobile_url(self):
        """点击label复制网址到剪贴板"""
        import socket
        from PySide6.QtWidgets import QApplication
        
        try:
            hostname = socket.gethostname()
            local_ip = socket.getaddrinfo(hostname, None, socket.AF_INET)[0][4][0]
        except:
            local_ip = "127.0.0.1"
        
        url = f"http://{local_ip}:{self.config.lan_port}"
        
        # 复制到剪贴板
        clipboard = QApplication.clipboard()
        clipboard.setText(url)
        
        # 显示提示
        QMessageBox.information(self.main_window, "提示", f"网址已复制到剪贴板：\n{url}")
    
    def _on_lan_port_changed(self, text):
        """局域网端口改变 - 只验证不重置，避免循环"""
        try:
            if text:
                port = int(text)
                if 1024 <= port <= 65535:
                    self.config.lan_port = port
                    # 更新OCR网页链接显示
                    self._update_ocr_link_label()
                    # 更新手机端网址显示
                    self._update_mobile_link_label()
        except ValueError:
            # 输入包含非数字字符，不处理（让用户继续输入）
            pass
    
    def _setup_connections(self):
        """设置信号连接"""
        # 首页按钮
        self.main_window.pushButton.clicked.connect(self._on_start_practice)
        self.main_window.pushButton_2.clicked.connect(lambda: self._goto_page(PAGE_CREATE))
        self.main_window.pushButton_3.clicked.connect(self._on_manage)
        self.main_window.pushButton_4.clicked.connect(lambda: self._goto_page(PAGE_SETTINGS))
        
        # 录入页面按钮
        self.main_window.pushButton_5.clicked.connect(self._on_save_question)
        self.main_window.pushButton_6.clicked.connect(self._on_back)
        
        # 管理页面按钮
        self.main_window.pushButton_7.clicked.connect(self._on_reload)
        self.main_window.pushButton_8.clicked.connect(self._on_back)
        self.main_window.pushButton_9.clicked.connect(self._on_filter)
        self.main_window.pushButton_10.clicked.connect(self._on_select_all)
        self.main_window.pushButton_11.clicked.connect(self._on_export)
        self.main_window.pushButton_12.clicked.connect(self._on_reset_accuracy)
        self.main_window.pushButton_16.clicked.connect(self._on_save_edits)  # 保存
        
        # 练习页面按钮
        self.main_window.pushButton_18.clicked.connect(self._on_practice_back)  # 返回首页（不保存数据）
        self.main_window.pushButton_19.clicked.connect(self._on_confirm_answer)  # 提交答案（confirm_answer）
        self.main_window.pushButton_20.clicked.connect(self._on_end_practice_and_report)  # 结束练习并生成报告
        self.main_window.pushButton_21.clicked.connect(self._on_chat)  # 将问题发送给AI
        
        # 报告页面按钮
        self.main_window.pushButton_22.clicked.connect(self._on_report_back)  # 返回首页（提示保管报告）
        self.main_window.pushButton_24.clicked.connect(self._on_export_report)  # 导出报告
        
        # 设置页面按钮
        self.main_window.pushButton_13.clicked.connect(self._on_back)  # 返回（不保存）
        self.main_window.pushButton_14.clicked.connect(self._on_apply_settings)  # 保存（基本设置需要重启才能生效）
        self.main_window.pushButton_15.clicked.connect(self._on_select_dir)  # 浏览
        self.main_window.pushButton_17.clicked.connect(self._on_backup)  # Backup
        self.main_window.pushButton_25.clicked.connect(self._on_import_backup)  # 导入备份
        self.main_window.pushButton_28.clicked.connect(self._on_open_ocr_window)  # OCR识别
        
        # 表格单元格修改验证
        self.main_window.tableWidget.cellChanged.connect(self._on_cell_changed)
    
    def _init_state(self):
        """初始化状态"""
        # 设置练习管理器回调
        self.practice_manager.set_callbacks(
            on_time_update=self._on_time_update,
            on_question_changed=self._on_question_changed
        )
    
    def eventFilter(self, obj, event):
        """事件过滤器"""
        # 检查对象是否已被删除
        try:
            # 尝试访问对象，如果已被删除会抛出异常
            _ = self.main_window.windowTitle()
        except RuntimeError:
            # 对象已被删除，直接返回
            return False
        
        if obj == self.main_window and event.type() == QEvent.Close:
            # 如果是重启，直接关闭不询问
            if hasattr(self, '_restart_required') and self._restart_required:
                return False  # 允许关闭
            
            # 防止递归调用
            if hasattr(self, '_closing') and self._closing:
                return False
            
            event.ignore()
            reply = QMessageBox.question(
                self.main_window, "提示",
                "是否确定退出? \n在退出前请确保所有数据均已保存，防止丢失!\n做题时必须点击\"生成报告\"按钮，否则数据不会保存!",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self._closing = True  # 标记正在关闭
                # 关闭OCR窗口（如果存在）
                if hasattr(self, 'ocr_window') and self.ocr_window:
                    try:
                        self.ocr_window.close()
                    except:
                        pass
                # 接受关闭事件，让Qt正常关闭窗口
                event.accept()
            return True
        
        # 处理OCR网页链接label点击事件
        try:
            if obj == self.main_window.label_26 and event.type() == QEvent.MouseButtonRelease:
                self._on_open_ocr_web()
                return True
        except RuntimeError:
            pass
        
        # 处理手机端网址label点击事件
        try:
            if obj == self.main_window.label_29 and event.type() == QEvent.MouseButtonRelease:
                self._on_copy_mobile_url()
                return True
        except RuntimeError:
            pass
        
        return super().eventFilter(obj, event)
    
    # ========== 页面导航 ==========
    
    def _goto_page(self, index: int):
        """跳转到指定页面"""
        self.main_window.stackedWidget.setCurrentIndex(index)
    
    def _on_back(self):
        """返回首页"""
        # 如果在练习中，先结束练习
        if self.practice_manager.current_index >= 0:
            # 如果按钮19被禁用，说明已经在最后一题提交了答案，直接返回不弹出提示
            if not self.main_window.pushButton_19.isEnabled():
                self.practice_manager.stop_practice()
                self.practice_manager.current_index = -1
                self.main_window.textEdit_11.clear()
            elif not self.practice_manager.is_finished():
                reply = QMessageBox.question(
                    self.main_window, "提示", "将会按照现有的进度生成报告! 是否继续?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    self._show_report()
                return
            else:
                self.practice_manager.stop_practice()
                self.practice_manager.current_index = -1
                self.main_window.textEdit_11.clear()
        
        # 检查是否有未保存的编辑
        if self.edit_mode or self.deleted_questions:
            reply = QMessageBox.question(
                self.main_window, "提示", "存在未保存的修改，是否保存?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            if reply == QMessageBox.Save:
                self._on_save_edits()
                return
            elif reply == QMessageBox.Cancel:
                return
            # Discard: 放弃修改，继续返回
        
        # 重置状态
        self.main_window.setWindowTitle(self.window_title)
        self._on_reload(from_code=True)
        self._init_settings_page()
        self.edit_mode = False
        self.editable_rows.clear()
        self.deleted_questions.clear()
        self._goto_page(PAGE_HOME)
    
    def _on_report_back(self):
        """练习报告页面返回首页"""
        # 如果已经导出过报告，直接返回不询问
        if hasattr(self, '_report_exported') and self._report_exported:
            self._do_report_back()
            return
        
        # 未导出报告，显示提示
        msg_box = QMessageBox(self.main_window)
        msg_box.setWindowTitle("提示")
        msg_box.setText("请妥善保管练习报告！")
        msg_box.setIcon(QMessageBox.Information)
        
        # 添加自定义按钮
        confirm_btn = msg_box.addButton("确认返回", QMessageBox.YesRole)
        cancel_btn = msg_box.addButton("取消", QMessageBox.NoRole)
        
        msg_box.exec()
        
        if msg_box.clickedButton() == confirm_btn:
            self._do_report_back()
        # 取消则留在报告页面
    
    def _do_report_back(self):
        """执行返回首页操作"""
        # 结束练习并返回首页
        self.practice_manager.stop_practice()
        self.practice_manager.current_index = -1
        self.main_window.textEdit_11.clear()
        self.main_window.setWindowTitle(self.window_title)
        self._on_reload(from_code=True)
        self._init_settings_page()
        self._goto_page(PAGE_HOME)
        # 重置报告导出标志
        self._report_exported = False
    
    def _on_export_report(self):
        """导出报告为txt文件"""
        # 获取报告内容
        report_text = self.main_window.textEdit_11.toPlainText()
        if not report_text:
            QMessageBox.warning(self.main_window, "警告", "没有可导出的报告内容！")
            return
        
        # 生成文件名（使用与backup相同的时间戳格式）
        from src.utils import get_timestamp
        import os
        
        timestamp = get_timestamp()
        report_filename = f"report_{timestamp}.txt"
        # 使用config中配置的output_dir
        output_dir = self.config.output_dir
        report_path = os.path.join(output_dir, report_filename)
        
        # 确保output目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            # 写入文件
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            
            QMessageBox.information(self.main_window, "提示", f"报告已导出到：\n{report_path}")
            
            # 标记报告已保存，返回首页时不再询问
            self._report_exported = True
            
        except Exception as e:
            QMessageBox.critical(self.main_window, "错误", f"导出报告失败：{str(e)}")
    
    # ========== 首页功能 ==========
    
    def _on_start_practice(self):
        """开始练习"""
        mode = self.btn_group_mode.checkedId()
        
        if mode == -1:
            QMessageBox.information(self.main_window, "提示", "你还没有选择练习模式!")
            return
        
        # 清空界面
        self._clear_practice_page()
        
        # 启动练习
        practice_mode = PracticeMode.ENDLESS if mode == 0 else PracticeMode.PAPER
        if self.practice_manager.start_practice(practice_mode):
            self._goto_page(PAGE_PRACTICE)
        else:
            QMessageBox.information(self.main_window, "提示", "需要保证题库中至少有一道题才能开始练习！")
    
    def _clear_practice_page(self):
        """清空练习页面"""
        self.main_window.textEdit_9.clear()
        self.main_window.textEdit_10.clear()
        
        # 启用提交按钮（可能被上一轮的最后一题禁用）
        self.main_window.pushButton_19.setEnabled(True)
        self.main_window.pushButton_19.setText("提交")
        
        self.btn_group_option.setExclusive(False)
        for btn in self.btn_group_option.buttons():
            btn.setChecked(False)
        self.btn_group_option.setExclusive(True)
        
        labels = [
            self.main_window.label_14, self.main_window.label_15,
            self.main_window.label_16, self.main_window.label_17,
            self.main_window.label_18, self.main_window.label_19,
            self.main_window.label_20, self.main_window.label_21,
            self.main_window.label_23
        ]
        texts = [
            "第 - 题，共 - 题（0.00%）", "单题用时：00: 00",
            "累计用时：00: 00", "第 - 次刷到该题",
            "过往正确率：0.00%", "当前练习模式：-",
            "所属套题：-", "正确答案：-",
            "这题选什么呀？"
        ]
        for lbl, txt in zip(labels, texts):
            lbl.setText(txt)
        
        for btn in [
            self.main_window.radioButton_18, self.main_window.radioButton_19,
            self.main_window.radioButton_20, self.main_window.radioButton_21
        ]:
            btn.setText("选项")
    
    # ========== 练习页面功能 ==========
    
    def _on_time_update(self, single_time: int, total_time: int):
        """时间更新回调"""
        self.main_window.label_15.setText(f"单题用时：{format_time(single_time)}")
        self.main_window.label_16.setText(f"累计用时：{format_time(total_time)}")
    
    def _on_question_changed(self, question: Question, current: int, total: int):
        """题目切换回调"""
        self.main_window.textEdit_9.setText(question.question)
        self.main_window.radioButton_18.setText(question.A)
        self.main_window.radioButton_19.setText(question.B)
        self.main_window.radioButton_20.setText(question.C)
        self.main_window.radioButton_21.setText(question.D)
        
        # 重置选项按钮状态
        self.btn_group_option.setExclusive(False)
        for btn in self.btn_group_option.buttons():
            btn.setChecked(False)
        self.btn_group_option.setExclusive(True)
        
        self.main_window.label_14.setText(
            f"第 {current} 题，共 {total} 题（{current/total*100:.2f}%）"
        )
        self.main_window.label_17.setText(f"第 {question.total + 1} 次刷到该题")
        self.main_window.label_18.setText(f"过往正确率：{format_accuracy(question.correct, question.total)}")
        self.main_window.label_19.setText(f"当前练习模式：{['无尽模式', '套题模式'][self.practice_manager.mode.value]}")
        self.main_window.label_20.setText(f"所属套题：{question.source}")
        self.main_window.label_21.setText("正确答案：-")
        self.main_window.label_23.setText("这题选什么呀？")
        self.main_window.textEdit_10.setText("")
        
        self.main_window.pushButton_19.setText("提交")
    
    def _on_confirm_answer(self):
        """确认答案"""
        # 如果已经显示答案且是最后一题，进入报告页面
        if self.practice_manager.showing_answer and self.practice_manager.is_last_question():
            self._show_report()
            return
        
        # 如果已经显示答案，切换到下一题
        if self.practice_manager.showing_answer:
            self.practice_manager.next_question()
            # 点击下一题后继续计时
            self.practice_manager.resume_timer()
            return
        
        # 获取选择的答案
        answer = self.btn_group_option.checkedId()
        if answer == -1:
            QMessageBox.information(self.main_window, "提示", "你没有选择任何选项!")
            return
        
        # 提交答案
        is_correct, message = self.practice_manager.submit_answer(answer)
        
        # 更新界面
        question = self.practice_manager.get_current_question()
        self.main_window.label_21.setText(f"正确答案：{question.answer}")
        self.main_window.textEdit_10.setText(question.analysis)
        self.main_window.label_23.setText(message)
        
        # 更新按钮文本为"下一题"
        self.main_window.pushButton_19.setText("下一题")
        
        # 提交答案后暂停计时
        self.practice_manager.pause_timer()
        
        # 如果是最后一题，禁用按钮
        if self.practice_manager.is_last_question():
            self.main_window.pushButton_19.setEnabled(False)
    
    def _on_end_practice_and_report(self):
        """结束练习并生成报告"""
        # 显示确认对话框
        reply = QMessageBox.question(
            self.main_window, "提示", "将会按照现有的进度生成报告! 是否继续?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self._show_report()
    
    def _on_practice_back(self):
        """练习页面返回首页（不保存数据）"""
        msg_box = QMessageBox(self.main_window)
        msg_box.setWindowTitle("提示")
        msg_box.setText("退出后将不保存任何数据，是否继续？")
        msg_box.setIcon(QMessageBox.Warning)
        
        # 添加自定义按钮
        confirm_btn = msg_box.addButton("继续", QMessageBox.YesRole)
        cancel_btn = msg_box.addButton("取消", QMessageBox.NoRole)
        
        msg_box.exec()
        
        if msg_box.clickedButton() == confirm_btn:
            # 不保存数据，直接返回首页
            self.practice_manager.stop_practice()
            self.practice_manager.current_index = -1
            self.main_window.textEdit_11.clear()
            self.main_window.setWindowTitle(self.window_title)
            self._on_reload(from_code=True)
            self._init_settings_page()
            self._goto_page(PAGE_HOME)
        # 取消则留在当前页面继续做题
    
    def _on_prev_question(self):
        """上一题"""
        QMessageBox.information(self.main_window, "提示", "此功能还没做")
    
    def _on_chat(self):
        """AI问答 - 打开AI对话窗口"""
        # 获取当前题目
        current_question = self.practice_manager.get_current_question()
        if not current_question:
            QMessageBox.warning(self.main_window, "警告", "请先开始练习并加载题目！")
            return

        # 获取当前使用的AI配置
        chat_ai_name = self.config.chat_ai_name
        ai_config = None
        for cfg in self.config.ai_configs:
            if cfg.name == chat_ai_name:
                ai_config = cfg
                break

        if not ai_config or not ai_config.api_key:
            QMessageBox.warning(self.main_window, "警告", "请先配置AI！\n在设置页面配置AI并选择用于提问的AI。")
            return

        # 获取现有解析
        current_analysis = current_question.analysis if current_question.analysis else ""

        # 打开AI对话窗口
        dialog = AIChatDialog(
            parent=self.main_window,
            question=current_question,
            ai_config=ai_config,
            current_analysis=current_analysis
        )

        # 连接保存信号
        dialog.analysis_saved.connect(lambda new_analysis: self._save_question_analysis(current_question, new_analysis))

        dialog.exec()

    def _save_question_analysis(self, question, new_analysis):
        """保存题目解析"""
        if new_analysis != question.analysis:
            # 更新题目解析
            question.analysis = new_analysis
            # 保存到数据管理器
            self.data_manager.save()
    
    def _show_report(self):
        """显示报告"""
        self.practice_manager.stop_practice()
        self._goto_page(PAGE_REPORT)
        
        # 同步计时器的总时间到统计中，确保报告中的累计用时与做题时显示的一致
        self.practice_manager.statistics.set_total_time(self.practice_manager.total_time)
        
        # 保存进度
        self.practice_manager.save_progress()
        
        # 显示报告
        self.main_window.textEdit_11.clear()
        self.main_window.textEdit_11.setText(self.practice_manager.statistics.generate_report_text())
    
    # ========== 录入页面功能 ==========
    
    def _on_save_question(self):
        """收集错题表单并保存"""
        # 先检查数据
        if self.btn_group_answer.checkedId() == -1:
            QMessageBox.information(self.main_window, "提示", "正确答案未选择！")
            return
        if self.btn_group_classification.checkedId() == -1:
            QMessageBox.information(self.main_window, "提示", "分类未选择！")
            return
        
        # 获取来源
        source = self.main_window.textEdit_6.toPlainText().strip()
        
        # 检查题目来源是否为空
        if not source:
            QMessageBox.information(self.main_window, "提示", "题目来源留空会默认填写\"无\"！")
            self.main_window.textEdit_6.setPlainText("无")
            return
        
        # 创建题目
        question = Question(
            question=self.main_window.textEdit.toPlainText(),
            A=self.main_window.textEdit_2.toPlainText(),
            B=self.main_window.textEdit_3.toPlainText(),
            C=self.main_window.textEdit_4.toPlainText(),
            D=self.main_window.textEdit_5.toPlainText(),
            answer=OPTIONS[self.btn_group_answer.checkedId()],
            classification=self.btn_group_classification.checkedId(),
            source=source,
            analysis=self.main_window.textEdit_7.toPlainText()
        )
        
        app_logger.debug(f"new: {question.to_dict()}")
        
        # 如果是新分类，那么要添加在self.main_window.comboBox中
        papers = self.data_manager.get_papers()
        if source not in papers:
            self.data_manager.papers.append(source)
            self.main_window.comboBox.addItem(source)
        
        # 保存新问题
        try:
            self.data_manager.questions.append(question)
            if self.data_manager.save():
                # 清空并重置表单
                for te in [self.main_window.textEdit, self.main_window.textEdit_2,
                          self.main_window.textEdit_3, self.main_window.textEdit_4,
                          self.main_window.textEdit_5, self.main_window.textEdit_6,
                          self.main_window.textEdit_7]:
                    te.setPlainText("")
                
                self.btn_group_answer.setExclusive(False)
                self.btn_group_classification.setExclusive(False)
                for btn in self.btn_group_answer.buttons():
                    btn.setChecked(False)
                for btn in self.btn_group_classification.buttons():
                    btn.setChecked(False)
                self.btn_group_answer.setExclusive(True)
                self.btn_group_classification.setExclusive(True)
                
                QMessageBox.information(self.main_window, "提示", "题目保存成功!")
                
                # 检查是否有OCR批量导入的下一道题
                if hasattr(self, 'ocr_window') and self.ocr_window:
                    self.ocr_window.check_and_import_next()
        except Exception as e:
            # 保存失败时写入临时文件
            from src.utils import get_timestamp
            import os
            temp_file = os.path.join(TEMP_DIR, f"temp_{get_timestamp()}.txt")
            os.makedirs(TEMP_DIR, exist_ok=True)
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(str([q.to_dict() for q in self.data_manager.questions]))
            QMessageBox.critical(self.main_window, "Error", f"{e}")
    
    # ========== 管理页面功能 ==========
    
    def _on_manage(self):
        """打开管理页面"""
        # 更新套题下拉菜单
        self._check_comboBox()
        self._on_filter()
        self._goto_page(PAGE_MANAGE)
    
    def _show_context_menu(self, position):
        """显示右键菜单"""
        clicked_item = self.main_window.tableWidget.itemAt(position)
        if not clicked_item:
            return

        # 获取选中的行
        selected_rows = set()
        for item in self.main_window.tableWidget.selectedItems():
            selected_rows.add(item.row())

        self.selected_rows = sorted(list(selected_rows))
        if not self.selected_rows:
            self.selected_rows = [clicked_item.row()]
            self.main_window.tableWidget.selectRow(clicked_item.row())

        # 更新"统一设置来源"菜单文本
        self.set_source_action.setText(f"统一设置来源 ({len(self.selected_rows)} 行)")

        self.context_menu.exec(self.main_window.tableWidget.mapToGlobal(position))
    
    def _on_edit_question(self):
        """编辑题目"""
        if not self.selected_rows:
            return
        
        for row in self.selected_rows:
            # 记录原始题目内容作为标识
            original_question = self.result_list[row].question
            for col in range(self.main_window.tableWidget.columnCount()):
                if col == COL_ACCURACY:
                    continue
                item = self.main_window.tableWidget.item(row, col)
                if item:
                    item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
            # 存储原始题目和行号的映射
            self.editable_rows.append((row, original_question))
        
        self.edit_mode = True
    
    def _on_delete_question(self):
        """删除题目"""
        if not self.selected_rows:
            return
        
        reply = QMessageBox.question(
            self.main_window, '确认删除',
            f"确定要删除选中的{len(self.selected_rows)}行吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.No:
            return
        
        # 倒序删除
        for row in sorted(self.selected_rows, reverse=True):
            try:
                # 获取要删除的题目
                question_to_delete = self.result_list[row]
                self.deleted_questions.append([row, question_to_delete])
                # 从 result_list 中移除
                self.result_list.pop(row)
                # 从 data_manager.questions 中移除
                for idx, q in enumerate(self.data_manager.questions):
                    if q.question == question_to_delete.question:
                        self.data_manager.questions.pop(idx)
                        break
            except Exception as e:
                app_logger.error(f"删除失败: {e}")
            self.main_window.tableWidget.removeRow(row)
        
        self._refresh_table(self.result_list)

    def _on_set_source_for_selected(self):
        """为选中的行统一设置来源（题目管理页面）"""
        if not self.selected_rows:
            QMessageBox.warning(self.main_window, "警告", "请先选择要设置的行!")
            return

        # 获取当前第一行的来源作为默认值
        default_source = ""
        first_row = min(self.selected_rows)
        source_item = self.main_window.tableWidget.item(first_row, COL_SOURCE)
        if source_item:
            default_source = source_item.text()

        # 弹出对话框
        from .dialogs import SourceDialog
        dialog = SourceDialog(self.main_window, current_source=default_source, title=f"统一设置来源 ({len(self.selected_rows)} 行)")
        if dialog.exec() != QDialog.Accepted:
            return

        new_source = dialog.get_source()

        # 为所有选中的行设置来源（只修改表格，不保存到数据）
        for row in self.selected_rows:
            self.main_window.tableWidget.setItem(row, COL_SOURCE, QTableWidgetItem(new_source))

        # 标记为编辑模式，需要点击保存才能生效
        if not self.edit_mode:
            self.edit_mode = True
            for row in self.selected_rows:
                original_question = self.result_list[row].question
                if (row, original_question) not in self.editable_rows:
                    self.editable_rows.append((row, original_question))

        QMessageBox.information(self.main_window, "完成", f"已为 {len(self.selected_rows)} 行设置来源!\n请点击保存按钮使修改生效。")

    def _on_save_edits(self):
        """点击题目管理页面的保存按钮后保存修改后的题目"""
        invalid = []  # 记录非法的类型
        
        def check(target, data):
            """检查各种参数是否合法"""
            if target == "answer":
                # 转换为大写后检查
                data_upper = data.upper()
                if data_upper in OPTIONS:
                    return data_upper
                elif "invalid-answer" not in invalid:
                    invalid.append("invalid-answer")
            if target == "classification":
                for idx, classification in enumerate(CLASSIFICATIONS):
                    if data == classification:
                        return idx
                if "invalid-classification" not in invalid:
                    invalid.append("invalid-classification")
            if target == "source":
                # 检查来源是否为空或全是空格
                if not data or data.strip() == "":
                    QMessageBox.information(self.main_window, "提示", "题目来源留空会默认填写\"无\"！")
                    # 找到对应的行并填充"无"
                    for r, orig_q in self.editable_rows:
                        if r == row:
                            self.main_window.tableWidget.item(row, COL_SOURCE).setText("无")
                            break
                    return None
                papers = self.data_manager.get_papers()
                if data in papers:
                    return data
                else:
                    return data
            return None
        
        # 获取新修改的表单并修改 data_manager.questions 和 result_list
        if len(self.editable_rows) != 0:
            for row, original_question in self.editable_rows:
                for idx, question in enumerate(self.data_manager.questions):
                    if question.question == original_question:
                        # 检查来源
                        source = check("source", self.main_window.tableWidget.item(row, COL_SOURCE).text())
                        if source is None:
                            # 来源为空，已提示用户并填充"无"，不保存
                            return
                        
                        modified_question = Question(
                            question=self.main_window.tableWidget.item(row, COL_QUESTION).text(),
                            A=self.main_window.tableWidget.item(row, COL_OPTION_A).text(),
                            B=self.main_window.tableWidget.item(row, COL_OPTION_B).text(),
                            C=self.main_window.tableWidget.item(row, COL_OPTION_C).text(),
                            D=self.main_window.tableWidget.item(row, COL_OPTION_D).text(),
                            answer=check("answer", self.main_window.tableWidget.item(row, COL_ANSWER).text()),
                            classification=check("classification", self.main_window.tableWidget.item(row, COL_CLASSIFICATION).text()),
                            source=source,
                            analysis=self.main_window.tableWidget.item(row, COL_ANALYSIS).text(),
                            total=self.result_list[row].total,
                            correct=self.result_list[row].correct
                        )
                        if len(invalid) != 0:
                            # 构建详细的错误信息
                            error_messages = []
                            if "invalid-answer" in invalid:
                                error_messages.append("答案只能是 A、B、C 或 D")
                            if "invalid-classification" in invalid:
                                error_messages.append("分类无效")
                            error_msg = "\n".join(error_messages)
                            QMessageBox.critical(
                                self.main_window, "Error",
                                f"第 {row + 1} 行存在非法数据，请修改后再保存！\n\n{error_msg}"
                            )
                            return
                        else:
                            self.data_manager.questions[idx] = modified_question
                            self.result_list[row] = modified_question
                            # 将单元格设置为不可编辑
                            for column in range(0, 10):
                                self.main_window.tableWidget.item(row, column).setFlags(
                                    Qt.ItemIsSelectable | Qt.ItemIsEnabled
                                )
                            app_logger.debug(f"保存的题目: {self.data_manager.questions[idx]}")
                        break
        
        # 清空编辑状态
        self.editable_rows.clear()
        self.selected_rows.clear()
        self.deleted_questions.clear()
        self.edit_mode = False
        
        # 保存到文件
        try:
            if self.data_manager.save():
                if len(invalid) == 0:
                    self._check_comboBox()
                    QMessageBox.information(self.main_window, "提示", "题库保存成功!")
        except Exception as e:
            # 保存失败时写入临时文件
            from src.utils import get_timestamp
            import os
            temp_file = os.path.join(TEMP_DIR, f"temp_{get_timestamp()}.txt")
            os.makedirs(TEMP_DIR, exist_ok=True)
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(str([q.to_dict() for q in self.data_manager.questions]))
            QMessageBox.critical(self.main_window, "Error", f"{e}")

    def _on_cell_changed(self, row, column):
        """表格单元格修改时的验证"""
        # 只在编辑模式下验证答案列
        if not self.edit_mode or column != COL_ANSWER:
            return
        
        item = self.main_window.tableWidget.item(row, column)
        if not item:
            return
        
        text = item.text().strip()
        if not text:
            return
        
        # 验证答案只能是A、B、C、D
        if text.upper() not in OPTIONS:
            # 恢复为空或提示用户
            QMessageBox.warning(
                self.main_window, "答案格式错误",
                f"第 {row + 1} 行的答案只能是 A、B、C 或 D！"
            )
            item.setText("")  # 清空非法输入
        else:
            # 自动转换为大写
            item.setText(text.upper())
    
    def _check_comboBox(self):
        """检查删除题目后，是否还存在该分类（实际上是重新添加了一遍分类）"""
        # 清空
        self.data_manager.papers = []
        self.main_window.comboBox.clear()
        # 重新添加
        for question in self.data_manager.questions:
            if question.source not in self.data_manager.papers:
                self.data_manager.papers.append(question.source)
        self.main_window.comboBox.addItem("Any")
        self.main_window.comboBox.addItems(self.data_manager.papers)
    
    def _on_filter(self):
        """筛选题目"""
        # 获取筛选条件
        checkboxes = [
            self.main_window.checkBox_10, self.main_window.checkBox_11,
            self.main_window.checkBox_12, self.main_window.checkBox_13,
            self.main_window.checkBox_14, self.main_window.checkBox_15,
            self.main_window.checkBox_16, self.main_window.checkBox_17,
            self.main_window.checkBox_18
        ]
        
        classifications = [i for i, cb in enumerate(checkboxes) if cb.isChecked()]
        accuracy_id = self.btn_group_accuracy.checkedId()
        paper = self.main_window.comboBox.currentText()
        
        # 转换正确率条件
        max_accuracy = None
        if accuracy_id != -1:
            max_accuracy = [0.25, 0.50, 0.75][accuracy_id]
        
        # 筛选
        self.result_list = self.data_manager.filter_questions(
            classifications=classifications if classifications else None,
            source=paper if paper != "Any" else None,
            max_accuracy=max_accuracy
        )
        
        self._refresh_table(self.result_list)
        self.main_window.setWindowTitle(f"{self.window_title} - {len(self.result_list)}个筛选结果")
    
    def _on_reset_accuracy(self):
        """重置正确率筛选"""
        self.btn_group_accuracy.setExclusive(False)
        for btn in self.btn_group_accuracy.buttons():
            btn.setChecked(False)
        self.btn_group_accuracy.setExclusive(True)
    
    def _on_select_all(self):
        """全选"""
        self.main_window.tableWidget.selectAll()
    
    def _on_export(self):
        """导出"""
        if self.deleted_questions or self.edit_mode:
            QMessageBox.information(self.main_window, "提示", "存在未保存的数据，请先保存再导出！")
            return
        
        # 获取选中的题目
        selected = []
        for range_obj in self.main_window.tableWidget.selectedRanges():
            for row in range(range_obj.topRow(), range_obj.bottomRow() + 1):
                if row < len(self.result_list):
                    selected.append(self.result_list[row])
        
        if not selected:
            QMessageBox.information(self.main_window, "提示", "你还没有选择任何题目!")
            return
        
        # 显示导出对话框
        dialog = ExportDialog(self.main_window, self.config.output_dir)
        if dialog.exec() != ExportDialog.Accepted:
            return
        
        # 执行导出
        path = dialog.get_path()
        filename = dialog.get_filename()
        title = dialog.get_title()
        options = dialog.get_export_options()
        fmt = dialog.get_format()
        
        os.makedirs(path, exist_ok=True)
        
        try:
            if fmt == 0:  # DOCX
                filepath = os.path.join(path, f"{filename}.docx")
                if ExportManager.export_to_docx(selected, filepath, title, options):
                    QMessageBox.information(self.main_window, "提示", f"文档{filename}.docx创建成功！")
            elif fmt == 1:  # PDF
                filepath = os.path.join(path, f"{filename}.pdf")
                if ExportManager.export_to_pdf(selected, filepath, title, options):
                    QMessageBox.information(self.main_window, "提示", f"文档{filename}.pdf创建成功！")
            elif fmt == 2:  # CSV
                filepath = os.path.join(path, f"{filename}.csv")
                if ExportManager.export_to_csv(selected, filepath):
                    QMessageBox.information(self.main_window, "提示", f"文档{filename}.csv创建成功！")
            else:
                QMessageBox.information(self.main_window, "提示", "你还没有选择任何文件类型!")
        except Exception as e:
            QMessageBox.critical(self.main_window, "错误", str(e))
    
    def _on_backup(self):
        """备份数据"""
        backup_path = self.data_manager.backup()
        if backup_path:
            QMessageBox.information(self.main_window, "提示", f"备份成功！\n{backup_path}")
        else:
            QMessageBox.critical(self.main_window, "错误", "备份失败!")

    def _on_import_backup(self):
        """导入备份数据"""
        from PySide6.QtWidgets import QFileDialog

        # 选择备份文件
        file_path, _ = QFileDialog.getOpenFileName(
            self.main_window,
            "选择备份文件",
            BACKUP_DIR,
            "备份文件 (*.zip);;所有文件 (*.*)"
        )

        if not file_path:
            return

        # 确认导入
        reply = QMessageBox.question(
            self.main_window,
            "确认导入",
            "导入备份将覆盖当前所有数据！\n\n" +
            "导入前会自动备份当前数据。\n" +
            "导入成功后需要重启应用才能生效。\n\n" +
            "是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 执行导入
        if self.data_manager.import_backup(file_path):
            QMessageBox.information(
                self.main_window,
                "导入成功",
                "备份导入成功！\n\n" +
                "应用将自动重启以加载新数据。\n" +
                "原数据已自动备份到 restore_points 目录。"
            )
            # 标记需要重启
            self._restart_required = True
            # 关闭主窗口，触发内部重启
            self.main_window.close()
        else:
            QMessageBox.critical(
                self.main_window,
                "导入失败",
                "备份导入失败！\n" +
                "请检查备份文件是否损坏。"
            )

    def _on_reload(self, from_code=False):
        """重新加载数据"""
        if self.data_manager.load():
            self.result_list = self.data_manager.get_all_questions()
            self._refresh_table(self.result_list)
            
            # 恢复删除的题目
            for row, question in self.deleted_questions:
                self.result_list.insert(row, question)
            self.deleted_questions.clear()
            
            if not from_code:
                QMessageBox.information(self.main_window, "提示", "数据已重新加载!")
    
    def _refresh_table(self, questions):
        """刷新表格"""
        table = self.main_window.tableWidget
        table.clearContents()
        table.setRowCount(len(questions))
        
        # 设置列宽
        table.setColumnWidth(COL_QUESTION, self.config.font_size * 100)
        for col in range(COL_OPTION_A, COL_OPTION_D + 1):
            table.setColumnWidth(col, self.config.font_size * 30)
        table.setColumnWidth(COL_ANSWER, self.config.font_size * 15)
        table.setColumnWidth(COL_CLASSIFICATION, self.config.font_size * 15)
        table.setColumnWidth(COL_ACCURACY, self.config.font_size * 15)
        table.setColumnWidth(COL_SOURCE, self.config.font_size * 30)
        table.setColumnWidth(COL_ANALYSIS, self.config.font_size * 50)
        
        # 填充数据
        for row, q in enumerate(questions):
            items = [
                (COL_QUESTION, q.question),
                (COL_OPTION_A, q.A),
                (COL_OPTION_B, q.B),
                (COL_OPTION_C, q.C),
                (COL_OPTION_D, q.D),
                (COL_ANSWER, q.answer),
                (COL_CLASSIFICATION, CLASSIFICATIONS[q.classification] if 0 <= q.classification < len(CLASSIFICATIONS) else "Error"),
                (COL_SOURCE, q.source),
                (COL_ANALYSIS, q.analysis)
            ]
            
            for col, text in items:
                item = QTableWidgetItem(str(text))
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                if col in (COL_ANSWER, COL_CLASSIFICATION):
                    item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row, col, item)

            # 正确率列
            accuracy_item = QTableWidgetItem(format_accuracy(q.correct, q.total))
            accuracy_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            accuracy_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, COL_ACCURACY, accuracy_item)
    
    # ========== 设置页面功能 ==========
    
    def _on_select_dir(self):
        """选择导出文件的目录"""
        # 获取当前配置的输出目录（保持原样，可能是相对路径）
        current_dir = self.main_window.lineEdit.text()
        
        # 如果为空，使用默认输出目录
        if not current_dir:
            from src.utils.constants import OUTPUT_DIR
            current_dir = OUTPUT_DIR
        
        selected = QFileDialog.getExistingDirectory(
            self.main_window, 
            "请选择导出文件的目录",
            current_dir  # 默认打开当前配置的目录
        )
        if selected:
            self.main_window.lineEdit.setText(selected)
    
    def _on_apply_settings(self):
        """应用设置"""
        # 基本设置
        self.config.font_size = self.main_window.spinBox.value()
        self.config.font_name = self.main_window.fontComboBox_2.currentText()
        self.config.output_dir = self.main_window.lineEdit.text()

        # 局域网端口设置 - 验证端口号范围
        try:
            port_text = self.main_window.lineEdit_2.text().strip()
            if not port_text:
                QMessageBox.warning(self.main_window, "警告", "端口号不能为空！")
                return
            
            port = int(port_text)
            if port < 1024:
                QMessageBox.warning(self.main_window, "警告", "端口号不能小于1024（1024以下端口需要管理员权限）！")
                return
            elif port > 65535:
                QMessageBox.warning(self.main_window, "警告", "端口号不能大于65535！")
                return
            else:
                self.config.lan_port = port
        except ValueError:
            QMessageBox.warning(self.main_window, "警告", "端口号必须是数字！")
            return

        # AI设置 - 从表格中读取当前的所有AI配置
        # 列顺序：名称(0)、baseurl(1)、模型(2)、key(3)
        self.config.ai_configs = []
        table = self.main_window.tableWidget_2
        for row in range(table.rowCount()):
            name_item = table.item(row, 0)
            base_url_item = table.item(row, 1)
            model_item = table.item(row, 2)
            if name_item and name_item.text():
                name = name_item.text()
                # 从映射中获取原始API Key，而不是从表格（表格中是脱敏的）
                api_key = self._ai_api_keys.get(name, "")
                self.config.ai_configs.append(AIConfig(
                    name=name,
                    base_url=base_url_item.text() if base_url_item else "",
                    api_key=api_key,
                    model=model_item.text() if model_item else "glm-4v-flash"
                ))

        # AI选择
        self.config.ocr_ai_name = self.main_window.comboBox_3.currentText()
        self.config.chat_ai_name = self.main_window.comboBox_4.currentText()

        if self.config.save():
            # 内部重启应用
            QMessageBox.information(self.main_window, "提示", "设置已保存，应用将重启!")
            # 标记需要重启
            self._restart_required = True
            # 关闭主窗口，触发内部重启
            self.main_window.close()
        else:
            QMessageBox.critical(self.main_window, "错误", "保存设置失败!")
    
    def show(self):
        """显示主窗口"""
        self.main_window.showMaximized()  # 最大化窗口（不是全屏）
