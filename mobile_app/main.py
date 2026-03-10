"""
英语单选 - 手机端APP
使用Kivy框架开发，支持Android系统
功能：
1. 从电脑端拉取可用AI列表
2. 上传图像进行OCR识别
3. 编辑识别结果并传回电脑端
"""

import os
import sys
import json
import requests
from io import BytesIO
from typing import List, Dict, Optional

# Kivy imports
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.image import Image
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.properties import StringProperty, ListProperty, ObjectProperty
from kivy.core.window import Window
from kivy.clock import Clock
from kivy.utils import platform

# 设置窗口大小（开发时模拟手机屏幕）
Window.size = (400, 800)


class ServerConfigScreen(Screen):
    """服务器配置页面"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_ui()
    
    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=20, spacing=10)
        
        # 标题
        title = Label(
            text='服务器配置',
            font_size='24sp',
            size_hint_y=None,
            height=50
        )
        layout.add_widget(title)
        
        # 说明文字
        info = Label(
            text='请输入电脑端的IP地址和端口\n确保手机和电脑在同一WiFi下',
            font_size='14sp',
            size_hint_y=None,
            height=60
        )
        layout.add_widget(info)
        
        # IP地址输入
        layout.add_widget(Label(text='电脑IP地址:', size_hint_y=None, height=30))
        self.ip_input = TextInput(
            hint_text='例如: 192.168.1.100',
            multiline=False,
            size_hint_y=None,
            height=50
        )
        layout.add_widget(self.ip_input)
        
        # 端口输入
        layout.add_widget(Label(text='端口:', size_hint_y=None, height=30))
        self.port_input = TextInput(
            hint_text='例如: 8080',
            multiline=False,
            input_filter='int',
            size_hint_y=None,
            height=50
        )
        self.port_input.text = '8080'
        layout.add_widget(self.port_input)
        
        # 连接按钮
        connect_btn = Button(
            text='连接电脑',
            size_hint_y=None,
            height=60,
            background_color=(0.2, 0.6, 1, 1)
        )
        connect_btn.bind(on_press=self.on_connect)
        layout.add_widget(connect_btn)
        
        # 状态标签
        self.status_label = Label(
            text='',
            font_size='14sp',
            size_hint_y=None,
            height=40,
            color=(1, 0.5, 0, 1)
        )
        layout.add_widget(self.status_label)
        
        # 填充空间
        layout.add_widget(Label())
        
        self.add_widget(layout)
    
    def on_connect(self, instance):
        """连接电脑"""
        ip = self.ip_input.text.strip()
        port = self.port_input.text.strip()
        
        if not ip:
            self.status_label.text = '请输入IP地址'
            return
        if not port:
            self.status_label.text = '请输入端口'
            return
        
        # 保存配置到app
        app = App.get_running_app()
        app.server_ip = ip
        app.server_port = int(port)
        app.base_url = f'http://{ip}:{port}'
        
        # 测试连接
        self.status_label.text = '正在连接...'
        Clock.schedule_once(self.test_connection, 0.1)
    
    def test_connection(self, dt):
        """测试连接"""
        app = App.get_running_app()
        try:
            response = requests.get(f'{app.base_url}/api/ping', timeout=5)
            if response.status_code == 200:
                self.status_label.text = '连接成功!'
                self.status_label.color = (0, 1, 0, 1)
                # 跳转到主页面
                self.manager.current = 'main'
            else:
                self.status_label.text = f'连接失败: {response.status_code}'
                self.status_label.color = (1, 0, 0, 1)
        except Exception as e:
            self.status_label.text = f'连接失败: {str(e)}'
            self.status_label.color = (1, 0, 0, 1)


class MainScreen(Screen):
    """主页面"""
    
    ai_list = ListProperty([])
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_ui()
    
    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # 标题栏
        header = BoxLayout(size_hint_y=None, height=50)
        title = Label(text='英语单选 - 手机端', font_size='20sp')
        header.add_widget(title)
        layout.add_widget(header)
        
        # AI列表区域
        ai_section = BoxLayout(orientation='vertical', size_hint_y=None, height=150)
        ai_section.add_widget(Label(text='可用AI列表:', size_hint_y=None, height=30))
        
        self.ai_label = Label(
            text='点击刷新获取AI列表',
            font_size='12sp'
        )
        ai_section.add_widget(self.ai_label)
        
        refresh_btn = Button(
            text='刷新AI列表',
            size_hint_y=None,
            height=50
        )
        refresh_btn.bind(on_press=self.refresh_ai_list)
        ai_section.add_widget(refresh_btn)
        
        layout.add_widget(ai_section)
        
        # 功能按钮区域
        func_layout = GridLayout(cols=1, spacing=10, size_hint_y=None, height=200)
        
        ocr_btn = Button(
            text='拍照/选择图片进行OCR识别',
            font_size='16sp',
            background_color=(0.2, 0.7, 0.3, 1)
        )
        ocr_btn.bind(on_press=self.go_to_ocr)
        func_layout.add_widget(ocr_btn)
        
        pending_btn = Button(
            text='查看待导入题目',
            font_size='16sp'
        )
        pending_btn.bind(on_press=self.go_to_pending)
        func_layout.add_widget(pending_btn)
        
        layout.add_widget(func_layout)
        
        # 状态信息
        self.status_label = Label(
            text='',
            font_size='12sp',
            size_hint_y=None,
            height=40
        )
        layout.add_widget(self.status_label)
        
        # 填充空间
        layout.add_widget(Label())
        
        self.add_widget(layout)
    
    def refresh_ai_list(self, instance):
        """刷新AI列表"""
        app = App.get_running_app()
        try:
            response = requests.get(f'{app.base_url}/api/ai-list', timeout=10)
            if response.status_code == 200:
                data = response.json()
                self.ai_list = data.get('ai_list', [])
                if self.ai_list:
                    ai_names = [ai['name'] for ai in self.ai_list]
                    self.ai_label.text = f'可用AI: {", ".join(ai_names)}'
                    app.ai_list = self.ai_list
                else:
                    self.ai_label.text = '暂无可用AI配置'
            else:
                self.ai_label.text = f'获取失败: {response.status_code}'
        except Exception as e:
            self.ai_label.text = f'获取失败: {str(e)}'
    
    def go_to_ocr(self, instance):
        """跳转到OCR页面"""
        self.manager.current = 'ocr'
    
    def go_to_pending(self, instance):
        """跳转到待导入页面"""
        self.manager.current = 'pending'


class OCRScreen(Screen):
    """OCR识别页面"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.selected_image_path = None
        self.build_ui()
    
    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # 标题
        header = BoxLayout(size_hint_y=None, height=50)
        back_btn = Button(text='返回', size_hint_x=None, width=80)
        back_btn.bind(on_press=self.go_back)
        header.add_widget(back_btn)
        header.add_widget(Label(text='OCR识别', font_size='20sp'))
        layout.add_widget(header)
        
        # 题号范围输入
        layout.add_widget(Label(text='题号范围 (可选):', size_hint_y=None, height=30))
        self.question_range_input = TextInput(
            hint_text='例如: 1-5,8 表示第1-5题和第8题',
            multiline=False,
            size_hint_y=None,
            height=50
        )
        layout.add_widget(self.question_range_input)
        
        # 选择图片按钮
        select_btn = Button(
            text='选择图片',
            size_hint_y=None,
            height=60,
            background_color=(0.2, 0.6, 1, 1)
        )
        select_btn.bind(on_press=self.select_image)
        layout.add_widget(select_btn)
        
        # 图片预览
        self.image_preview = Image(
            source='',
            size_hint_y=None,
            height=200
        )
        layout.add_widget(self.image_preview)
        
        # 开始识别按钮
        self.ocr_btn = Button(
            text='开始OCR识别',
            size_hint_y=None,
            height=60,
            background_color=(0.2, 0.7, 0.3, 1),
            disabled=True
        )
        self.ocr_btn.bind(on_press=self.start_ocr)
        layout.add_widget(self.ocr_btn)
        
        # 状态标签
        self.status_label = Label(
            text='请先选择图片',
            font_size='14sp',
            size_hint_y=None,
            height=40
        )
        layout.add_widget(self.status_label)
        
        # 填充空间
        layout.add_widget(Label())
        
        self.add_widget(layout)
    
    def go_back(self, instance):
        """返回主页面"""
        self.manager.current = 'main'
    
    def select_image(self, instance):
        """选择图片"""
        # 在Android上使用文件选择器
        if platform == 'android':
            from android.permissions import request_permissions, Permission
            request_permissions([Permission.READ_EXTERNAL_STORAGE])
        
        # 创建文件选择弹窗
        content = BoxLayout(orientation='vertical')
        filechooser = FileChooserListView(
            path=os.path.expanduser('~'),
            filters=['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp']
        )
        content.add_widget(filechooser)
        
        btn_layout = BoxLayout(size_hint_y=None, height=50)
        select_btn = Button(text='选择')
        cancel_btn = Button(text='取消')
        btn_layout.add_widget(select_btn)
        btn_layout.add_widget(cancel_btn)
        content.add_widget(btn_layout)
        
        popup = Popup(title='选择图片', content=content, size_hint=(0.9, 0.9))
        
        def on_select(instance):
            if filechooser.selection:
                self.selected_image_path = filechooser.selection[0]
                self.image_preview.source = self.selected_image_path
                self.ocr_btn.disabled = False
                self.status_label.text = '已选择图片，点击开始识别'
                popup.dismiss()
        
        def on_cancel(instance):
            popup.dismiss()
        
        select_btn.bind(on_press=on_select)
        cancel_btn.bind(on_press=on_cancel)
        popup.open()
    
    def start_ocr(self, instance):
        """开始OCR识别"""
        if not self.selected_image_path:
            self.status_label.text = '请先选择图片'
            return
        
        app = App.get_running_app()
        self.status_label.text = '正在上传图片...'
        
        # 在后台线程中执行上传
        Clock.schedule_once(self.do_ocr, 0.1)
    
    def do_ocr(self, dt):
        """执行OCR"""
        app = App.get_running_app()
        
        try:
            # 读取图片
            with open(self.selected_image_path, 'rb') as f:
                image_data = f.read()
            
            # 上传图片
            files = {'image': ('image.jpg', image_data, 'image/jpeg')}
            data = {
                'question_range': self.question_range_input.text.strip()
            }
            
            response = requests.post(
                f'{app.base_url}/api/ocr',
                files=files,
                data=data,
                timeout=120
            )
            
            if response.status_code == 200:
                result = response.json()
                questions = result.get('questions', [])
                if questions:
                    # 保存到待导入列表
                    app.pending_questions = questions
                    self.status_label.text = f'识别成功，共{len(questions)}道题'
                    # 跳转到待导入页面
                    Clock.schedule_once(lambda dt: setattr(self.manager, 'current', 'pending'), 1)
                else:
                    self.status_label.text = '未识别到题目'
            else:
                self.status_label.text = f'识别失败: {response.status_code}'
        except Exception as e:
            self.status_label.text = f'识别失败: {str(e)}'


class PendingScreen(Screen):
    """待导入题目页面"""
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.question_widgets = []
        self.build_ui()
    
    def build_ui(self):
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        # 标题
        header = BoxLayout(size_hint_y=None, height=50)
        back_btn = Button(text='返回', size_hint_x=None, width=80)
        back_btn.bind(on_press=self.go_back)
        header.add_widget(back_btn)
        header.add_widget(Label(text='待导入题目', font_size='20sp'))
        main_layout.add_widget(header)
        
        # 滚动区域
        scroll = ScrollView()
        self.questions_layout = BoxLayout(
            orientation='vertical',
            spacing=10,
            size_hint_y=None
        )
        self.questions_layout.bind(
            minimum_height=self.questions_layout.setter('height')
        )
        scroll.add_widget(self.questions_layout)
        main_layout.add_widget(scroll)
        
        # 底部按钮
        btn_layout = BoxLayout(size_hint_y=None, height=60, spacing=10)
        
        add_btn = Button(
            text='添加题目',
            background_color=(0.2, 0.6, 1, 1)
        )
        add_btn.bind(on_press=self.add_empty_question)
        btn_layout.add_widget(add_btn)
        
        send_btn = Button(
            text='发送到电脑',
            background_color=(0.2, 0.7, 0.3, 1)
        )
        send_btn.bind(on_press=self.send_to_computer)
        btn_layout.add_widget(send_btn)
        
        main_layout.add_widget(btn_layout)
        
        self.add_widget(main_layout)
    
    def on_enter(self):
        """进入页面时刷新题目列表"""
        self.refresh_questions()
    
    def refresh_questions(self):
        """刷新题目列表"""
        # 清除现有内容
        self.questions_layout.clear_widgets()
        self.question_widgets = []
        
        app = App.get_running_app()
        questions = getattr(app, 'pending_questions', [])
        
        for i, q in enumerate(questions):
            self.add_question_widget(i, q)
    
    def add_question_widget(self, index, question):
        """添加题目编辑控件"""
        box = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            height=400,
            padding=10
        )
        
        # 题目标题
        header = BoxLayout(size_hint_y=None, height=30)
        header.add_widget(Label(text=f'题目 {index + 1}', font_size='16sp'))
        delete_btn = Button(text='删除', size_hint_x=None, width=80)
        delete_btn.bind(on_press=lambda x, idx=index: self.delete_question(idx))
        header.add_widget(delete_btn)
        box.add_widget(header)
        
        # 问题内容
        box.add_widget(Label(text='问题:', size_hint_y=None, height=25))
        question_input = TextInput(
            text=question.get('question', ''),
            multiline=True,
            size_hint_y=None,
            height=60
        )
        box.add_widget(question_input)
        
        # 选项
        options = {}
        for opt in ['A', 'B', 'C', 'D']:
            box.add_widget(Label(text=f'选项{opt}:', size_hint_y=None, height=25))
            opt_input = TextInput(
                text=question.get(opt, ''),
                multiline=False,
                size_hint_y=None,
                height=40
            )
            box.add_widget(opt_input)
            options[opt] = opt_input
        
        # 答案
        box.add_widget(Label(text='答案:', size_hint_y=None, height=25))
        answer_input = TextInput(
            text=question.get('answer', ''),
            multiline=False,
            size_hint_y=None,
            height=40
        )
        box.add_widget(answer_input)
        
        # 分类
        box.add_widget(Label(text='分类:', size_hint_y=None, height=25))
        classification_input = TextInput(
            text=question.get('classification', ''),
            multiline=False,
            size_hint_y=None,
            height=40
        )
        box.add_widget(classification_input)
        
        # 来源
        box.add_widget(Label(text='来源:', size_hint_y=None, height=25))
        source_input = TextInput(
            text=question.get('source', ''),
            multiline=False,
            size_hint_y=None,
            height=40
        )
        box.add_widget(source_input)
        
        # 解析
        box.add_widget(Label(text='解析:', size_hint_y=None, height=25))
        analysis_input = TextInput(
            text=question.get('analysis', ''),
            multiline=True,
            size_hint_y=None,
            height=60
        )
        box.add_widget(analysis_input)
        
        # 保存引用
        self.question_widgets.append({
            'question': question_input,
            'A': options['A'],
            'B': options['B'],
            'C': options['C'],
            'D': options['D'],
            'answer': answer_input,
            'classification': classification_input,
            'source': source_input,
            'analysis': analysis_input
        })
        
        # 分隔线
        separator = Label(
            text='─' * 40,
            size_hint_y=None,
            height=20,
            color=(0.5, 0.5, 0.5, 1)
        )
        box.add_widget(separator)
        
        self.questions_layout.add_widget(box)
    
    def add_empty_question(self, instance):
        """添加空白题目"""
        empty_question = {
            'question': '',
            'A': '',
            'B': '',
            'C': '',
            'D': '',
            'answer': '',
            'classification': '',
            'source': '',
            'analysis': ''
        }
        app = App.get_running_app()
        if not hasattr(app, 'pending_questions'):
            app.pending_questions = []
        app.pending_questions.append(empty_question)
        self.add_question_widget(len(app.pending_questions) - 1, empty_question)
    
    def delete_question(self, index):
        """删除题目"""
        app = App.get_running_app()
        if hasattr(app, 'pending_questions') and 0 <= index < len(app.pending_questions):
            app.pending_questions.pop(index)
            self.refresh_questions()
    
    def collect_questions(self):
        """收集所有题目数据"""
        questions = []
        for widgets in self.question_widgets:
            question = {
                'question': widgets['question'].text,
                'A': widgets['A'].text,
                'B': widgets['B'].text,
                'C': widgets['C'].text,
                'D': widgets['D'].text,
                'answer': widgets['answer'].text.upper(),
                'classification': widgets['classification'].text,
                'source': widgets['source'].text,
                'analysis': widgets['analysis'].text
            }
            questions.append(question)
        return questions
    
    def send_to_computer(self, instance):
        """发送到电脑"""
        questions = self.collect_questions()
        
        if not questions:
            self.show_popup('提示', '没有题目可发送')
            return
        
        # 验证题目
        invalid = []
        for i, q in enumerate(questions):
            if not q['question'] or not q['answer']:
                invalid.append(str(i + 1))
        
        if invalid:
            self.show_popup('验证失败', f'以下题目不完整: {", ".join(invalid)}')
            return
        
        # 发送到电脑
        app = App.get_running_app()
        try:
            response = requests.post(
                f'{app.base_url}/api/import',
                json={'questions': questions},
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                imported = result.get('imported', 0)
                self.show_popup('成功', f'成功导入 {imported} 道题目')
                # 清空待导入列表
                app.pending_questions = []
                self.refresh_questions()
            else:
                self.show_popup('失败', f'导入失败: {response.status_code}')
        except Exception as e:
            self.show_popup('错误', f'发送失败: {str(e)}')
    
    def show_popup(self, title, message):
        """显示弹窗"""
        popup = Popup(
            title=title,
            content=Label(text=message),
            size_hint=(0.8, 0.4)
        )
        popup.open()
    
    def go_back(self, instance):
        """返回主页面"""
        self.manager.current = 'main'


class EnglishQuizApp(App):
    """英语单选APP"""
    
    server_ip = StringProperty('')
    server_port = 8080
    base_url = StringProperty('')
    ai_list = ListProperty([])
    pending_questions = ListProperty([])
    
    def build(self):
        # 创建屏幕管理器
        sm = ScreenManager()
        sm.add_widget(ServerConfigScreen(name='config'))
        sm.add_widget(MainScreen(name='main'))
        sm.add_widget(OCRScreen(name='ocr'))
        sm.add_widget(PendingScreen(name='pending'))
        return sm


if __name__ == '__main__':
    EnglishQuizApp().run()
