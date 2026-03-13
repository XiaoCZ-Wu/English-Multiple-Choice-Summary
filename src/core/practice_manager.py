"""
练习管理模块
负责练习模式的管理和题目准备
"""

import random
import threading
import time
from typing import List, Optional, Callable
from enum import Enum, auto

from src.models import Question, DataManager
from .statistics import PracticeStatistics


class PracticeMode(Enum):
    """练习模式枚举"""
    ENDLESS = 0  # 无尽模式
    PAPER = 1    # 套题模式


class PracticeManager:
    """练习管理器"""

    def __init__(self, data_manager: DataManager):
        self.data_manager = data_manager
        self.statistics = PracticeStatistics()
        
        # 练习状态
        self.mode: Optional[PracticeMode] = None
        self.prepared_questions: List[Question] = []
        self.current_index: int = -1
        self.single_time: int = 0
        self.total_time: int = 0
        self.showing_answer: bool = False
        
        # 计时线程
        self._timer_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()  # 暂停事件
        
        # 回调函数
        self._on_time_update: Optional[Callable[[int, int], None]] = None
        self._on_question_changed: Optional[Callable[[Question, int, int], None]] = None

    def set_callbacks(
        self,
        on_time_update: Optional[Callable[[int, int], None]] = None,
        on_question_changed: Optional[Callable[[Question, int, int], None]] = None
    ):
        """设置回调函数"""
        self._on_time_update = on_time_update
        self._on_question_changed = on_question_changed

    def start_practice(self, mode: PracticeMode, paper_name: Optional[str] = None) -> bool:
        """开始练习"""
        self.mode = mode
        self.statistics.clear()
        
        # 准备题目
        if mode == PracticeMode.ENDLESS:
            all_questions = self.data_manager.get_all_questions()
            self.prepared_questions = random.sample(all_questions, len(all_questions))
        elif mode == PracticeMode.PAPER and paper_name:
            self.prepared_questions = self.data_manager.get_questions_by_source(paper_name)
        else:
            return False
        
        if not self.prepared_questions:
            return False
        
        # 重置状态
        self.current_index = -1
        self.single_time = 0
        self.total_time = 0
        self.showing_answer = False
        
        # 启动计时
        self._start_timer()
        
        # 切换到第一题
        self.next_question()
        return True

    def _start_timer(self):
        """启动计时线程"""
        self._stop_event.clear()
        self._pause_event.clear()  # 清除暂停状态，开始计时
        self._timer_thread = threading.Thread(target=self._timer_loop, daemon=True)
        self._timer_thread.start()

    def _timer_loop(self):
        """计时循环"""
        while not self._stop_event.is_set() and self.current_index < len(self.prepared_questions):
            # 如果处于暂停状态，等待继续
            if self._pause_event.is_set():
                time.sleep(0.1)
                continue
            time.sleep(1)
            self.single_time += 1
            self.total_time += 1
            if self._on_time_update:
                self._on_time_update(self.single_time, self.total_time)

    def pause_timer(self):
        """暂停计时"""
        self._pause_event.set()

    def resume_timer(self):
        """继续计时"""
        self._pause_event.clear()

    def stop_practice(self):
        """停止练习"""
        # 如果当前题目未提交，扣除当前题目的用时
        if not self.showing_answer and self.single_time > 0:
            self.total_time -= self.single_time
            if self.total_time < 0:
                self.total_time = 0
        self._stop_event.set()
        if self._timer_thread and self._timer_thread.is_alive():
            self._timer_thread.join(timeout=1)

    def next_question(self):
        """切换到下一题"""
        self.current_index += 1
        self.single_time = 0
        self.showing_answer = False
        
        if self._on_question_changed and self.current_index < len(self.prepared_questions):
            question = self.prepared_questions[self.current_index]
            self._on_question_changed(
                question,
                self.current_index + 1,
                len(self.prepared_questions)
            )

    def submit_answer(self, selected_option: int) -> tuple[bool, str]:
        """提交答案，返回(是否正确, 反馈信息)"""
        if self.current_index >= len(self.prepared_questions):
            return False, ""
        
        question = self.prepared_questions[self.current_index]
        options = ["A", "B", "C", "D"]
        selected_answer = options[selected_option]
        is_correct = selected_answer == question.answer.upper()
        
        # 记录统计
        self.statistics.record_question(
            is_correct,
            question.classification,
            self.single_time
        )
        
        # 更新题目数据
        question.record_answer(is_correct)
        
        # 生成反馈信息
        if is_correct:
            message = "正确!\n后面同学!"
        else:
            message = "错误!\n都白讲了!\n来抬头我再说一遍→"
        
        self.showing_answer = True
        return is_correct, message

    def is_finished(self) -> bool:
        """检查是否完成所有题目"""
        return self.current_index >= len(self.prepared_questions)

    def is_last_question(self) -> bool:
        """检查是否是最后一题"""
        return self.current_index == len(self.prepared_questions) - 1

    def get_current_question(self) -> Optional[Question]:
        """获取当前题目"""
        if 0 <= self.current_index < len(self.prepared_questions):
            return self.prepared_questions[self.current_index]
        return None

    def get_progress(self) -> float:
        """获取进度百分比"""
        if not self.prepared_questions:
            return 0.0
        return (self.current_index + 1) / len(self.prepared_questions) * 100

    def save_progress(self) -> bool:
        """保存练习进度（保存题目数据）"""
        return self.data_manager.save()
