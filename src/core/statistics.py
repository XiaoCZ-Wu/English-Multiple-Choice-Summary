"""
统计功能模块
负责练习统计和报告生成
"""

from dataclasses import dataclass, field
from typing import List, Dict
from collections import defaultdict

from src.utils import CLASSIFICATIONS, format_time_with_unit


@dataclass
class QuestionStat:
    """单题统计"""
    is_correct: bool
    classification: int
    time_spent: int  # 秒


@dataclass
class ClassificationStat:
    """分类统计"""
    name: str
    correct: int = 0
    total: int = 0

    @property
    def accuracy(self) -> float:
        if self.total == 0:
            return 0.0
        return self.correct / self.total


class PracticeStatistics:
    """练习统计管理器"""

    def __init__(self):
        self.question_stats: List[QuestionStat] = []
        self.total_time: int = 0
        self._classification_stats: Dict[int, ClassificationStat] = {}
        self._init_classification_stats()

    def _init_classification_stats(self):
        """初始化分类统计"""
        for idx, name in enumerate(CLASSIFICATIONS):
            self._classification_stats[idx] = ClassificationStat(name=name)

    def record_question(self, is_correct: bool, classification: int, time_spent: int):
        """记录一道题的统计"""
        stat = QuestionStat(is_correct, classification, time_spent)
        self.question_stats.append(stat)
        self.total_time += time_spent
        
        # 更新分类统计
        if classification in self._classification_stats:
            self._classification_stats[classification].total += 1
            if is_correct:
                self._classification_stats[classification].correct += 1

    def set_total_time(self, total_time: int):
        """设置总用时（用于同步计时器的总时间）"""
        self.total_time = total_time

    def get_total_correct(self) -> int:
        """获取总正确数"""
        return sum(1 for s in self.question_stats if s.is_correct)

    def get_total_questions(self) -> int:
        """获取总题数"""
        return len(self.question_stats)

    def get_overall_accuracy(self) -> float:
        """获取总体正确率"""
        total = self.get_total_questions()
        if total == 0:
            return 0.0
        return self.get_total_correct() / total

    def get_average_time(self) -> int:
        """获取平均用时（秒）"""
        total = self.get_total_questions()
        if total == 0:
            return 0
        return self.total_time // total

    def get_classification_stats(self) -> List[ClassificationStat]:
        """获取所有分类统计"""
        return list(self._classification_stats.values())

    def generate_report_text(self) -> str:
        """生成报告文本"""
        lines = []
        
        # 总体统计
        lines.append(f"{'累计用时'.center(20)}{format_time_with_unit(self.total_time)}")
        lines.append(f"{'平均用时'.center(20)}{format_time_with_unit(self.get_average_time())}")
        lines.append(f"{'题目总数'.center(20)}{self.get_total_questions()}")
        lines.append("=" * 80)
        
        # 分类统计表格
        from tabulate import tabulate
        table_data = []
        for stat in self.get_classification_stats():
            table_data.append([
                stat.name,
                stat.correct,
                stat.total,
                f"{stat.accuracy * 100:.2f}%"
            ])
        
        lines.append(tabulate(
            table_data,
            headers=["题目类型", "答对", "共计", "正确率"],
            tablefmt="grid"
        ))
        
        return "\n".join(lines)

    def clear(self):
        """清空统计"""
        self.question_stats.clear()
        self.total_time = 0
        self._classification_stats.clear()
        self._init_classification_stats()
