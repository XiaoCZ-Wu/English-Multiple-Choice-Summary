"""
题目数据模型
定义题目数据结构和相关操作
"""

from dataclasses import dataclass, field, asdict
from typing import Optional


@dataclass
class Question:
    """题目数据类"""
    question: str
    A: str
    B: str
    C: str
    D: str
    answer: str
    classification: int
    source: str
    analysis: str
    total: int = 0
    correct: int = 0

    def to_dict(self) -> dict:
        """转换为字典"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "Question":
        """从字典创建实例"""
        return cls(**data)

    def get_accuracy(self) -> float:
        """获取正确率"""
        if self.total == 0:
            return 0.0
        return self.correct / self.total

    def record_answer(self, is_correct: bool):
        """记录一次答题结果"""
        self.total += 1
        if is_correct:
            self.correct += 1

    def is_valid(self) -> bool:
        """验证题目数据是否有效"""
        from src.utils import validate_answer, validate_classification
        return (
            self.question and
            self.A and self.B and self.C and self.D and
            validate_answer(self.answer) and
            validate_classification(self.classification)
        )
