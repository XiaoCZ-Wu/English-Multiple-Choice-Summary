"""
辅助函数模块
存放各种工具函数
"""

from datetime import datetime
from typing import Optional


def format_time(seconds: int) -> str:
    """将秒数格式化为 MM: SS 格式"""
    minutes = int(seconds / 60)
    secs = seconds % 60
    return f"{str(minutes).zfill(2)}: {str(secs).zfill(2)}"


def format_time_with_unit(seconds: int) -> str:
    """将秒数格式化为 XmXs 格式"""
    return f"{int(seconds / 60)}m{seconds % 60}s"


def calculate_accuracy(correct: int, total: int) -> float:
    """计算正确率，处理除零错误"""
    try:
        return correct / total if total > 0 else 0.0
    except (ZeroDivisionError, TypeError):
        return 0.0


def format_accuracy(correct: int, total: int) -> str:
    """格式化正确率为百分比字符串"""
    accuracy = calculate_accuracy(correct, total)
    return f"{accuracy * 100:.2f}%"


def get_classification_name(classification_id: int) -> str:
    """根据分类ID获取分类名称"""
    from .constants import CLASSIFICATIONS
    if 0 <= classification_id < len(CLASSIFICATIONS):
        return CLASSIFICATIONS[classification_id]
    return "Error"


def get_classification_id(name: str) -> Optional[int]:
    """根据分类名称获取分类ID"""
    from .constants import CLASSIFICATIONS
    for idx, classification in enumerate(CLASSIFICATIONS):
        if classification == name:
            return idx
    return None


def get_timestamp() -> str:
    """获取当前时间戳字符串"""
    return datetime.now().strftime('%y%m%d_%H%M%S')


def validate_answer(answer: str) -> bool:
    """验证答案是否合法（A/B/C/D），不区分大小写"""
    from .constants import OPTIONS
    return answer.upper() in OPTIONS


def validate_classification(classification_id: int) -> bool:
    """验证分类ID是否合法"""
    from .constants import CLASSIFICATION_COUNT
    return 0 <= classification_id < CLASSIFICATION_COUNT
