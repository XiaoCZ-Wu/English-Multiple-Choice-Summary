"""
数据模型模块
"""

from .question import Question
from .config import AppConfig, AIConfig
from .data_manager import DataManager

__all__ = [
    'Question',
    'AppConfig',
    'AIConfig',
    'DataManager',
]
