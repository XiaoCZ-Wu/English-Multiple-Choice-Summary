"""
常量定义模块
存放项目中使用的所有常量
"""

import os
import sys

# 题目分类
CLASSIFICATIONS = [
    "交际用语",
    "词义辨析",
    "时态",
    "非谓语动词",
    "定语从句",
    "状语从句",
    "情态动词",
    "名词性从句",
    "代词"
]

# 分类数量
CLASSIFICATION_COUNT = len(CLASSIFICATIONS)

# 选项字母
OPTIONS = ["A", "B", "C", "D"]

# 练习模式
PRACTICE_MODE_ENDLESS = 0
PRACTICE_MODE_PAPER = 1

# 表格列索引
COL_QUESTION = 0
COL_OPTION_A = 1
COL_OPTION_B = 2
COL_OPTION_C = 3
COL_OPTION_D = 4
COL_ANSWER = 5
COL_CLASSIFICATION = 6
COL_ACCURACY = 7
COL_SOURCE = 8
COL_ANALYSIS = 9

# 页面索引
PAGE_HOME = 0
PAGE_CREATE = 1
PAGE_MANAGE = 2
PAGE_SETTINGS = 3
PAGE_PRACTICE = 4
PAGE_REPORT = 5


def get_base_dir():
    """获取基础目录（支持正常环境和 PyInstaller 打包环境）"""
    if getattr(sys, 'frozen', False):
        # 运行在打包后的 exe 中
        return sys._MEIPASS
    else:
        # 运行在普通 Python 环境中
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    base_dir = get_base_dir()
    return os.path.join(base_dir, relative_path)


# 获取 src 目录
SRC_DIR = get_resource_path("src")
ROOT_DIR = get_base_dir()

# 文件路径（基于 src 目录的绝对路径）
DATA_DIR = get_resource_path("src/data")
QUESTIONS_FILE = get_resource_path("src/data/questions.json")
CONFIG_FILE = get_resource_path("src/data/config.json")
BACKUP_DIR = get_resource_path("backup")
OUTPUT_DIR = get_resource_path("output")
TEMP_DIR = get_resource_path("temp")
OCR_TEMP_DIR = get_resource_path("temp/ocr")
UI_DIR = get_resource_path("src/ui_dir")
UI_FILE = get_resource_path("src/ui_dir/main.ui")

# 默认配置
DEFAULT_CONFIG = {
    "font-name": "Microsoft YaHei",
    "font-size": 11,
    "output-dir": OUTPUT_DIR
}
