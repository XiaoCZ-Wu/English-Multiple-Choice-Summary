"""
英语单选题练习系统 - 程序入口

重构后的项目入口文件
原文件: main-pyside6-beta-v1.2.py

使用方法:
    python main.py

项目结构:
    src/            - 源代码目录
    ├── models/     - 数据模型
    ├── core/       - 核心业务逻辑
    ├── ui/         - 用户界面
    └── utils/      - 工具函数
"""

import sys
import os
from pathlib import Path

# 获取项目根目录
root_dir = Path(__file__).parent.absolute()

# 添加项目根目录和src目录到Python路径
sys.path.insert(0, str(root_dir))
sys.path.insert(0, str(root_dir / 'src'))

# 导入并运行主程序
from src.main import main

if __name__ == '__main__':
    main()
