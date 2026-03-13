"""
程序入口
"""

import os
import sys

# 添加项目根目录到Python路径
# 支持正常环境和 PyInstaller 打包环境
if getattr(sys, 'frozen', False):
    # 运行在打包后的 exe 中
    base_dir = sys._MEIPASS
else:
    # 运行在普通 Python 环境中
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, base_dir)

# 初始化日志系统
from src.utils.logger import setup_logging, app_logger
setup_logging()

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

from src.ui import MainWindow


def main():
    """主函数 - 支持内部重启"""
    # 创建全局 QApplication
    app = QApplication(sys.argv)
    
    # 设置应用程序属性，确保正确清理
    app.setQuitOnLastWindowClosed(True)
    
    while True:
        # 创建并显示主窗口
        window = MainWindow()
        window.show()
        
        # 运行应用，等待退出
        exit_code = app.exec()
        
        # 检查是否需要重启
        if hasattr(window, '_restart_required') and window._restart_required:
            app_logger.info("应用正在内部重启...")
            # 删除窗口实例
            del window
            # 处理剩余事件，确保窗口完全关闭
            app.processEvents()
            # 继续循环，创建新的窗口实例
            continue
        else:
            # 正常退出
            break
    
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
