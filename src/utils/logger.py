"""
日志模块
提供统一的日志记录功能
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path


class AppLogger:
    """应用日志管理器"""
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if AppLogger._initialized:
            return
        
        self.logger = logging.getLogger("EnglishQuizApp")
        self.logger.setLevel(logging.DEBUG)
        self.handlers = []
        AppLogger._initialized = True
    
    def setup(self, log_dir: str = None, log_to_file: bool = True, log_to_console: bool = True):
        """
        设置日志
        
        Args:
            log_dir: 日志文件目录，默认为程序目录下的 logs
            log_to_file: 是否输出到文件
            log_to_console: 是否输出到控制台
        """
        # 清除现有处理器
        self.clear_handlers()
        
        # 设置格式
        formatter = logging.Formatter(
            '[%(asctime)s][%(levelname)s][%(name)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台输出
        if log_to_console:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.DEBUG)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
            self.handlers.append(console_handler)
        
        # 文件输出
        if log_to_file:
            if log_dir is None:
                # 获取程序根目录
                if getattr(sys, 'frozen', False):
                    base_dir = os.path.dirname(sys.executable)
                else:
                    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                log_dir = os.path.join(base_dir, 'log')
            
            os.makedirs(log_dir, exist_ok=True)
            
            # 按日期命名日志文件 (使用 YYMMDD_HHMMSS 格式，与 backup 一致)
            log_file = os.path.join(log_dir, f"app_{datetime.now().strftime('%y%m%d_%H%M%S')}.log")
            
            file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='a')
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)
            self.handlers.append(file_handler)
            
            self.logger.info(f"日志文件: {log_file}")
    
    def clear_handlers(self):
        """清除所有处理器"""
        for handler in self.handlers:
            self.logger.removeHandler(handler)
            handler.close()
        self.handlers.clear()
    
    def debug(self, msg: str):
        """调试日志"""
        self.logger.debug(msg)
    
    def info(self, msg: str):
        """信息日志"""
        self.logger.info(msg)
    
    def warning(self, msg: str):
        """警告日志"""
        self.logger.warning(msg)
    
    def error(self, msg: str):
        """错误日志"""
        self.logger.error(msg)
    
    def critical(self, msg: str):
        """严重错误日志"""
        self.logger.critical(msg)
    
    def exception(self, msg: str):
        """异常日志（自动包含堆栈信息）"""
        self.logger.exception(msg)


# 全局日志实例
app_logger = AppLogger()


def setup_logging(log_dir: str = None, log_to_file: bool = True, log_to_console: bool = True):
    """
    初始化日志系统
    
    使用示例:
        from src.utils.logger import setup_logging, app_logger
        setup_logging()
        app_logger.info("应用启动")
    """
    app_logger.setup(log_dir, log_to_file, log_to_console)


def get_logger():
    """获取日志实例"""
    return app_logger
