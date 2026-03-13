"""
截图工具
支持全屏截图和区域截图
"""

import os
from typing import Optional, Callable

from PySide6.QtCore import Qt, QRect, QPoint
from PySide6.QtGui import QPixmap, QPainter, QPen, QColor, QKeyEvent, QMouseEvent, QPaintEvent
from PySide6.QtWidgets import QWidget, QApplication

from src.utils import app_logger


class ScreenshotWidget(QWidget):
    """截图工具窗口"""
    
    def __init__(self, on_screenshot_taken: Optional[Callable[[QPixmap], None]] = None, parent=None):
        super().__init__(parent)
        
        self.on_screenshot_taken = on_screenshot_taken
        self.screenshot: Optional[QPixmap] = None
        self.start_pos: Optional[QPoint] = None
        self.end_pos: Optional[QPoint] = None
        self.is_drawing = False
        
        # 设置窗口属性（窗口标志在take_screenshot中统一设置）
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # 获取屏幕截图
        self._capture_screen()
    
    def _capture_screen(self):
        """捕获全屏（支持多显示器）"""
        # 获取所有屏幕
        screens = QApplication.screens()
        if not screens:
            return
        
        # 计算所有屏幕的并集
        total_rect = screens[0].geometry()
        for screen in screens[1:]:
            total_rect = total_rect.united(screen.geometry())
        
        # 创建一个足够大的pixmap来容纳所有屏幕
        self.screenshot = QPixmap(total_rect.size())
        self.screenshot.fill(Qt.black)
        
        # 绘制每个屏幕的内容
        painter = QPainter(self.screenshot)
        for screen in screens:
            screen_geometry = screen.geometry()
            # 计算相对于total_rect的偏移
            offset_x = screen_geometry.x() - total_rect.x()
            offset_y = screen_geometry.y() - total_rect.y()
            # 捕获屏幕并绘制到对应位置
            screen_pixmap = screen.grabWindow(0)
            painter.drawPixmap(offset_x, offset_y, screen_pixmap)
        painter.end()
        
        # 设置窗口几何区域为所有屏幕的并集
        self.setGeometry(total_rect)
    
    def paintEvent(self, event: QPaintEvent):
        """绘制事件"""
        if not self.screenshot:
            return
        
        painter = QPainter(self)
        
        # 绘制屏幕截图（半透明）
        painter.drawPixmap(0, 0, self.screenshot)
        
        # 绘制半透明遮罩
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100))
        
        # 如果有选区，绘制选区内的清晰图像
        if self.start_pos and self.end_pos:
            rect = self._get_selection_rect()
            if rect:
                # 绘制选区内的清晰图像
                painter.drawPixmap(rect, self.screenshot, rect)
                
                # 绘制选区边框
                pen = QPen(QColor(0, 150, 255), 2)
                painter.setPen(pen)
                painter.drawRect(rect)
                
                # 绘制选区尺寸提示
                size_text = f"{rect.width()} x {rect.height()}"
                painter.setPen(QColor(255, 255, 255))
                painter.drawText(rect.topLeft() + QPoint(5, -5), size_text)
    
    def _get_selection_rect(self) -> Optional[QRect]:
        """获取选区矩形"""
        if not self.start_pos or not self.end_pos:
            return None
        
        x1, y1 = self.start_pos.x(), self.start_pos.y()
        x2, y2 = self.end_pos.x(), self.end_pos.y()
        
        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)
        
        return QRect(left, top, right - left, bottom - top)
    
    def mousePressEvent(self, event: QMouseEvent):
        """鼠标按下"""
        if event.button() == Qt.LeftButton:
            self.start_pos = event.pos()
            self.end_pos = event.pos()
            self.is_drawing = True
            self.update()
    
    def mouseMoveEvent(self, event: QMouseEvent):
        """鼠标移动"""
        if self.is_drawing:
            self.end_pos = event.pos()
            self.update()
    
    def mouseReleaseEvent(self, event: QMouseEvent):
        """鼠标释放"""
        if event.button() == Qt.LeftButton and self.is_drawing:
            self.end_pos = event.pos()
            self.is_drawing = False
            self._take_screenshot()
    
    def keyPressEvent(self, event: QKeyEvent):
        """按键事件"""
        if event.key() == Qt.Key_Escape:
            self.close()
    
    def _take_screenshot(self):
        """执行截图"""
        rect = self._get_selection_rect()
        if rect and rect.width() > 10 and rect.height() > 10:
            # 截取选区
            cropped = self.screenshot.copy(rect)
            
            if self.on_screenshot_taken:
                self.on_screenshot_taken(cropped)
        
        self.close()


def take_screenshot(on_screenshot_taken: Optional[Callable[[QPixmap], None]] = None, parent=None):
    """
    开始截图
    
    Args:
        on_screenshot_taken: 截图完成后的回调函数，参数为QPixmap
        parent: 父窗口（保留参数但不再使用，避免被模态对话框阻塞）
    """
    # 创建截图窗口（不设置父窗口，避免被模态对话框阻塞）
    widget = ScreenshotWidget(on_screenshot_taken, None)
    
    # 设置窗口标志
    # Qt.FramelessWindowHint - 无边框
    # Qt.WindowStaysOnTopHint - 保持在最前面
    # Qt.Tool - 工具窗口，不显示在任务栏
    # Qt.WindowDoesNotAcceptFocus - 不获取焦点（避免干扰）
    widget.setWindowFlags(
        Qt.FramelessWindowHint | 
        Qt.WindowStaysOnTopHint | 
        Qt.Tool |
        Qt.WindowDoesNotAcceptFocus
    )
    
    # 显示窗口
    widget.show()
    
    # 强制激活和提升到最前
    widget.activateWindow()
    widget.raise_()
    
    # 设置鼠标追踪
    widget.setMouseTracking(True)
    
    return widget


if __name__ == "__main__":
    # 测试
    import sys
    app = QApplication(sys.argv)
    
    def on_screenshot(pixmap: QPixmap):
        app_logger.info(f"截图尺寸: {pixmap.width()} x {pixmap.height()}")
        # 保存截图
        pixmap.save("screenshot.png")
        app_logger.info("截图已保存到 screenshot.png")
    
    take_screenshot(on_screenshot)
    sys.exit(app.exec())
