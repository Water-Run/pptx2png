r"""
PPT转图片工具的图形界面程序

:file: pptx2img-exe.py
:author: WaterRun
:time: 2025-12-28
"""

import sys
import os
import shutil
import tempfile
from typing import Any

import pythoncom
import win32com.client

try:
    import pptx2img
except ImportError:
    print("Error: pptx2img module missing.")
    pptx2img = None

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QScrollArea, QFileDialog,
    QDialog, QGridLayout, QFrame, QLineEdit, QStackedWidget,
    QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QPixmap, QIcon, QColor, QDesktopServices, QPainter, QPainterPath


# ==================== 多语言配置 ====================

LANG_TEXTS: dict[str, dict[str, str]] = {
    'zh': {
        'title': 'pptx2img',
        'by_author': 'by WaterRun',
        'file_info': '文件信息',
        'no_file': '未选择文件',
        'slide_count': '共 {0} 张幻灯片',
        'export_settings': '导出设置',
        'output_to': '输出至:',
        'browse': '浏览...',
        'scale': '倍率:',
        'scale_display': '显示 (默认)',
        'scale_1x': '1x',
        'scale_2x': '2x',
        'scale_3x': '3x',
        'scale_5x': '5x',
        'selected': '已选 {0} / {1}',
        'select_all': '全选',
        'select_none': '全不选',
        'export': '导出',
        'export_n': '导出 ({0})',
        'placeholder': '点击或拖拽 PPT 文件到此处',
        'loading': '正在加载...',
        'exporting': '正在导出...',
        'export_success': '导出成功',
        'export_success_msg': '成功导出 {0} 张图片到:\n{1}',
        'open_folder': '打开文件夹',
        'close': '关闭',
        'error': '错误',
        'export_failed': '导出失败:\n{0}',
        'select_ppt': '选择 PPT 文件',
        'select_output': '选择输出文件夹',
        'slide_n': '幻灯片 {0}',
        'output_path_ph': '输出路径...',
        'language': '语言',
        'lang_zh': '中文',
        'lang_en': 'English'
    },
    'en': {
        'title': 'pptx2img',
        'by_author': 'by WaterRun',
        'file_info': 'File Info',
        'no_file': 'No file selected',
        'slide_count': '{0} slides in total',
        'export_settings': 'Export Settings',
        'output_to': 'Output to:',
        'browse': 'Browse...',
        'scale': 'Scale:',
        'scale_display': 'Display (Default)',
        'scale_1x': '1x',
        'scale_2x': '2x',
        'scale_3x': '3x',
        'scale_5x': '5x',
        'selected': 'Selected {0} / {1}',
        'select_all': 'Select All',
        'select_none': 'Select None',
        'export': 'Export',
        'export_n': 'Export ({0})',
        'placeholder': 'Click or drag PPT file here',
        'loading': 'Loading...',
        'exporting': 'Exporting...',
        'export_success': 'Export Successful',
        'export_success_msg': 'Successfully exported {0} images to:\n{1}',
        'open_folder': 'Open Folder',
        'close': 'Close',
        'error': 'Error',
        'export_failed': 'Export failed:\n{0}',
        'select_ppt': 'Select PPT File',
        'select_output': 'Select Output Folder',
        'slide_n': 'Slide {0}',
        'output_path_ph': 'Output path...',
        'language': 'Language',
        'lang_zh': '中文',
        'lang_en': 'English'
    }
}


# ==================== 样式与配置 ====================

COLOR_PPT_ORANGE: str = "#D24726"
COLOR_PPT_DARK: str = "#A4371E"
COLOR_BG_SIDEBAR: str = "#F3F3F3"
COLOR_BG_CONTENT: str = "#FFFFFF"
COLOR_TEXT_MAIN: str = "#2D2D2D"
COLOR_TEXT_SUB: str = "#757575"
BORDER_COLOR: str = "#C0C0C0"

GLOBAL_STYLESHEET: str = f"""
    QWidget {{
        font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
        font-size: 14px;
        color: {COLOR_TEXT_MAIN};
        background-color: {COLOR_BG_CONTENT};
    }}
    
    QToolTip {{
        background-color: #FFFFE0;
        color: black;
        border: 1px solid {COLOR_TEXT_SUB};
        padding: 5px;
        font-size: 12px;
    }}

    QPushButton {{
        background-color: white;
        border: 1px solid {BORDER_COLOR};
        border-radius: 3px;
        padding: 6px 16px;
        color: {COLOR_TEXT_MAIN};
        min-height: 24px;
    }}
    QPushButton:hover {{
        border-color: {COLOR_PPT_ORANGE};
        color: {COLOR_PPT_ORANGE};
        background-color: #FFF5F2;
    }}
    QPushButton:pressed {{
        background-color: {COLOR_PPT_ORANGE};
        color: white;
    }}
    
    .PrimaryButton {{
        background-color: {COLOR_PPT_ORANGE};
        color: white;
        border: none;
        font-weight: bold;
        font-size: 15px;
        border-radius: 4px;
        padding: 10px;
        min-height: 40px;
    }}
    .PrimaryButton:hover {{ 
        background-color: {COLOR_PPT_DARK}; 
    }}
    .PrimaryButton:disabled {{ 
        background-color: #CCCCCC; 
        color: #888888; 
    }}

    QLineEdit {{
        background-color: white;
        color: black;
        border: 1px solid {BORDER_COLOR};
        border-radius: 4px;
        padding: 6px 10px;
        selection-background-color: {COLOR_PPT_ORANGE};
        selection-color: white;
    }}
    QLineEdit:focus {{
        border: 2px solid {COLOR_PPT_ORANGE};
        padding: 5px 9px;
    }}
    QLineEdit:read-only {{ 
        background-color: #F9F9F9; 
        color: #555; 
    }}

    QRadioButton {{
        spacing: 8px;
        color: {COLOR_TEXT_MAIN};
        background: transparent;
    }}
    QRadioButton::indicator {{
        width: 16px;
        height: 16px;
        border-radius: 8px;
        border: 2px solid {BORDER_COLOR};
        background: white;
    }}
    QRadioButton::indicator:checked {{
        border: 2px solid {COLOR_PPT_ORANGE};
        background: white;
    }}
    QRadioButton:hover {{
        color: {COLOR_PPT_ORANGE};
    }}
    
    QScrollArea {{ 
        border: none; 
        background: transparent; 
    }}
    QScrollBar:vertical {{
        background: #F5F5F5; 
        width: 10px; 
        margin: 0;
        border-radius: 5px;
    }}
    QScrollBar::handle:vertical {{
        background: #D0D0D0; 
        min-height: 30px;
        border-radius: 5px;
    }}
    QScrollBar::handle:vertical:hover {{ 
        background: #A0A0A0; 
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
"""

MODERN_COMBOBOX_STYLE: str = f"""
    QComboBox {{
        background-color: white;
        color: {COLOR_TEXT_MAIN};
        border: 1px solid {BORDER_COLOR};
        border-radius: 4px;
        padding: 6px 30px 6px 10px;
        min-height: 28px;
        font-size: 14px;
    }}
    QComboBox:hover {{
        border: 1px solid {COLOR_PPT_ORANGE};
        background-color: #FFFBFA;
    }}
    QComboBox:focus {{
        border: 2px solid {COLOR_PPT_ORANGE};
        padding: 5px 29px 5px 9px;
    }}
    QComboBox::drop-down {{
        subcontrol-origin: padding;
        subcontrol-position: center right;
        width: 24px;
        border: none;
        background: transparent;
    }}
    QComboBox::down-arrow {{
        image: none;
        border-left: 4px solid transparent;
        border-right: 4px solid transparent;
        border-top: 5px solid {COLOR_TEXT_SUB};
        margin-right: 6px;
    }}
    QComboBox::down-arrow:hover {{
        border-top-color: {COLOR_PPT_ORANGE};
    }}
    QComboBox QAbstractItemView {{
        background-color: white;
        color: {COLOR_TEXT_MAIN};
        border: 1px solid {COLOR_PPT_ORANGE};
        selection-background-color: {COLOR_PPT_ORANGE};
        selection-color: white;
        outline: none;
        padding: 4px;
        border-radius: 4px;
    }}
    QComboBox QAbstractItemView::item {{
        min-height: 32px;
        padding: 4px 10px;
        border-radius: 3px;
    }}
    QComboBox QAbstractItemView::item:hover {{
        background-color: #FFF5F2;
        color: {COLOR_PPT_ORANGE};
    }}
    QComboBox QAbstractItemView::item:selected {{
        background-color: {COLOR_PPT_ORANGE};
        color: white;
    }}
"""


# ==================== 工具函数 ====================

def merge_ranges(indices: list[int]) -> list[list[int]]:
    r"""
    将索引列表合并为连续范围列表
    
    :param indices: 索引列表, 如 [1, 2, 3, 5, 6, 8]
    :return list[list[int]]: 范围列表, 如 [[1, 3], [5, 6], [8, 8]]
    """
    if not indices:
        return []
    
    sorted_indices: list[int] = sorted(set(indices))
    ranges: list[list[int]] = []
    start: int = sorted_indices[0]
    end: int = sorted_indices[0]
    
    for i in range(1, len(sorted_indices)):
        if sorted_indices[i] == end + 1:
            end = sorted_indices[i]
        else:
            ranges.append([start, end])
            start = sorted_indices[i]
            end = sorted_indices[i]
    
    ranges.append([start, end])
    return ranges


# ==================== 自定义对话框 ====================

class CustomMessageDialog(QDialog):
    r"""
    自定义消息对话框, 采用PPT风格: 白底、橙色边框、直角矩形按钮
    """
    
    def __init__(
        self,
        parent: QWidget | None,
        title: str,
        message: str,
        buttons: list[tuple[str, str]],
        icon_type: str = "info"
    ) -> None:
        r"""
        初始化自定义消息对话框
        
        :param parent: 父窗口
        :param title: 对话框标题
        :param message: 消息内容
        :param buttons: 按钮列表, 每项为 (按钮文本, 按钮标识)
        :param icon_type: 图标类型, "info" 或 "error"
        """
        super().__init__(parent)
        self.clicked_button: str = ""
        self._setup_ui(title, message, buttons, icon_type)
    
    def _setup_ui(
        self,
        title: str,
        message: str,
        buttons: list[tuple[str, str]],
        icon_type: str
    ) -> None:
        r"""
        设置对话框UI
        
        :param title: 对话框标题
        :param message: 消息内容
        :param buttons: 按钮列表
        :param icon_type: 图标类型
        """
        self.setWindowTitle(title)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setMinimumWidth(400)
        self.setMaximumWidth(500)
        
        self.setStyleSheet(f"""
            CustomMessageDialog {{
                background-color: white;
                border: 2px solid {COLOR_PPT_ORANGE};
            }}
        """)
        
        main_layout: QVBoxLayout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        title_bar: QWidget = QWidget()
        title_bar.setFixedHeight(36)
        title_bar.setStyleSheet(f"background-color: {COLOR_PPT_ORANGE};")
        title_layout: QHBoxLayout = QHBoxLayout(title_bar)
        title_layout.setContentsMargins(12, 0, 8, 0)
        
        title_label: QLabel = QLabel(title)
        title_label.setStyleSheet("color: white; font-weight: bold; font-size: 14px; background: transparent;")
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        close_btn: QPushButton = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: white;
                font-size: 14px;
                border-radius: 0px;
            }
            QPushButton:hover {
                background: rgba(255,255,255,0.2);
            }
        """)
        close_btn.clicked.connect(self.reject)
        title_layout.addWidget(close_btn)
        
        main_layout.addWidget(title_bar)
        
        content_widget: QWidget = QWidget()
        content_widget.setStyleSheet("background-color: white;")
        content_layout: QVBoxLayout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(24, 20, 24, 20)
        content_layout.setSpacing(16)
        
        msg_layout: QHBoxLayout = QHBoxLayout()
        msg_layout.setSpacing(16)
        
        icon_label: QLabel = QLabel()
        icon_label.setFixedSize(40, 40)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_color: str = COLOR_PPT_ORANGE if icon_type == "info" else "#E81123"
        icon_char: str = "✓" if icon_type == "info" else "✕"
        icon_label.setStyleSheet(f"""
            QLabel {{
                background-color: {icon_color};
                color: white;
                font-size: 20px;
                font-weight: bold;
                border-radius: 20px;
            }}
        """)
        icon_label.setText(icon_char)
        msg_layout.addWidget(icon_label, 0, Qt.AlignmentFlag.AlignTop)
        
        msg_label: QLabel = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet(f"color: {COLOR_TEXT_MAIN}; font-size: 14px; background: transparent;")
        msg_layout.addWidget(msg_label, 1)
        
        content_layout.addLayout(msg_layout)
        
        btn_layout: QHBoxLayout = QHBoxLayout()
        btn_layout.setSpacing(12)
        btn_layout.addStretch()
        
        for i, (btn_text, btn_id) in enumerate(buttons):
            btn: QPushButton = QPushButton(btn_text)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setMinimumWidth(90)
            btn.setFixedHeight(32)
            
            if i == 0:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {COLOR_PPT_ORANGE};
                        color: white;
                        border: none;
                        border-radius: 0px;
                        padding: 6px 20px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: {COLOR_PPT_DARK};
                    }}
                    QPushButton:pressed {{
                        background-color: #8B2E15;
                    }}
                """)
            else:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: white;
                        color: {COLOR_TEXT_MAIN};
                        border: 1px solid {BORDER_COLOR};
                        border-radius: 0px;
                        padding: 6px 20px;
                    }}
                    QPushButton:hover {{
                        border-color: {COLOR_PPT_ORANGE};
                        color: {COLOR_PPT_ORANGE};
                    }}
                    QPushButton:pressed {{
                        background-color: #F5F5F5;
                    }}
                """)
            
            btn.clicked.connect(lambda checked, bid=btn_id: self._on_button_clicked(bid))
            btn_layout.addWidget(btn)
        
        content_layout.addLayout(btn_layout)
        main_layout.addWidget(content_widget)
        
        self._drag_pos: Any = None
        title_bar.mousePressEvent = self._title_mouse_press
        title_bar.mouseMoveEvent = self._title_mouse_move
    
    def _title_mouse_press(self, event: Any) -> None:
        r"""
        标题栏鼠标按下事件处理
        
        :param event: 鼠标事件
        """
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
    
    def _title_mouse_move(self, event: Any) -> None:
        r"""
        标题栏鼠标移动事件处理
        
        :param event: 鼠标事件
        """
        if event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()
    
    def _on_button_clicked(self, button_id: str) -> None:
        r"""
        按钮点击事件处理
        
        :param button_id: 按钮标识
        """
        self.clicked_button = button_id
        self.accept()
    
    def get_clicked_button(self) -> str:
        r"""
        获取被点击的按钮标识
        
        :return str: 按钮标识
        """
        return self.clicked_button


# ==================== 线程工作类 ====================

class LoadThread(QThread):
    r"""
    加载PPT预览图的工作线程
    """
    
    finished: pyqtSignal = pyqtSignal(bool, str, dict)

    def __init__(self, path: str) -> None:
        r"""
        初始化加载线程
        
        :param path: PPT文件路径
        """
        super().__init__()
        self.path: str = path

    def run(self) -> None:
        r"""
        执行加载任务
        """
        pythoncom.CoInitialize()
        pres: Any = None
        ppt_app: Any = None
        temp_dir: str | None = None
        try:
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                ppt_app.WindowState = 2
            except Exception:
                ...
            
            pres = ppt_app.Presentations.Open(
                os.path.abspath(self.path),
                ReadOnly=True,
                WithWindow=False
            )
            
            temp_dir = tempfile.mkdtemp(prefix="pptx2img_prev_")
            slides_info: list[dict[str, Any]] = []
            
            w: float = pres.PageSetup.SlideWidth
            h: float = pres.PageSetup.SlideHeight
            count: int = pres.Slides.Count
            
            preview_w: int = 640
            preview_h: int = int(preview_w * h / w)
            
            for i in range(1, count + 1):
                slide: Any = pres.Slides(i)
                img_path: str = os.path.join(temp_dir, f"p_{i}.jpg")
                slide.Export(img_path, "JPG", preview_w, preview_h)
                slides_info.append({
                    'index': i,
                    'path': img_path,
                    'selected': True
                })
            
            data: dict[str, Any] = {
                'path': self.path,
                'temp_dir': temp_dir,
                'size': (w, h),
                'slides': slides_info
            }
            pres.Close()
            self.finished.emit(True, "OK", data)
            
        except Exception as e:
            self.finished.emit(False, str(e), {})
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            if pres:
                try:
                    pres.Close()
                except Exception:
                    ...
            if ppt_app:
                try:
                    ppt_app.Quit()
                except Exception:
                    ...
        finally:
            pythoncom.CoUninitialize()


class ExportThread(QThread):
    r"""
    导出图片的工作线程, 支持合并连续范围优化
    """
    
    progress: pyqtSignal = pyqtSignal(int, int)
    finished: pyqtSignal = pyqtSignal(bool, str, int)

    def __init__(
        self,
        ppt_path: str,
        indices: list[int],
        out_dir: str,
        scale: int
    ) -> None:
        r"""
        初始化导出线程
        
        :param ppt_path: PPT文件路径
        :param indices: 要导出的幻灯片索引列表
        :param out_dir: 输出目录
        :param scale: 导出倍率, 0表示使用显示倍率
        """
        super().__init__()
        self.ppt_path: str = ppt_path
        self.indices: list[int] = indices
        self.out_dir: str = out_dir
        self.scale: int = scale

    def run(self) -> None:
        r"""
        执行导出任务
        """
        pythoncom.CoInitialize()
        try:
            if not pptx2img:
                raise ImportError("pptx2img library not found!")

            total: int = len(self.indices)
            count: int = 0
            
            ranges: list[list[int]] = merge_ranges(self.indices)
            lib_scale: int | None = self.scale if self.scale > 0 else None
            
            for slide_range in ranges:
                pptx2img.topng(
                    pptx=self.ppt_path,
                    output_dir=self.out_dir,
                    slide_range=slide_range,
                    scale=lib_scale
                )
                
                range_count: int = slide_range[1] - slide_range[0] + 1
                count += range_count
                self.progress.emit(count, total)
            
            self.finished.emit(True, self.out_dir, count)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished.emit(False, str(e), 0)
        finally:
            pythoncom.CoUninitialize()


# ==================== UI 组件 ====================

class TitleBar(QWidget):
    r"""
    自定义橙色标题栏, 居中显示软件名, 右侧包含最小化和关闭按钮
    """
    
    def __init__(self, parent: QMainWindow) -> None:
        r"""
        初始化标题栏
        
        :param parent: 父窗口
        """
        super().__init__(parent)
        self.main_window: QMainWindow = parent
        self._drag_pos: Any = None
        
        self.setObjectName("TitleBar")
        self.setFixedHeight(40)
        self.setAutoFillBackground(True)
        
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self) -> None:
        r"""
        设置标题栏UI布局
        """
        layout: QHBoxLayout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        layout.addStretch(1)
        
        self.lbl_title: QLabel = QLabel("pptx2img")
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_title)
        
        layout.addStretch(1)
        
        self.btn_minimize: QPushButton = QPushButton("─")
        self.btn_minimize.setFixedSize(45, 40)
        self.btn_minimize.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_minimize.clicked.connect(self.main_window.showMinimized)
        layout.addWidget(self.btn_minimize)
        
        self.btn_close: QPushButton = QPushButton("✕")
        self.btn_close.setFixedSize(45, 40)
        self.btn_close.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_close.clicked.connect(self.main_window.close)
        layout.addWidget(self.btn_close)
    
    def _apply_styles(self) -> None:
        r"""
        应用标题栏样式
        """
        self.setStyleSheet(f"""
            #TitleBar {{
                background-color: {COLOR_PPT_ORANGE};
            }}
        """)
        
        self.lbl_title.setStyleSheet("""
            QLabel {
                color: white;
                font-weight: bold;
                font-size: 16px;
                background: transparent;
                padding: 0;
                margin: 0;
            }
        """)
        
        minimize_style: str = """
            QPushButton {
                background: transparent;
                border: none;
                color: white;
                font-size: 18px;
                padding: 0;
                margin: 0;
                border-radius: 0px;
            }
            QPushButton:hover {
                background: rgba(255,255,255,0.15);
            }
        """
        self.btn_minimize.setStyleSheet(minimize_style)
        
        close_style: str = """
            QPushButton {
                background: transparent;
                border: none;
                color: white;
                font-size: 18px;
                padding: 0;
                margin: 0;
                border-radius: 0px;
            }
            QPushButton:hover {
                background: #E81123;
            }
        """
        self.btn_close.setStyleSheet(close_style)
    
    def paintEvent(self, event: Any) -> None:
        r"""
        绘制事件, 确保背景色正确渲染
        
        :param event: 绘制事件
        """
        painter: QPainter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor(COLOR_PPT_ORANGE))
        super().paintEvent(event)
    
    def mousePressEvent(self, event: Any) -> None:
        r"""
        鼠标按下事件, 用于拖拽窗口
        
        :param event: 鼠标事件
        """
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.main_window.frameGeometry().topLeft()
            event.accept()
    
    def mouseMoveEvent(self, event: Any) -> None:
        r"""
        鼠标移动事件, 用于拖拽窗口
        
        :param event: 鼠标事件
        """
        if event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos:
            self.main_window.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()


class Sidebar(QWidget):
    r"""
    左侧控制面板, 包含语言切换、文件信息、导出设置等
    """
    
    def __init__(self, parent: "MainWindow") -> None:
        r"""
        初始化侧边栏
        
        :param parent: 主窗口
        """
        super().__init__(parent)
        self.main_window: "MainWindow" = parent
        
        self.setObjectName("Sidebar")
        self.setFixedWidth(260)
        self.setStyleSheet(f"""
            #Sidebar {{
                background-color: {COLOR_BG_SIDEBAR};
                border-right: 1px solid {BORDER_COLOR};
            }}
        """)
        
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        r"""
        设置侧边栏UI布局
        """
        self.layout: QVBoxLayout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 30, 20, 30)
        self.layout.setSpacing(10)
        
        self._setup_logo()
        self._setup_language_section()
        self._setup_file_info_section()
        self._setup_export_settings_section()
        self._setup_action_section()
    
    def _setup_logo(self) -> None:
        r"""
        设置Logo和软件信息区域
        """
        self.lbl_logo: QLabel = QLabel()
        self.lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_logo.setStyleSheet("background: transparent;")
        self._load_logo()
        self.layout.addWidget(self.lbl_logo)
        
        self.layout.addSpacing(5)
        
        self.lbl_name: QLabel = QLabel()
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_name.setStyleSheet(f"""
            font-weight: bold;
            font-size: 18px;
            color: {COLOR_TEXT_MAIN};
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_name)
        
        self.lbl_author: QLabel = QLabel()
        self.lbl_author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_author.setStyleSheet(f"""
            font-size: 12px;
            color: {COLOR_TEXT_SUB};
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_author)
        
        self.lbl_github: QLabel = QLabel(
            f'<a href="https://github.com/Water-Run/pptx2img" '
            f'style="color:{COLOR_PPT_ORANGE}; text-decoration:none; font-weight:bold;">GitHub</a>'
        )
        self.lbl_github.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_github.setOpenExternalLinks(True)
        self.lbl_github.setStyleSheet("background: transparent;")
        self.lbl_github.setCursor(Qt.CursorShape.PointingHandCursor)
        self.layout.addWidget(self.lbl_github)
        
        line: QFrame = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet(f"color: {BORDER_COLOR}; margin-top: 15px; margin-bottom: 15px;")
        self.layout.addWidget(line)
    
    def _load_logo(self) -> None:
        r"""
        加载Logo图片，支持高DPI显示
        """
        logo_path: str = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            '..',
            'logo.png'
        )
        if os.path.exists(logo_path):
            # 获取设备像素比
            screen = QApplication.primaryScreen()
            dpr: float = screen.devicePixelRatio() if screen else 1.0
            
            # 目标逻辑尺寸
            logical_size: int = 80
            # 实际需要的物理像素尺寸
            physical_size: int = int(logical_size * dpr)
            
            logo_pix: QPixmap = QPixmap(logo_path)
            # 缩放到物理像素尺寸
            scaled_pix: QPixmap = logo_pix.scaled(
                physical_size, physical_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            # 设置设备像素比，告诉Qt这是高分辨率图片
            scaled_pix.setDevicePixelRatio(dpr)
            
            self.lbl_logo.setPixmap(scaled_pix)
        else:
            self.lbl_logo.setText("Logo")
    
    def _setup_language_section(self) -> None:
        r"""
        设置语言选择区域
        """
        self.lbl_language: QLabel = QLabel()
        self.lbl_language.setStyleSheet(f"""
            font-weight: bold;
            color: {COLOR_PPT_ORANGE};
            margin-bottom: 5px;
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_language)
        
        lang_layout: QHBoxLayout = QHBoxLayout()
        lang_layout.setSpacing(15)
        
        self.radio_zh: QRadioButton = QRadioButton()
        self.radio_en: QRadioButton = QRadioButton()
        self.radio_zh.setChecked(True)
        
        self.lang_group: QButtonGroup = QButtonGroup()
        self.lang_group.addButton(self.radio_zh, 0)
        self.lang_group.addButton(self.radio_en, 1)
        self.lang_group.buttonClicked.connect(self._on_language_changed)
        
        lang_layout.addWidget(self.radio_zh)
        lang_layout.addWidget(self.radio_en)
        lang_layout.addStretch()
        self.layout.addLayout(lang_layout)
        self.layout.addSpacing(10)
    
    def _setup_file_info_section(self) -> None:
        r"""
        设置文件信息区域
        """
        self.lbl_grp_info: QLabel = QLabel()
        self.lbl_grp_info.setStyleSheet(f"""
            font-weight: bold;
            color: {COLOR_PPT_ORANGE};
            margin-bottom: 5px;
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_grp_info)
        
        self.lbl_filename: QLabel = QLabel()
        self.lbl_filename.setWordWrap(True)
        self.lbl_filename.setStyleSheet("""
            font-weight: bold;
            font-size: 13px;
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_filename)
        
        self.lbl_slide_count: QLabel = QLabel("-")
        self.lbl_slide_count.setStyleSheet(f"""
            color: {COLOR_TEXT_SUB};
            font-size: 12px;
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_slide_count)
        self.layout.addSpacing(10)
    
    def _setup_export_settings_section(self) -> None:
        r"""
        设置导出设置区域
        """
        self.lbl_grp_settings: QLabel = QLabel()
        self.lbl_grp_settings.setStyleSheet(f"""
            font-weight: bold;
            color: {COLOR_PPT_ORANGE};
            margin-bottom: 5px;
            background: transparent;
        """)
        self.layout.addWidget(self.lbl_grp_settings)
        
        self.lbl_output: QLabel = QLabel()
        self.lbl_output.setStyleSheet("background: transparent;")
        self.layout.addWidget(self.lbl_output)
        
        path_layout: QHBoxLayout = QHBoxLayout()
        self.entry_out_dir: QLineEdit = QLineEdit()
        self.entry_out_dir.setReadOnly(True)
        self.btn_browse: QPushButton = QPushButton()
        self.btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_browse.clicked.connect(self.main_window.choose_output_dir)
        path_layout.addWidget(self.entry_out_dir, 1)
        path_layout.addWidget(self.btn_browse)
        self.layout.addLayout(path_layout)
        
        self.lbl_scale: QLabel = QLabel()
        self.lbl_scale.setStyleSheet("background: transparent;")
        self.layout.addWidget(self.lbl_scale)
        
        self.combo_scale: QComboBox = QComboBox()
        self.combo_scale.setStyleSheet(MODERN_COMBOBOX_STYLE)
        self.combo_scale.setCurrentIndex(0)
        self.layout.addWidget(self.combo_scale)
        
        self.layout.addStretch()
    
    def _setup_action_section(self) -> None:
        r"""
        设置操作按钮区域
        """
        self.lbl_sel_info: QLabel = QLabel()
        self.lbl_sel_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sel_info.setStyleSheet("background: transparent;")
        self.layout.addWidget(self.lbl_sel_info)

        btn_layout: QHBoxLayout = QHBoxLayout()
        self.btn_all: QPushButton = QPushButton()
        self.btn_none: QPushButton = QPushButton()
        self.btn_all.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_none.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_all.clicked.connect(self.main_window.select_all)
        self.btn_none.clicked.connect(self.main_window.select_none)
        btn_layout.addWidget(self.btn_all)
        btn_layout.addWidget(self.btn_none)
        self.layout.addLayout(btn_layout)
        
        self.btn_export: QPushButton = QPushButton()
        self.btn_export.setProperty("class", "PrimaryButton")
        self.btn_export.setFixedHeight(45)
        self.btn_export.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_export.clicked.connect(self.main_window.start_export)
        self.btn_export.setEnabled(False)
        self.layout.addWidget(self.btn_export)
    
    def _on_language_changed(self) -> None:
        r"""
        语言切换事件处理
        """
        self.main_window.update_language()
    
    def update_texts(self, lang: str) -> None:
        r"""
        更新所有文本为指定语言
        
        :param lang: 语言代码, "zh" 或 "en"
        """
        t: dict[str, str] = LANG_TEXTS[lang]
        self.lbl_name.setText(t['title'])
        self.lbl_author.setText(t['by_author'])
        self.lbl_language.setText(t['language'])
        self.radio_zh.setText(t['lang_zh'])
        self.radio_en.setText(t['lang_en'])
        self.lbl_grp_info.setText(t['file_info'])
        self.lbl_grp_settings.setText(t['export_settings'])
        self.lbl_output.setText(t['output_to'])
        self.btn_browse.setText(t['browse'])
        self.lbl_scale.setText(t['scale'])
        self.btn_all.setText(t['select_all'])
        self.btn_none.setText(t['select_none'])
        self.entry_out_dir.setPlaceholderText(t['output_path_ph'])
        
        self.combo_scale.clear()
        self.combo_scale.addItems([
            t['scale_display'],
            t['scale_1x'],
            t['scale_2x'],
            t['scale_3x'],
            t['scale_5x']
        ])
        self.combo_scale.setCurrentIndex(0)


class SlideCard(QFrame):
    r"""
    幻灯片卡片组件, 显示预览图和选择状态
    """
    
    toggled: pyqtSignal = pyqtSignal()
    
    def __init__(self, info: dict[str, Any], lang: str = 'zh') -> None:
        r"""
        初始化幻灯片卡片
        
        :param info: 幻灯片信息字典, 包含 index, path, selected
        :param lang: 语言代码
        """
        super().__init__()
        self.info: dict[str, Any] = info
        self.lang: str = lang
        
        self.setFixedSize(220, 160)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self._setup_ui()
        self._update_style()
    
    def _setup_ui(self) -> None:
        r"""
        设置卡片UI布局
        """
        layout: QVBoxLayout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 5)
        layout.setSpacing(5)
        
        self.lbl_img: QLabel = QLabel()
        self.lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_img.setStyleSheet("background: #EEE; border-radius: 4px;")
        
        if os.path.exists(self.info['path']):
            pix: QPixmap = QPixmap(self.info['path'])
            self.lbl_img.setPixmap(pix.scaled(
                204, 115,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            ))
        
        self.lbl_idx: QLabel = QLabel()
        self.lbl_idx.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(self.lbl_img)
        layout.addWidget(self.lbl_idx)
    
    def set_language(self, lang: str) -> None:
        r"""
        设置语言
        
        :param lang: 语言代码
        """
        self.lang = lang
        self._update_style()
    
    def _update_style(self) -> None:
        r"""
        更新卡片样式和文本
        """
        t: dict[str, str] = LANG_TEXTS[self.lang]
        self.lbl_idx.setText(t['slide_n'].format(self.info['index']))
        
        if self.info['selected']:
            self.setStyleSheet(f"""
                SlideCard {{
                    background: white;
                    border: 3px solid {COLOR_PPT_ORANGE};
                    border-radius: 6px;
                }}
            """)
            self.lbl_idx.setStyleSheet(f"""
                color: {COLOR_PPT_ORANGE};
                font-weight: bold;
                background: transparent;
                font-size: 12px;
            """)
        else:
            self.setStyleSheet(f"""
                SlideCard {{
                    background: white;
                    border: 1px solid {BORDER_COLOR};
                    border-radius: 6px;
                }}
                SlideCard:hover {{
                    border-color: {COLOR_PPT_ORANGE};
                }}
            """)
            self.lbl_idx.setStyleSheet(f"""
                color: {COLOR_TEXT_SUB};
                background: transparent;
                font-size: 12px;
            """)
    
    def mousePressEvent(self, event: Any) -> None:
        r"""
        鼠标点击事件, 切换选择状态
        
        :param event: 鼠标事件
        """
        if event.button() == Qt.MouseButton.LeftButton:
            self.info['selected'] = not self.info['selected']
            self._update_style()
            self.toggled.emit()


class PlaceholderWidget(QWidget):
    r"""
    占位组件, 显示文件拖拽提示
    """
    
    clicked: pyqtSignal = pyqtSignal()
    
    def __init__(self, parent: QWidget | None = None, lang: str = 'zh') -> None:
        r"""
        初始化占位组件
        
        :param parent: 父组件
        :param lang: 语言代码
        """
        super().__init__(parent)
        self.lang: str = lang
        self._setup_ui()
    
    def _setup_ui(self) -> None:
        r"""
        设置占位组件UI
        """
        layout: QVBoxLayout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        lbl_icon: QLabel = QLabel("＋")
        lbl_icon.setStyleSheet(f"""
            QLabel {{
                font-size: 80px;
                color: {COLOR_PPT_ORANGE};
                border: 3px dashed {COLOR_PPT_ORANGE};
                border-radius: 20px;
                padding: 20px 60px;
                background: #FFF5F3;
            }}
            QLabel:hover {{
                background: #FFEBE8;
            }}
        """)
        lbl_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_icon.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.lbl_text: QLabel = QLabel()
        self.lbl_text.setStyleSheet("font-size: 18px; color: #555; background: transparent;")
        self.lbl_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(lbl_icon)
        layout.addWidget(self.lbl_text)
        
        self._update_text()
    
    def set_language(self, lang: str) -> None:
        r"""
        设置语言
        
        :param lang: 语言代码
        """
        self.lang = lang
        self._update_text()
    
    def _update_text(self) -> None:
        r"""
        更新提示文本
        """
        self.lbl_text.setText(LANG_TEXTS[self.lang]['placeholder'])
    
    def mousePressEvent(self, event: Any) -> None:
        r"""
        鼠标点击事件
        
        :param event: 鼠标事件
        """
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()


# ==================== 主窗口 ====================

class MainWindow(QMainWindow):
    r"""
    应用程序主窗口
    """
    
    def __init__(self) -> None:
        r"""
        初始化主窗口
        """
        super().__init__()
        self.current_lang: str = 'zh'
        self.ppt_data: dict[str, Any] | None = None
        self.cards: list[SlideCard] = []
        self.loader: LoadThread | None = None
        self.exporter: ExportThread | None = None
        
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(1100, 750)
        
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            self.move(
                (screen_geometry.width() - self.width()) // 2,
                (screen_geometry.height() - self.height()) // 2
            )
        
        self._setup_ui()
        self.update_language()
    
    def _setup_ui(self) -> None:
        r"""
        设置主窗口UI
        """
        self.main_container: QFrame = QFrame()
        self.main_container.setObjectName("MainContainer")
        self.main_container.setStyleSheet(f"""
            QFrame#MainContainer {{
                background-color: {COLOR_BG_CONTENT};
                border: 2px solid {COLOR_PPT_ORANGE};
                border-radius: 10px;
            }}
        """)
        self.setCentralWidget(self.main_container)
        
        root_layout: QVBoxLayout = QVBoxLayout(self.main_container)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        
        self.title_bar: TitleBar = TitleBar(self)
        root_layout.addWidget(self.title_bar)
        
        body_layout: QHBoxLayout = QHBoxLayout()
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(0)
        
        self.sidebar: Sidebar = Sidebar(self)
        body_layout.addWidget(self.sidebar)
        
        self.right_stack: QStackedWidget = QStackedWidget()
        self.right_stack.setStyleSheet(f"""
            background-color: {COLOR_BG_CONTENT};
            border-bottom-right-radius: 10px;
        """)
        
        self.page_empty: PlaceholderWidget = PlaceholderWidget(lang=self.current_lang)
        self.page_empty.clicked.connect(self.open_file_dialog)
        self.right_stack.addWidget(self.page_empty)
        
        self.scroll_area: QScrollArea = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.grid_container: QWidget = QWidget()
        self.grid_container.setStyleSheet("background: transparent;")
        self.grid_layout: QGridLayout = QGridLayout(self.grid_container)
        self.grid_layout.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )
        self.grid_layout.setSpacing(25)
        self.grid_layout.setContentsMargins(30, 30, 30, 30)
        self.scroll_area.setWidget(self.grid_container)
        self.right_stack.addWidget(self.scroll_area)
        
        self.loading_lbl: QLabel = QLabel()
        self.loading_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_lbl.setStyleSheet(f"""
            font-size: 24px;
            color: {COLOR_PPT_ORANGE};
            font-weight: bold;
            background: transparent;
        """)
        self.right_stack.addWidget(self.loading_lbl)
        
        body_layout.addWidget(self.right_stack)
        root_layout.addLayout(body_layout)
        
        self.setAcceptDrops(True)
    
    def paintEvent(self, event: Any) -> None:
        r"""
        绘制事件, 绘制窗口圆角边框
        
        :param event: 绘制事件
        """
        painter: QPainter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        path: QPainterPath = QPainterPath()
        path.setFillRule(Qt.FillRule.WindingFill)
        path.addRoundedRect(0.0, 0.0, float(self.width()), 40.0, 10.0, 10.0)
        painter.fillPath(path, QColor(COLOR_PPT_ORANGE))
        painter.fillRect(0, 20, self.width(), 20, QColor(COLOR_PPT_ORANGE))
    
    def update_language(self) -> None:
        r"""
        更新界面语言
        """
        self.current_lang = 'zh' if self.sidebar.radio_zh.isChecked() else 'en'
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        
        self.sidebar.update_texts(self.current_lang)
        self.loading_lbl.setText(t['loading'])
        self.page_empty.set_language(self.current_lang)
        
        for card in self.cards:
            card.set_language(self.current_lang)
        
        self._update_stats()
        
        if self.ppt_data:
            self.sidebar.lbl_filename.setText(os.path.basename(self.ppt_data['path']))
        else:
            self.sidebar.lbl_filename.setText(t['no_file'])
    
    def dragEnterEvent(self, event: Any) -> None:
        r"""
        拖拽进入事件
        
        :param event: 拖拽事件
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: Any) -> None:
        r"""
        放下事件
        
        :param event: 拖拽事件
        """
        for url in event.mimeData().urls():
            path: str = url.toLocalFile()
            if path.lower().endswith(('.pptx', '.ppt')):
                self._load_file(path)
                break
        event.acceptProposedAction()
    
    def open_file_dialog(self) -> None:
        r"""
        打开文件选择对话框
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        path, _ = QFileDialog.getOpenFileName(
            self,
            t['select_ppt'],
            "",
            "PowerPoint (*.pptx *.ppt)"
        )
        if path:
            self._load_file(path)
    
    def _load_file(self, path: str) -> None:
        r"""
        加载PPT文件
        
        :param path: 文件路径
        """
        self.right_stack.setCurrentIndex(2)
        self.sidebar.btn_export.setEnabled(False)
        self.loader = LoadThread(path)
        self.loader.finished.connect(self._on_load_finished)
        self.loader.start()
    
    def _on_load_finished(
        self,
        success: bool,
        msg: str,
        data: dict[str, Any]
    ) -> None:
        r"""
        文件加载完成回调
        
        :param success: 是否成功
        :param msg: 消息
        :param data: 加载的数据
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        
        if success:
            self.ppt_data = data
            self.sidebar.lbl_filename.setText(os.path.basename(data['path']))
            default_out: str = os.path.join(
                os.path.dirname(data['path']),
                "pptx2img"
            )
            self.sidebar.entry_out_dir.setText(default_out)
            self.sidebar.entry_out_dir.setToolTip(default_out)
            
            self._populate_grid()
            self.right_stack.setCurrentIndex(1)
            self.sidebar.btn_export.setEnabled(True)
        else:
            self._show_error_dialog(t['error'], msg)
            self.right_stack.setCurrentIndex(0)
        
        self._update_stats()
    
    def _populate_grid(self) -> None:
        r"""
        填充幻灯片网格
        """
        for i in reversed(range(self.grid_layout.count())):
            widget = self.grid_layout.itemAt(i)
            if widget:
                w = widget.widget()
                if w:
                    w.setParent(None)
        
        self.cards = []
        
        if not self.ppt_data:
            return
        
        slides: list[dict[str, Any]] = self.ppt_data['slides']
        cols: int = 3
        
        for i, info in enumerate(slides):
            card: SlideCard = SlideCard(info, self.current_lang)
            card.toggled.connect(self._update_stats)
            self.cards.append(card)
            self.grid_layout.addWidget(card, i // cols, i % cols)
    
    def _update_stats(self) -> None:
        r"""
        更新统计信息
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        
        if not self.ppt_data:
            self.sidebar.lbl_filename.setText(t['no_file'])
            self.sidebar.lbl_slide_count.setText("-")
            self.sidebar.lbl_sel_info.setText("")
            self.sidebar.btn_export.setText(t['export'])
            self.sidebar.btn_export.setEnabled(False)
            return
        
        total: int = len(self.ppt_data['slides'])
        sel: int = sum(1 for s in self.ppt_data['slides'] if s['selected'])
        
        self.sidebar.lbl_slide_count.setText(t['slide_count'].format(total))
        self.sidebar.lbl_sel_info.setText(t['selected'].format(sel, total))
        self.sidebar.btn_export.setEnabled(sel > 0)
        self.sidebar.btn_export.setText(
            t['export_n'].format(sel) if sel > 0 else t['export']
        )
    
    def choose_output_dir(self) -> None:
        r"""
        选择输出目录
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        path: str = QFileDialog.getExistingDirectory(
            self,
            t['select_output'],
            self.sidebar.entry_out_dir.text()
        )
        if path:
            self.sidebar.entry_out_dir.setText(path)
            self.sidebar.entry_out_dir.setToolTip(path)
    
    def select_all(self) -> None:
        r"""
        全选所有幻灯片
        """
        for card in self.cards:
            card.info['selected'] = True
            card._update_style()
        self._update_stats()
    
    def select_none(self) -> None:
        r"""
        取消全选
        """
        for card in self.cards:
            card.info['selected'] = False
            card._update_style()
        self._update_stats()
    
    def start_export(self) -> None:
        r"""
        开始导出
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        scale_map: dict[int, int] = {0: 0, 1: 1, 2: 2, 3: 3, 4: 5}
        scale: int = scale_map.get(self.sidebar.combo_scale.currentIndex(), 0)
        
        indices: list[int] = [
            c.info['index'] for c in self.cards if c.info['selected']
        ]
        out_dir: str = self.sidebar.entry_out_dir.text()
        
        if not indices or not self.ppt_data:
            return
        
        self.sidebar.btn_export.setEnabled(False)
        self.sidebar.btn_export.setText(t['exporting'])
        
        self.exporter = ExportThread(
            self.ppt_data['path'],
            indices,
            out_dir,
            scale
        )
        self.exporter.progress.connect(
            lambda c, tot: self.sidebar.btn_export.setText(f"{c} / {tot}")
        )
        self.exporter.finished.connect(self._on_export_finished)
        self.exporter.start()
    
    def _on_export_finished(
        self,
        success: bool,
        result: str,
        count: int
    ) -> None:
        r"""
        导出完成回调
        
        :param success: 是否成功
        :param result: 结果路径或错误信息
        :param count: 导出数量
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        self._update_stats()
        
        if success:
            dialog: CustomMessageDialog = CustomMessageDialog(
                self,
                t['export_success'],
                t['export_success_msg'].format(count, result),
                [
                    (t['open_folder'], 'open'),
                    (t['close'], 'close')
                ],
                "info"
            )
            dialog.exec()
            
            if dialog.get_clicked_button() == 'open':
                QDesktopServices.openUrl(QUrl.fromLocalFile(result))
        else:
            self._show_error_dialog(t['error'], t['export_failed'].format(result))
    
    def _show_error_dialog(self, title: str, message: str) -> None:
        r"""
        显示错误对话框
        
        :param title: 标题
        :param message: 错误消息
        """
        t: dict[str, str] = LANG_TEXTS[self.current_lang]
        dialog: CustomMessageDialog = CustomMessageDialog(
            self,
            title,
            message,
            [(t['close'], 'close')],
            "error"
        )
        dialog.exec()


# ==================== 程序入口 ====================

def main() -> None:
    r"""
    程序入口函数
    """
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app: QApplication = QApplication(sys.argv)
    app.setStyleSheet(GLOBAL_STYLESHEET)
    
    icon_path: str = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        '..',
        'logo.png'
    )
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window: MainWindow = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()