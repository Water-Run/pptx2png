import sys
import os
import shutil
import tempfile
import pythoncom
import win32com.client

# 核心库调用
try:
    import pptx2img
except ImportError:
    print("Error: pptx2img module missing.")
    pptx2img = None

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QScrollArea, QFileDialog,
    QMessageBox, QGridLayout, QFrame, QLineEdit, QStackedWidget,
    QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QPixmap, QFont, QIcon, QColor, QDesktopServices, QPainter, QPainterPath

# --- 多语言配置 ---
LANG_TEXTS = {
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

# --- 样式与配置 ---
COLOR_PPT_ORANGE = "#D24726"
COLOR_PPT_DARK = "#A4371E"
COLOR_BG_SIDEBAR = "#F3F3F3"
COLOR_BG_CONTENT = "#FFFFFF"
COLOR_TEXT_MAIN = "#2D2D2D"
COLOR_TEXT_SUB = "#757575"
BORDER_COLOR = "#C0C0C0"

# 现代化全局样式表
GLOBAL_STYLESHEET = f"""
    QWidget {{
        font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
        font-size: 14px;
        color: {COLOR_TEXT_MAIN};
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
    QRadioButton::indicator:checked::after {{
        content: '';
        width: 8px;
        height: 8px;
        border-radius: 4px;
        background: {COLOR_PPT_ORANGE};
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

# 现代化 ComboBox 样式（参考 PPT 风格）
MODERN_COMBOBOX_STYLE = f"""
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

# --- 工具函数：合并连续范围 ---
def merge_ranges(indices):
    """将索引列表合并为连续范围列表
    例: [1,2,3,5,6,8] -> [[1,3], [5,6], [8,8]]
    """
    if not indices:
        return []
    
    sorted_indices = sorted(set(indices))
    ranges = []
    start = sorted_indices[0]
    end = sorted_indices[0]
    
    for i in range(1, len(sorted_indices)):
        if sorted_indices[i] == end + 1:
            end = sorted_indices[i]
        else:
            ranges.append([start, end])
            start = sorted_indices[i]
            end = sorted_indices[i]
    
    ranges.append([start, end])
    return ranges

# --- 线程工作类 ---

class LoadThread(QThread):
    """加载预览图"""
    finished = pyqtSignal(bool, str, dict)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        pythoncom.CoInitialize()
        pres = None
        ppt_app = None
        temp_dir = None
        try:
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                ppt_app.WindowState = 2
            except: pass

            pres = ppt_app.Presentations.Open(os.path.abspath(self.path), ReadOnly=True, WithWindow=False)
            
            temp_dir = tempfile.mkdtemp(prefix="pptx2img_prev_")
            slides_info = []
            
            w = pres.PageSetup.SlideWidth
            h = pres.PageSetup.SlideHeight
            count = pres.Slides.Count
            
            preview_w = 640
            preview_h = int(preview_w * h / w)
            
            for i in range(1, count + 1):
                slide = pres.Slides(i)
                img_path = os.path.join(temp_dir, f"p_{i}.jpg")
                slide.Export(img_path, "JPG", preview_w, preview_h)
                slides_info.append({
                    'index': i,
                    'path': img_path,
                    'selected': True
                })
            
            data = {
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
                try: pres.Close()
                except: pass
            if ppt_app:
                try: ppt_app.Quit()
                except: pass
        finally:
            pythoncom.CoUninitialize()

class ExportThread(QThread):
    """优化的导出线程（合并连续范围）"""
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(bool, str, int)

    def __init__(self, ppt_path, indices, out_dir, scale):
        super().__init__()
        self.ppt_path = ppt_path
        self.indices = indices
        self.out_dir = out_dir
        self.scale = scale

    def run(self):
        pythoncom.CoInitialize()
        try:
            if not pptx2img:
                raise ImportError("pptx2img library not found!")

            total = len(self.indices)
            count = 0
            
            # 合并连续范围以优化速度
            ranges = merge_ranges(self.indices)
            
            lib_scale = self.scale if self.scale > 0 else None
            
            for slide_range in ranges:
                # 调用库函数导出范围
                pptx2img.topng(
                    pptx=self.ppt_path, 
                    output_dir=self.out_dir, 
                    slide_range=slide_range, 
                    scale=lib_scale
                )
                
                # 更新进度
                range_count = slide_range[1] - slide_range[0] + 1
                count += range_count
                self.progress.emit(count, total)
            
            self.finished.emit(True, self.out_dir, count)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished.emit(False, str(e), 0)
        finally:
            pythoncom.CoUninitialize()

# --- UI 组件 ---

class TitleBar(QWidget):
    """修复的橙色标题栏（居中软件名，右侧关闭按钮）"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setFixedHeight(40)
        self.setStyleSheet(f"background-color: {COLOR_PPT_ORANGE};")
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 左侧占位（保持居中）
        layout.addStretch(1)
        
        # 中间标题
        self.lbl_title = QLabel("pptx2img")
        self.lbl_title.setStyleSheet("""
            color: white; 
            font-weight: bold; 
            font-size: 16px; 
            background: transparent;
            padding: 0;
            margin: 0;
        """)
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_title)
        
        # 右侧占位 + 关闭按钮
        layout.addStretch(1)
        
        # 最小化按钮
        btn_min_style = """
            QPushButton { 
                background: transparent; 
                border: none; 
                color: white; 
                font-size: 18px;
                padding: 0;
                margin: 0;
            }
            QPushButton:hover { background: rgba(255,255,255,0.15); }
        """
        btn_min = QPushButton("─")
        btn_min.setFixedSize(45, 40)
        btn_min.setStyleSheet(btn_min_style)
        btn_min.clicked.connect(self.parent.showMinimized)
        layout.addWidget(btn_min)
        
        # 关闭按钮
        btn_close_style = """
            QPushButton { 
                background: transparent; 
                border: none; 
                color: white; 
                font-size: 18px;
                padding: 0;
                margin: 0;
            }
            QPushButton:hover { background: #E81123; }
        """
        btn_close = QPushButton("✕")
        btn_close.setFixedSize(45, 40)
        btn_close.setStyleSheet(btn_close_style)
        btn_close.clicked.connect(self.parent.close)
        layout.addWidget(btn_close)
        
        self.setLayout(layout)
        self._drag_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.parent.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos:
            self.parent.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()

class Sidebar(QWidget):
    """左侧控制区（含语言切换）"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setObjectName("Sidebar")
        self.setFixedWidth(260)
        self.setStyleSheet(f"#Sidebar {{ background-color: {COLOR_BG_SIDEBAR}; border-right: 1px solid {BORDER_COLOR}; }}")
        
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(20, 30, 20, 30)
        self.layout.setSpacing(10)
        
        # Logo
        self.lbl_logo = QLabel()
        self.lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_logo.setStyleSheet("background: transparent;")
        self.load_logo()
        self.layout.addWidget(self.lbl_logo)
        
        # 软件信息
        self.layout.addSpacing(5)
        self.lbl_name = QLabel()
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_name.setStyleSheet(f"font-weight: bold; font-size: 18px; color: {COLOR_TEXT_MAIN}; background: transparent;")
        self.layout.addWidget(self.lbl_name)
        
        self.lbl_author = QLabel()
        self.lbl_author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_author.setStyleSheet(f"font-size: 12px; color: {COLOR_TEXT_SUB}; background: transparent;")
        self.layout.addWidget(self.lbl_author)
        
        # GitHub 链接
        self.lbl_github = QLabel('<a href="https://github.com/Water-Run/pptx2img" style="color:#D24726; text-decoration:none; font-weight:bold;">GitHub</a>')
        self.lbl_github.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_github.setOpenExternalLinks(True)
        self.lbl_github.setStyleSheet("background: transparent;")
        self.lbl_github.setCursor(Qt.CursorShape.PointingHandCursor)
        self.layout.addWidget(self.lbl_github)
        
        # 分隔线
        line1 = QFrame()
        line1.setFrameShape(QFrame.Shape.HLine)
        line1.setStyleSheet(f"color: {BORDER_COLOR}; margin-top: 15px; margin-bottom: 15px;")
        self.layout.addWidget(line1)
        
        # 语言选择
        self.lbl_language = QLabel()
        self.lbl_language.setStyleSheet(f"font-weight: bold; color: {COLOR_PPT_ORANGE}; margin-bottom: 5px; background: transparent;")
        self.layout.addWidget(self.lbl_language)
        
        lang_layout = QHBoxLayout()
        lang_layout.setSpacing(15)
        self.radio_zh = QRadioButton()
        self.radio_en = QRadioButton()
        self.radio_zh.setChecked(True)
        
        self.lang_group = QButtonGroup()
        self.lang_group.addButton(self.radio_zh, 0)
        self.lang_group.addButton(self.radio_en, 1)
        self.lang_group.buttonClicked.connect(self.on_language_changed)
        
        lang_layout.addWidget(self.radio_zh)
        lang_layout.addWidget(self.radio_en)
        lang_layout.addStretch()
        self.layout.addLayout(lang_layout)
        
        self.layout.addSpacing(10)
        
        # 文件信息
        self.lbl_grp_info = QLabel()
        self.lbl_grp_info.setStyleSheet(f"font-weight: bold; color: {COLOR_PPT_ORANGE}; margin-bottom: 5px; background: transparent;")
        self.layout.addWidget(self.lbl_grp_info)
        
        self.lbl_filename = QLabel()
        self.lbl_filename.setWordWrap(True)
        self.lbl_filename.setStyleSheet("font-weight: bold; font-size: 13px; background: transparent;")
        self.layout.addWidget(self.lbl_filename)
        
        self.lbl_slide_count = QLabel("-")
        self.lbl_slide_count.setStyleSheet(f"color: {COLOR_TEXT_SUB}; font-size: 12px; background: transparent;")
        self.layout.addWidget(self.lbl_slide_count)
        
        self.layout.addSpacing(10)
        
        # 导出设置
        self.lbl_grp_settings = QLabel()
        self.lbl_grp_settings.setStyleSheet(f"font-weight: bold; color: {COLOR_PPT_ORANGE}; margin-bottom: 5px; background: transparent;")
        self.layout.addWidget(self.lbl_grp_settings)
        
        self.lbl_output = QLabel()
        self.layout.addWidget(self.lbl_output)
        
        path_layout = QHBoxLayout()
        self.entry_out_dir = QLineEdit()
        self.entry_out_dir.setReadOnly(True)
        self.btn_browse = QPushButton()
        self.btn_browse.clicked.connect(self.parent.choose_output_dir)
        path_layout.addWidget(self.entry_out_dir, 1)
        path_layout.addWidget(self.btn_browse)
        self.layout.addLayout(path_layout)
        
        self.lbl_scale = QLabel()
        self.layout.addWidget(self.lbl_scale)
        self.combo_scale = QComboBox()
        self.combo_scale.setStyleSheet(MODERN_COMBOBOX_STYLE)
        self.combo_scale.setCurrentIndex(0)  # 默认"显示"
        self.layout.addWidget(self.combo_scale)
        
        self.layout.addStretch()
        
        # 操作区
        self.lbl_sel_info = QLabel()
        self.lbl_sel_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sel_info.setStyleSheet("background: transparent;")
        self.layout.addWidget(self.lbl_sel_info)

        btn_layout = QHBoxLayout()
        self.btn_all = QPushButton()
        self.btn_none = QPushButton()
        self.btn_all.clicked.connect(self.parent.select_all)
        self.btn_none.clicked.connect(self.parent.select_none)
        btn_layout.addWidget(self.btn_all)
        btn_layout.addWidget(self.btn_none)
        self.layout.addLayout(btn_layout)
        
        self.btn_export = QPushButton()
        self.btn_export.setProperty("class", "PrimaryButton")
        self.btn_export.setFixedHeight(45)
        self.btn_export.clicked.connect(self.parent.start_export)
        self.btn_export.setEnabled(False)
        self.layout.addWidget(self.btn_export)
        
        self.setLayout(self.layout)

    def load_logo(self):
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'logo.png')
        if os.path.exists(logo_path):
            logo_pix = QPixmap(logo_path)
            self.lbl_logo.setPixmap(logo_pix.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            self.lbl_logo.setText("Logo")

    def on_language_changed(self):
        self.parent.update_language()

    def update_texts(self, lang):
        """更新所有文本"""
        t = LANG_TEXTS[lang]
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
        
        # 更新倍率选项
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
    """幻灯片卡片"""
    toggled = pyqtSignal()
    
    def __init__(self, info, lang='zh'):
        super().__init__()
        self.info = info
        self.lang = lang
        self.setFixedSize(220, 160)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 5)
        layout.setSpacing(5)
        
        self.lbl_img = QLabel()
        self.lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_img.setStyleSheet("background: #EEE; border-radius: 4px;")
        if os.path.exists(info['path']):
            pix = QPixmap(info['path'])
            self.lbl_img.setPixmap(pix.scaled(204, 115, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        self.lbl_idx = QLabel()
        self.lbl_idx.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(self.lbl_img)
        layout.addWidget(self.lbl_idx)
        self.setLayout(layout)
        self.update_ui()
    
    def set_language(self, lang):
        self.lang = lang
        self.update_ui()
        
    def update_ui(self):
        t = LANG_TEXTS[self.lang]
        self.lbl_idx.setText(t['slide_n'].format(self.info['index']))
        
        if self.info['selected']:
            self.setStyleSheet(f"""
                SlideCard {{
                    background: white;
                    border: 3px solid {COLOR_PPT_ORANGE};
                    border-radius: 6px;
                }}
            """)
            self.lbl_idx.setStyleSheet(f"color: {COLOR_PPT_ORANGE}; font-weight: bold; background: transparent; font-size: 12px;")
        else:
            self.setStyleSheet(f"""
                SlideCard {{
                    background: white;
                    border: 1px solid {BORDER_COLOR};
                    border-radius: 6px;
                }}
                SlideCard:hover {{ border-color: {COLOR_PPT_ORANGE}; }}
            """)
            self.lbl_idx.setStyleSheet(f"color: {COLOR_TEXT_SUB}; background: transparent; font-size: 12px;")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.info['selected'] = not self.info['selected']
            self.update_ui()
            self.toggled.emit()

class PlaceholderWidget(QWidget):
    clicked = pyqtSignal()
    def __init__(self, parent=None, lang='zh'):
        super().__init__(parent)
        self.lang = lang
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        lbl_icon = QLabel("＋")
        lbl_icon.setStyleSheet(f"""
            QLabel {{
                font-size: 80px; 
                color: {COLOR_PPT_ORANGE};
                border: 3px dashed {COLOR_PPT_ORANGE};
                border-radius: 20px;
                padding: 20px 60px;
                background: #FFF5F3;
            }}
            QLabel:hover {{ background: #FFEBE8; }}
        """)
        
        self.lbl_text = QLabel()
        self.lbl_text.setStyleSheet("font-size: 18px; color: #555;")
        
        layout.addWidget(lbl_icon)
        layout.addWidget(self.lbl_text)
        self.setLayout(layout)
        self.update_text()

    def set_language(self, lang):
        self.lang = lang
        self.update_text()

    def update_text(self):
        self.lbl_text.setText(LANG_TEXTS[self.lang]['placeholder'])

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

# --- 主窗口 ---

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_lang = 'zh'  # 默认中文
        self.ppt_data = None
        self.cards = []
        
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(1100, 750)
        
        screen = QApplication.primaryScreen().availableGeometry()
        self.move((screen.width() - self.width()) // 2, (screen.height() - self.height()) // 2)
        
        self.setup_ui()
        self.update_language()

    def setup_ui(self):
        self.main_container = QFrame()
        self.main_container.setStyleSheet(f"""
            QFrame#MainContainer {{
                background-color: {COLOR_BG_CONTENT};
                border: 2px solid {COLOR_PPT_ORANGE};
                border-radius: 10px;
            }}
        """)
        self.main_container.setObjectName("MainContainer")
        self.setCentralWidget(self.main_container)
        
        root_layout = QVBoxLayout(self.main_container)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        
        self.title_bar = TitleBar(self)
        root_layout.addWidget(self.title_bar)
        
        body_layout = QHBoxLayout()
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(0)
        
        self.sidebar = Sidebar(self)
        body_layout.addWidget(self.sidebar)
        
        self.right_stack = QStackedWidget()
        self.right_stack.setStyleSheet(f"background-color: {COLOR_BG_CONTENT}; border-bottom-right-radius: 10px;")
        
        self.page_empty = PlaceholderWidget(lang=self.current_lang)
        self.page_empty.clicked.connect(self.open_file_dialog)
        self.right_stack.addWidget(self.page_empty)
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.grid_container = QWidget()
        self.grid_layout = QGridLayout()
        self.grid_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.grid_layout.setSpacing(25)
        self.grid_layout.setContentsMargins(30, 30, 30, 30)
        self.grid_container.setLayout(self.grid_layout)
        self.scroll_area.setWidget(self.grid_container)
        self.right_stack.addWidget(self.scroll_area)
        
        self.loading_lbl = QLabel()
        self.loading_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.loading_lbl.setStyleSheet(f"font-size: 24px; color: {COLOR_PPT_ORANGE}; font-weight: bold;")
        self.right_stack.addWidget(self.loading_lbl)
        
        body_layout.addWidget(self.right_stack)
        root_layout.addLayout(body_layout)
        
        self.setAcceptDrops(True)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.setFillRule(Qt.FillRule.WindingFill)
        path.addRoundedRect(0, 0, self.width(), 40, 10, 10)
        painter.fillPath(path, QColor(COLOR_PPT_ORANGE))
        painter.fillRect(0, 20, self.width(), 20, QColor(COLOR_PPT_ORANGE))

    def update_language(self):
        """切换语言"""
        self.current_lang = 'zh' if self.sidebar.radio_zh.isChecked() else 'en'
        t = LANG_TEXTS[self.current_lang]
        
        # 更新侧边栏
        self.sidebar.update_texts(self.current_lang)
        
        # 更新加载文本
        self.loading_lbl.setText(t['loading'])
        
        # 更新占位页面
        self.page_empty.set_language(self.current_lang)
        
        # 更新卡片
        for card in self.cards:
            card.set_language(self.current_lang)
        
        # 更新统计信息
        self.update_stats()
        
        # 更新文件名显示
        if self.ppt_data:
            self.sidebar.lbl_filename.setText(os.path.basename(self.ppt_data['path']))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): 
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.pptx', '.ppt')):
                self.load_file(path)
                break
        event.acceptProposedAction()

    def open_file_dialog(self):
        t = LANG_TEXTS[self.current_lang]
        path, _ = QFileDialog.getOpenFileName(self, t['select_ppt'], "", "PowerPoint (*.pptx *.ppt)")
        if path: 
            self.load_file(path)

    def load_file(self, path):
        self.right_stack.setCurrentIndex(2)
        self.sidebar.btn_export.setEnabled(False)
        self.loader = LoadThread(path)
        self.loader.finished.connect(self.on_load_finished)
        self.loader.start()

    def on_load_finished(self, success, msg, data):
        t = LANG_TEXTS[self.current_lang]
        if success:
            self.ppt_data = data
            self.sidebar.lbl_filename.setText(os.path.basename(data['path']))
            default_out = os.path.join(os.path.dirname(data['path']), "pptx2img")
            self.sidebar.entry_out_dir.setText(default_out)
            self.sidebar.entry_out_dir.setToolTip(default_out)
            
            self.populate_grid()
            self.right_stack.setCurrentIndex(1)
            self.sidebar.btn_export.setEnabled(True)
        else:
            QMessageBox.critical(self, t['error'], msg)
            self.right_stack.setCurrentIndex(0)
        self.update_stats()

    def populate_grid(self):
        for i in reversed(range(self.grid_layout.count())): 
            self.grid_layout.itemAt(i).widget().setParent(None)
        self.cards = []
        slides = self.ppt_data['slides']
        cols = 3
        for i, info in enumerate(slides):
            card = SlideCard(info, self.current_lang)
            card.toggled.connect(self.update_stats)
            self.cards.append(card)
            self.grid_layout.addWidget(card, i // cols, i % cols)

    def update_stats(self):
        if not self.ppt_data: 
            return
        t = LANG_TEXTS[self.current_lang]
        total = len(self.ppt_data['slides'])
        sel = sum(1 for s in self.ppt_data['slides'] if s['selected'])
        self.sidebar.lbl_slide_count.setText(t['slide_count'].format(total))
        self.sidebar.lbl_sel_info.setText(t['selected'].format(sel, total))
        self.sidebar.btn_export.setEnabled(sel > 0)
        self.sidebar.btn_export.setText(t['export_n'].format(sel) if sel > 0 else t['export'])
        
        # 如果没有文件，显示默认文本
        if not self.ppt_data:
            self.sidebar.lbl_filename.setText(t['no_file'])

    def choose_output_dir(self):
        t = LANG_TEXTS[self.current_lang]
        path = QFileDialog.getExistingDirectory(self, t['select_output'], self.sidebar.entry_out_dir.text())
        if path: 
            self.sidebar.entry_out_dir.setText(path)
            self.sidebar.entry_out_dir.setToolTip(path)

    def select_all(self):
        for c in self.cards: 
            c.info['selected'] = True
            c.update_ui()
        self.update_stats()

    def select_none(self):
        for c in self.cards: 
            c.info['selected'] = False
            c.update_ui()
        self.update_stats()

    def start_export(self):
        t = LANG_TEXTS[self.current_lang]
        scale_map = {0: 0, 1: 1, 2: 2, 3: 3, 4: 5}
        scale = scale_map.get(self.sidebar.combo_scale.currentIndex(), 0)
        
        indices = [c.info['index'] for c in self.cards if c.info['selected']]
        out_dir = self.sidebar.entry_out_dir.text()
        
        if not indices: 
            return

        self.sidebar.btn_export.setEnabled(False)
        self.sidebar.btn_export.setText(t['exporting'])
        
        self.exporter = ExportThread(self.ppt_data['path'], indices, out_dir, scale)
        self.exporter.progress.connect(lambda c, tot: self.sidebar.btn_export.setText(f"{c} / {tot}"))
        self.exporter.finished.connect(self.on_export_finished)
        self.exporter.start()

    def on_export_finished(self, success, res, count):
        t = LANG_TEXTS[self.current_lang]
        self.update_stats()
        if success:
            box = QMessageBox(self)
            box.setWindowTitle(t['export_success'])
            box.setText(t['export_success_msg'].format(count, res))
            box.setIcon(QMessageBox.Icon.Information)
            btn_open = box.addButton(t['open_folder'], QMessageBox.ButtonRole.ActionRole)
            box.addButton(t['close'], QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() == btn_open:
                QDesktopServices.openUrl(QUrl.fromLocalFile(res))
        else:
            QMessageBox.critical(self, t['error'], t['export_failed'].format(res))

if __name__ == "__main__":
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    app.setStyleSheet(GLOBAL_STYLESHEET)
    
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'logo.png')
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())