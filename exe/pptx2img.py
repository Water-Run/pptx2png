import sys
import os
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QComboBox, 
                             QScrollArea, QFileDialog, QMessageBox, QToolTip)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QImage, QPalette, QColor, QFont, QCursor
import win32com.client
from PIL import Image
import io
import traceback

# PowerPoint Orange Color
POWERPOINT_ORANGE = "#D24726"

class LoadingThread(QThread):
    """Thread for loading PowerPoint presentation"""
    finished = pyqtSignal(bool, str, object)  # success, message, presentation_data
    
    def __init__(self, pptx_path):
        super().__init__()
        self.pptx_path = pptx_path
        
    def run(self):
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(self.pptx_path, WithWindow=False)
            
            # Extract slide previews
            slides_data = []
            temp_dir = os.path.join(os.path.dirname(self.pptx_path), ".pptx2img_temp")
            os.makedirs(temp_dir, exist_ok=True)
            
            for i in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides(i)
                temp_path = os.path.join(temp_dir, f"preview_{i}.png")
                slide.Export(temp_path, "PNG", 320, 180)
                slides_data.append({
                    'index': i,
                    'preview_path': temp_path,
                    'selected': True  # Default: all selected
                })
            
            self.finished.emit(True, "Success", {
                'powerpoint': powerpoint,
                'presentation': presentation,
                'slides': slides_data,
                'pptx_path': self.pptx_path,
                'temp_dir': temp_dir
            })
        except Exception as e:
            error_msg = f"{str(e)}\n{traceback.format_exc()}"
            self.finished.emit(False, error_msg, None)


class ExportThread(QThread):
    """Thread for exporting slides to PNG"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, presentation, slides, output_dir, scale):
        super().__init__()
        self.presentation = presentation
        self.slides = slides
        self.output_dir = output_dir
        self.scale = scale
        
    def run(self):
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            
            slide_width = self.presentation.PageSetup.SlideWidth
            slide_height = self.presentation.PageSetup.SlideHeight
            
            target_w = int(slide_width * self.scale)
            target_h = int(slide_height * self.scale)
            
            count = 0
            for slide_data in self.slides:
                if slide_data['selected']:
                    i = slide_data['index']
                    slide = self.presentation.Slides(i)
                    image_name = f"Slide_{i}.png"
                    image_path = os.path.join(self.output_dir, image_name)
                    
                    slide.Export(image_path, "PNG", target_w, target_h)
                    count += 1
                    self.progress.emit(f"Exported: {image_name}")
            
            self.finished.emit(True, f"Successfully exported {count} slides to:\n{self.output_dir}")
        except Exception as e:
            self.finished.emit(False, f"Export failed:\n{str(e)}")


class SlidePreviewWidget(QWidget):
    """Widget for individual slide preview thumbnail"""
    clicked = pyqtSignal(int)
    
    def __init__(self, slide_data, parent=None):
        super().__init__(parent)
        self.slide_data = slide_data
        self.setFixedSize(160, 120)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Preview image
        self.image_label = QLabel()
        pixmap = QPixmap(self.slide_data['preview_path'])
        self.image_label.setPixmap(pixmap.scaled(150, 85, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.image_label.setAlignment(Qt.AlignCenter)
        
        # Slide number
        self.number_label = QLabel(f"Slide {self.slide_data['index']}")
        self.number_label.setAlignment(Qt.AlignCenter)
        self.number_label.setStyleSheet(f"color: {POWERPOINT_ORANGE}; font-weight: bold;")
        
        layout.addWidget(self.image_label)
        layout.addWidget(self.number_label)
        self.setLayout(layout)
        
        self.update_style()
        
    def update_style(self):
        if self.slide_data['selected']:
            self.setStyleSheet(f"""
                SlidePreviewWidget {{
                    background-color: {POWERPOINT_ORANGE};
                    border: 3px solid {POWERPOINT_ORANGE};
                    border-radius: 5px;
                }}
            """)
            self.number_label.setStyleSheet("color: white; font-weight: bold;")
        else:
            self.setStyleSheet("""
                SlidePreviewWidget {
                    background-color: #f0f0f0;
                    border: 2px solid #cccccc;
                    border-radius: 5px;
                }
            """)
            self.number_label.setStyleSheet(f"color: {POWERPOINT_ORANGE}; font-weight: bold;")
    
    def mousePressEvent(self, event):
        self.clicked.emit(self.slide_data['index'])
        
    def toggle_selection(self):
        self.slide_data['selected'] = not self.slide_data['selected']
        self.update_style()


class MainWindow(QMainWindow):
    def __init__(self, pptx_path, presentation_data):
        super().__init__()
        self.pptx_path = pptx_path
        self.presentation_data = presentation_data
        self.slides = presentation_data['slides']
        self.current_slide_index = 0
        
        self.setWindowTitle("pptx2img")
        self.setGeometry(100, 100, 1200, 800)
        
        self.setup_ui()
        
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Top Bar
        top_bar = self.create_top_bar()
        
        # Content Area
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(10, 10, 10, 10)
        
        # Left: Main Preview
        left_widget = self.create_main_preview()
        
        # Right: Thumbnail List
        right_widget = self.create_thumbnail_list()
        
        content_layout.addWidget(left_widget, 7)
        content_layout.addWidget(right_widget, 3)
        
        # Bottom Bar
        bottom_bar = self.create_bottom_bar()
        
        main_layout.addWidget(top_bar)
        main_layout.addLayout(content_layout)
        main_layout.addWidget(bottom_bar)
        
        central_widget.setLayout(main_layout)
        
    def create_top_bar(self):
        top_bar = QWidget()
        top_bar.setStyleSheet(f"background-color: {POWERPOINT_ORANGE};")
        top_bar.setFixedHeight(60)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(15, 10, 15, 10)
        
        # Logo (using text as placeholder)
        logo_label = QLabel("📊")
        logo_label.setStyleSheet("color: white; font-size: 32px;")
        
        # Title
        title_label = QLabel("pptx2img")
        title_label.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
        
        # Spacer
        layout.addWidget(logo_label)
        layout.addWidget(title_label)
        layout.addStretch()
        
        # Author and GitHub
        author_label = QLabel("by WaterRun")
        author_label.setStyleSheet("color: white; font-size: 14px;")
        
        github_label = QLabel('<a href="https://github.com/Water-Run/pptx2img" style="color: white;">GitHub</a>')
        github_label.setOpenExternalLinks(True)
        github_label.setStyleSheet("font-size: 14px;")
        
        layout.addWidget(author_label)
        layout.addSpacing(10)
        layout.addWidget(github_label)
        
        top_bar.setLayout(layout)
        return top_bar
    
    def create_main_preview(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        self.main_preview_label = QLabel()
        self.main_preview_label.setAlignment(Qt.AlignCenter)
        self.main_preview_label.setStyleSheet("background-color: #2b2b2b; border: 2px solid #cccccc;")
        self.main_preview_label.setMinimumSize(600, 400)
        
        layout.addWidget(self.main_preview_label)
        widget.setLayout(layout)
        
        self.update_main_preview()
        return widget
    
    def create_thumbnail_list(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        label = QLabel("Slides")
        label.setStyleSheet(f"color: {POWERPOINT_ORANGE}; font-size: 16px; font-weight: bold;")
        label.setAlignment(Qt.AlignCenter)
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        scroll_content = QWidget()
        self.thumbnail_layout = QVBoxLayout()
        self.thumbnail_layout.setSpacing(10)
        
        self.thumbnail_widgets = []
        for slide_data in self.slides:
            thumb = SlidePreviewWidget(slide_data)
            thumb.clicked.connect(self.on_thumbnail_clicked)
            self.thumbnail_widgets.append(thumb)
            self.thumbnail_layout.addWidget(thumb)
        
        self.thumbnail_layout.addStretch()
        scroll_content.setLayout(self.thumbnail_layout)
        scroll_area.setWidget(scroll_content)
        
        layout.addWidget(label)
        layout.addWidget(scroll_area)
        widget.setLayout(layout)
        
        return widget
    
    def create_bottom_bar(self):
        bottom_bar = QWidget()
        bottom_bar.setStyleSheet("background-color: #f5f5f5;")
        bottom_bar.setFixedHeight(80)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(15, 10, 15, 10)
        
        # Go Button
        self.go_button = QPushButton("GO")
        self.go_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {POWERPOINT_ORANGE};
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px 30px;
            }}
            QPushButton:hover {{
                background-color: #b33b1e;
            }}
        """)
        self.go_button.setCursor(QCursor(Qt.PointingHandCursor))
        self.go_button.clicked.connect(self.on_go_clicked)
        self.go_button.setToolTip("<i>Ctrl+A: Select All<br>Ctrl+Shift+A: Deselect All</i>")
        
        # Quality Selector
        quality_label = QLabel("Quality:")
        quality_label.setStyleSheet("font-size: 14px;")
        
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["DISPLAY", "1.0x", "2.0x", "3.0x", "5.0x", "10.0x"])
        self.quality_combo.setCurrentIndex(0)
        self.quality_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                font-size: 14px;
                border: 2px solid #cccccc;
                border-radius: 3px;
            }
        """)
        
        # Output Info
        output_dir = os.path.join(os.path.dirname(self.pptx_path), "pptx2img")
        output_label = QLabel(f"Output: {output_dir}")
        output_label.setStyleSheet("font-size: 12px; color: #666666;")
        
        layout.addWidget(self.go_button)
        layout.addSpacing(20)
        layout.addWidget(quality_label)
        layout.addWidget(self.quality_combo)
        layout.addStretch()
        layout.addWidget(output_label)
        
        bottom_bar.setLayout(layout)
        return bottom_bar
    
    def update_main_preview(self):
        if 0 <= self.current_slide_index < len(self.slides):
            preview_path = self.slides[self.current_slide_index]['preview_path']
            pixmap = QPixmap(preview_path)
            scaled_pixmap = pixmap.scaled(
                self.main_preview_label.width() - 20,
                self.main_preview_label.height() - 20,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.main_preview_label.setPixmap(scaled_pixmap)
    
    def on_thumbnail_clicked(self, index):
        self.current_slide_index = index - 1
        self.update_main_preview()
        
        # Toggle selection
        self.thumbnail_widgets[self.current_slide_index].toggle_selection()
    
    def on_go_clicked(self):
        # Get selected slides
        selected_slides = [s for s in self.slides if s['selected']]
        
        if not selected_slides:
            QMessageBox.warning(self, "No Selection", "Please select at least one slide to export.")
            return
        
        # Get scale factor
        quality_text = self.quality_combo.currentText()
        if quality_text == "DISPLAY":
            scale = None  # Auto-resolution mode
        else:
            scale = float(quality_text.replace('x', ''))
        
        # Prepare output directory
        output_dir = os.path.join(os.path.dirname(self.pptx_path), "pptx2img")
        
        # Create export thread
        if scale is None:
            # Use display mode (2x default)
            import ctypes
            try:
                user32 = ctypes.windll.user32
                screen_w = user32.GetSystemMetrics(0)
                screen_h = user32.GetSystemMetrics(1)
                screen_long = max(screen_w, screen_h)
                
                slide_width = self.presentation_data['presentation'].PageSetup.SlideWidth
                slide_height = self.presentation_data['presentation'].PageSetup.SlideHeight
                slide_ratio = slide_width / slide_height
                
                target_long = screen_long * 2
                if slide_ratio >= 1:
                    scale = target_long / slide_width
                else:
                    scale = target_long / slide_height
            except:
                scale = 2.0
        
        self.export_thread = ExportThread(
            self.presentation_data['presentation'],
            self.slides,
            output_dir,
            scale
        )
        self.export_thread.progress.connect(self.on_export_progress)
        self.export_thread.finished.connect(self.on_export_finished)
        
        self.go_button.setEnabled(False)
        self.go_button.setText("Exporting...")
        self.export_thread.start()
    
    def on_export_progress(self, message):
        print(message)
    
    def on_export_finished(self, success, message):
        self.go_button.setEnabled(True)
        self.go_button.setText("GO")
        
        if success:
            QMessageBox.information(self, "Export Complete", message)
        else:
            QMessageBox.critical(self, "Export Failed", message)
    
    def keyPressEvent(self, event):
        # Ctrl+A: Select All
        if event.modifiers() == Qt.ControlModifier and event.key() == Qt.Key_A:
            for thumb in self.thumbnail_widgets:
                thumb.slide_data['selected'] = True
                thumb.update_style()
        
        # Ctrl+Shift+A: Deselect All
        elif event.modifiers() == (Qt.ControlModifier | Qt.ShiftModifier) and event.key() == Qt.Key_A:
            for thumb in self.thumbnail_widgets:
                thumb.slide_data['selected'] = False
                thumb.update_style()
    
    def closeEvent(self, event):
        # Cleanup
        try:
            self.presentation_data['presentation'].Close()
            
            # Clean temp directory
            import shutil
            temp_dir = self.presentation_data['temp_dir']
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except:
            pass
        
        event.accept()


def main():
    app = QApplication(sys.argv)
    
    print("Select a PowerPoint presentation (.pptx) to export")
    print("选取需要导出的.pptx演示文稿")
    
    # Wait 1.5 seconds before showing file dialog
    QTimer.singleShot(1500, lambda: None)
    time.sleep(1.5)
    
    # File dialog
    pptx_path, _ = QFileDialog.getOpenFileName(
        None,
        "Select PowerPoint File",
        "",
        "PowerPoint Files (*.pptx)"
    )
    
    if not pptx_path:
        print("No file selected. Exiting.")
        sys.exit(0)
    
    # Loading window
    loading_window = QMessageBox()
    loading_window.setWindowTitle("Loading")
    loading_window.setText("Loading presentation, please wait...")
    loading_window.setStandardButtons(QMessageBox.NoButton)
    loading_window.show()
    
    # Start loading thread
    def on_loading_finished(success, message, data):
        loading_window.close()
        
        if success:
            print("Loading successful: Entering GUI")
            window = MainWindow(pptx_path, data)
            window.show()
        else:
            print("Loading failed")
            print("加载失败")
            print(f"Error: {message}")
            QMessageBox.critical(None, "Loading Failed", f"Failed to load presentation:\n{message}")
            sys.exit(1)
    
    loading_thread = LoadingThread(pptx_path)
    loading_thread.finished.connect(on_loading_finished)
    loading_thread.start()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()