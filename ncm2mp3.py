import sys
import os
import threading
from queue import Queue, Empty
from ncmdump import dump
# 确保PyQt6已正确安装，并且环境变量配置无误
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QLabel, QPushButton, QProgressBar, QFileDialog,
                             QListWidget, QMessageBox, QListWidgetItem, QHBoxLayout,
                             QRadioButton, QButtonGroup, QScrollArea)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, pyqtSlot, QMetaObject, Q_ARG
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from mutagen.id3 import ID3, TRCK, TIT2, TPE1, TALB, APIC

class SignalManager(QObject):
    progress_updated = pyqtSignal(str, int)
    conversion_complete = pyqtSignal(str)
    error_occurred = pyqtSignal(str, str)

class NCMConverter:
    def __init__(self):
        pass

    def decrypt_file(self, file_path, output_format='mp3'):
        try:
            # 使用ncmdump库解密NCM文件
            flac_path = os.path.splitext(file_path)[0] + '.flac'
            mp3_path = os.path.splitext(file_path)[0] + '.mp3'
            
            # 如果目标文件已存在，先删除
            if os.path.exists(flac_path):
                os.remove(flac_path)
            if os.path.exists(mp3_path):
                os.remove(mp3_path)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(flac_path)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 调用dump函数进行NCM解密，生成FLAC文件
            meta_data = dump(file_path, flac_path)
            
            # 根据选择的格式进行处理
            if output_format == 'mp3' or output_format == 'both':
                # 将FLAC转换为MP3，使用subprocess来执行ffmpeg，并重定向输出
                import subprocess
                ffmpeg_cmd = ['ffmpeg', '-loglevel', 'error', '-i', flac_path, '-y', 
                            '-acodec', 'libmp3lame', '-ab', '320k', mp3_path]
                subprocess.run(ffmpeg_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                if output_format == 'mp3':
                    # 如果只需要MP3，删除FLAC文件
                    os.remove(flac_path)
                    return mp3_path, meta_data
                else:
                    # 同时保留两种格式
                    return [flac_path, mp3_path], meta_data
            else:
                # 只保留FLAC格式
                return flac_path, meta_data

        except Exception as e:
            # 清理可能存在的临时文件
            if os.path.exists(flac_path):
                os.remove(flac_path)
            if os.path.exists(mp3_path):
                os.remove(mp3_path)
            raise Exception(f'文件转换失败: {str(e)}')

class DropArea(QWidget):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        layout = QVBoxLayout()
        self.label = QLabel('拖放NCM文件到这里')
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()
                if url.toLocalFile().lower().endswith('.ncm')]
        if files:
            self.files_dropped.emit(files)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('NCM转MP3工具')
        self.setMinimumSize(800, 600)
        self.setup_ui()
        self.converter = NCMConverter()
        self.signal_manager = SignalManager()
        self.conversion_queue = Queue()
        self.files_converting = set()
        self.total_files = 0
        self.completed_files = 0
        self.converted_files = []

        self.signal_manager.progress_updated.connect(self.update_progress)
        self.signal_manager.conversion_complete.connect(self.conversion_completed)
        self.signal_manager.error_occurred.connect(self.show_error)

    def setup_ui(self):
        # 按钮样式
        button_style = """QPushButton {
            background-color: #409EFF;
            color: white;
            border: none;
            padding: 8px 20px;
            border-radius: 4px;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #66B1FF;
        }
        QPushButton:pressed {
            background-color: #3a8ee6;
        }
        QPushButton[class='danger'] {
            background-color: #F56C6C;
        }
        QPushButton[class='danger']:hover {
            background-color: #f78989;
        }
        """

        # 进度条样式
        progress_style = """QProgressBar {
            border: 1px solid #E4E7ED;
            border-radius: 3px;
            background-color: #EBEEF5;
            text-align: center;
            height: 20px;
        }
        QProgressBar::chunk {
            background-color: #409EFF;
            border-radius: 2px;
        }
        """

        # 文件列表样式
        list_style = """QListWidget {
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            background-color: white;
            padding: 5px;
        }
        QListWidget::item {
            border-bottom: 1px solid #EBEEF5;
            padding: 8px 5px;
        }
        QListWidget::item:last {
            border-bottom: none;
        }
        QListWidget QScrollBar:vertical {
            border: none;
            background: #F2F6FC;
            width: 6px;
            margin: 0px 0px 0px 0px;
        }
        QListWidget QScrollBar::handle:vertical {
            background: #C0C4CC;
            border-radius: 3px;
            min-height: 20px;
        }
        QListWidget QScrollBar::add-line:vertical, QListWidget QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """

        # 单选按钮样式
        radio_style = """QRadioButton {
            spacing: 5px;
            color: #606266;
        }
        QRadioButton::indicator {
            width: 18px;
            height: 18px;
        }
        QRadioButton::indicator:unchecked {
            border: 2px solid #DCDFE6;
            background-color: white;
            border-radius: 9px;
        }
        QRadioButton::indicator:checked {
            border: 2px solid #409EFF;
            background-color: #409EFF;
            border-radius: 9px;
        }
        """

        # 应用样式
        self.setStyleSheet(button_style + progress_style + list_style + radio_style)

        # 设置窗口样式
        self.setStyleSheet(self.styleSheet() + """
            QMainWindow {
                background-color: #F2F6FC;
            }
            QLabel {
                color: #606266;
                font-size: 14px;
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)  # 设置边距
        layout.setSpacing(15)  # 设置组件间距

        # 拖放区域
        self.drop_area = DropArea()
        self.drop_area.files_dropped.connect(self.add_files)
        layout.addWidget(self.drop_area)

        # 转换格式选择
        format_group = QWidget()
        format_layout = QHBoxLayout(format_group)
        self.format_group = QButtonGroup()

        self.flac_only_radio = QRadioButton('仅转换为FLAC')
        self.mp3_only_radio = QRadioButton('仅转换为MP3')
        self.both_formats_radio = QRadioButton('同时生成FLAC和MP3')

        self.format_group.addButton(self.flac_only_radio)
        self.format_group.addButton(self.mp3_only_radio)
        self.format_group.addButton(self.both_formats_radio)

        format_layout.addWidget(self.flac_only_radio)
        format_layout.addWidget(self.mp3_only_radio)
        format_layout.addWidget(self.both_formats_radio)

        # 默认选择MP3格式
        self.mp3_only_radio.setChecked(True)

        layout.addWidget(format_group)

        # 文件列表
        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        # 转换信息显示区域
        # 信息标签和滚动区域样式
        info_label_style = """QLabel {
            background-color: white;
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            padding: 10px;
            color: #606266;
            font-family: 'Microsoft YaHei';
            font-size: 12px;
            line-height: 1.5;
        }"""

        scroll_area_style = """QScrollArea {
            border: 1px solid #DCDFE6;
            border-radius: 4px;
            background-color: white;
        }
        QScrollBar:vertical {
            border: none;
            background: #F2F6FC;
            width: 6px;
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            background: #C0C4CC;
            border-radius: 3px;
            min-height: 20px;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }"""

        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        self.info_label.setMinimumHeight(100)  # 设置最小高度
        self.info_label.setStyleSheet(info_label_style)
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.info_label)
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(200)  # 增加滚动区域的最小高度
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)  # 根据需要显示垂直滚动条
        scroll_area.setStyleSheet(scroll_area_style)
        layout.addWidget(scroll_area)

        # 按钮
        button_layout = QHBoxLayout()  # 改为水平布局
        button_layout.setSpacing(10)  # 设置按钮之间的间距
        self.select_button = QPushButton('选择文件')
        self.select_button.setFixedSize(120, 36)  # 设置固定大小
        self.select_button.clicked.connect(self.select_files)
        self.convert_button = QPushButton('开始转换')
        self.convert_button.setFixedSize(120, 36)  # 设置固定大小
        self.convert_button.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.convert_button)
        layout.addLayout(button_layout)

        # 进度条
        self.progress_bars = {}

    def add_files(self, files):
        for file in files:
            if file not in self.files_converting and os.path.exists(file):
                # 创建列表项
                item = QListWidgetItem()
                self.file_list.addItem(item)
                
                # 创建容器widget和水平布局
                container = QWidget()
                layout = QHBoxLayout(container)
                layout.setContentsMargins(5, 2, 5, 2)  # 设置边距
                
                # 创建文件信息标签
                file_info = QLabel(f"文件名: {os.path.basename(file)}\n路径: {file}")
                file_info.setWordWrap(True)  # 允许文本换行
                file_info.setFixedHeight(60)  # 增加高度以容纳两行文本
                layout.addWidget(file_info, stretch=7)  # 分配70%的空间
                
                # 创建进度条
                progress_bar = QProgressBar()
                progress_bar.setRange(0, 100)
                progress_bar.setFixedHeight(20)  # 设置进度条高度
                layout.addWidget(progress_bar, stretch=2)  # 减少进度条占用空间
                
                # 创建删除按钮
                delete_button = QPushButton('删除')
                delete_button.setFixedSize(80, 28)  # 增加按钮宽度
                delete_button.setProperty('class', 'danger')  # 设置危险按钮样式
                delete_button.clicked.connect(lambda checked, f=file: self.remove_file(f))
                delete_button.setCursor(Qt.CursorShape.PointingHandCursor)  # 设置鼠标悬停样式
                layout.addWidget(delete_button)
                
                # 保存进度条引用
                self.progress_bars[file] = progress_bar
                
                # 设置列表项的大小
                item.setSizeHint(container.sizeHint())
                self.file_list.setItemWidget(item, container)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, '选择NCM文件', '',
                                              'NCM Files (*.ncm)')
        self.add_files(files)

    def start_conversion(self):
        if self.file_list.count() == 0:
            QMessageBox.warning(self, '警告', '请先添加要转换的文件')
            return

        # 验证所有文件路径的有效性
        invalid_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if isinstance(widget, QWidget):
                file_info = widget.findChild(QLabel)
                if file_info:
                    file_path = file_info.text().split('路径: ')[-1]
                    if not os.path.exists(file_path):
                        invalid_files.append(os.path.basename(file_path))

        if invalid_files:
            QMessageBox.critical(self, '错误',
                               f'以下文件不存在或无法访问:\n{"".join(invalid_files)}')
            return

        self.total_files = self.file_list.count()
        self.completed_files = 0
        self.converted_files = []

        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if isinstance(widget, QWidget):
                file_info = widget.findChild(QLabel)
                if file_info:
                    file_path = file_info.text().split('路径: ')[-1]
                    if file_path not in self.files_converting:
                        self.conversion_queue.put(file_path)

        # 启动转换线程
        max_threads = min(os.cpu_count(), 4)  # 最多使用4个线程
        for _ in range(max_threads):
            thread = threading.Thread(target=self.conversion_worker)
            thread.daemon = True
            thread.start()

    def conversion_worker(self):
        while True:
            try:
                file_path = self.conversion_queue.get_nowait()
            except Empty:
                break

            if file_path in self.files_converting:
                continue

            self.files_converting.add(file_path)
            try:
                # 获取选择的输出格式
                output_format = 'flac' if self.flac_only_radio.isChecked() else \
                               'mp3' if self.mp3_only_radio.isChecked() else 'both'

                # 转换文件
                output_paths, meta_data = self.converter.decrypt_file(file_path, output_format)
                
                # 使用QMetaObject.invokeMethod确保UI更新在主线程中执行
                QMetaObject.invokeMethod(self, 'update_progress',
                                      Qt.ConnectionType.QueuedConnection,
                                      Q_ARG(str, file_path),
                                      Q_ARG(int, 100))

                # 处理输出路径列表
                if not isinstance(output_paths, list):
                    output_paths = [output_paths]

                # 写入元数据到所有MP3文件
                for output_path in output_paths:
                    if output_path.lower().endswith('.mp3') and meta_data:
                        try:
                            audio = ID3(output_path)
                        except:
                            audio = ID3()
                        if 'musicName' in meta_data:
                            audio.add(TIT2(encoding=3, text=meta_data['musicName']))
                        if 'artist' in meta_data:
                            audio.add(TPE1(encoding=3, text=meta_data['artist']))
                        if 'album' in meta_data:
                            audio.add(TALB(encoding=3, text=meta_data['album']))
                        if 'track' in meta_data:
                            audio.add(TRCK(encoding=3, text=str(meta_data['track'])))
                        audio.save(output_path)

                # 记录转换成功的文件信息
                success_info = {
                    'original': file_path,
                    'outputs': output_paths
                }
                QMetaObject.invokeMethod(self, 'add_converted_file',
                                      Qt.ConnectionType.QueuedConnection,
                                      Q_ARG(dict, success_info))

                # 在主线程中移除已完成的文件
                QMetaObject.invokeMethod(self, 'remove_completed_file',
                                      Qt.ConnectionType.QueuedConnection,
                                      Q_ARG(str, file_path))

            except Exception as e:
                QMetaObject.invokeMethod(self, 'show_error',
                                      Qt.ConnectionType.QueuedConnection,
                                      Q_ARG(str, file_path),
                                      Q_ARG(str, str(e)))

            finally:
                self.files_converting.remove(file_path)
                self.conversion_queue.task_done()
                QMetaObject.invokeMethod(self, 'check_all_completed',
                                      Qt.ConnectionType.QueuedConnection)

    @pyqtSlot(dict)
    def add_converted_file(self, success_info):
        self.converted_files.append(success_info)
        self.completed_files += 1
        # 更新信息显示区域，累积显示所有文件的转换信息
        current_text = self.info_label.text()
        info_text = f'原文件: {os.path.basename(success_info["original"])}\n'
        info_text += f'原文件路径: {success_info["original"]}\n'
        for path in success_info["outputs"]:
            info_text += f'输出文件: {os.path.basename(path)}\n'
            info_text += f'输出路径: {path}\n'
        info_text += '-------------------\n'
        self.info_label.setText(current_text + info_text)

    @pyqtSlot()
    def check_all_completed(self):
        if self.completed_files == self.total_files:
            QMessageBox.information(self, '完成', '所有文件转换成功')
            # 保持转换记录显示
            self.converted_files = []
            

    @pyqtSlot(str)
    def show_success_message(self, message):
        QMessageBox.information(self, '完成', message)

    @pyqtSlot(str)
    def remove_completed_file(self, file_path):
        for i in range(self.file_list.count()):
            if self.file_list.item(i).text() == file_path:
                self.file_list.takeItem(i)
                break

    @pyqtSlot(str, int)
    def update_progress(self, file_path, progress):
        if file_path in self.progress_bars:
            self.progress_bars[file_path].setValue(progress)

    def remove_file(self, file_path):
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            widget = self.file_list.itemWidget(item)
            if isinstance(widget, QWidget):
                file_info = widget.findChild(QLabel)
                if file_info and file_info.text().endswith(file_path):
                    if file_path in self.progress_bars:
                        del self.progress_bars[file_path]
                    self.file_list.takeItem(i)
                    break

    def conversion_completed(self, file_path):
        QMessageBox.information(self, '完成', f'{os.path.basename(file_path)} 转换完成')
        # 从列表中移除已完成的文件
        for i in range(self.file_list.count()):
            if self.file_list.item(i).text() == file_path:
                self.file_list.takeItem(i)
                break

    @pyqtSlot(str, str)
    def show_error(self, file_path, error_message):
        QMessageBox.critical(self, '错误',
                           f'转换 {os.path.basename(file_path)} 时发生错误:\n{error_message}')

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()