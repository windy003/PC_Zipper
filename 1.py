import os
import sys
import zipfile
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                            QWidget, QFileDialog, QProgressBar, QTextEdit, 
                            QLabel, QHBoxLayout, QListView, QTreeView, QSplitter)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QAbstractListModel, QModelIndex, QTimer
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QIcon

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

class ZipWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, mode, source_path, dest_path=None):
        super().__init__()
        self.mode = mode
        self.source_path = source_path
        self.dest_path = dest_path

    def run(self):
        try:
            if self.mode == 'compress':
                self._compress_folder()
            elif self.mode == 'extract':
                self._extract_zip()
        except Exception as e:
            self.error.emit(str(e))
        self.finished.emit()

    def _compress_folder(self):
        source_path = Path(self.source_path)
        zip_path = source_path.with_suffix('.zip')
        
        # 获取所有文件的总大小
        total_size = sum(f.stat().st_size for f in source_path.rglob('*') if f.is_file())
        processed_size = 0

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(source_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, source_path)
                    zipf.write(file_path, arcname)
                    
                    processed_size += os.path.getsize(file_path)
                    progress = int((processed_size / total_size) * 100)
                    self.progress.emit(progress)

    def _extract_zip(self):
        zip_path = Path(self.source_path)
        extract_path = self.dest_path or zip_path.with_suffix('')
        
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            total_size = sum(info.file_size for info in zipf.infolist())
            extracted_size = 0
            
            for info in zipf.infolist():
                zipf.extract(info, extract_path)
                extracted_size += info.file_size
                progress = int((extracted_size / total_size) * 100)
                self.progress.emit(progress)

class ZipCache:
    def __init__(self):
        self.cache = {}
        self.max_cache = 10  # 最多缓存10个文件的信息
        
    def get_info(self, zip_path):
        if zip_path in self.cache:
            return self.cache[zip_path]
            
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            info = {
                'files': zipf.infolist(),
                'total_size': sum(info.file_size for info in zipf.infolist()),
                'total_files': len(zipf.infolist())
            }
            
        if len(self.cache) >= self.max_cache:
            # 删除最早的缓存
            oldest = next(iter(self.cache))
            del self.cache[oldest]
            
        self.cache[zip_path] = info
        return info

class ZipLoadWorker(QThread):
    """异步加载ZIP文件信息的工作线程"""
    finished = pyqtSignal(object)
    progress = pyqtSignal(str)
    
    def __init__(self, zip_path):
        super().__init__()
        self.zip_path = zip_path
        
    def run(self):
        try:
            self.progress.emit("正在加载文件信息...")
            with zipfile.ZipFile(self.zip_path, 'r') as zipf:
                info = {
                    'files': zipf.infolist(),
                    'total_size': sum(info.file_size for info in zipf.infolist()),
                    'total_files': len(zipf.infolist())
                }
            self.finished.emit(info)
        except Exception as e:
            self.progress.emit(f"加载失败: {str(e)}")

class ZipListModel(QAbstractListModel):
    def __init__(self):
        super().__init__()
        self.infolist = []
        self.batch_size = 100  # 每批显示的文件数
        self.current_batch = 0
        
    def rowCount(self, parent=QModelIndex()):
        return len(self.infolist)
        
    def data(self, index, role):
        if not index.isValid() or role != Qt.ItemDataRole.DisplayRole:
            return None
            
        info = self.infolist[index.row()]
        size = info.file_size
        date = f"{info.date_time[0]}-{info.date_time[1]:02d}-{info.date_time[2]:02d}"
        return f"{info.filename:<40} {size:>10,d} 字节  {date}"
    
    def load_batch(self, file_list):
        """分批加载文件列表"""
        start = self.current_batch * self.batch_size
        end = start + self.batch_size
        batch = file_list[start:end]
        
        if batch:
            self.beginInsertRows(QModelIndex(), len(self.infolist), 
                               len(self.infolist) + len(batch) - 1)
            self.infolist.extend(batch)
            self.endInsertRows()
            self.current_batch += 1
            return True
        return False

class ZipTreeModel(QStandardItemModel):
    def __init__(self):
        super().__init__()
        self.setHorizontalHeaderLabels(['名称', '大小', '修改日期'])
        self.folder_icon = None  # 可以添加文件夹图标
        self.file_icon = None    # 可以添加文件图标
        
    def load_zip_content(self, zip_path):
        self.clear()
        self.setHorizontalHeaderLabels(['名称', '大小', '修改日期'])
        root = self.invisibleRootItem()
        
        try:
            with zipfile.ZipFile(zip_path, 'r') as zipf:
                # 创建目录树结构
                folders = {}
                
                # 首先添加所有文件夹
                for info in zipf.infolist():
                    path_parts = info.filename.split('/')
                    current_path = ''
                    parent_item = root
                    
                    # 处理路径中的每一部分
                    for i, part in enumerate(path_parts):
                        if not part:  # 跳过空文件夹名
                            continue
                            
                        current_path = current_path + part if not current_path else current_path + '/' + part
                        
                        if current_path not in folders:
                            # 创建新的文件夹项
                            folder_item = QStandardItem(part)
                            size_item = QStandardItem('')  # 文件夹大小暂时留空
                            date_item = QStandardItem('')  # 文件夹日期暂时留空
                            
                            # 如果是文件而不是文件夹
                            if i == len(path_parts) - 1 and not info.filename.endswith('/'):
                                size = info.file_size
                                date = f"{info.date_time[0]}-{info.date_time[1]:02d}-{info.date_time[2]:02d}"
                                size_item.setText(f"{size:,d} 字节")
                                date_item.setText(date)
                            
                            parent_item.appendRow([folder_item, size_item, date_item])
                            folders[current_path] = folder_item
                            
                        parent_item = folders[current_path]
                
                return True
        except Exception as e:
            self.clear()
            return False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ZIP文件管理器")
        
        # 设置应用图标
        icon_path = get_resource_path("icon.ico")
        self.setWindowIcon(QIcon(icon_path))
        
        # 创建主布局
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # 创建按钮 (添加快捷键 Alt+C/X/P)
        btn_layout = QHBoxLayout()
        self.compress_btn = QPushButton("压缩文件夹(&C)")
        self.extract_btn = QPushButton("解压ZIP文件(&X)")
        self.preview_btn = QPushButton("预览ZIP文件(&P)")
        
        # 设置工具提示
        self.compress_btn.setToolTip("快捷键: Alt+C")
        self.extract_btn.setToolTip("快捷键: Alt+X")
        self.preview_btn.setToolTip("快捷键: Alt+P")
        
        btn_layout.addWidget(self.compress_btn)
        btn_layout.addWidget(self.extract_btn)
        btn_layout.addWidget(self.preview_btn)
        layout.addLayout(btn_layout)
        
        # 创建进度条
        self.progress_label = QLabel("进度:")
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        
        # 创建树状视图
        self.tree_view = QTreeView()
        self.tree_view.setAlternatingRowColors(True)  # 交替行颜色
        self.tree_view.setSortingEnabled(True)        # 允许排序
        self.tree_model = ZipTreeModel()
        self.tree_view.setModel(self.tree_model)
        
        # 设置列宽
        self.tree_view.setColumnWidth(0, 300)  # 名称列
        self.tree_view.setColumnWidth(1, 100)  # 大小列
        self.tree_view.setColumnWidth(2, 150)  # 日期列
        
        layout.addWidget(self.tree_view)
        
        # 连接信号
        self.compress_btn.clicked.connect(self.compress_folder)
        self.extract_btn.clicked.connect(self.extract_zip)
        self.preview_btn.clicked.connect(self.preview_zip)
        
        self.current_worker = None
        self.status_label = QLabel()
        self.statusBar().addWidget(self.status_label)

    def compress_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "选择要压缩的文件夹")
        if folder_path:
            self.start_worker('compress', folder_path)

    def extract_zip(self):
        zip_path, _ = QFileDialog.getOpenFileName(self, "选择ZIP文件", "", "ZIP文件 (*.zip)")
        if zip_path:
            extract_path = QFileDialog.getExistingDirectory(self, "选择解压目标文件夹")
            if extract_path:
                self.start_worker('extract', zip_path, extract_path)

    def preview_zip(self):
        zip_path, _ = QFileDialog.getOpenFileName(self, "选择ZIP文件", "", "ZIP文件 (*.zip)")
        if zip_path:
            self.status_label.setText("正在加载...")
            if self.tree_model.load_zip_content(zip_path):
                self.status_label.setText("加载完成")
                # 展开根节点
                self.tree_view.expandToDepth(0)
            else:
                self.status_label.setText("加载失败")

    def start_worker(self, mode, source_path, dest_path=None):
        self.progress_bar.setValue(0)
        self.current_worker = ZipWorker(mode, source_path, dest_path)
        self.current_worker.progress.connect(self.update_progress)
        self.current_worker.finished.connect(self.on_worker_finished)
        self.current_worker.error.connect(self.on_worker_error)
        self.current_worker.start()
        
        # 禁用按钮
        self.compress_btn.setEnabled(False)
        self.extract_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def on_worker_finished(self):
        # 重新启用按钮
        self.compress_btn.setEnabled(True)
        self.extract_btn.setEnabled(True)
        self.preview_btn.setEnabled(True)
        
        if self.progress_bar.value() == 100:
            self.statusBar().showMessage("操作完成！")

    def on_worker_error(self, error_msg):
        self.statusBar().showMessage(f"错误: {error_msg}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    icon_path = get_resource_path("icon.ico")
    app.setWindowIcon(QIcon(icon_path))
    
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec()) 