from __future__ import annotations
import os
import sys
import threading
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal, QSize, QUrl
from PySide6.QtGui import QIcon, QImage, QPixmap, QDesktopServices, QFont
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QAbstractItemView, QLineEdit, QFileDialog,
    QComboBox, QProgressBar, QTextEdit, QMessageBox, QGroupBox)

from .core import (ImageSettings, read_image, natural_key, detect_engines, images_to_pdf, word_to_pdf,
                   docx_font_warnings, Cancelled, ConversionError, LOG_DIR, log, fs_path)


class FileList(QListWidget):
    files_dropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setIconSize(QSize(62, 54))
        self.setAlternatingRowColors(True)
        self.setMinimumHeight(160)

    def dragEnterEvent(self, event):
        if self.isEnabled() and event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        self.dragEnterEvent(event)

    def dropEvent(self, event):
        if self.isEnabled() and event.mimeData().hasUrls():
            self.files_dropped.emit([u.toLocalFile() for u in event.mimeData().urls()])
            event.acceptProposedAction()


class Job(QThread):
    update = Signal(int, str, str)
    result = Signal(object)
    message = Signal(str)
    summary = Signal(str)

    def __init__(self, mode, paths, output, name, settings, engine):
        super().__init__()
        self.mode, self.paths, self.output = mode, paths, output
        self.name, self.settings, self.engine = name, settings, engine
        self.cancel = threading.Event()

    def run(self):
        successes = 0
        failures = 0
        try:
            if self.mode == 'image':
                states = ['等待'] * len(self.paths)
                def update(i, state, path):
                    states[i] = state
                    self.update.emit(i, state, path)
                try:
                    result = images_to_pdf(self.paths, self.output, self.name, self.settings, self.cancel, update)
                    self.result.emit(result)
                    for i, p in enumerate(self.paths):
                        self.update.emit(i, '成功', p)
                    self.summary.emit(f'合并成功 · {result.pages} 页')
                except Cancelled as exc:
                    for i, p in enumerate(self.paths):
                        self.update.emit(i, '已取消，未生成合并 PDF', p)
                    self.summary.emit(str(exc))
                except Exception as exc:
                    for i, p in enumerate(self.paths):
                        if not states[i].startswith('失败'):
                            self.update.emit(i, '未生成 PDF，合并任务失败', p)
                    log.error('Image job failed type=%s', type(exc).__name__)
                    self.message.emit(str(exc))
                    self.summary.emit('合并失败，未生成成品。请移除或替换失败项后重试。')
                return
            for i, p in enumerate(self.paths):
                if self.cancel.is_set():
                    self.update.emit(i, '未执行（已取消）', p)
                    break
                self.update.emit(i, '处理中', p)
                try:
                    for warning in docx_font_warnings(p):
                        self.message.emit(Path(p).name + '：' + warning)
                    result = word_to_pdf(p, self.output, self.engine, self.cancel)
                    self.result.emit(result)
                    self.update.emit(i, '成功', p)
                    successes += 1
                except Cancelled:
                    self.update.emit(i, '已取消', p)
                    break
                except Exception as exc:
                    failures += 1
                    self.update.emit(i, '失败', p)
                    self.message.emit(Path(p).name + '：' + str(exc))
                    log.error('Word job failed file=%s type=%s', Path(p).name, type(exc).__name__)
            if self.cancel.is_set():
                for j in range(i+1, len(self.paths)):
                    self.update.emit(j, '未执行（已取消）', self.paths[j])
                self.summary.emit(f'已取消 · 成功 {successes} 份，失败 {failures} 份；未完成项未生成 PDF。')
            else:
                self.summary.emit(f'处理结束 · 成功 {successes} 份，失败 {failures} 份')
        except Exception as exc:
            self.message.emit(str(exc))
            self.summary.emit('任务失败；请查看详细信息。')


class ConvertPage(QWidget):
    busy_changed = Signal(bool)

    def __init__(self, mode):
        super().__init__()
        self.mode = mode
        self.job = None
        self.busy = False
        self.automatic_output = True
        self.engines = detect_engines()
        root = QVBoxLayout(self)
        root.setSpacing(10)
        intro = QLabel('添加 Word 文档，每份分别导出 PDF；支持 .docx 和 .doc。' if mode == 'word' else '添加 JPG / JPEG / PNG，按列表顺序合并，每张图片一页。')
        intro.setWordWrap(True)
        root.addWidget(intro)
        self.controls = QWidget()
        controls_layout = QVBoxLayout(self.controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)
        bar = QHBoxLayout()
        self.add_button = self.button(bar, '添加文件', self.pick_files)
        self.button(bar, '移除选中', self.remove_selected)
        self.button(bar, '清空列表', self.clear_files)
        if mode == 'image':
            self.button(bar, '上移 ↑', lambda: self.move(-1))
            self.button(bar, '下移 ↓', lambda: self.move(1))
            self.button(bar, '自然排序', self.sort_files)
        bar.addStretch()
        controls_layout.addLayout(bar)
        self.files = FileList()
        self.files.setToolTip('可从资源管理器拖入文件；支持 Ctrl / Shift 多选。')
        self.files.files_dropped.connect(self.add_files)
        controls_layout.addWidget(self.files, 1)
        hint = QLabel('也可将文件拖入上方列表。悬停文件可查看完整路径。')
        hint.setObjectName('muted')
        controls_layout.addWidget(hint)
        settings = QHBoxLayout()
        if mode == 'word':
            settings.addWidget(QLabel('转换引擎'))
            self.engine = QComboBox()
            self.engine.addItems(list(self.engines))
            if not self.engines:
                self.engine.addItem('未检测到引擎：请安装 Microsoft Word 或 LibreOffice')
            settings.addWidget(self.engine, 1)
            self.button(settings, '重新检测', self.refresh_engines)
        else:
            self.page_mode = QComboBox()
            self.page_mode.addItems(['A4 页面', '按图片尺寸'])
            self.orientation = QComboBox()
            self.orientation.addItems(['每张自动方向', '纵向', '横向'])
            self.margin = QComboBox()
            self.margin.addItems(['无边距 0 mm', '小边距 10 mm', '标准边距 20 mm'])
            self.margin.setCurrentIndex(1)
            for label, widget in [('页面', self.page_mode), ('方向', self.orientation), ('边距', self.margin)]:
                settings.addWidget(QLabel(label))
                settings.addWidget(widget)
            self.page_mode.currentIndexChanged.connect(self.settings_changed)
        controls_layout.addLayout(settings)
        note = QLabel('原排版由所选引擎导出；字体缺失或引擎差异可能影响效果。密码或权限保护文档不处理。' if mode == 'word' else '图片完整居中，不裁剪、不拉伸；EXIF 自动纠正，透明区域铺白，不降采样。\n按图片尺寸：使用有效且横纵一致的 DPI（10–2400），否则 300 DPI；忽略 A4 方向和边距。')
        note.setWordWrap(True)
        note.setObjectName('muted')
        controls_layout.addWidget(note)
        if mode == 'image':
            row = QHBoxLayout()
            row.addWidget(QLabel('PDF 名称'))
            self.pdf_name = QLineEdit('合并图片.pdf')
            row.addWidget(self.pdf_name)
            controls_layout.addLayout(row)
        row = QHBoxLayout()
        row.addWidget(QLabel('输出目录'))
        self.output = QLineEdit()
        self.output.setPlaceholderText('默认：首个输入文件所在目录 / PDF输出')
        self.output.textEdited.connect(lambda: setattr(self, 'automatic_output', False))
        row.addWidget(self.output, 1)
        self.button(row, '选择…', self.pick_output)
        controls_layout.addLayout(row)
        root.addWidget(self.controls, 1)
        actions = QHBoxLayout()
        self.start_button = self.button(actions, '开始转换', self.start)
        self.start_button.setObjectName('primary')
        self.cancel_button = self.button(actions, '取消', self.cancel_job)
        self.cancel_button.setEnabled(False)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        actions.addWidget(self.progress, 1)
        root.addLayout(actions)
        self.status = QLabel('等待添加文件')
        self.status.setWordWrap(True)
        root.addWidget(self.status)
        self.results = QListWidget()
        self.results.setMaximumHeight(110)
        self.results.setMinimumHeight(65)
        self.results.setToolTip('转换结果：选择一项后打开 PDF。')
        self.results.itemSelectionChanged.connect(self.refresh)
        self.results.itemDoubleClicked.connect(lambda _: self.open_pdf())
        root.addWidget(self.results)
        row = QHBoxLayout()
        self.open_button = self.button(row, '打开 PDF', self.open_pdf)
        self.folder_button = self.button(row, '打开输出文件夹', self.open_folder)
        self.details_toggle = QPushButton('详细信息 / 错误日志 ▸')
        self.details_toggle.setCheckable(True)
        row.addStretch()
        row.addWidget(self.details_toggle)
        root.addLayout(row)
        self.details = QTextEdit()
        self.details.setReadOnly(True)
        self.details.setMaximumHeight(110)
        self.details.hide()
        self.details_toggle.toggled.connect(self.details.setVisible)
        root.addWidget(self.details)
        self.refresh()

    @staticmethod
    def button(layout, text, callback):
        button = QPushButton(text)
        button.clicked.connect(callback)
        layout.addWidget(button)
        return button

    def refresh_engines(self):
        self.engines = detect_engines()
        self.engine.clear()
        self.engine.addItems(list(self.engines) or ['未检测到引擎：请安装 Microsoft Word 或 LibreOffice'])
        self.refresh()

    def settings_changed(self):
        a4 = self.page_mode.currentIndex() == 0
        self.orientation.setEnabled(a4)
        self.margin.setEnabled(a4)

    def pick_files(self):
        filt = 'Word 文档 (*.docx *.doc)' if self.mode == 'word' else '图片 (*.jpg *.jpeg *.png)'
        paths, _ = QFileDialog.getOpenFileNames(self, '添加文件', '', filt)
        self.add_files(paths)

    def add_files(self, paths):
        if self.busy:
            return
        errors = []
        known = {self.files.item(i).data(Qt.ItemDataRole.UserRole)['path'].casefold() for i in range(self.files.count())}
        allowed = ('.docx', '.doc') if self.mode == 'word' else ('.jpg', '.jpeg', '.png')
        for path in paths:
            p = Path(path).resolve()
            if not path or not fs_path(p).is_file() or p.suffix.lower() not in allowed:
                errors.append(f'{p.name}：不支持此类型或文件不存在。支持 {", ".join(allowed)}。')
                continue
            if str(p).casefold() in known:
                continue
            info = {'path': str(p), 'label': p.name, 'valid': True}
            item = QListWidgetItem()
            if self.mode == 'image':
                try:
                    im, dpi = read_image(p)
                    info['label'] += f'  ·  {im.width} × {im.height} px'
                    im.thumbnail((100, 70))
                    data = im.tobytes('raw', 'RGB')
                    qi = QImage(data, im.width, im.height, im.width*3, QImage.Format.Format_RGB888).copy()
                    item.setIcon(QIcon(QPixmap.fromImage(qi)))
                    item.setSizeHint(QSize(100, 62))
                except ConversionError as exc:
                    info['valid'] = False
                    info['label'] += '  ·  无法读取，请移除或替换'
                    errors.append(str(exc))
            item.setData(Qt.ItemDataRole.UserRole, info)
            item.setText(info['label'])
            item.setToolTip(str(p))
            self.files.addItem(item)
            known.add(str(p).casefold())
        self.update_default_output()
        self.refresh()
        if errors:
            self.add_message('\n'.join(errors))
            QMessageBox.warning(self, '部分文件需要处理', '\n'.join(errors[:5]) + ('\n其余错误见详细信息。' if len(errors)>5 else ''))

    def update_default_output(self):
        if self.automatic_output:
            self.output.setText(str(Path(self.files.item(0).data(Qt.ItemDataRole.UserRole)['path']).parent/'PDF输出') if self.files.count() else '')

    def remove_selected(self):
        for item in self.files.selectedItems():
            self.files.takeItem(self.files.row(item))
        self.update_default_output()
        self.refresh()

    def clear_files(self):
        self.files.clear()
        self.update_default_output()
        self.refresh()

    def move(self, delta):
        rows = sorted([self.files.row(item) for item in self.files.selectedItems()], reverse=delta>0)
        selected = set(rows)
        for row in rows:
            nxt = row+delta
            if 0 <= nxt < self.files.count() and nxt not in selected:
                item = self.files.takeItem(row)
                self.files.insertItem(nxt, item)
                item.setSelected(True)
                selected.remove(row)
                selected.add(nxt)

    def sort_files(self):
        items = [self.files.takeItem(0) for _ in range(self.files.count())]
        items.sort(key=lambda item: natural_key(item.data(Qt.ItemDataRole.UserRole)['path']))
        for item in items:
            self.files.addItem(item)

    def pick_output(self):
        directory = QFileDialog.getExistingDirectory(self, '选择输出文件夹', self.output.text())
        if directory:
            self.output.setText(directory)
            self.automatic_output = False

    def refresh(self):
        count = self.files.count()
        valid = count > 0 and all(self.files.item(i).data(Qt.ItemDataRole.UserRole)['valid'] for i in range(count))
        self.start_button.setEnabled(not self.busy and valid and (self.mode == 'image' or bool(self.engines)))
        self.open_button.setEnabled(bool(self.results.selectedItems()))
        self.folder_button.setEnabled(bool(self.results.count()))

    def start(self):
        if self.busy or not self.start_button.isEnabled():
            return
        if not self.output.text().strip():
            QMessageBox.warning(self, '需要输出目录', '请选择输出文件夹。')
            return
        paths = [self.files.item(i).data(Qt.ItemDataRole.UserRole)['path'] for i in range(self.files.count())]
        settings = ImageSettings()
        name, engine = '', ''
        if self.mode == 'image':
            settings = ImageSettings(['a4', 'image'][self.page_mode.currentIndex()], ['auto', 'portrait', 'landscape'][self.orientation.currentIndex()], [0, 10, 20][self.margin.currentIndex()])
            name = self.pdf_name.text()
        else:
            engine = self.engines[self.engine.currentText()]
        self.results.clear()
        self.details.clear()
        for i in range(self.files.count()):
            item = self.files.item(i)
            item.setText(item.data(Qt.ItemDataRole.UserRole)['label'] + '  ·  等待')
        self.busy = True
        self.controls.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.progress.setRange(0, len(paths))
        self.progress.setValue(0)
        self.status.setText('正在转换…' + (f' 使用 {self.engine.currentText()}' if self.mode == 'word' else ''))
        self.busy_changed.emit(True)
        self.job = Job(self.mode, paths, self.output.text(), name, settings, engine)
        self.job.update.connect(self.on_update)
        self.job.result.connect(self.on_result)
        self.job.message.connect(self.add_message)
        self.job.summary.connect(self.status.setText)
        self.job.finished.connect(self.on_finished)
        self.refresh()
        self.job.start()

    def on_update(self, index, state, path):
        item = self.files.item(index)
        item.setText(item.data(Qt.ItemDataRole.UserRole)['label'] + '  ·  ' + state)
        if state == '处理中':
            self.status.setText(f'正在处理 {index+1}/{self.files.count()}：{Path(path).name}')
        else:
            self.progress.setValue(index+1)

    def on_result(self, result):
        item = QListWidgetItem(f'{Path(result.path).name}  ·  {result.pages} 页  ·  {result.size/1024:,.1f} KB\n{result.path}')
        item.setData(Qt.ItemDataRole.UserRole, result.path)
        item.setToolTip(result.path)
        self.results.addItem(item)
        self.results.setCurrentItem(item)

    def add_message(self, text):
        self.details.append(text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
        self.details_toggle.setChecked(True)

    def on_finished(self):
        self.busy = False
        self.controls.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.busy_changed.emit(False)
        self.refresh()

    def cancel_job(self):
        if self.job and self.busy:
            self.job.cancel.set()
            self.cancel_button.setEnabled(False)
            self.status.setText('已请求取消：当前 Word / LibreOffice 文件结束后停止（最长约 180 秒），未提交 PDF 将清理。' if self.mode == 'word' else '已请求取消：当前图片处理结束后停止，临时文件将清理。')

    def open_pdf(self):
        if self.results.currentItem():
            self.open_path(self.results.currentItem().data(Qt.ItemDataRole.UserRole))

    def open_folder(self):
        if self.results.currentItem():
            self.open_path(str(Path(self.results.currentItem().data(Qt.ItemDataRole.UserRole)).parent))

    def open_path(self, path):
        if not Path(path).exists() or not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
            QMessageBox.warning(self, '无法打开', '文件可能已移动，或系统尚未配置 PDF 阅读器。请在资源管理器中打开。')


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('本地文件转 PDF')
        self.resize(1020, 810)
        self.setMinimumSize(850, 720)
        body = QWidget()
        root = QVBoxLayout(body)
        title = QLabel('本地文件转 PDF')
        title.setObjectName('title')
        root.addWidget(title)
        subtitle = QLabel('文件在本机转换  ·  原文件保留  ·  同名自动编号')
        subtitle.setObjectName('muted')
        root.addWidget(subtitle)
        self.tabs = QTabWidget()
        self.pages = [ConvertPage('word'), ConvertPage('image')]
        for i, (page, name) in enumerate(zip(self.pages, ['Word 转 PDF', '图片转 PDF'])):
            self.tabs.addTab(page, name)
            page.busy_changed.connect(lambda busy, index=i: self.tabs.setTabEnabled(1-index, not busy))
        root.addWidget(self.tabs, 1)
        footer = QLabel('日志位置：' + str(LOG_DIR) + '（不记录文档正文）')
        footer.setObjectName('muted')
        footer.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        root.addWidget(footer)
        self.setCentralWidget(body)

    def closeEvent(self, event):
        running = [p for p in self.pages if p.busy]
        if running:
            choice = QMessageBox.question(self, '任务仍在进行', '当前任务尚未结束。是否请求取消？\n取消完成后请再次关闭窗口。', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if choice == QMessageBox.StandardButton.Yes:
                for page in running:
                    page.cancel_job()
            event.ignore()
        else:
            event.accept()


def configure_application(app):
    app.setApplicationName('LocalFilePDF')
    app.setFont(QFont('Microsoft YaHei UI', 10))
    app.setStyle('Fusion')
    app.setStyleSheet('''
        QMainWindow, QWidget { background: #f6f7fa; color: #172233; }
        QLabel#title { font-size: 23px; font-weight: 650; margin: 5px 0; }
        QLabel#muted { color: #626e80; font-size: 12px; }
        QPushButton { padding: 7px 13px; border: 1px solid #ccd3df; border-radius: 5px; background: white; }
        QPushButton:hover { background: #edf3ff; border-color: #6c91d2; }
        QPushButton:disabled { color: #9ca4b0; background: #eef0f4; }
        QPushButton#primary { background: #2463cf; color: white; border-color: #2463cf; font-weight: bold; }
        QPushButton#primary:disabled { background: #b2bfda; border-color: #b2bfda; }
        QLineEdit, QComboBox, QListWidget, QTextEdit { background: white; alternate-background-color: #f1f5fb; border: 1px solid #cdd5e0; border-radius: 4px; padding: 5px; }
        QListWidget::item:selected { background: #dce9ff; color: #162642; }
        QTabBar::tab { padding: 10px 25px; background: #e8ecf3; }
        QTabBar::tab:selected { background: white; color: #1858bd; }
        QProgressBar { border: 1px solid #d3dbe6; border-radius: 4px; text-align: center; min-height: 22px; }
        QProgressBar::chunk { background: #84aceb; }
    ''')
def main():
    app = QApplication(sys.argv)
    configure_application(app)
    window = Window()
    window.show()
    sys.exit(app.exec())
