import os
from pathlib import Path
import time
from PySide6.QtCore import Qt, QMimeData, QUrl, QPointF
from PySide6.QtGui import QDropEvent
from PySide6.QtWidgets import QApplication
from localpdf.gui import Window, configure_application
from tests.make_fixtures import create


def test_gui_order_drop_and_real_conversion(tmp_path):
    app = QApplication.instance() or QApplication([])
    configure_application(app)
    window = Window()
    page = window.pages[1]
    window.tabs.setCurrentIndex(1)
    window.show()
    app.processEvents()
    assert not page.start_button.isEnabled()
    samples = create()
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(samples/n)) for n in ['10.jpg', '2.jpg', '1.jpg']])
    event = QDropEvent(QPointF(20, 20), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    page.files.dropEvent(event)
    assert page.files.count() == 3 and page.start_button.isEnabled()
    page.sort_files()
    assert [Path(page.files.item(i).data(Qt.ItemDataRole.UserRole)['path']).name for i in range(3)] == ['1.jpg','2.jpg','10.jpg']
    page.files.item(2).setSelected(True)
    page.move(-1)
    page.move(-1)
    assert Path(page.files.item(0).data(Qt.ItemDataRole.UserRole)['path']).name == '10.jpg'
    page.output.setText(str(tmp_path))
    page.start()
    assert not page.start_button.isEnabled() and page.cancel_button.isEnabled()
    assert not window.tabs.isTabEnabled(0)
    job = page.job
    page.start()
    assert page.job is job
    deadline = time.time()+30
    ticks = 0
    while page.busy and time.time() < deadline:
        app.processEvents()
        time.sleep(.01)
        ticks += 1
    assert not page.busy and page.results.count() == 1 and ticks > 0
    assert '成功' in page.status.text()
    assert page.open_button.isEnabled()
    assert window.tabs.isTabEnabled(0)
    qa = Path('tests/验证输出')
    qa.mkdir(exist_ok=True)
    window.grab().save(str(qa/'界面-图片转换完成.png'))
    page.clear_files()
    assert not page.start_button.isEnabled()
    window.close()


def test_gui_missing_engine_invalid_files_and_close(monkeypatch, tmp_path):
    import localpdf.gui as gui
    from PySide6.QtGui import QCloseEvent
    from PySide6.QtWidgets import QMessageBox
    from threading import Event
    from types import SimpleNamespace
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(gui, 'detect_engines', lambda: {})
    monkeypatch.setattr(QMessageBox, 'warning', lambda *args: QMessageBox.StandardButton.Ok)
    window = Window()
    samples = create()
    word, page = window.pages
    word.add_files([str(samples/'中文 English 图表.docx')])
    assert not word.start_button.isEnabled()
    page.add_files([str(samples/'损坏.png'), str(samples/'1.jpg'), str(samples/'损坏.docx')])
    assert page.files.count() == 2 and not page.start_button.isEnabled()
    page.files.item(0).setSelected(True)
    page.remove_selected()
    assert page.start_button.isEnabled()
    page.busy = True
    page.job = SimpleNamespace(cancel=Event())
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.StandardButton.Yes)
    event = QCloseEvent()
    window.closeEvent(event)
    assert not event.isAccepted() and page.job.cancel.is_set()
    page.busy = False
    window.close()
