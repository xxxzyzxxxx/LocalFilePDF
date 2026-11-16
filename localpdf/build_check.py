"""Developer-only packaged UI smoke check using real widgets and converters."""
import json
from pathlib import Path
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication
from .gui import Window, configure_application


def run(samples):
    samples = Path(samples).resolve()
    output = samples.parent / '打包验证'
    output.mkdir(exist_ok=True)
    app = QApplication([])
    configure_application(app)
    window = Window()
    window.show()
    report = {}
    stage = [0]

    def begin():
        page = window.pages[1]
        window.tabs.setCurrentIndex(1)
        page.add_files([str(samples/n) for n in ['1.jpg', '2.jpg', 'EXIF 旋转.jpg', '透明 图片.png']])
        page.output.setText(str(output))
        page.start()
        page.job.finished.connect(image_finished)

    def image_finished():
        page = window.pages[1]
        report['image'] = {'results': page.results.count(), 'summary': page.status.text()}
        window.grab().save(str(output/'EXE-图片.png'))
        QTimer.singleShot(200, begin_word)

    def begin_word():
        page = window.pages[0]
        window.tabs.setCurrentIndex(0)
        page.add_files([str(samples/'中文 English 图表.docx'), str(samples/'页眉 页脚 分页 公式.docx')])
        page.output.setText(str(output))
        page.start()
        if page.job:
            page.job.finished.connect(word_finished)
        else:
            word_finished()

    def word_finished():
        page = window.pages[0]
        report['word'] = {'results': page.results.count(), 'summary': page.status.text(), 'details': page.details.toPlainText()}
        window.grab().save(str(output/'EXE-Word.png'))
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        window.close()
        app.exit(0 if report['image']['results'] == 1 and report['word']['results'] == 2 else 1)

    QTimer.singleShot(300, begin)
    return app.exec()
