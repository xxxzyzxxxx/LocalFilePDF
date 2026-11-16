import hashlib
import os
from pathlib import Path
import threading

import fitz
import pytest
from PIL import Image
from pypdf import PdfReader

from localpdf.core import *
from tests.make_fixtures import create


@pytest.fixture(scope='session')
def samples():
    return create()


def test_natural_order():
    assert sorted(['10.jpg', '2.jpg', '1.jpg'], key=natural_key) == ['1.jpg', '2.jpg', '10.jpg']


def test_exif_and_alpha(samples):
    im, dpi = read_image(samples/'EXIF 旋转.jpg')
    assert im.size == (240, 480)
    assert im.getpixel((120, 30))[0] < 150  # purple side becomes top
    im, dpi = read_image(samples/'透明 图片.png')
    assert im.mode == 'RGB' and im.getpixel((0, 0)) == (255, 255, 255)


def test_pages_order_pixels_and_collision(samples, tmp_path):
    paths = [samples/n for n in ['10.jpg', '1.jpg', '2.jpg', 'EXIF 旋转.jpg', '透明 图片.png']]
    result = images_to_pdf(paths, tmp_path/'中文 输出 with spaces', '合并.pdf', ImageSettings())
    assert result.pages == 5
    doc = fitz.open(result.path)
    assert doc[0].rect.width > doc[0].rect.height
    assert doc[2].rect.width < doc[2].rect.height
    assert doc[3].rect.width < doc[3].rect.height
    # Center colors assert exact selected page order; embedded pixel dimensions assert no downsampling.
    for i, channel in enumerate([2, 0, 1]):
        pix = doc[i].get_pixmap()
        color = pix.pixel(pix.width//2, pix.height//2)
        assert color[channel] > max(color[j] for j in range(3) if j != channel)
    for i, path in enumerate(paths):
        expected, _ = read_image(path)
        entry = doc[i].get_images()[0]
        assert tuple(entry[2:4]) == expected.size
    doc.close()
    before = Path(result.path).read_bytes()
    again = images_to_pdf(paths, Path(result.path).parent, '合并.pdf', ImageSettings())
    assert Path(again.path).name == '合并 (1).pdf'
    assert Path(result.path).read_bytes() == before


@pytest.mark.parametrize('orientation,landscape', [('portrait', False), ('landscape', True), ('auto', True)])
@pytest.mark.parametrize('margin', [0, 10, 20])
def test_a4_fit(samples, tmp_path, orientation, landscape, margin):
    result = images_to_pdf([samples/'1.jpg'], tmp_path, 'a4', ImageSettings('a4', orientation, margin))
    with fitz.open(result.path) as doc:
        page = doc[0]
        assert (page.rect.width > page.rect.height) == landscape
        box = page.get_image_rects(page.get_images()[0][0])[0]
        assert abs(box.width/box.height - 2) < .001
        m = margin*72/25.4
        assert box.x0 >= m-.01 and box.y0 >= m-.01
        assert abs((page.rect.width-box.width)/2-box.x0) < .01


def test_dpi_rules(samples, tmp_path):
    im = Image.new('RGB', (600, 300), 'red')
    for filename, dpi, expected in [('a.png', (150, 150), 288), ('b.png', (72, 96), 144), ('c.png', None, 144)]:
        path = tmp_path/filename
        im.save(path, **({'dpi': dpi} if dpi else {}))
        result = images_to_pdf([path], tmp_path, filename, ImageSettings('image'))
        assert abs(float(PdfReader(result.path).pages[0].mediabox.width)-expected) < .1


def test_broken_no_partial(samples, tmp_path):
    with pytest.raises(ConversionError, match='损坏.png'):
        images_to_pdf([samples/'1.jpg', samples/'损坏.png'], tmp_path, 'broken', ImageSettings())
    assert not list(tmp_path.iterdir())


def test_cancel_during_work(samples, tmp_path):
    event = threading.Event()
    def progress(i, state, path):
        if state.startswith('已处理'):
            event.set()
    with pytest.raises(Cancelled):
        images_to_pdf([samples/'1.jpg', samples/'2.jpg'], tmp_path, 'cancel', ImageSettings(), event, progress)
    assert not list(tmp_path.iterdir())


def test_unwritable_and_invalid_pdf(samples, tmp_path):
    blocker = tmp_path/'file-not-directory'
    blocker.write_text('existing')
    with pytest.raises(ConversionError, match='无法写入'):
        images_to_pdf([samples/'1.jpg'], blocker/'output', 'x', ImageSettings())
    with pytest.raises(ConversionError, match='无效'):
        validate_pdf(blocker)
    assert blocker.read_text() == 'existing'


def test_long_unicode_path(samples, tmp_path):
    folder = fs_path(tmp_path / ('中文路径 空格 '*16).strip() / ('子目录 '*15).strip())
    folder.mkdir(parents=True)
    path = folder / ('较长文件名 '*23 + '.png')
    Image.new('RGB', (50, 90), 'blue').save(path)
    result = images_to_pdf([path], folder/'PDF输出', '测试', ImageSettings())
    assert result.pages == 1 and len(result.path) > 260


@pytest.mark.parametrize('name', ['', '../bad', 'CON', 'x'*240, 'a|b'])
def test_bad_output_name(name):
    with pytest.raises(ConversionError):
        output_name(name)


def test_real_permission_denied(samples, tmp_path):
    import win32security
    import win32api
    import win32con
    import ntsecuritycon
    folder = tmp_path/'权限拒绝测试'
    folder.mkdir()
    old = win32security.GetFileSecurity(str(folder), win32security.DACL_SECURITY_INFORMATION)
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_QUERY)
    sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    acl = win32security.ACL()
    acl.AddAccessDeniedAce(win32security.ACL_REVISION, ntsecuritycon.FILE_ADD_FILE | ntsecuritycon.FILE_ADD_SUBDIRECTORY, sid)
    acl.AddAccessAllowedAce(win32security.ACL_REVISION, ntsecuritycon.FILE_ALL_ACCESS, sid)
    descriptor = win32security.SECURITY_DESCRIPTOR()
    descriptor.SetSecurityDescriptorDacl(1, acl, 0)
    try:
        win32security.SetFileSecurity(str(folder), win32security.DACL_SECURITY_INFORMATION, descriptor)
        with pytest.raises(ConversionError, match='无法写入'):
            images_to_pdf([samples/'1.jpg'], folder, '不能创建', ImageSettings())
    finally:
        win32security.SetFileSecurity(str(folder), win32security.DACL_SECURITY_INFORMATION, old)
        token.Close()
    assert not list(folder.iterdir())


def test_locked_input_and_existing_output(samples, tmp_path):
    import win32file
    import win32con
    locked = tmp_path/'被占用.png'
    Image.new('RGB', (20, 20), 'red').save(locked)
    handle = win32file.CreateFile(str(locked), win32con.GENERIC_READ, 0, None, win32con.OPEN_EXISTING, 0, None)
    try:
        with pytest.raises(ConversionError, match='无法读取'):
            images_to_pdf([locked], tmp_path, 'x', ImageSettings())
    finally:
        handle.Close()
    existing = tmp_path/'相同名称.pdf'
    existing.write_bytes(b'existing-file-must-not-change')
    handle = win32file.CreateFile(str(existing), win32con.GENERIC_READ, 0, None, win32con.OPEN_EXISTING, 0, None)
    try:
        result = images_to_pdf([samples/'1.jpg'], tmp_path, '相同名称', ImageSettings())
        assert Path(result.path).name == '相同名称 (1).pdf'
    finally:
        handle.Close()
    assert existing.read_bytes() == b'existing-file-must-not-change'


@pytest.mark.skipif(os.environ.get('RUN_WORD_TESTS') != '1', reason='Requires installed Microsoft Word')
def test_real_word_batch_and_existing_instance(samples, tmp_path):
    import pythoncom
    import win32com.client
    import psutil
    import win32process
    from localpdf.gui import Job
    from PySide6.QtWidgets import QApplication
    appqt = QApplication.instance() or QApplication([])
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Add()
        doc.Content.Text = 'USER UNSAVED SENTINEL 12345'
        _, pid = win32process.GetWindowThreadProcessId(doc.Windows(1).Hwnd)
        paths = [samples/'中文 English 图表.docx', samples/'损坏.docx', samples/'页眉 页脚 分页 公式.docx']
        hashes = [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths]
        job = Job('word', [str(p) for p in paths], str(tmp_path), '', ImageSettings(), 'word')
        results, statuses, summaries = [], [], []
        job.result.connect(results.append)
        job.update.connect(lambda i, s, p: statuses.append((i, s)))
        job.summary.connect(summaries.append)
        job.run()
        assert len(results) == 2
        assert (1, '失败') in statuses and (2, '成功') in statuses
        assert [r.pages for r in results] == [1, 2]
        assert '失败 1' in summaries[-1]
        assert psutil.pid_exists(pid) and doc.Content.Text.strip() == 'USER UNSAVED SENTINEL 12345'
        assert not doc.Saved  # original unsaved user state unchanged
        assert hashes == [hashlib.sha256(p.read_bytes()).hexdigest() for p in paths]
        qa = Path('tests/验证输出')
        qa.mkdir(exist_ok=True)
        for r in results:
            shutil.copy2(r.path, qa/Path(r.path).name)
            with fitz.open(r.path) as pdf:
                for i, page in enumerate(pdf):
                    page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2)).save(str(qa/(Path(r.path).stem+f'-{i+1}.png')))
                text = ''.join(page.get_text() for page in pdf)
                assert '单元格' in text and 'English' in text
                if r.pages == 2:
                    assert '第二页' in pdf[1].get_text()
                    assert all('测试页眉' in p.get_text() and '页码' in p.get_text() for p in pdf)
                    import unicodedata
                    assert 'mc' in unicodedata.normalize('NFKC', pdf[1].get_text())
        # Password-protected input must fail promptly, never bypass protection.
        encrypted = tmp_path/'加密.docx'
        doc.Password = 'sample-password'
        doc.SaveAs2(str(encrypted), FileFormat=16)
        assert not zipfile.is_zipfile(encrypted)
        with pytest.raises(ConversionError):
            word_to_pdf(encrypted, tmp_path, 'word', timeout=30)
        assert not (tmp_path/'加密.pdf').exists()
        # Legacy DOC export and real in-flight cancellation.
        legacy = tmp_path/'旧格式.doc'
        doc.Password = ''
        doc.SaveAs2(str(legacy), FileFormat=0)
        assert word_to_pdf(legacy, tmp_path, 'word').pages == 1
        event = threading.Event()
        timer = threading.Timer(.3, event.set)
        timer.start()
        with pytest.raises(Cancelled):
            word_to_pdf(paths[0], tmp_path, 'word', event)
        timer.join()
        assert not list(tmp_path.glob('.localpdf-*'))
    finally:
        if doc is not None:
            doc.Close(0)
        word.Quit(0)
        pythoncom.CoUninitialize()
