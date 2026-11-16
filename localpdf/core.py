from __future__ import annotations

import io
import json
import logging
from logging.handlers import RotatingFileHandler
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import warnings
import uuid
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass

from PIL import Image, ImageOps
from pypdf import PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4

LOG_DIR = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'LocalFilePDF' / 'logs'
LOG_DIR.mkdir(parents=True, exist_ok=True)
log = logging.getLogger('localpdf')
if not log.handlers:
    handler = RotatingFileHandler(LOG_DIR / 'app.log', maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    log.addHandler(handler)
    log.setLevel(logging.INFO)


class ConversionError(Exception):
    pass


class Cancelled(ConversionError):
    pass


def fs_path(path):
    """Use Windows extended paths without changing system long-path policy."""
    value = os.path.abspath(os.fspath(path))
    if sys.platform == 'win32' and not value.startswith('\\\\?\\'):
        value = ('\\\\?\\UNC\\' + value[2:]) if value.startswith('\\\\') else ('\\\\?\\' + value)
    return Path(value)


def display_path(path):
    value = str(path)
    if value.startswith('\\\\?\\UNC\\'):
        return '\\\\' + value[8:]
    return value[4:] if value.startswith('\\\\?\\') else value


@dataclass
class Result:
    path: str
    pages: int
    size: int


@dataclass
class ImageSettings:
    mode: str = 'a4'
    orientation: str = 'auto'
    margin_mm: float = 10


def check_cancel(event):
    if event.is_set():
        raise Cancelled('已取消；未提交的临时 PDF 已清理。')


def natural_key(path):
    return [int(s) if s.isdigit() else s.casefold() for s in re.split(r'(\d+)', Path(path).name)]


def detect_engines():
    engines = {}
    if sys.platform == 'win32':
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CLSID'):
                engines['Microsoft Word'] = 'word'
        except OSError:
            pass
    candidates = [shutil.which('soffice')]
    for key in ('PROGRAMFILES', 'PROGRAMFILES(X86)'):
        candidates.append(str(Path(os.environ.get(key, 'C:/Program Files')) / 'LibreOffice/program/soffice.exe'))
    for p in candidates:
        if p and Path(p).is_file():
            engines['LibreOffice'] = p
            break
    return engines


def read_image(path):
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(fs_path(path)) as src:
                if src.format not in ('JPEG', 'PNG'):
                    raise ValueError('文件内容不是 JPG 或 PNG')
                dpi = src.info.get('dpi', (300, 300))
                # One isotropic DPI preserves pixel aspect ratio. Invalid/unequal values use 300.
                try:
                    x, y = float(dpi[0]), float(dpi[1])
                    dpi_value = x if math.isfinite(x) and math.isfinite(y) and 10 <= x <= 2400 and abs(x-y) < 1 else 300
                except (TypeError, ValueError, IndexError):
                    dpi_value = 300
                src.load()
                image = ImageOps.exif_transpose(src)
                if image.mode in ('RGBA', 'LA') or 'transparency' in image.info:
                    rgba = image.convert('RGBA')
                    image = Image.new('RGB', rgba.size, 'white')
                    image.paste(rgba, mask=rgba.getchannel('A'))
                else:
                    image = image.convert('RGB')
                return image, dpi_value
    except Exception as exc:
        raise ConversionError(f'无法读取图片“{Path(path).name}”：文件损坏、格式不符、过大或无读取权限。\n技术信息：{type(exc).__name__}: {exc}') from exc


def validate_pdf(path):
    try:
        with open(path, 'rb') as stream:
            reader = PdfReader(stream, strict=True)
            if reader.is_encrypted or len(reader.pages) < 1:
                raise ValueError('PDF 加密或没有页面')
            for page in reader.pages:
                if float(page.mediabox.width) <= 0 or float(page.mediabox.height) <= 0:
                    raise ValueError('页面尺寸无效')
                content = page.get_contents()
                if content is not None:
                    content.get_data()
            return len(reader.pages)
    except Exception as exc:
        raise ConversionError(f'生成的 PDF 无效，未保存为成品：{type(exc).__name__}') from exc


def output_name(name):
    name = name.strip()
    if name.lower().endswith('.pdf'):
        name = name[:-4]
    if not name or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith(('.', ' ')):
        raise ConversionError('输出名称不能为空，不能包含路径或 Windows 不允许的字符。')
    if re.fullmatch(r'(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?', name, re.I):
        raise ConversionError('此名称是 Windows 保留名称，请换一个名称。')
    # Leave room for extension, counter, and the 255 UTF-16 component limit.
    if len(name.encode('utf-16-le')) // 2 > 230:
        raise ConversionError('输出名称过长，请缩短到 230 个字符以内。')
    return name


def prepare_output(directory):
    directory = fs_path(directory)
    try:
        directory.mkdir(parents=True, exist_ok=True)
        # NamedTemporaryFile retries PermissionError on Windows when os.access
        # misleadingly reports writable despite ACL denial. One exclusive probe fails promptly.
        probe = directory / ('.localpdf-probe-' + uuid.uuid4().hex)
        with open(probe, 'xb'):
            pass
        probe.unlink()
    except OSError as exc:
        raise ConversionError('输出目录无法写入：请检查权限、剩余空间，或选择其他文件夹。') from exc
    return directory


def commit_pdf(temp, directory, stem, event):
    pages = validate_pdf(temp)
    check_cancel(event)
    stem = output_name(stem)
    for i in range(10000):
        target = directory / (stem + (f' ({i})' if i else '') + '.pdf')
        if target.exists():
            continue
        try:
            # Windows rename is atomic and refuses replacement; temp is on the same volume.
            if sys.platform == 'win32':
                os.rename(temp, target)
            else:
                os.link(temp, target)
                os.unlink(temp)
            return Result(display_path(target), pages, target.stat().st_size)
        except FileExistsError:
            continue
        except OSError as exc:
            raise ConversionError('无法保存 PDF：目录不可写、磁盘空间不足或目标文件被占用。') from exc
    raise ConversionError('同名文件过多，请更改输出名称。')


def images_to_pdf(paths, directory, name, settings, event=None, progress=None):
    event = event or threading.Event()
    progress = progress or (lambda *args: None)
    if not paths:
        raise ConversionError('请先添加图片。')
    stem = output_name(name)
    directory = prepare_output(directory)
    if settings.mode not in ('a4', 'image') or settings.orientation not in ('auto', 'portrait', 'landscape') or not 0 <= settings.margin_mm <= 30:
        raise ConversionError('页面设置无效。')
    with tempfile.TemporaryDirectory(prefix='.localpdf-', dir=directory) as tmp:
        target = Path(tmp) / 'output.pdf'
        pdf = canvas.Canvas(str(target), pageCompression=1)
        for i, path in enumerate(paths):
            check_cancel(event)
            progress(i, '处理中', str(path))
            try:
                im, dpi = read_image(path)
                w, h = im.size
                if settings.mode == 'a4':
                    pw, ph = A4
                    if settings.orientation == 'landscape' or (settings.orientation == 'auto' and w > h):
                        pw, ph = ph, pw
                    margin = settings.margin_mm * 72 / 25.4
                    scale = min((pw - 2*margin)/w, (ph - 2*margin)/h)
                    dw, dh = w*scale, h*scale
                else:
                    pw, ph = w*72/dpi, h*72/dpi
                    if max(pw, ph) > 14400:
                        raise ConversionError('按图片尺寸生成的页面超过 200 英寸，请使用 A4 模式。')
                    dw, dh = pw, ph
                pdf.setPageSize((pw, ph))
                # No resizing or JPEG recompression: decoded RGB is embedded with lossless Flate.
                pdf.drawImage(ImageReader(im), (pw-dw)/2, (ph-dh)/2, dw, dh)
                pdf.showPage()
                im.close()
                progress(i, '已处理，待合并保存', str(path))
            except Exception:
                progress(i, '失败，请移除或替换后重试', str(path))
                raise
        check_cancel(event)
        pdf.save()
        return commit_pdf(target, directory, stem, event)


def docx_font_warnings(path):
    if Path(path).suffix.lower() != '.docx':
        return []
    try:
        with zipfile.ZipFile(fs_path(path)) as z:
            root = ET.fromstring(z.read('word/fontTable.xml'))
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        fonts = {e.get(ns+'name') for e in root if e.get(ns+'name')}
        import winreg
        installed = []
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts') as key:
                    for i in range(winreg.QueryInfoKey(key)[1]):
                        installed.append(winreg.EnumValue(key, i)[0].casefold())
            except OSError:
                pass
        missing = sorted(f for f in fonts if not any(f.casefold() in n for n in installed))
        return ['文档字体表中以下字体可能未安装（含未使用字体或别名，可能误报）：' + '、'.join(missing)] if missing else []
    except Exception:
        return []


def worker_command(request):
    if getattr(sys, 'frozen', False):
        return [sys.executable, '--word-worker', str(request)]
    return [sys.executable, str(Path(__file__).resolve().parents[1] / 'app.py'), '--word-worker', str(request)]


def cleanup_owned_word(pidfile):
    try:
        import psutil
        data = json.loads(Path(pidfile).read_text())
        proc = psutil.Process(data['pid'])
        if abs(proc.create_time() - data['created']) < 0.01 and proc.name().lower() == 'winword.exe':
            proc.kill()
            proc.wait(timeout=10)
    except Exception:
        pass


def word_to_pdf(path, directory, engine, event=None, timeout=180):
    event = event or threading.Event()
    path = fs_path(path)
    check_cancel(event)
    if path.suffix.lower() not in ('.docx', '.doc') or not path.is_file():
        raise ConversionError('Word 输入不存在或不是 .docx / .doc 文件。')
    if path.suffix.lower() == '.docx':
        try:
            with zipfile.ZipFile(path) as archive:
                if 'word/document.xml' not in archive.namelist():
                    raise ValueError('missing document')
        except (zipfile.BadZipFile, ValueError):
            raise ConversionError('DOCX 已加密、损坏或文件内容与扩展名不符。请在 Word 中检查；本工具不解密文档。')
        except OSError as exc:
            raise ConversionError('无法读取 Word 文档：文件被占用或没有读取权限。') from exc
    directory = prepare_output(directory)
    with tempfile.TemporaryDirectory(prefix='.localpdf-', dir=directory) as tmp:
        tmp = Path(tmp)
        target = tmp / 'document.pdf'
        proc = None
        try:
            if engine == 'word':
                request = tmp / 'request.json'
                response = tmp / 'response.json'
                request.write_text(json.dumps({'source': display_path(path), 'target': display_path(target), 'response': str(response), 'pidfile': str(tmp/'wordpid.json')}, ensure_ascii=False), encoding='utf-8')
                command = worker_command(request)
            else:
                if not Path(engine).is_file():
                    raise ConversionError('未找到 LibreOffice，请安装后重新启动工具。')
                # Work from a private copy and a private profile, never an existing LO session.
                source = tmp / ('document' + path.suffix.lower())
                shutil.copy2(path, source)
                command = [engine, '-env:UserInstallation='+Path(display_path(tmp/'profile')).as_uri(), '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', display_path(tmp), display_path(source)]
            proc = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            start = time.monotonic()
            while proc.poll() is None:
                if time.monotonic() - start > timeout:
                    raise ConversionError(f'转换超过 {timeout} 秒，已停止本工具的转换实例。文档可能需要密码、存在弹窗或过于复杂。')
                # Word cancellation is deliberately deferred to the end of this file.
                time.sleep(.1)
            check_cancel(event)
            if engine == 'word':
                if not response.exists():
                    raise ConversionError('Word 未正常完成导出。请确认 Office 已激活且可以正常打开此文档。')
                reply = json.loads(response.read_text(encoding='utf-8'))
                if not reply['ok']:
                    raise ConversionError(reply['error'])
            if not target.exists():
                raise ConversionError('转换引擎未生成 PDF。文档可能已加密、损坏或无法读取；请在 Word / LibreOffice 中检查。')
            return commit_pdf(target, directory, path.stem, event)
        finally:
            if proc and proc.poll() is None:
                import psutil
                try:
                    for child in psutil.Process(proc.pid).children(recursive=True):
                        child.kill()
                    proc.kill()
                    proc.wait(timeout=10)
                except (psutil.Error, subprocess.TimeoutExpired):
                    pass
            if engine == 'word':
                cleanup_owned_word(tmp/'wordpid.json')


def word_worker(request_path):
    import pythoncom
    import win32com.client
    import win32process
    import psutil
    req = json.loads(Path(request_path).read_text(encoding='utf-8'))
    pythoncom.CoInitialize()
    app = doc = probe = None
    owned = False
    answer = {'ok': False, 'error': 'Word 转换失败。'}
    try:
        before = set(psutil.pids())
        app = win32com.client.DispatchEx('Word.Application')
        probe = app.Documents.Add(Visible=False)
        _, pid = win32process.GetWindowThreadProcessId(probe.Windows(1).Hwnd)
        probe.Close(SaveChanges=0)
        probe = None
        if pid in before:
            raise ConversionError('Word 未创建独立进程，为保护已打开的文档，本次转换已停止。')
        owned = True
        Path(req['pidfile']).write_text(json.dumps({'pid': pid, 'created': psutil.Process(pid).create_time()}))
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        # Explicit noninteractive password prevents a password prompt; no password bypass.
        doc = app.Documents.Open(FileName=req['source'], ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
                                 PasswordDocument='LocalPDF-NoPassword', PasswordTemplate='LocalPDF-NoPassword',
                                 WritePasswordDocument='LocalPDF-NoPassword', WritePasswordTemplate='LocalPDF-NoPassword',
                                 Revert=False, Visible=False, OpenAndRepair=False, NoEncodingDialog=True)
        if doc.ProtectionType != -1:
            raise ConversionError('此 Word 文档存在编辑或权限保护。本工具不绕过保护，请先使用正常授权方式处理。')
        doc.ExportAsFixedFormat(OutputFileName=req['target'], ExportFormat=17, OpenAfterExport=False, OptimizeFor=0,
                                Range=0, Item=0, IncludeDocProps=False, KeepIRM=True, CreateBookmarks=1,
                                DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False)
        answer = {'ok': True}
    except Exception as exc:
        # Do not log exception text from Office, which can contain document data.
        code = getattr(exc, 'hresult', None)
        log.error('Word export failed file=%s type=%s hresult=%s', Path(req['source']).name, type(exc).__name__, code)
        answer = {'ok': False, 'error': str(exc) if isinstance(exc, ConversionError) else f'Word 无法导出此文档：可能已加密、损坏、被权限保护，或 Office 无法正常工作。请先在 Word 中检查。技术代码：{code}'}
    finally:
        if probe is not None:
            try:
                probe.Close(SaveChanges=0)
            except Exception:
                pass
        if owned:
            try:
                if doc is not None:
                    doc.Close(SaveChanges=0)
            except Exception:
                pass
            try:
                app.Quit(SaveChanges=0)
            except Exception:
                pass
        doc = app = None
        pythoncom.CoUninitialize()
        Path(req['response']).write_text(json.dumps(answer, ensure_ascii=False), encoding='utf-8')
