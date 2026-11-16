from pathlib import Path
from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1] / 'tests' / '测试材料 中文 spaces'


def create():
    ROOT.mkdir(parents=True, exist_ok=True)
    for name, size, color in [('1.jpg', (900, 450), 'red'), ('2.jpg', (450, 900), 'green'), ('10.jpg', (700, 500), 'blue')]:
        im = Image.new('RGB', size, color)
        draw = ImageDraw.Draw(im)
        draw.rectangle((15, 15, size[0]-15, size[1]-15), outline='white', width=10)
        draw.text((40, 40), name, fill='white', font_size=50)
        im.save(ROOT/name, quality=95, dpi=(300, 300))
    im = Image.new('RGB', (480, 240), 'orange')
    ImageDraw.Draw(im).rectangle((0, 0, 120, 240), fill='purple')
    exif = Image.Exif()
    exif[274] = 6
    im.save(ROOT/'EXIF 旋转.jpg', exif=exif)
    im = Image.new('RGBA', (500, 500), (0, 0, 0, 0))
    ImageDraw.Draw(im).ellipse((100, 100, 400, 400), fill=(0, 80, 240, 180))
    im.save(ROOT/'透明 图片.png')
    (ROOT/'损坏.png').write_bytes(b'not an image')
    (ROOT/'损坏.docx').write_bytes(b'not a docx')
    for name, advanced in [('中文 English 图表.docx', False), ('页眉 页脚 分页 公式.docx', True)]:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Microsoft YaHei'
        style.font.size = Pt(11)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        doc.add_heading('本地 PDF 转换测试', 0)
        doc.add_paragraph('中文段落：字体、图片与表格应保持完整。English text 12345.')
        doc.add_picture(str(ROOT/'1.jpg'), width=Inches(4.5))
        table = doc.add_table(rows=3, cols=3)
        table.style = 'Table Grid'
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                cell.text = f'单元格 {i+1}-{j+1}'
        if advanced:
            doc.sections[0].header.paragraphs[0].text = '测试页眉 Header'
            footer = doc.sections[0].footer.paragraphs[0]
            footer.add_run('页码 ')
            field = OxmlElement('w:fldSimple')
            field.set(qn('w:instr'), 'PAGE')
            footer._p.append(field)
            doc.add_page_break()
            doc.add_heading('第二页 公式与分页', 1)
            doc.add_paragraph('此内容必须出现在第二页。Equation below:')
            p = doc.add_paragraph()
            math = OxmlElement('m:oMath')
            run = OxmlElement('m:r')
            text = OxmlElement('m:t')
            text.text = 'E = mc²'
            run.append(text)
            math.append(run)
            p._p.append(math)
        doc.save(ROOT/name)
    return ROOT


if __name__ == '__main__':
    print(create())
