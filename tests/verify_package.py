"""Run the portable EXE with no Python/Conda on PATH, then independently inspect PDFs."""
import json
import os
from pathlib import Path
import subprocess
import fitz

root = Path(__file__).resolve().parents[1]
exe = root/'dist/LocalFilePDF/LocalFilePDF.exe'
samples = root/'tests/测试材料 中文 spaces'
env = dict(os.environ)
for key in ['PYTHONHOME', 'PYTHONPATH', 'QT_QPA_PLATFORM', 'QT_PLUGIN_PATH']:
    env.pop(key, None)
env['PATH'] = os.environ['SystemRoot'] + ';' + os.environ['SystemRoot'] + r'\System32'
result = subprocess.run([str(exe), '--verify-build', str(samples)], cwd=os.environ['TEMP'], env=env, timeout=240)
assert result.returncode == 0, result.returncode
out = root/'tests/打包验证'
report = json.loads((out/'report.json').read_text(encoding='utf-8'))
assert report['image']['results'] == 1 and report['word']['results'] == 2
pdfs = {}
for p in out.glob('*.pdf'):
    with fitz.open(p) as doc:
        pdfs[p.name] = {'pages': len(doc), 'bytes': p.stat().st_size}
        for i, page in enumerate(doc):
            page.get_pixmap(matrix=fitz.Matrix(1, 1)).save(str(out/(p.stem+f'-page-{i+1}.png')))
assert sorted(v['pages'] for v in pdfs.values()) == [1, 2, 4]
(out/'independent-validation.json').write_text(json.dumps(pdfs, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps({'exe_exit': result.returncode, 'pdfs': pdfs}, ensure_ascii=False, indent=2))
