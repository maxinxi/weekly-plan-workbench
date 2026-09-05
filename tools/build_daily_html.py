"""Build the offline HTML and standalone Python from the exact original engine.

No runtime CDN, npm, WebAssembly or rewritten layout engine is required.
"""
from pathlib import Path
import base64
import gzip
import json

ROOT = Path(__file__).resolve().parents[1]
server = (ROOT/'daily_plan_server.py').read_text(encoding='utf-8')
original = (ROOT/'daily_plan_legacy.py').read_text(encoding='utf-8-sig')
template = (ROOT/'tools/daily-workbench.template.html').read_text(encoding='utf-8')
files = {'daily_plan_server.py':server,'daily_plan_legacy.py':original}
payload = base64.b64encode(gzip.compress(json.dumps(files,ensure_ascii=False).encode(),mtime=0)).decode()
# Standalone Python writes its embedded source and HTML to a private temporary app folder.
# Its output/backup directory is stable and independent of this app folder.
def standalone(html):
    packed = base64.b64encode(gzip.compress(json.dumps({**files,'日计划工作台.html':html},ensure_ascii=False).encode(),mtime=0)).decode()
    return '''# -*- coding: utf-8 -*-
"""日计划工作台：双击启动；已内嵌原版 Python 与网页，不修改原表。"""
import base64, gzip, json, os, pathlib, subprocess, sys, tempfile
DATA = "'''+packed+'''"
def main():
    folder = pathlib.Path(tempfile.mkdtemp(prefix='daily-workbench-app-'))
    try:
        for name, content in json.loads(gzip.decompress(base64.b64decode(DATA))).items():
            (folder/name).write_text(content,encoding='utf-8')
        subprocess.run([sys.executable,str(folder/'daily_plan_server.py'),*sys.argv[1:]],check=True)
    finally:
        import shutil
        shutil.rmtree(folder)
if __name__ == '__main__':
    main()
'''
first_html = template.replace('__PAYLOAD__',json.dumps({'gzip':payload}))
py = standalone(first_html)
# This copy is only for the offline HTML's optional Python download button.
files['standalone.py'] = py
full_payload = base64.b64encode(gzip.compress(json.dumps(files,ensure_ascii=False).encode(),mtime=0)).decode()
html = template.replace('__PAYLOAD__',json.dumps({'gzip':full_payload}))
(ROOT/'日计划工作台.html').write_text(html,encoding='utf-8')
(ROOT/'日计划处理.py').write_text(standalone(html),encoding='utf-8')
print('Built 日计划工作台.html and 日计划处理.py (embedded original Python; no network dependencies)')
