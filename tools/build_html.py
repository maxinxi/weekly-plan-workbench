"""Re-embed the maintained browser processor, without changing bundled dependencies."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
html_path = ROOT / '周计划工作台.html'
html = html_path.read_text(encoding='utf-8')
start = '// BEGIN SOURCE PROCESSOR'
end = '// END SOURCE PROCESSOR'
processor = (ROOT / 'source-processor.js').read_text(encoding='utf-8')
block = start + '\n' + processor + '''
  const sourceProcessor = createSourceProcessor(import_exceljs.default);
  async function preprocessSourceBuffer(buffer, filename) {
    return sourceProcessor.preprocessBuffer(buffer, filename);
  }
  async function writeProcessedSourceXlsx(buffer, parsed, measuresPlus = 1) {
    return sourceProcessor.processBuffer(buffer, parsed?.sourceName || "表1.xlsx", measuresPlus);
  }
''' + end + '\n\n'
if start in html:
    html = re.sub(re.escape(start) + r'.*?' + re.escape(end) + r'\n\n', lambda _: block, html, count=1, flags=re.S)
else:
    a = html.index('  function getMergeRowSpan(')
    b = html.index('  function mergeDateColumns(', a)
    html = html[:a] + block + html[b:]
    html = html.replace('    DEFAULT_MEASURES_PLUS: () => DEFAULT_MEASURES_PLUS,',
                        '    sourceProcessor: () => sourceProcessor,\n    DEFAULT_MEASURES_PLUS: () => DEFAULT_MEASURES_PLUS,')
# Desktop Excel enforces CT_SheetPr child order. Bundled ExcelJS had pageSetUpPr
# before outlinePr: real workbooks carrying outline properties became unreadable.
html = html.replace(
    'r2 = this.map.pageSetUpPr.render(e2, t2.pageSetup) || r2, r2 = this.map.outlinePr.render(e2, t2.outlineProperties) || r2',
    'r2 = this.map.outlinePr.render(e2, t2.outlineProperties) || r2, r2 = this.map.pageSetUpPr.render(e2, t2.pageSetup) || r2')
# Reuse JSZip already inside the bundle; no additional dependency or download.
html = html.replace('const n = { Workbook: e("./doc/workbook") },',
                    'const n = { Workbook: e("./doc/workbook"), SourceZip: e(441) },')
html_path.write_text(html, encoding='utf-8')
print('Updated embedded source processor')
