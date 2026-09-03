# -*- coding: utf-8 -*-
"""
周计划工具（V5功能移植版 - 管控措施行高可调版）
功能：
1. 周计划明细 Excel 输出（基于V5的优化逻辑）
2. 三四五级风险汇总 TXT 输出
3. 周计划项目汇总文本输出

行高：
- 含「管控措施」的内容格：估算高度 + MEASURES_HEIGHT_PLUS（默认 +1 磅）
- 只改文件顶部的 MEASURES_HEIGHT_PLUS：1.0=+1，2.0=+2，0=不加
- 已取消原脚本「全表统一 -1」
- 超长内容（>409.5磅）仍自动插行续行
"""
import os
import re
import glob
import sys
import math
import traceback
from datetime import datetime, timedelta
from collections import OrderedDict
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break, RowBreak
import warnings
warnings.filterwarnings('ignore')

# 可选导入 win32com（仅 Windows 可用，用于 Excel 原生 AutoFit）
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# -------------------- 环境初始化 --------------------
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

ERROR_LOG_PATH = "周计划导出错误日志.txt"
_error_log_inited = False
_error_count = 0

# Excel 单行最大行高（磅），超过此值需要插行续行
EXCEL_MAX_ROW_HEIGHT = 409.5
# A3 横向可打印高度估算（磅）：297mm / 25.4 * 72 ≈ 842，减页边距后约 820
A3_LANDSCAPE_PRINTABLE_HEIGHT = 820.0
PAGE_HEIGHT_SAFETY_FACTOR = 0.96

# ============================================================
# ★★★ 行高微调：只改这一处 ★★★
# 含「管控措施」的内容单元格 = 估算行高 + 此值（磅）
#   1.0 → +1     2.0 → +2     0 → 不加
# 其他单元格不加，也不再全表统一减 1。
# ============================================================
MEASURES_HEIGHT_PLUS = 1.0

# -------------------- 工具函数 --------------------
def safe_print(msg, error=False):
    """安全打印（避免编码崩溃）"""
    prefix = "❌ " if error else ""
    try:
        print(f"{prefix}{msg}")
    except:
        clean = re.sub(r'[\x1b\[0-9;]*[mGKHF]', '', msg)
        print(clean.encode('gbk', errors='ignore').decode('gbk', errors='ignore'))

def _ensure_error_log():
    global _error_log_inited
    if _error_log_inited:
        return
    _error_log_inited = True
    header = f"\n{'='*60}\n"
    header += f"❌ 错误日志 | 启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    header += f"💻 环境: Python {sys.version.split()[0]} | 目录: {os.getcwd()}\n"
    header += f"{'='*60}\n"
    try:
        with open(ERROR_LOG_PATH, 'w', encoding='utf-8') as f:
            f.write(header)
    except Exception:
        pass

def log_error(file_name, error_type, error_msg, traceback_str=""):
    global _error_count
    _error_count += 1
    _ensure_error_log()
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_entry = f"\n[{timestamp}] 源文件: {file_name}\n"
    log_entry += f"  🚨 类型: {error_type}\n"
    log_entry += f"  💡 描述: {error_msg}\n"
    if traceback_str:
        log_entry += f"  📋 堆栈:\n{traceback_str[:500]}\n"
    log_entry += "-" * 60
    try:
        with open(ERROR_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(log_entry)
    except:
        print(f"⚠️  日志写入失败！请记录: {error_msg}")

def unmerge_and_fill(ws):
    """拆分合并单元格并填充原值。
    先收集所有合并区域的值，再统一 unmerge，最后统一赋值，
    避免重叠合并区域导致 MergedCell 赋值失败。
    """
    merged = list(ws.merged_cells.ranges)
    if not merged:
        return ws
    # 1. 收集所有合并区域的值
    fill_data = []
    for rng in merged:
        min_row, min_col, max_row, max_col = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        try:
            val = ws.cell(min_row, min_col).value
        except Exception:
            val = None
        fill_data.append((min_row, min_col, max_row, max_col, val))
    # 2. 先 unmerge 所有区域
    for rng in merged:
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass
    # 3. 再统一赋值
    for min_row, min_col, max_row, max_col, val in fill_data:
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                try:
                    ws.cell(r, c).value = val
                except Exception:
                    pass
    return ws

def normalize_content(s):
    if pd.isna(s): return ""
    return re.sub(r'\s+', ' ', str(s)).strip()

def normalize_clock_text(value):
    """规范化 H:MM/HH:MM，允许 24:00，非法值原样返回。"""
    text = str(value).strip()
    m = re.fullmatch(r'(\d{1,2}):(\d{2})', text)
    if not m:
        return text
    hour, minute = map(int, m.groups())
    if not (0 <= minute <= 59 and 0 <= hour <= 24):
        return text
    if hour == 24 and minute != 0:
        return text
    return f"{hour:02d}:{minute:02d}"

def canonicalize_time(s):
    """规范化时间字符串，供去重使用。兼容日期 / 或 - 分隔，兼容遗漏"-"的情况。"""
    if pd.isna(s):
        return ""
    text = str(s).strip()
    # 优先匹配带 "-" 的标准格式
    pattern = r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})?\s*(\d{1,2}:\d{2})?'
    m = re.search(pattern, text)
    if not m:
        # 兼容遗漏"-"的情况：日期 时间 日期 时间
        pattern_no_dash = r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s+(\d{1,2}:\d{2})\s+(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s+(\d{1,2}:\d{2})'
        m = re.search(pattern_no_dash, text)
    if not m:
        return text
    d1, t1, d2, t2 = m.groups()
    d1 = d1.replace('-', '/')
    d2 = (d2 or d1).replace('-', '/')
    try:
        d1_norm = datetime.strptime(d1, "%Y/%m/%d").strftime("%Y/%m/%d")
        d2_norm = datetime.strptime(d2, "%Y/%m/%d").strftime("%Y/%m/%d")
    except ValueError:
        return text
    t1_norm = normalize_clock_text(t1)
    t2_default = "18:00" if d1_norm != d2_norm else "17:30"
    t2_norm = normalize_clock_text(t2 or t2_default)
    return f"{d1_norm}|{t1_norm}|{d2_norm}|{t2_norm}"

def extract_daily_time_segments(time_str):
    """将作业时间拆为每天的时间段；兼容无 - 分隔符（漏填时段）以及遗漏"-"的完整时段。"""
    if pd.isna(time_str):
        return {}
    s = str(time_str).strip()
    # 模式A：完整格式 "YYYY/M/D H:MM - YYYY/M/D H:MM"（带 - 分隔）
    pat_a = r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s*(\d{1,2}:\d{2})\s*[-–]\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})?\s*(\d{1,2}:\d{2})?'
    m = re.search(pat_a, s)
    if not m:
        # 模式A2：遗漏"-"的完整格式 "YYYY/M/D H:MM YYYY/M/D H:MM"
        pat_a2 = r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s+(\d{1,2}:\d{2})\s+(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s+(\d{1,2}:\d{2})'
        m = re.search(pat_a2, s)
    if m:
        d1, t1, d2, t2 = m.groups()
        d1 = d1.replace('-', '/')
        d2 = (d2 or d1).replace('-', '/')
        try:
            start_date = datetime.strptime(d1, "%Y/%m/%d")
            end_date = datetime.strptime(d2, "%Y/%m/%d")
        except ValueError:
            return {}
        if end_date < start_date:
            return {}
        t1 = normalize_clock_text(t1)
        t2 = normalize_clock_text(t2) if t2 else None
    else:
        # 模式B：只有日期，无时间段（漏填）——按整天 08:30-18:00
        pat_b = r'(\d{4}[/-]\d{1,2}[/-]\d{1,2})'
        m = re.search(pat_b, s)
        if not m:
            return {}
        d1 = m.group(1).replace('-', '/')
        try:
            start_date = datetime.strptime(d1, "%Y/%m/%d")
            end_date = start_date
        except ValueError:
            return {}
        t1, t2 = "08:30", "18:00"
    segments = {}
    current = start_date
    while current <= end_date:
        if start_date == end_date:
            start_t, end_t = t1, (t2 or "17:30")
        elif current == start_date:
            start_t, end_t = t1, "18:00"
        else:
            start_t, end_t = "08:30", "18:00"
        segments[current.date()] = (start_t, end_t)
        current += timedelta(days=1)
    return segments

def format_job_content(content):
    parts = re.split(r'[;；]', content)
    parts = [part.strip() for part in parts if part.strip()]
    return parts

def find_excel_files():
    """查找源表1文件，排除程序自身生成的周计划明细文件。"""
    files = [
        f for f in glob.glob("*表1*.xlsx")
        if os.path.isfile(f)
        and not f.startswith('~$')
        and "（周计划明细）" not in f
        and "(周计划明细)" not in f
    ]
    return files

# ======================================
# 专业排序权重定义
# ======================================
PROF_ORDER = [
    ("变电", ["变电"]),
    ("输电", ["输电"]),
    ("配电", ["配电"]),
    ("配网工程", ["配网工程", "配网", "配网施工", "配网检修", "配网运维"]),
    ("信通", ["信通", "通信", "信号", "信息通信"]),
    ("营销", ["营销", "业扩", "客户服务", "计量", "增容"])
]

def get_prof_priority(spec):
    if not spec:
        return len(PROF_ORDER)
    s = str(spec)
    for idx, (_, keys) in enumerate(PROF_ORDER):
        for k in keys:
            if k in s:
                return idx
    return len(PROF_ORDER)

def get_standard_profession(spec):
    if pd.isna(spec) or not str(spec).strip():
        return "未分类"
    s = str(spec).strip()
    if "营销" in s:
        return "营销"
    for prof_name, keys in PROF_ORDER:
        if prof_name == "营销":
            continue
        if any(k in s for k in keys):
            return prof_name
    marketing_keys = ["业扩", "客户服务", "高压计量", "高压增容", "计量", "增容"]
    if any(k in s for k in marketing_keys):
        return "营销"
    return s

def is_marketing_profession(spec):
    return get_standard_profession(spec) == "营销"

def clean_text_value(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text

def get_level_priority(person_text):
    if not person_text:
        return 3
    s = str(person_text)
    if "领导" in s:
        return 1
    management_keywords = ["主任", "专责", "管理", "组长", "主管"]
    if any(k in s for k in management_keywords):
        return 2
    return 3

def clean_attachment_text(text):
    if pd.isna(text) or str(text).strip().lower() in ('nan', 'none', ''):
        return ""
    s = str(text).strip()
    s = re.sub(r"^附件(?:\s*[-—]?\s*\d+)+\s*[-—]?\s*", "", s)
    s = re.sub(r"^附件\s*[-—]?\s*", "", s)
    return s.strip()

def find_column_by_keywords(columns, keywords, required=True):
    for col in columns:
        col_clean = str(col).strip().replace('\n', '').replace(' ', '').replace('\r', '')
        for kw in keywords:
            if kw in col_clean:
                return col
    if required:
        raise KeyError(f"未找到匹配关键词 {keywords} 的列。")
    return None

def extract_risk_level(risk_val):
    if not risk_val:
        return None
    text = str(risk_val)
    for level in ("五级", "四级", "三级"):
        if level in text:
            return level
    return None

def infer_base_year_from_time_column(df, col_time):
    years = []
    for v in df[col_time].dropna():
        m = re.search(r'(\d{4})/\d{1,2}/\d{1,2}', str(v))
        if m:
            years.append(int(m.group(1)))
    return min(years) if years else datetime.now().year

def extract_date_range_from_title(title, base_year):
    m = re.search(r'(\d{1,2})月(\d{1,2})日(?:[—\-~～]|至)(\d{1,2})月(\d{1,2})日', str(title))
    if not m:
        raise ValueError("标题中未找到日期范围")
    m1, d1, m2, d2 = map(int, m.groups())
    start = datetime(base_year, m1, d1)
    end = datetime(base_year, m2, d2)
    if end < start:
        end = datetime(base_year + 1, m2, d2)
    return f"{m1}.{d1}", f"{m2}.{d2}", start, end

# ======================================
# 公共数据层
# ======================================
def load_source_df(src):
    wb = load_workbook(src, data_only=True)
    ws = wb.active
    title = str(ws["A1"].value).strip() if ws["A1"].value is not None else ""
    unmerge_and_fill(ws)
    headers = [cell.value for cell in ws[2]]
    data = [tuple(cell.value for cell in row) for row in ws.iter_rows(min_row=3)]
    df = pd.DataFrame(data, columns=headers)
    df = df.dropna(how='all').reset_index(drop=True)
    return df, title

def resolve_columns(df):
    cols = df.columns.tolist()
    return {
        'content': find_column_by_keywords(cols, ['作业内容', '工作内容', '作业项目', '作业任务']),
        'spec': find_column_by_keywords(cols, ['专业']),
        'risk': find_column_by_keywords(cols, ['风险等级', '作业风险', '风险级别']),
        'power_plan': find_column_by_keywords(cols, ['关联停电计划'], required=False),
        'time': df.columns[9] if len(df.columns) > 9 else None,
        'person': df.columns[17] if len(df.columns) > 17 else None,
        'project': df.columns[1] if len(df.columns) > 1 else None,
        'measures': find_column_by_keywords(
            cols, ['管控措施', '控制措施', '防控措施', '风险措施'], required=False
        ),
    }

def _safe_row_value(row, col):
    if col is None or col not in row.index:
        return None
    return row[col]

def build_standard_records(df, cols):
    col_content = cols['content']
    col_time = cols['time']
    if col_time is None:
        return []

    df_valid = df.dropna(subset=[col_content, col_time], how='any').reset_index(drop=True)

    records = []
    seen = set()

    for _, row in df_valid.iterrows():
        content_raw = clean_text_value(_safe_row_value(row, col_content))
        content_norm = normalize_content(_safe_row_value(row, col_content))
        time_raw = _safe_row_value(row, col_time)
        time_canon = canonicalize_time(time_raw)

        project_raw = clean_text_value(_safe_row_value(row, cols['project']))
        project_norm = normalize_content(_safe_row_value(row, cols['project']))
        project_display = project_raw[:-1] if project_raw.endswith("工程。") else project_raw

        spec_raw = clean_text_value(_safe_row_value(row, cols['spec']))
        spec_norm = normalize_content(_safe_row_value(row, cols['spec']))
        std_spec = get_standard_profession(spec_raw)
        is_marketing = std_spec == "营销"

        risk_raw = clean_text_value(_safe_row_value(row, cols['risk']))
        risk_norm = normalize_content(_safe_row_value(row, cols['risk']))
        risk_level = extract_risk_level(risk_raw)

        person_raw = clean_attachment_text(_safe_row_value(row, cols['person']))
        person_norm = normalize_content(_safe_row_value(row, cols['person']))

        measures_raw = clean_text_value(_safe_row_value(row, cols.get('measures')))
        measures_norm = normalize_content(_safe_row_value(row, cols.get('measures')))

        power_raw = clean_text_value(_safe_row_value(row, cols['power_plan']))
        power_norm = normalize_content(_safe_row_value(row, cols['power_plan']))
        has_power_plan = power_raw.strip().lower() not in ('否', '空', 'none', 'nan', '')

        if is_marketing:
            full_content = content_raw
        elif (project_display and "工程" in project_display
              and "：" not in project_display and ":" not in project_display
              and not content_raw.startswith(project_display + "：")
              and not content_raw.startswith(project_display + ":")):
            full_content = f"{project_display}：{content_raw}"
        else:
            full_content = content_raw

        l_rank = get_level_priority(person_raw)
        p_rank = get_prof_priority(spec_raw)
        is_low_voltage = "低压计量" in content_raw or "低压计量" in project_display

        dedup_key = (content_norm, time_canon, project_norm, spec_norm,
                     risk_norm, person_norm, power_norm, measures_norm)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        records.append({
            'content_raw': content_raw,
            'content_norm': content_norm,
            'full_content': full_content,
            'project_raw': project_raw,
            'project_display': project_display,
            'spec_raw': spec_raw,
            'std_spec': std_spec,
            'is_marketing': is_marketing,
            'risk_raw': risk_raw,
            'risk_level': risk_level,
            'time_raw': time_raw,
            'time_canon': time_canon,
            'person_raw': person_raw,
            'measures_raw': measures_raw,
            'has_power_plan': has_power_plan,
            'l_rank': l_rank,
            'p_rank': p_rank,
            'is_low_voltage': is_low_voltage,
        })

    return records

def expand_to_daily(records):
    daily = []
    for rec in records:
        segments = extract_daily_time_segments(rec['time_raw'])
        for seg_day, (st, et) in segments.items():
            daily.append({
                **rec,
                'date': seg_day,
                'start_time': st,
                'end_time': et,
            })
    return daily

# ======================================
# 源表预处理
# ======================================
def _is_yellow_fill(cell):
    """判断单元格是否标黄（标准黄或常见浅黄色）。"""
    try:
        fill = cell.fill
        if fill is None or fill.fgColor is None:
            return False
        rgb = fill.fgColor.rgb
        if rgb is None:
            return False
        rgb = str(rgb).upper()
        # 标准黄 FFFF00，以及常见浅黄
        if rgb.endswith("FFFF00") or rgb == "00FFFF00":
            return True
        if len(rgb) == 8:
            r = int(rgb[2:4], 16)
            g = int(rgb[4:6], 16)
            b = int(rgb[6:8], 16)
            return r >= 220 and g >= 200 and b <= 170
    except Exception:
        pass
    return False

def _has_strikethrough(cell):
    """判断单元格是否有删除线（整格或字符级）。"""
    try:
        if cell.font and cell.font.strikethrough:
            return True
    except Exception:
        pass
    return False

def normalize_majority_format(ws, max_row, max_col):
    """多数派格式统一：修正起始行为第3行，禁止修改表头和标题。"""
    from collections import Counter
    DATA_START_ROW = 3  # 跳过标题和表头
    changed_total = 0

    for col in range(1, max_col + 1):
        cells = []
        for row in range(DATA_START_ROW, max_row + 1):
            cell = ws.cell(row, col)
            if cell.value is None or str(cell.value).strip() == "":
                continue
            cells.append((row, cell))

        if len(cells) < 2:
            continue

        for prop_name in ("horizontal", "vertical", "font_name", "font_size"):
            observed = []
            for row, cell in cells:
                if prop_name == "horizontal":
                    val = cell.alignment.horizontal
                elif prop_name == "vertical":
                    val = cell.alignment.vertical
                elif prop_name == "font_name":
                    val = cell.font.name
                elif prop_name == "font_size":
                    val = cell.font.size
                if val is not None:
                    observed.append((row, cell, val))

            if len(observed) < 2:
                continue

            dominant_val, dominant_cnt = Counter(v for _, _, v in observed).most_common(1)[0]
            if dominant_cnt * 2 <= len(observed):
                continue

            for row, cell, cur_val in observed:
                if cur_val == dominant_val:
                    continue
                if prop_name in ("horizontal", "vertical"):
                    old = cell.alignment
                    cell.alignment = Alignment(
                        horizontal=dominant_val if prop_name == "horizontal" else old.horizontal,
                        vertical=dominant_val if prop_name == "vertical" else old.vertical,
                        wrap_text=old.wrap_text,
                    )
                elif prop_name in ("font_name", "font_size"):
                    old = cell.font
                    cell.font = Font(
                        name=dominant_val if prop_name == "font_name" else old.name,
                        size=dominant_val if prop_name == "font_size" else old.size,
                        bold=old.bold,
                        italic=old.italic,
                        color=old.color,
                    )
                changed_total += 1

    return changed_total
def normalize_cell_line_breaks(ws, max_row, max_col):
    """统一单元格内换行格式：
    1. 统一换行符为 \\n（\\r\\n、\\r 转为 \\n）
    2. 清理每行首尾空白
    3. 合并连续多个空行为单个空行
    4. 去掉首尾空行
    数据从第3行开始。返回修改的单元格数。
    """
    changed = 0
    DATA_START_ROW = 1
    for row in range(DATA_START_ROW, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            if cell.value is None or not isinstance(cell.value, str):
                continue
            original = cell.value
            if '\n' not in original and '\r' not in original:
                continue
            # 统一换行符
            text = original.replace('\r\n', '\n').replace('\r', '\n')
            # 清理每行首尾空白
            lines = [line.strip() for line in text.split('\n')]
            # 合并连续空行为单个空行
            cleaned = []
            prev_empty = False
            for line in lines:
                if line == '':
                    if not prev_empty:
                        cleaned.append(line)
                    prev_empty = True
                else:
                    cleaned.append(line)
                    prev_empty = False
            # 去掉首尾空行
            while cleaned and cleaned[0] == '':
                cleaned.pop(0)
            while cleaned and cleaned[-1] == '':
                cleaned.pop()
            new_text = '\n'.join(cleaned)
            if new_text != original:
                cell.value = new_text
                changed += 1
    if changed > 0:
        safe_print(f"  📝 换行统一：修正了 {changed} 个单元格的换行格式（统一换行符/清理空行）")
    return changed

def merge_name_phone(ws, max_row, max_col):
    """
    合并单元格内换行的姓名与电话（无空格紧贴拼接）。
    1. 修正起始行为第3行，严禁触碰第1行大标题与第2行业务表头。
    2. 优先动态识别表头含"人"、"负责人"、"联系"等关键词的列，避免全表误伤。
    """
    changed = 0
    DATA_START_ROW = 3  # 严格从数据行开始，保护标题与表头
    phone_pattern = re.compile(r'^(1\d{10}|0\d{2,3}-?\d{7,8})$')

    # 动态定位人员/联系方式相关列
    target_cols = []
    for c in range(1, max_col + 1):
        header_val = str(ws.cell(2, c).value or "").strip()
        if any(kw in header_val for kw in ("人", "联系", "电话", "到岗", "监护")):
            target_cols.append(c)

    # 若表头未识别到，降级扫描前20列，但避开前3列（序号、工程、专业）
    if not target_cols:
        target_cols = list(range(4, max_col + 1))

    for col in target_cols:
        for row in range(DATA_START_ROW, max_row + 1):
            cell = ws.cell(row, col)
            if not cell.value or not isinstance(cell.value, str):
                continue
            original = cell.value
            if '\n' not in original and '\r' not in original:
                continue

            text = original.replace('\r\n', '\n').replace('\r', '\n')
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            if len(lines) < 2:
                continue

            # 校验是否存在匹配电话号码的行
            if any(phone_pattern.match(line) for line in lines):
                # 按照业务规则：姓名与电话紧贴拼接，不加空格
                new_text = ''.join(lines)
                if new_text != original:
                    cell.value = new_text
                    changed += 1

    safe_print(f"  📞 姓名电话合并：修正了 {changed} 个单元格（无空格紧贴合并）")
    return changed
def fix_missing_time_dash(ws, max_row):
    """时间列破折号补齐：严格限定在识别出的时间列执行，未识别则退出。"""
    DATA_START_ROW = 3
    max_col = min(ws.max_column, 20)
    time_col = None

    for r in (2, 1):
        for c in range(1, max_col + 1):
            val = str(ws.cell(r, c).value or "")
            if len(val) <= 20 and "时间" in val:
                time_col = c
                break
        if time_col:
            break

    if not time_col:
        safe_print("  ⏰ 时间补齐：未明确识别到时间列表头，放弃自动补齐以防篡改工作内容")
        return 0

    changed = 0
    for row in range(DATA_START_ROW, max_row + 1):
        cell = ws.cell(row, time_col)
        original = str(cell.value or "").strip()
        time_matches = list(re.finditer(r'\d{1,2}:\d{2}', original))
        if len(time_matches) < 2:
            continue

        first_end = time_matches[0].end()
        second_start = time_matches[1].start()
        between = original[first_end:second_start]
        has_separator = any(sep in between for sep in ('-', '–', '—', '至', '到', '~', '～'))
        if not has_separator:
            rest = original[first_end:]
            cell.value = original[:first_end] + '-' + rest.lstrip()
            changed += 1

    return changed
def preprocess_source_file(src):
    """
    源表预处理（始终执行，就地插行并保存，维持原文件名）：
    1. 创建"源文件备份"文件夹，复制原始源文件（原始上报文件始终有备份）
    2. 检测并删除示例行/填报人行（仅按文本含"例"/"示例"剔除，不再因黄色填充删行）
    3. 单元格内换行统一：统一换行符、清理多余空行
    4. 姓名+电话合并（无空格紧贴拼接，仅处理人员/联系相关列）
    5. 时间列补齐遗漏的"-"分隔符（严格限定时间列，未识别则跳过）
    6. 对源表【管控措施】列超长整行插行 + 均分行高（未超长保持原行高）
    7. 保存回原始源文件（文件名不变）
    返回：原始源文件路径（已就地插行处理）
    """
    import shutil
    global _processed_source_cache

    # 单次预处理判定：同一文件只清洗一次，避免 1,2,3 批量任务重复触发
    if src in _processed_source_cache:
        safe_print(f"  🔁 源表 {os.path.basename(src)} 已预处理过，直接复用，跳过")
        return src

    safe_print(f"  🔍 源表预处理开始...")

    # 1. 备份原始文件到"源文件备份"文件夹
    backup_dir = "源文件备份"
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    backup_path = os.path.join(backup_dir, os.path.basename(src))
    try:
        shutil.copy2(src, backup_path)
        safe_print(f"  📦 已备份原始文件到: {backup_dir}/{os.path.basename(src)}")
    except Exception as e:
        safe_print(f"  ⚠️  备份失败: {e}，继续预处理")

    try:
        wb = load_workbook(src)
        ws = wb.active
    except Exception as e:
        safe_print(f"  ⚠️  预处理跳过：无法加载文件 {e}")
        return src

    example_row = None
    filler_row = None
    example_reason = ""

    # 2. 检测示例行/填报人行（前20行）
    #    仅依据文本是否明确包含"例"字或"示例"来剔除模板示例行，不因黄色填充删行
    scan_range = min(20, ws.max_row)
    for r in range(1, scan_range + 1):
        a_val = ws.cell(r, 1).value
        b_val = ws.cell(r, 2).value

        is_example = False
        reasons = []
        # A列含"例"字且长度<=5（序号列是短文本，排除长标题误判）
        if a_val and isinstance(a_val, str) and len(a_val) <= 5 and ("例" in a_val or "示例" in a_val):
            is_example = True
            reasons.append(f"含'例'字('{a_val}')")
        if _has_strikethrough(ws.cell(r, 1)):
            is_example = True
            reasons.append("有删除线")

        if is_example and example_row is None:
            example_row = r
            example_reason = "、".join(reasons)
            safe_print(f"    📌 检测到示例行：第{r}行（{example_reason}）")

        if b_val:
            b_clean = str(b_val).strip()
            if b_clean in ("填报人及联系方式", "填报人"):
                filler_row = r
                safe_print(f"    📌 检测到填报人行：第{r}行（B列='{b_clean}'）")

    # 先删除示例行/填报人行（避免影响后续统计）
    if example_row is not None or filler_row is not None:
        rows_to_delete = sorted(set(r for r in [example_row, filler_row] if r is not None), reverse=True)
        for r in rows_to_delete:
            try:
                ws.delete_rows(r)
            except Exception as e:
                safe_print(f"  ⚠️  删除第{r}行失败: {e}")
        deleted_desc = []
        if example_row:
            deleted_desc.append(f"示例行(第{example_row}行, {example_reason})")
        if filler_row:
            deleted_desc.append(f"填报人行(第{filler_row}行)")
        safe_print(f"  🔧 已删除 {'、'.join(deleted_desc)}")
    else:
        safe_print(f"  ✅ 未检测到示例行和填报人行，跳过删除（扫描了前{scan_range}行）")

    format_max_col = min(ws.max_column, 20)

    # 3. 单元格内换行统一
    normalize_cell_line_breaks(ws, ws.max_row, format_max_col)

    # 4. 姓名+电话合并
    merge_name_phone(ws, ws.max_row, format_max_col)

    # 5. 时间列补齐"-"
    fix_missing_time_dash(ws, ws.max_row)

    # 6. 对源表【管控措施】列超长整行插行（就地，未超长保持原行高）
    try:
        insert_overflow_rows_by_measures(ws, start_row=3)
    except Exception as e:
        safe_print(f"  ⚠️  管控措施列插行处理失败: {e}")

    # 7. 保存回原始源文件（文件名不变，维持原名）
    try:
        if not os.access(src, os.W_OK):
            safe_print(f"  ⚠️  检测到文件为只读，正在解除只读属性...")
            try:
                import stat
                os.chmod(src, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass
            try:
                import subprocess
                subprocess.run(['attrib', '-r', src], shell=True, capture_output=True)
            except Exception:
                pass
        wb.save(src)
        wb.close()
        safe_print(f"  ✅ 预处理完成：源表已就地插行并保存（文件名不变）{os.path.basename(src)}")
    except Exception as e:
        safe_print(f"  ❌ 保存源文件失败: {e}")
        safe_print(f"  💡 可能原因：文件正在被Excel打开 / 文件为只读 / 权限不足")
        try:
            wb.close()
        except Exception:
            pass
        return src

    _processed_source_cache.add(src)
    return src
def cleanup_preprocessed_file(path):
    """预处理已改为就地修改源文件，不再产生临时文件，此函数保留为空操作。"""
    pass


# ======================================
# A3 打印 + 不跨行分页 + 插行续行
# ======================================
def setup_a3_print(ws, total_rows, total_cols):
    """
    设置 A3 横向打印：窄页边距、1页宽适配、不打印网格线、重复标题行。
    借鉴参考文件的 setup_print_page 逻辑，用 openpyxl 实现。
    """
    # 纸张：A3 (openpyxl 中 PAPERSIZE_A3 = 8)
    ws.page_setup.paperSize = 8
    ws.page_setup.orientation = 'landscape'

    # 适配为1页宽，高度自动
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 不打印网格线
    ws.print_options.gridLines = False

    # 窄页边距（英寸）
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.15
    ws.page_margins.bottom = 0.15
    ws.page_margins.header = 0.05
    ws.page_margins.footer = 0.05

    # 打印区域
    last_col_letter = get_column_letter(total_cols)
    ws.print_area = f"A1:{last_col_letter}{total_rows}"

    # 每页重复大标题行
    ws.print_title_rows = '1:1'

def autofit_rows_with_excel(filepath, plus_height=2.0, start_row=3, content_col_start=2):
    """
    用 Excel 原生 AutoFit（win32com）调整行高，再加 plus_height 磅余量。
    借鉴日计划 AutoFitRowEx 的临时工作表测量法：
    1. 先整体 AutoFit 作为基础
    2. 逐行用临时工作表测量内容单元格的真实高度（不受合并单元格影响）
    3. 取较大值 + 余量设置行高

    仅 Windows 可用，需要安装 pywin32。

    参数:
        filepath: Excel 文件路径
        plus_height: AutoFit 后增加的行高余量（磅）
        start_row: 从第几行开始调整（跳过标题和表头）
        content_col_start: 内容列起始列（默认第2列，第1列是日期）
    返回:
        (success, adjusted, message)
    """
    if not HAS_WIN32COM:
        return False, 0, "未安装 pywin32，无法使用 Excel 原生 AutoFit（请运行 pip install pywin32）"

    excel = None
    wb = None
    temp_ws = None
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        wb = excel.Workbooks.Open(os.path.abspath(filepath))
        ws = wb.Worksheets(1)

        # 创建临时工作表（借鉴日计划 create_autofitrowex_temp_sheet）
        temp_ws = wb.Worksheets.Add()
        temp_ws.Name = "__autofit_temp__"
        temp_ws.Visible = False

        # 获取使用的行数和列数
        used_rows = ws.UsedRange.Rows.Count
        used_cols = ws.UsedRange.Columns.Count
        adjusted = 0

        # 第一步：先整体 AutoFit 作为基础
        try:
            ws.Rows(f"{start_row}:{used_rows}").AutoFit()
        except Exception:
            pass

        # 第二步：逐行用临时工作表测量内容单元格的真实高度
        for r in range(start_row, used_rows + 1):
            try:
                max_measured_height = 0.0

                # 遍历内容列，找到有内容的锚点单元格（合并单元格的左上角）
                for c in range(content_col_start, used_cols + 1):
                    cell = ws.Cells(r, c)
                    value = cell.Value
                    if value is None or str(value).strip() == "":
                        continue

                    # 检查是否是合并单元格的锚点（MergeArea 的左上角）
                    try:
                        merge_area = cell.MergeArea
                        if merge_area.Cells(1, 1).Address != cell.Address:
                            continue  # 不是锚点，跳过
                    except Exception:
                        pass

                    # 获取该列的列宽
                    try:
                        col_width = ws.Columns(c).ColumnWidth
                    except Exception:
                        col_width = 25.0

                    # 把内容复制到临时工作表（借鉴日计划 prepare_temp_cell_from_source）
                    temp_cell = temp_ws.Cells(1, 1)
                    temp_cell.Value = value
                    try:
                        temp_cell.Font.Name = cell.Font.Name
                        temp_cell.Font.Size = cell.Font.Size
                        temp_cell.Font.Bold = cell.Font.Bold
                        temp_cell.WrapText = True
                        temp_cell.HorizontalAlignment = cell.HorizontalAlignment
                        temp_cell.VerticalAlignment = cell.VerticalAlignment
                    except Exception:
                        pass
                    temp_ws.Columns(1).ColumnWidth = col_width

                    # AutoFit 临时工作表的第1行，测量真实高度
                    try:
                        temp_ws.Rows(1).AutoFit()
                    except Exception:
                        pass
                    try:
                        measured_h = float(temp_ws.Rows(1).RowHeight)
                        if measured_h > 0:
                            max_measured_height = max(max_measured_height, measured_h)
                    except Exception:
                        pass

                    # 清空临时单元格
                    temp_cell.Value = ""

                # 取 AutoFit 后的行高和测量高度的较大值，加余量
                try:
                    autofit_h = float(ws.Rows(r).RowHeight)
                except Exception:
                    autofit_h = 0.0
                final_h = max(autofit_h, max_measured_height) + plus_height
                final_h = min(409.5, final_h)  # 不超过单行上限

                if final_h > 0:
                    ws.Rows(r).RowHeight = final_h
                    adjusted += 1
            except Exception:
                continue

        # 删除临时工作表
        try:
            temp_ws.Delete()
            temp_ws = None
        except Exception:
            pass

        wb.Save()
        wb.Close()
        wb = None
        excel.Quit()
        excel = None
        pythoncom.CoUninitialize()
        return True, adjusted, f"Excel 原生 AutoFit（临时表测量法）完成，调整 {adjusted} 行，+{plus_height}磅余量"

    except Exception as e:
        if temp_ws:
            try:
                temp_ws.Delete()
            except Exception:
                pass
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False, 0, f"Excel 原生 AutoFit 失败: {str(e)}"

def _get_row_height(ws, row, default=15.0):
    """获取行高，未设置时返回默认值。"""
    h = ws.row_dimensions[row].height
    if h is None or h <= 0:
        return default
    return float(h)

def _estimate_text_height(text, col_width_chars, font_size=11):
    """
    估算文本在给定列宽下需要的行高（磅）。
    col_width_chars: Excel 列宽单位（约等于字符数）
    借鉴参考文件的 AutoFitRowEx 思路，基于字符数估算。
    """
    if not text:
        return 15.0
    lines = str(text).split('\n')
    total_lines = 0
    # 中文字符约占 2.3 个列宽单位（Excel 实际渲染中文字符更宽）
    for line in lines:
        # 计算可见字符宽度（中文算2.3个宽度单位，英文/数字算1.1）
        line_width = 0
        for ch in line:
            line_width += 2.3 if '\u4e00' <= ch <= '\u9fff' else 1.1
        line_lines = max(1, int((line_width + col_width_chars - 1) // col_width_chars))
        total_lines += line_lines
    # 每行约 font_size * 1.5 + 3 磅（Excel 实际行高包括字体+上下间距）
    # 乘以1.15安全系数，确保打印完整
    return max(15.0, total_lines * (font_size * 1.5 + 3) * 1.15)

def _cell_has_measures(cell):
    """单元格文本是否含「管控措施」（关键字段判定）。"""
    v = cell.value
    return bool(v) and ("管控措施" in str(v))


def _estimate_needed_height(cell, col_width_chars):
    """估算单元格所需行高。含管控措施时加上 MEASURES_HEIGHT_PLUS。"""
    if not cell.value or not str(cell.value).strip():
        return 0.0
    h = _estimate_text_height(cell.value, col_width_chars)
    if _cell_has_measures(cell):
        h += MEASURES_HEIGHT_PLUS
    return h


def _apply_autofit_plus_height(ws, start_row, end_row, content_cols, col_widths, plus=0.5, force_plus_one_rows=None, minus_rows=None, skip_rows=None):
    """
    模拟 Excel 双击边框自动调整行高（AutoFit）+ plus 磅余量。
    基于内容估算每行所需高度，取该行所有内容列中的最大值，再加 plus 磅余量。
    skip_rows: 跳过行高设置的行号集合（如插行区域，由插行函数的_apply_height_to_rows统一设置）。
    minus_rows: 需要-1的行号集合（如插行产生的续行）。
    force_plus_one_rows: 强制+1的行号集合（保留兼容）。

    借鉴参考文件 simple_autofit_rows 的逻辑：AutoFit 后 +plus_height，
    但 openpyxl 无法调用 Excel 原生 AutoFit，故用字符数估算替代。
    """
    if force_plus_one_rows is None:
        force_plus_one_rows = set()
    if minus_rows is None:
        minus_rows = set()
    if skip_rows is None:
        skip_rows = set()
    for row in range(start_row, end_row + 1):
        # 插行区域的行：由插行函数的_apply_height_to_rows统一设置，不覆盖
        if row in skip_rows:
            continue
        max_height = 15.0
        for col in content_cols:
            cell = ws.cell(row, col)
            if cell.value and str(cell.value).strip():
                w = col_widths.get(col, 25)
                h = _estimate_text_height(cell.value, w)
                max_height = max(max_height, h)
        # 只设置未超过单行上限的行；超过的由插行续行函数处理
        if max_height <= EXCEL_MAX_ROW_HEIGHT:
            if row in minus_rows:
                actual_plus = -1.0
            elif row in force_plus_one_rows:
                actual_plus = 1.0
            else:
                actual_plus = plus
            ws.row_dimensions[row].height = max_height + actual_plus

def _get_merge_bounds(ws, row, col):
    """获取单元格所在的合并区域边界（等价于参考文件 get_cell_merge_bounds）。
    返回 (is_anchor, min_row, min_col, max_row, max_col)
    is_anchor: 该单元格是否是合并区域的左上角起点
    """
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            is_anchor = (row == rng.min_row and col == rng.min_col)
            return is_anchor, rng.min_row, rng.min_col, rng.max_row, rng.max_col
    return True, row, col, row, col

def _apply_height_to_rows(ws, start_row, end_row, required_height):
    """给行范围均匀分配行高，借鉴参考文件 apply_height_to_merge_area 的逻辑。
    考虑 Excel 单行最大高度 409.5 磅的限制：某行加不下时，把剩余高度分配给后续行。
    """
    row_count = end_row - start_row + 1
    if row_count <= 0:
        return
    current_total = sum(_get_row_height(ws, r, default=15) for r in range(start_row, end_row + 1))
    if required_height <= current_total:
        return
    diff = required_height - current_total
    leftover = 0.0
    for offset in range(row_count):
        r = start_row + offset
        current = _get_row_height(ws, r, default=15)
        add_height = diff / row_count + leftover / (row_count - offset)
        if current + add_height > EXCEL_MAX_ROW_HEIGHT:
            actual_add = EXCEL_MAX_ROW_HEIGHT - current
            leftover += add_height - actual_add
        else:
            actual_add = add_height
            leftover = 0.0
        if actual_add > 0:
            ws.row_dimensions[r].height = current + actual_add
# ============================================================
# ★ 源表【管控措施】列超长整行插行（处理后源表）
# ============================================================
_processed_source_cache = set()  # 单次预处理判定：同一源文件只清洗一次


def insert_overflow_rows_by_measures(ws, start_row=3):
    """
    只针对【管控措施】列超长进行全列贯穿式插行分摊
    ws: 源表 Worksheet
    start_row: 数据起始行（跳过标题和表头，通常为第3行）
    """
    max_col = ws.max_column

    # 1. 动态定位【管控措施】列（仅扫描第2-3行表头，排除第1行大标题）
    #    优先精确匹配"管控措施"，未命中再放宽到"措施/关键风险"，避免误锁"关键风险点"列
    measures_col = None
    for kw in (("管控措施", "防范措施"), ("措施", "关键风险")):
        for c in range(1, max_col + 1):
            for r in (2, 3):
                val = str(ws.cell(r, c).value or "").strip()
                if val and len(val) <= 20 and any(k in val for k in kw):
                    measures_col = c
                    break
            if measures_col:
                break
        if measures_col:
            break

    if not measures_col:
        print("  ⚠️ 未找到【管控措施】列，跳过插行处理")
        return 0
    print(f"  📌 锁定【管控措施】列：第 {measures_col} 列 ({get_column_letter(measures_col)}列)")
    col_width = ws.column_dimensions[get_column_letter(measures_col)].width or 30.0
    # 估算用列宽加上限，避免超宽列（如 255.4）导致行数/高度严重低估
    est_width = max(30.0, min(col_width, 60.0))
    total_inserted = 0
    row = start_row
    while row <= ws.max_row:
        # 若当前行是管控措施列合并区的非首行（从属行），跳过避免重复插行
        is_sub_merged = any(
            rng.min_row < row <= rng.max_row and rng.min_col <= measures_col <= rng.max_col
            for rng in ws.merged_cells.ranges
        )
        if is_sub_merged:
            row += 1
            continue
        cell_val = ws.cell(row, measures_col).value
        text = str(cell_val or "").strip()
        # 估算该单元格文本所需总高度（简易字符算法，无需逐字精确win32）
        needed_height = 15.0
        if text:
            lines = text.split('\n')
            total_lines = 0
            for line in lines:
                line_width = sum(2.2 if '\u4e00' <= ch <= '\u9fff' else 1.0 for ch in line)
                total_lines += max(1, math.ceil(line_width / est_width))
            needed_height = max(15.0, total_lines * 18.0 * 1.15)
        # 超出 Excel 单行上限 409.5 磅，触发整行插行
        if needed_height > EXCEL_MAX_ROW_HEIGHT:
            rows_needed = math.ceil(needed_height / EXCEL_MAX_ROW_HEIGHT)
            rows_to_insert = rows_needed - 1
            insert_at = row + 1
            end_insert_row = insert_at + rows_to_insert - 1

            # 2. 记录当前所有合并区 + 该行独立单格列（插行前）
            all_merges = [(mr.min_row, mr.min_col, mr.max_row, mr.max_col)
                          for mr in list(ws.merged_cells.ranges)]
            single_cols = []
            for c in range(1, max_col + 1):
                hit = any(mr.min_row <= row <= mr.max_row and mr.min_col <= c <= mr.max_col
                          for mr in list(ws.merged_cells.ranges))
                if not hit:
                    single_cols.append(c)

            # 3. 解除全部合并（先解除再插行；master 值保留在单元格中）
            for (min_r, min_c, max_r, max_c) in all_merges:
                try:
                    ws.unmerge_cells(start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c)
                except Exception:
                    pass

            # 4. 执行物理插行
            ws.insert_rows(insert_at, amount=rows_to_insert)

            # 5. 重新合并：原合并区按偏移调整（插行前保留/延伸/下移），独立单格列纵向合并
            new_merges = []
            for (min_r, min_c, max_r, max_c) in all_merges:
                if max_r < row:
                    new_merges.append((min_r, min_c, max_r, max_c))                      # 之前（标题/表头）
                elif min_r <= row <= max_r:
                    if min_r >= start_row:
                        new_merges.append((min_r, min_c, max_r + rows_to_insert, max_c)) # 涉及插行行：块延伸
                    else:
                        new_merges.append((min_r, min_c, max_r, max_c))                  # 表头区：不变
                else:
                    new_merges.append((min_r + rows_to_insert, min_c,
                                       max_r + rows_to_insert, max_c))                    # 之后：整体下移
            for c in single_cols:
                new_merges.append((row, c, end_insert_row, c))                           # 独立单格：与新续行合并
            # 去重（同一单元格范围不重复合并）
            seen = set()
            for m in new_merges:
                if m in seen:
                    continue
                seen.add(m)
                min_r, min_c, max_r, max_c = m
                try:
                    ws.merge_cells(start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c)
                except Exception:
                    pass
                # 6. 复制 master 样式到块内所有行（防止打印断线）
                src_cell = ws.cell(min_r, min_c)
                if src_cell.has_style:
                    for r_new in range(min_r, max_r + 1):
                        dst_cell = ws.cell(r_new, min_c)
                        dst_cell.font = src_cell.font.copy()
                        dst_cell.border = src_cell.border.copy()
                        dst_cell.fill = src_cell.fill.copy()
                        dst_cell.alignment = src_cell.alignment.copy()

            # 7. 均匀分配行高到每一个物理行（每行都不超过 409.5）
            avg_height = round(needed_height / rows_needed, 1)
            for r_sub in range(row, end_insert_row + 1):
                ws.row_dimensions[r_sub].height = avg_height  # 设 height 后 customHeight 自动为 True

            # 8. 管控措施列插行区域强制居左对齐 + 垂直居中 + 自动换行（长文本更易读）
            for r_sub in range(row, end_insert_row + 1):
                cell = ws.cell(r_sub, measures_col)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            print(f"  📏 第{row}行【管控措施】超长({needed_height:.0f}磅)，已插入{rows_to_insert}行续行，所有列全量合并，每行分摊{avg_height}磅")
            total_inserted += rows_to_insert
            row = end_insert_row + 1
        else:
            # 不插行：保持原行高（不覆盖），避免压缩上报表精心设置的行高
            row += 1
    return total_inserted


def export_processed_source(cleaned_path, src):
    """对清洗副本执行【管控措施】列超长插行，导出（第X周）（处理后的源表）.xlsx"""
    try:
        wb2 = load_workbook(cleaned_path)
    except Exception as e:
        safe_print(f"  ⚠️  无法加载清洗副本进行插行: {e}")
        return None
    ws2 = wb2.active
    inserted = insert_overflow_rows_by_measures(ws2, start_row=3)
    m = re.search(r'第\s*(\d+)\s*周', os.path.basename(src))
    if m:
        out_name = f"（第{m.group(1)}周）（处理后的源表）.xlsx"
    else:
        base = os.path.splitext(os.path.basename(src))[0]
        out_name = f"{base}（处理后的源表）.xlsx"
    out_path = os.path.join(os.path.dirname(src) or '.', out_name)
    try:
        wb2.save(out_path)
        wb2.close()
        safe_print(f"  📄 已生成处理后源表（管控措施列插行 {inserted} 行）：{out_name}")
    except Exception as e:
        safe_print(f"  ❌ 保存处理后源表失败: {e}")
        try:
            wb2.close()
        except Exception:
            pass
        return None
    return out_path



def insert_overflow_continuation_rows(ws, start_row, end_row, content_col_start, content_col_end, col_widths):
    """
    超长内容插行续行：当工作内容单元格文本超过 Excel 单行最大高度(409.5磅)时，
    插入续行并自动扩展所有已有合并范围（含A列日期），同时均匀分配行高。

    完整借鉴参考文件：
    - autofit_rows：高度估算 + 插行判断
    - merge_inserted_rows_with_upper：插行后遍历所有列，合并起点自动向下扩展
    - apply_height_to_merge_area：合并区域均匀分配行高，考虑409.5磅上限

    返回：插入的总行数（用于修正后续行号）
    """
    total_inserted = 0
    continuation_rows = set()  # 记录插行产生的续行行号
    inserted_area_rows = set()  # 记录插行区域所有行号（原行+续行），这些行不由_autofit覆盖
    max_col = content_col_end
    row = start_row

    while row <= end_row + total_inserted:
        max_needed_height = 0.0

        # 检查内容列范围内每个单元格所需高度（含管控措施的格自动 +MEASURES_HEIGHT_PLUS）
        row_has_measures = False
        for col in range(content_col_start, content_col_end + 1):
            cell = ws.cell(row, col)
            if cell.value and str(cell.value).strip():
                w = col_widths.get(col, 25)
                h = _estimate_needed_height(cell, w)
                max_needed_height = max(max_needed_height, h)
                if _cell_has_measures(cell):
                    row_has_measures = True

        if max_needed_height > EXCEL_MAX_ROW_HEIGHT:
            rows_needed = int((max_needed_height + EXCEL_MAX_ROW_HEIGHT - 1) // EXCEL_MAX_ROW_HEIGHT)
            rows_to_insert = rows_needed - 1

            if rows_to_insert > 0:
                insert_at = row + 1
                end_insert_row = insert_at + rows_to_insert - 1
                ws.insert_rows(insert_at, amount=rows_to_insert)

                # === 借鉴 merge_inserted_rows_with_upper ===
                # 遍历所有列，检查上方单元格(row)是否是合并起点；
                # 如果是，先取消旧合并，再将合并范围向下扩展到 end_insert_row。
                # 这会自动处理 A 列日期跨多组合并的扩展，以及内容列的续行合并。
                expanded_cols = set()
                for c in range(1, max_col + 1):
                    is_anchor, min_r, min_c, max_r, max_c = _get_merge_bounds(ws, row, c)
                    if not is_anchor:
                        continue
                    if max_r < row:
                        continue
                    # 取消旧合并
                    old_range = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
                    try:
                        ws.unmerge_cells(old_range)
                    except Exception:
                        pass
                    # 创建扩展后的合并
                    new_end_row = max(max_r, end_insert_row)
                    try:
                        ws.merge_cells(
                            start_row=min_r, start_column=min_c,
                            end_row=new_end_row, end_column=max_c
                        )
                        expanded_cols.add(c)
                    except Exception:
                        pass

                # 内容列中仍是单独单元格的（未被上面的合并覆盖），手动合并续行
                for col in range(content_col_start, content_col_end + 1):
                    if col in expanded_cols:
                        continue
                    is_anchor, min_r, min_c, max_r, max_c = _get_merge_bounds(ws, row, col)
                    if min_r == row and max_r == row:
                        try:
                            ws.merge_cells(
                                start_row=row, start_column=col,
                                end_row=end_insert_row, end_column=col
                            )
                        except Exception:
                            pass

                # === 借鉴 apply_height_to_merge_area ===
                # 均匀分配行高，考虑单行409.5磅上限
                _apply_height_to_rows(ws, row, end_insert_row, max_needed_height)

                # 非合并列：复制边框/字体/对齐样式到续行，保持表格网格完整
                for c in range(1, max_col + 1):
                    is_anchor, min_r, min_c, max_r, max_c = _get_merge_bounds(ws, row, c)
                    if min_r <= row <= max_r and min_c <= c <= max_c and (max_r > row or max_c > c):
                        continue  # 在合并区域内，跳过
                    src_cell = ws.cell(row, c)
                    for r in range(row + 1, end_insert_row + 1):
                        dst_cell = ws.cell(r, c)
                        if src_cell.has_style:
                            dst_cell.font = src_cell.font.copy()
                            dst_cell.border = src_cell.border.copy()
                            dst_cell.fill = src_cell.fill.copy()
                            dst_cell.alignment = src_cell.alignment.copy()

                safe_print(f"  📏 插行续行：第{row}行内容超长(约{max_needed_height:.0f}磅)，插入{rows_to_insert}行续行")

                total_inserted += rows_to_insert
                # 记录续行行号（row是原行，row+1到end_insert_row是续行）
                for r in range(row + 1, end_insert_row + 1):
                    continuation_rows.add(r)
                # 记录插行区域所有行号（原行+续行），这些行不由_autofit覆盖
                for r in range(row, end_insert_row + 1):
                    inserted_area_rows.add(r)
                row = end_insert_row + 1
                continue
        else:
            # 不超长的行：直接写入估算高度（管控措施格已含 MEASURES_HEIGHT_PLUS）
            if max_needed_height > 15.0:
                ws.row_dimensions[row].height = max_needed_height
            elif row_has_measures and MEASURES_HEIGHT_PLUS:
                ws.row_dimensions[row].height = 15.0 + MEASURES_HEIGHT_PLUS

        row += 1

    return total_inserted, continuation_rows, inserted_area_rows

def apply_no_split_page_breaks(ws, start_row, end_row):
    """
    不跨行分页：基于行高估算 A3 每页可容纳高度，
    以周计划明细的"组"为单位（每组3行：表头+人员+内容；无计划日为1行），
    确保分页线不落在组内部。
    借鉴参考文件的 apply_page_layout_no_split 逻辑，用 openpyxl 手动分页符实现。

    返回：插入的分页符行号列表
    """
    page_height = A3_LANDSCAPE_PRINTABLE_HEIGHT * PAGE_HEIGHT_SAFETY_FACTOR
    title_height = _get_row_height(ws, 1, default=45)

    # 构建行组：识别每组的起始和结束行
    groups = []
    row = start_row
    while row <= end_row:
        # 无计划日：A列有日期标签，B列空（只有1行）
        a_val = ws.cell(row, 1).value
        b_val = ws.cell(row, 2).value
        if a_val and not b_val:
            h = _get_row_height(ws, row, default=15)
            groups.append((row, row, h))
            row += 1
        else:
            # 有计划的组：3行一组（表头+人员+内容）
            group_end = min(row + 2, end_row)
            h = sum(_get_row_height(ws, r, default=25) for r in range(row, group_end + 1))
            groups.append((row, group_end, h))
            row = group_end + 1

    # 逐组装页，超页时在组前插入分页符
    current_height = title_height
    break_rows = []

    for g_start, g_end, g_height in groups:
        if current_height + g_height > page_height and current_height > title_height + 10:
            break_rows.append(g_start)
            current_height = title_height + g_height
        else:
            current_height += g_height

    # 清除旧分页符，添加新的（openpyxl 中水平分页符属性为 row_breaks，类型 RowBreak）
    ws.row_breaks = RowBreak()
    for br in break_rows:
        ws.row_breaks.append(Break(id=br))

    if break_rows:
        safe_print(f"  📄 不跨行分页：已在 {len(break_rows)} 处插入分页符（行号: {break_rows[:5]}{'...' if len(break_rows) > 5 else ''}）")

    return break_rows

# ======================================
# 统计打印
# ======================================
def print_statistics(daily_data, start_date, end_date):
    print("\n📊 统计概况:")
    total_items = sum(len(plans) for plans in daily_data.values())
    print(f"📅 日期范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")
    print(f"🔢 总计划项数: {total_items}")

    print("\n📆 每日计划分布:")
    max_count = 0
    max_date = None
    total_daily_plans = 0
    for i in range((end_date - start_date).days + 1):
        day = start_date + timedelta(days=i)
        count = len(daily_data.get(day, []))
        total_daily_plans += count
        weekday = "一二三四五六日"[day.weekday()]
        print(f"  {day.strftime('%m-%d')} 星期{weekday}: {count}项计划")
        if count > max_count:
            max_count = count
            max_date = day
    print(f"\n📈 本周计划总数（按天统计）: {total_daily_plans}项")
    if max_count > 0:
        print(f"🏆 计划最多的日期: {max_date.strftime('%m-%d')} ({max_count}项)")

    prof_stats = {}
    for day_plans in daily_data.values():
        for plan in day_plans:
            spec = plan["spec"] or "未分类"
            prof_stats[spec] = prof_stats.get(spec, 0) + 1
    if prof_stats:
        print("\n📋 专业分布:")
        for spec, count in sorted(prof_stats.items(), key=lambda x: x[1], reverse=True):
            print(f"  {spec}: {count}项")

    risk_stats = {}
    for day_plans in daily_data.values():
        for plan in day_plans:
            risk_val = plan["risk_val"] or "未分类"
            risk_stats[risk_val] = risk_stats.get(risk_val, 0) + 1
    if risk_stats:
        print("\n⚠️ 风险等级分布:")
        for risk_val, count in sorted(risk_stats.items(), key=lambda x: x[1], reverse=True):
            print(f"  {risk_val}: {count}项")

# -------------------- V5周计划明细功能 --------------------
def generate_weekly_plan_v5(all_files):
    safe_print("\n📝 正在生成周计划明细（V5增强版）...")
    if not all_files:
        safe_print("⚠️ 没有找到可处理的文件")
        return

    for src in all_files:
        # ---- 源表预处理 ----
        processed_src = preprocess_source_file(src)
        try:
            safe_print(f"\n📁 处理文件: {src}")

            df, title = load_source_df(processed_src)
            if df.empty:
                safe_print("⚠️ 未找到有效数据")
                continue

            cols = resolve_columns(df)
            if cols['time'] is None:
                safe_print("⚠️ 时间列获取失败")
                continue
            if cols.get('measures'):
                safe_print(f"  📌 已识别管控措施列：{cols['measures']}（行高 +{MEASURES_HEIGHT_PLUS:g}磅，改文件顶部 MEASURES_HEIGHT_PLUS）")
            else:
                safe_print("  ⚠️ 未找到「管控措施」列，内容行将按估算高度写入（无加磅）")

            records = build_standard_records(df, cols)
            daily_records = expand_to_daily(records)

            base_year = infer_base_year_from_time_column(df, cols['time'])
            try:
                start_short, end_short, start_date, end_date = extract_date_range_from_title(title, base_year)
            except Exception as e:
                safe_print(f"⚠️ 解析日期范围失败: {e}")
                continue

            # 明细特有过滤：低压计量不显示 + 日期范围内
            daily_data = {}
            for rec in daily_records:
                if rec['is_low_voltage']:
                    continue
                day_dt = datetime.combine(rec['date'], datetime.min.time())
                if not (start_date <= day_dt <= end_date):
                    continue

                job_desc_parts = [
                    rec['full_content'],
                    f"专业：{rec['spec_raw']}",
                    f"作业风险等级：{rec['risk_raw']}",
                    f"作业时间：{rec['start_time']}-{rec['end_time']}"
                ]
                if rec.get('measures_raw'):
                    job_desc_parts.append(f"管控措施：{rec['measures_raw']}")
                if rec['has_power_plan']:
                    job_desc_parts.append("【关联月度停电计划】")
                job_desc = "\n".join(job_desc_parts)

                daily_data.setdefault(day_dt, []).append({
                    "person": rec['person_raw'],
                    "job_desc": job_desc,
                    "spec": rec['spec_raw'],
                    "risk_val": rec['risk_raw'],
                    "has_measures": bool(rec.get('measures_raw')),
                    "l_rank": rec['l_rank'],
                    "p_rank": rec['p_rank'],
                })

            print_statistics(daily_data, start_date, end_date)
            generate_excel_output_v5(daily_data, start_date, end_date, src)
            safe_print(f"✅ 文件处理完成: {src}")

        except Exception as e:
            log_error(src, "周计划明细生成失败(V5)", str(e), traceback.format_exc())
            safe_print(f"❌ 文件处理失败: {src} | {e}", error=True)
        finally:
            cleanup_preprocessed_file(processed_src)

def generate_excel_output_v5(daily_data, start_date, end_date, src):
    """
    生成格式化Excel输出文件（V5版）
    - 每天的计划按每7个一组分组输出
    - A3横向打印、不跨行分页、超长内容插行续行
    """
    PLANS_PER_GROUP = 7
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "周计划明细"

    total_cols = PLANS_PER_GROUP + 1

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    no_border = Border()
    header_font = Font(bold=True, name="微软雅黑")
    header_alignment = Alignment(horizontal="center", vertical="center")
    data_font = Font(name="微软雅黑")
    center_align = Alignment(wrap_text=True, horizontal="center", vertical="center")
    left_align = Alignment(wrap_text=True, horizontal="left", vertical="center")

    # 标题行
    title_text = f"{start_date.month}月{start_date.day}日-{end_date.month}月{end_date.day}日计划明细表"
    ws_out.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    title_cell = ws_out.cell(1, 1, title_text)
    title_cell.font = Font(size=18, bold=True, name="微软雅黑")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_out.row_dimensions[1].height = 45
    for col_idx in range(1, total_cols + 1):
        ws_out.cell(1, col_idx).border = no_border

    current_row = 2
    data_start_row = 2

    num_days = (end_date - start_date).days + 1
    for i in range(num_days):
        day = start_date + timedelta(days=i)
        weekday_idx = (start_date.weekday() + i) % 7
        weekday = "一二三四五六日"[weekday_idx]
        label = f"星期{weekday} {day.month}月{day.day}日"

        day_plans = daily_data.get(day, [])
        day_plans.sort(key=lambda x: (x["l_rank"], x["p_rank"]))

        if not day_plans:
            date_cell = ws_out.cell(current_row, 1, label)
            date_cell.font = data_font
            date_cell.alignment = Alignment(horizontal="center", vertical="center")
            for col_idx in range(1, total_cols + 1):
                ws_out.cell(current_row, col_idx).border = border
            current_row += 1
            continue

        num_groups = (len(day_plans) + PLANS_PER_GROUP - 1) // PLANS_PER_GROUP
        total_day_rows = num_groups * 3

        date_start_row = current_row
        date_end_row = current_row + total_day_rows - 1
        ws_out.merge_cells(start_row=date_start_row, start_column=1,
                          end_row=date_end_row, end_column=1)
        date_cell = ws_out.cell(date_start_row, 1, label)
        date_cell.font = data_font
        date_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for r in range(date_start_row, date_end_row + 1):
            top_side = thin if r == date_start_row else None
            bottom_side = thin if r == date_end_row else None
            ws_out.cell(r, 1).border = Border(
                left=thin, right=thin, top=top_side, bottom=bottom_side
            )

        for group_idx in range(num_groups):
            group_start = group_idx * PLANS_PER_GROUP
            group_end = min(group_start + PLANS_PER_GROUP, len(day_plans))
            group_size = group_end - group_start
            group_plans = day_plans[group_start:group_end]

            # 第1行：表头行
            header_row = current_row
            ws_out.row_dimensions[header_row].height = 25
            for j in range(group_size):
                col_idx = j + 2
                plan_label = f"计划{group_start + j + 1}"
                header_cell = ws_out.cell(header_row, col_idx, plan_label)
                header_cell.font = header_font
                header_cell.alignment = header_alignment
                header_cell.border = border
            for col_idx in range(group_size + 2, total_cols + 1):
                ws_out.cell(header_row, col_idx).border = border
            current_row += 1

            # 第2行：人员行
            person_row = current_row
            for j in range(group_size):
                col_idx = j + 2
                plan = group_plans[j]
                person_cell = ws_out.cell(person_row, col_idx)
                person_cell.value = plan["person"]
                person_cell.font = data_font
                person_cell.alignment = center_align
                person_cell.border = border
            for col_idx in range(group_size + 2, total_cols + 1):
                ws_out.cell(person_row, col_idx).border = border
            current_row += 1

            # 第3行：工作内容行
            job_row = current_row
            for j in range(group_size):
                col_idx = j + 2
                plan = group_plans[j]
                job_cell = ws_out.cell(job_row, col_idx)
                job_cell.value = plan["job_desc"]
                job_cell.font = data_font
                job_cell.alignment = left_align
                job_cell.border = border
            for col_idx in range(group_size + 2, total_cols + 1):
                ws_out.cell(job_row, col_idx).border = border
            current_row += 1

    data_end_row = current_row - 1

    # 列宽
    ws_out.column_dimensions["A"].width = 18
    col_widths = {}
    for col in range(2, total_cols + 1):
        col_letter = get_column_letter(col)
        ws_out.column_dimensions[col_letter].width = 25
        col_widths[col] = 25

    # ---- 插行续行：工作内容列(B-H)超长(>409.5磅)时插行，同时接管所有行高设置 ----
    # 插行代码对超长行插行并用_apply_height_to_rows均匀分配行高；
    # 对不超长的行也基于估算设置行高。插行代码完全接管行高设置。
    safe_print("  📏 正在插行续行并设置行高（插行代码接管所有行高）...")
    inserted, continuation_rows, inserted_area_rows = insert_overflow_continuation_rows(
        ws_out, data_start_row, data_end_row,
        content_col_start=2, content_col_end=total_cols,
        col_widths=col_widths
    )
    data_end_row += inserted

    # ---- 管控措施行高：估算值已含 MEASURES_HEIGHT_PLUS，不再全表统一 -1 ----
    plus_count = 0
    sample_plus = []
    for row in range(data_start_row, data_end_row + 1):
        has_m = False
        for col in range(2, total_cols + 1):
            if _cell_has_measures(ws_out.cell(row, col)):
                has_m = True
                break
        if not has_m:
            continue
        plus_count += 1
        h = ws_out.row_dimensions[row].height
        if len(sample_plus) < 5:
            sample_plus.append(f"第{row}行={h:.1f}磅" if h else f"第{row}行=默认")
    safe_print(
        f"  ✅ 管控措施行高：已对 {plus_count} 行内容写入 估算+{MEASURES_HEIGHT_PLUS:g}磅"
        f"（改文件顶部 MEASURES_HEIGHT_PLUS 即可调成 +2 / +0）"
    )
    if sample_plus:
        safe_print(f"  📊 管控措施样例行高: {', '.join(sample_plus)}")
    sample_rows = []
    for row in range(data_start_row, min(data_start_row + 5, data_end_row + 1)):
        h = ws_out.row_dimensions[row].height
        sample_rows.append(f"第{row}行={h:.1f}磅" if h else f"第{row}行=默认")
    safe_print(f"  📊 前5行行高: {', '.join(sample_rows)}")

    # ---- 不跨行分页：以组为单位插入手动分页符 ----
    safe_print("  📄 正在计算不跨行分页位置...")
    apply_no_split_page_breaks(ws_out, data_start_row, data_end_row)

    # ---- A3 打印设置 ----
    setup_a3_print(ws_out, data_end_row, total_cols)
    safe_print("  🖨️ 已设置 A3 横向打印（窄页边距、1页宽适配、不跨行分页）")

    # 保存
    filename = os.path.splitext(os.path.basename(src))[0]
    out_name = f"{filename}（周计划明细）.xlsx"
    wb_out.save(out_name)
    safe_print(f"💾 已保存: {out_name}")

# -------------------- 三四五级 TXT --------------------
def generate_risk_txt(all_files):
    safe_print("\n📝 正在生成三四五级风险汇总 TXT...")
    if not all_files:
        safe_print("⚠️ 没有找到可处理的文件")
        return

    risk_data = {"三级": [], "四级": [], "五级": []}

    for src in all_files:
        processed_src = preprocess_source_file(src)
        try:
            df, _ = load_source_df(processed_src)
            if df.empty:
                continue
            cols = resolve_columns(df)
            records = build_standard_records(df, cols)
            daily_records = expand_to_daily(records)

            for rec in daily_records:
                risk_level = rec['risk_level']
                if not risk_level:
                    continue
                risk_data[risk_level].append((rec['date'], rec['full_content']))

        except Exception as e:
            log_error(src, "三四五级TXT生成失败", str(e), traceback.format_exc())
        finally:
            cleanup_preprocessed_file(processed_src)

    lines = []
    for level in ["三级", "四级", "五级"]:
        lines.append(f"{level}作业风险管控情况\n")
        jobs = sorted(risk_data[level], key=lambda x: x[0])
        if not jobs:
            lines.append(f"暂无{level}风险作业\n\n")
            continue
        date_groups = OrderedDict()
        for d, job in jobs:
            date_groups.setdefault(d, []).append(job)
        for day_dt, jobs in date_groups.items():
            lines.append(f"{day_dt.month}月{day_dt.day}日")
            for job in jobs:
                for part in format_job_content(job):
                    lines.append(part)
            lines.append("")

    txt_path = "三四五级风险汇总(简化版).txt"
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    safe_print(f"💾 已保存: {os.path.abspath(txt_path)}")

# -------------------- 周计划项目汇总 TXT --------------------
def generate_summary_txt(all_files):
    safe_print("\n📝 正在生成周计划项目汇总...")
    if not all_files:
        safe_print("⚠️ 没有找到可处理的文件")
        return

    total_count = 0
    spec_counts = {}
    risk_counts = {}
    yx_high_count = 0
    yx_low_count = 0

    for src in all_files:
        processed_src = preprocess_source_file(src)
        try:
            df, _ = load_source_df(processed_src)
            if df.empty:
                continue
            cols = resolve_columns(df)
            records = build_standard_records(df, cols)

            for rec in records:
                total_count += 1
                std_spec = rec['std_spec']
                spec_counts[std_spec] = spec_counts.get(std_spec, 0) + 1

                risk = rec['risk_level'] or rec['risk_raw']
                if risk:
                    risk_counts[risk] = risk_counts.get(risk, 0) + 1

                if std_spec == '营销':
                    content = rec['content_raw']
                    raw_spec = rec['spec_raw']
                    high_keywords = ['高压计量', '高压业扩', '业扩', '高压增容', '增容']
                    if any(kw in raw_spec or kw in content for kw in high_keywords):
                        yx_high_count += 1
                    elif '低压计量' in raw_spec or '低压计量' in content:
                        yx_low_count += 1

        except Exception as e:
            log_error(src, "汇总统计生成失败", str(e), traceback.format_exc())
            safe_print(f"⚠️ 文件处理异常: {src}", error=True)
        finally:
            cleanup_preprocessed_file(processed_src)

    std_prof_names = [name for name, _ in PROF_ORDER]
    spec_parts = []
    for prof_name in std_prof_names:
        if prof_name not in spec_counts:
            continue
        if prof_name == '营销':
            detail_parts = []
            if yx_high_count:
                detail_parts.append(f"高压类{yx_high_count}项")
            if yx_low_count:
                detail_parts.append(f"低压计量{yx_low_count}项")
            detail = f"（{'、'.join(detail_parts)}）" if detail_parts else ""
            spec_parts.append(f"营销专业{spec_counts[prof_name]}项{detail}")
        else:
            spec_parts.append(f"{prof_name}专业{spec_counts[prof_name]}项")
    for prof_name, cnt in spec_counts.items():
        if prof_name not in std_prof_names:
            spec_parts.append(f"{prof_name}专业{cnt}项")
    spec_text = "、".join(spec_parts) if spec_parts else "暂无专业分类"

    risk_parts = []
    for risk in ('三级', '四级', '五级'):
        if risk in risk_counts:
            risk_parts.append(f"{risk}作业风险{risk_counts[risk]}项")
    for risk, cnt in risk_counts.items():
        if risk not in ('三级', '四级', '五级'):
            risk_parts.append(f"{risk}{cnt}项")
    risk_text = "、".join(risk_parts) if risk_parts else "无三级、四级、五级风险作业"

    summary_text = f"本周共安排作业计划{total_count}项，包括{spec_text}，其中{risk_text}。"

    txt_path = "周计划项目汇总.txt"
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(summary_text)
    safe_print(f"\n📊 汇总结果:")
    safe_print(summary_text)
    safe_print(f"\n💾 已保存: {os.path.abspath(txt_path)}")

# -------------------- 主程序菜单 --------------------
def cleanup_error_log():
    if _error_count == 0 and os.path.exists(ERROR_LOG_PATH):
        try:
            os.remove(ERROR_LOG_PATH)
        except Exception:
            pass

def main():
    safe_print("📌 周计划工具（V5功能移植版）")
    safe_print("====================")
    safe_print("功能选项:")
    safe_print("1. 周计划明细 Excel 输出（V5增强版）")
    safe_print("2. 三四五级风险汇总 TXT 输出")
    safe_print("3. 周计划项目汇总文本输出")
    safe_print("4. 退出")
    choice = input("请输入数字选择功能（可一次输入多个，用逗号分隔，如1,2,3）: ")
    choices = [c.strip() for c in re.split(r"[,，、\s]+", choice) if c.strip()]
    all_files = find_excel_files()
    if "1" in choices:
        generate_weekly_plan_v5(all_files)
    if "2" in choices:
        generate_risk_txt(all_files)
    if "3" in choices:
        generate_summary_txt(all_files)
    cleanup_error_log()
    safe_print("\n✅ 所有任务完成")
    try:
        input("\n按回车键退出...")
    except Exception:
        pass

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        safe_print(f"\n❌ 程序异常退出: {e}", error=True)
        import traceback
        traceback.print_exc()
        try:
            input("\n按回车键退出...")
        except Exception:
            pass
