# -*- coding: utf-8 -*-
"""
最终处理版4号模式A汇总增强5.2-V3.5全语料语义校准成品版

本版调整：
V3.5重点：继承V3.4规则；增加可持续扩充的审核语料库、混合检索和来源证明硬门槛。
1. 已审核的相同原文可复用确认小结；相似案例只提供结构，不复制地点、线路、设备或动作。
2. 输出中的工程性质和关键动作必须由当前原文直接支持，或由已审核动作组合推导。
3. 无来源的“联络工程、台架安装、迁改”等候选拒绝采用并进入待审核库。
4. 绝缘化缺陷治理继续沿用V3.4严格证据链，不受逐字来源检查影响。
5. 新增原始填报质量检查：术语错写可规范化，站线漏字仅在词库有证据时补写，设备身份冲突只告警不代选。
6. 控制台和运行报告增加“原始填报错误提醒报告”，并另存独立TXT便于通报。
1. 仅在模式A缩写版和模式B领导总览中删除无业务价值的济宁、泗水等行政前缀。
2. 正式工程名称先建立保护区，禁止行政地域清理误伤工程名称。
3. 0.4kV相同动作多设备合并，台架台区压缩为台架，普通台区及配电室设备名完整保留。
4. 同设备重复记录去重；确有多处设备或多条支线身份时才概括为“0.4kV多条支线”。
5. 10kV和带电作业仅在动作单一、完全一致且无关键拓扑关系时合并。
6. 领导摘要取消字符硬截断；识别不完整时回退到语义完整文本。
7. 主线路、支线、台区和变电站按原文身份分别计数，类别混合时禁止统一概括。
V2.8重点：恢复到岗到位生成弹窗/文件链路；领导摘要“配合工作”必须保留核心动作。
1. 保留运行报告/调试日志：每次运行生成“处理报告_时间.txt”。
2. 保留 A3 横向 + 窄页边距打印设置。
3. 保留合并单元格/数据区域边框重刷。
4. 分页方式从“插入可见空白行”改为“添加手动水平分页符”，避免表格中间出现多条空白线。
5. 文本提取增加 nan/空白行过滤，避免分页空白行进入“汇总.txt”。
6. 三类正式表在原流程前按列执行严格多数派排版统一，备份目录完全跳过。
7. 到岗到位人员列按最长单行 AutoFit 后增加横向留白，改善 O/P 列提前换行。
8. 风险管控增加行高上下余量，最终外框和表头分隔线加强、内部网格保持细线。
9. 全程保留原表所有列的字体颜色，不再把非“工作内容”列红字转黑；删除线仍在预处理阶段清除。
10. 到岗到位 N/O 列自动识别姓名与手机号并统一为“姓名在上、手机号在下”；P列保持原结构并允许自动换行。
11. 小结中同一源项目只进入一个汇总分支，避免高压计量明细与普通营销动作重复输出。
12. 到岗到位最终框线改为按序号记录块收口，续行内部不画横线，并检查分页线是否落在记录边界。
13. 到岗到位按真实适配宽度缩放比例计算分页，不再把 11%-12% 强制降为 10%。
14. 每个打印页末行在最终合并和框线完成后单独落实中等粗底线，避免合并区域底框漏印。
15. 现场计划按本机次日校验日期，允许周末/周一提前计划；异常先报告并询问，确认后才修改标题。
16. 三类正式表强制校验二级、三级风险：未标黄时自动补黄；四级、五级不标黄。
17. 三类表格统一使用细框线，不再加粗外框、表头下沿、记录边界或分页末线。
18. 原计划表重排序号后A列底色跟随同一行专业列；风险管控和到岗到位打印表正文取消营销绿色，风险黄色保留。

重要说明：
- 本脚本仍然会直接修改当前目录下的 xlsx/xlsm 文件。
- 新增了自动备份功能，原文件会先复制到“自动备份_时间”文件夹。
- 分页功能目前属于“测试版”，请先用样本文件验证打印预览效果。
"""

import os
import re
import sys
import warnings
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
import shutil
import math
import traceback
import time
from collections import Counter

try:
    from 模式B语义检索库 import (
        collect_pending_case as v35_collect_pending_case,
        exact_approved_summary as v35_exact_approved_summary,
        provenance_violations as v35_provenance_violations,
        search_cases as v35_search_reviewed_cases,
    )
    V35_REVIEW_LIBRARY_AVAILABLE = True
except Exception:
    V35_REVIEW_LIBRARY_AVAILABLE = False
    v35_collect_pending_case = None
    v35_exact_approved_summary = None
    v35_provenance_violations = None
    v35_search_reviewed_cases = None

# ================= 用户可调整配置区 =================
ENABLE_BACKUP = True
SKIP_BACKUP_WHEN_BACKUP_DIR_EXISTS = True
ENABLE_PAGE_SETUP = True
ENABLE_AUTO_PAGE_LAYOUT = True
ENABLE_INSERT_SPACER_ROWS = False   # 第4版默认不插入可见空白行，只加手动分页符
ENABLE_MANUAL_PAGE_BREAKS = True
ENABLE_REPEAT_TITLE_ROWS = True   # 每页重复第1-2行：大标题 + 关键字段表头
REPEAT_TITLE_ROWS_TEXT = "$1:$2"
ENABLE_AUTOFITROWEX_HEIGHT = True
ENABLE_AUTOFITROWEX_OVERFLOW_ROWS = True  # 单元格文本超过 Excel 单行高度上限时，仿 AutoFitRowEx 插入续行并向上合并
ENABLE_DAOGANG_POST_INSERT_COMPACT_AUTOFIT = False  # 到岗到位插入续行后默认不再全量二次测量，避免重复耗时
ENABLE_DAOGANG_VERSION_PROMPT = True
ENABLE_MAJORITY_FORMAT_NORMALIZATION = True
PRINT_MAX_COLUMN_S = 19

# A3 横向 + 窄页边距
PAGE_PAPER_SIZE_A3 = 8       # Excel xlPaperA3
PAGE_ORIENTATION_LANDSCAPE = 2  # Excel xlLandscape

# Excel 窄边距常用值；上下和页眉页脚适当收紧，减少到岗到位版分页膨胀。
MARGIN_LEFT_INCH = 0.25
MARGIN_RIGHT_INCH = 0.25
MARGIN_TOP_INCH = 0.15
MARGIN_BOTTOM_INCH = 0.15
MARGIN_HEADER_INCH = 0.05
MARGIN_FOOTER_INCH = 0.05

# 分页估算安全系数。越小越保守，插入空白行可能越多。
PAGE_HEIGHT_SAFETY_FACTOR = 1.00
DAOGANG_PAGE_HEIGHT_SAFETY_FACTOR = 0.96  # 给 Excel 导出时的隐藏分页误差预留约一行

# 自动插入空白行设置
AUTO_SPACER_MARKER = "__AUTO_PAGE_SPACER_ROW__"
AUTO_SPACER_HELPER_COL = 60   # 隐藏辅助列，用来标记自动插入的空白行，避免重复插入
DEFAULT_SPACER_ROW_HEIGHT = 18
MIN_SPACER_ROW_HEIGHT = 8
MAX_SPACER_ROW_HEIGHT = 45
DEFAULT_INSERTED_CONTINUATION_ROW_HEIGHT = 15
DEFAULT_AUTOFIT_PLUS_HEIGHT = 2
SITE_AUTOFIT_PLUS_HEIGHT = 2.0
RISK_AUTOFIT_PLUS_HEIGHT = 2.0
RISK_COMPACT_AUTOFIT_PLUS_HEIGHT = 2.0
MIN_AUTOFIT_PLUS_HEIGHT = 0.5
MAX_COMPACT_AUTOFIT_PLUS_HEIGHT = 1.5

# 表格设置
DATA_START_ROW = 3
BORDER_START_ROW = 2
AUTOFITROWEX_START_ROW = 3
EXCEL_MAX_ROW_HEIGHT = 409.5
AUTOFITROWEX_TEMP_SHEET = "Temp_ForAdjustRowHeightAddIn"
AUTOFITROWEX_PRINT_MODE = True
AUTOFITROWEX_DEFAULT_EXTRA_HEIGHT = 1.04
AUTOFITROWEX_MS_WIDTH_RATIO = 0.88
# 到岗到位版的超高续行只重点测这些长文本列：
# C 工作内容、D 关键风险点及管控措施。T 督查计划不进打印区，不再撑高整行。
DAOGANG_AUTOFIT_MEASURE_COLUMNS = {3, 4, 14, 15, 16}
DAOGANG_FONT_SIZE = 45
DAOGANG_HEADER_ROW_HEIGHT = 204
DAOGANG_SIDEBAR_MIN_COLUMN_WIDTH = 90.0
DAOGANG_SIDEBAR_MAX_COLUMN_WIDTH = 104.0
DAOGANG_SIDEBAR_WIDTH_PER_CHAR = 7.6
DAOGANG_SIDEBAR_WIDTH_PADDING = 12.0
PERSONNEL_NORMAL_MIN_COLUMN_WIDTH = 24.0
PERSONNEL_NORMAL_MAX_COLUMN_WIDTH = 120.0
PERSONNEL_DAOGANG_MIN_COLUMN_WIDTH = 44.0
PERSONNEL_DAOGANG_MAX_COLUMN_WIDTH = 255.0  # Excel列宽上限；人员列不再因120上限被迫折行
PERSONNEL_COLUMN_WIDTH_PADDING_RATIO = 1.12
PERSONNEL_COLUMN_WIDTH_PADDING_MIN = 3.0
DAOGANG_SEQUENCE_MIN_COLUMN_WIDTH = 18.0
DAOGANG_SEQUENCE_MAX_COLUMN_WIDTH = 30.0
DAOGANG_SEQUENCE_WIDTH_PADDING_RATIO = 1.18
MIN_EXCEL_PRINT_SCALE_PERCENT = 10.0
MAX_EXCEL_PRINT_SCALE_PERCENT = 100.0
MIN_PRINT_ROW_COMPACT_FACTOR = 0.96
# 到岗到位专用：续行判断允许吸收 AutoFitRowEx 的约4%打印测量安全余量。
# 仅影响“是否需要多插一行”的临界判断；风险管控仍保持原来的严格 +1 行策略。
DAOGANG_OVERFLOW_ROW_CAPACITY_FACTOR = 1.04
AUTOFITROWEX_FONT_SETTINGS = {
    "Calibri": (1.05, 0.88),
    "Segoe UI": (1.05, 0.88),
}

# Excel 边框常量
XL_CONTINUOUS = 1
XL_THIN = 2
XL_MEDIUM = -4138
XL_EDGE_LEFT = 7
XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_EDGE_RIGHT = 10
XL_INSIDE_VERTICAL = 11
XL_INSIDE_HORIZONTAL = 12
XL_NONE = -4142
XL_SOLID = 1
XL_CENTER = -4108
XL_CENTER_VERTICAL = -4108
XL_COLOR_WHITE = 16777215
XL_COLOR_BLACK = 0
XL_COLOR_YELLOW = 65535

# ================= 运行报告/日志 =================
REPORT_LINES = []
REPORT_PATH = None
CURRENT_BACKUP_DIR = None
BACKUP_SKIP_LOGGED = False
SOURCE_ENTRY_ISSUES = []
SOURCE_ENTRY_ISSUE_KEYS = set()


def log(msg=""):
    """同时输出到控制台和运行报告。"""
    text = str(msg)
    print(text, flush=True)
    REPORT_LINES.append(text)


def log_red(msg=""):
    """控制台红字报警，同时写入运行报告。"""
    text = str(msg)
    try:
        print(f"\033[91m{text}\033[0m", flush=True)
    except Exception:
        print(text, flush=True)
    REPORT_LINES.append(text)


def log_section(title):
    line = "=" * 18 + f" {title} " + "=" * 18
    log("\n" + line)


def log_sub(title):
    log(f"\n--- {title} ---")


def register_source_entry_issue(
    issue_code,
    source_text,
    message,
    suggestion,
    auto_corrected=False,
    source_ref="",
    corrected_text="",
):
    """登记原表填报问题；同一原文同一问题只报告一次。"""
    normalized_source = re.sub(r"\s+", "", str(source_text or ""))
    key = (str(issue_code), normalized_source)
    if key in SOURCE_ENTRY_ISSUE_KEYS:
        return
    SOURCE_ENTRY_ISSUE_KEYS.add(key)
    SOURCE_ENTRY_ISSUES.append({
        "issue_code": str(issue_code),
        "source_ref": str(source_ref or "未提供行号"),
        "message": str(message),
        "suggestion": str(suggestion),
        "auto_corrected": bool(auto_corrected),
        "source_text": str(source_text or "").strip(),
        "corrected_text": str(corrected_text or "").strip(),
    })


def audit_source_entry_text(raw_text, source_ref=""):
    """检查原始填报错误；返回供摘要使用的纠正文本，不改动Excel原单元格。"""
    source = str(raw_text or "")
    corrected = source

    if "提接" in corrected:
        corrected = corrected.replace("提接", "T接")
        register_source_entry_issue(
            "TERM_T_CONNECTION",
            source,
            "发现将电力线路“T接”错写为“提接”。",
            "摘要已统一改为“T接”；请填报人员同步修正原表。",
            auto_corrected=True,
            source_ref=source_ref,
            corrected_text=corrected,
        )

    # 当前站线词库和既有语料均能证明主线路名称为“10kV马头山线”。
    missing_line_pattern = r"10kV马头山(?!线)大套村支"
    if re.search(missing_line_pattern, corrected):
        corrected = re.sub(missing_line_pattern, "10kV马头山线大套村支", corrected)
        register_source_entry_issue(
            "MISSING_MAIN_LINE_SUFFIX",
            source,
            "发现“10kV马头山大套村支”漏写主线路名称中的“线”字。",
            "依据站线词库补为“10kV马头山线大套村支”；请填报人员核对并修正原表。",
            auto_corrected=True,
            source_ref=source_ref,
            corrected_text=corrected,
        )

    if "南顶西村3号台架变" in source and "南百顶西村3号台区" in source:
        register_source_entry_issue(
            "DEVICE_IDENTITY_CONFLICT",
            source,
            "工程标题写“南顶西村3号台架变”，正文写“南百顶西村3号台区”，设备身份不一致。",
            "程序不自动选择设备名称；请填报人员确认正确名称后修改原表。",
            auto_corrected=False,
            source_ref=source_ref,
        )

    return corrected


def prepare_summary_source_entries(records):
    """为缩写版和模式B准备纠错文本；完整版继续保留原始填报。"""
    for record in records:
        order = record.get("order") or "?"
        source_row = record.get("source_row")
        source_ref = f"计划第{order}项"
        if source_row:
            source_ref += f"（Excel第{source_row}行）"
        record["_summary_work"] = audit_source_entry_text(
            record.get("work"), source_ref=source_ref
        )


def summary_source_text(record):
    return record.get("_summary_work") or record.get("work") or ""


def write_source_entry_issue_report(cwd):
    """输出控制台汇总，并生成便于通报的独立错误提醒报告。"""
    report_path = os.path.join(cwd, "原始填报错误提醒报告.txt")
    log_section("原始填报错误提醒报告")
    if not SOURCE_ENTRY_ISSUES:
        log("未发现已定义的原始填报错误。")
        if os.path.isfile(report_path):
            try:
                os.remove(report_path)
            except Exception:
                pass
        return None

    corrected_count = sum(item["auto_corrected"] for item in SOURCE_ENTRY_ISSUES)
    pending_count = len(SOURCE_ENTRY_ISSUES) - corrected_count
    lines = [
        "原始填报错误提醒报告",
        f"共发现{len(SOURCE_ENTRY_ISSUES)}项：摘要已自动纠正{corrected_count}项，需人工核实{pending_count}项。",
        "说明：程序仅纠正缩写版和模式B的摘要输入，不改动Excel原单元格；完整版保留原始填报用于追溯。",
        "",
    ]
    log_red(
        f"共发现{len(SOURCE_ENTRY_ISSUES)}项原始填报问题："
        f"摘要已自动纠正{corrected_count}项，需人工核实{pending_count}项。"
    )
    for index, item in enumerate(SOURCE_ENTRY_ISSUES, 1):
        status = "摘要已自动纠正" if item["auto_corrected"] else "未自动纠正，需人工核实"
        console_text = (
            f"{index}. 【{item['source_ref']}】【{status}】{item['message']} "
            f"处理建议：{item['suggestion']}"
        )
        log_red(console_text)
        compact_source = re.sub(r"\s+", " ", item["source_text"]).strip()
        lines.extend([
            f"{index}. {item['source_ref']}｜{status}",
            f"问题：{item['message']}",
            f"建议：{item['suggestion']}",
            f"原文：{compact_source}",
        ])
        if item["corrected_text"]:
            compact_corrected = re.sub(r"\s+", " ", item["corrected_text"]).strip()
            lines.append(f"摘要纠正后：{compact_corrected}")
        lines.append("")
    with open(report_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines).rstrip() + "\n")
    log(f"原始填报错误提醒报告已生成：{report_path}")
    return report_path


def step_timer(label):
    log(f"  >>> 开始：{label}")
    return time.perf_counter()


def step_done(label, start):
    log(f"  <<< 完成：{label}，耗时 {time.perf_counter() - start:.1f} 秒")


def configure_excel_silent(excel, wb=None):
    """关闭自动化处理期间的保存、兼容性和链接提示。"""
    for attr, value in (
        ("Visible", False),
        ("ScreenUpdating", False),
        ("DisplayAlerts", False),
        ("AskToUpdateLinks", False),
        ("EnableEvents", False),
        ("AlertBeforeOverwriting", False),
        ("Calculation", -4135),  # xlCalculationManual
        ("CalculateBeforeSave", False),
        ("AutomationSecurity", 3),  # msoAutomationSecurityForceDisable
    ):
        try:
            setattr(excel, attr, value)
        except Exception:
            pass
    if wb is not None:
        try:
            wb.CheckCompatibility = False
        except Exception:
            pass


def create_excel_application():
    """为每个文件创建独立的 Excel 自动化实例，避免复用卡住的旧实例。"""
    try:
        return win32.DispatchEx("Excel.Application")
    except Exception:
        return win32.gencache.EnsureDispatch("Excel.Application")


def save_report(cwd):
    global REPORT_PATH
    try:
        if REPORT_PATH is None:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            REPORT_PATH = os.path.join(cwd, f"处理报告_{ts}.txt")
        with open(REPORT_PATH, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(REPORT_LINES))
        print(f"\n运行报告已保存：{REPORT_PATH}")
    except Exception as e:
        print(f"保存运行报告失败：{e}")


# ================= 清除 win32com 缓存 =================
def clear_win32com_cache():
    """清除win32com缓存目录以解决COM对象缓存问题"""
    try:
        gen_path = win32.gencache.GetGeneratePath()
        log(f"尝试清除win32com缓存: {gen_path}")
        if os.path.exists(gen_path):
            shutil.rmtree(gen_path)
            log("已成功清除win32com缓存目录")
            return True
        else:
            log("win32com缓存目录不存在，无需清除")
            return False
    except Exception as e:
        log(f"清除缓存失败: {e}")
        return False


# ================= 检测文件是否被占用 =================
def is_file_locked(filepath):
    """检查文件是否被其他进程占用"""
    if not os.path.exists(filepath):
        return True
    try:
        with open(filepath, 'r+b') as f:
            f.read(1)
        return False
    except (PermissionError, OSError):
        return True


# ================= 强制关闭所有 Excel 进程 =================
def kill_excel_processes():
    try:
        import psutil
        killed = 0
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
                try:
                    proc.kill()
                    killed += 1
                except Exception:
                    pass
        if killed:
            log(f"已强制关闭 {killed} 个 Excel 进程")
    except ImportError:
        log("错误：缺少 psutil 模块！请运行：pip install psutil")
        exit(1)
    except Exception as e:
        log(f"关闭 Excel 进程时出错: {e}")


# ================= 工具函数 =================
def remove_phone_and_noise(text):
    if not text:
        return ""
    s = str(text).strip()
    s = re.sub(r'\d{7,}', '', s)
    s = re.sub(r'[：:]\s*$', '', s)
    return s.strip()


def parse_datetime_from_cell(val):
    if val is None or str(val).strip() in ["", "nan"]:
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    m = re.search(r'(\d{1,2})[:：](\d{1,2})', s)
    if m:
        try:
            return datetime(2000, 1, 1, int(m.group(1)), int(m.group(2)))
        except Exception:
            return None
    return None


def format_HHMM(dt):
    if dt is None:
        return ""
    return dt.strftime("%H:%M")


def safe_float(value, default=0.0):
    try:
        if value is None:
            return default
        return float(value)
    except Exception:
        return default


def is_blank_value(value):
    """统一判断 Excel / pandas 里的空值，避免 nan 行进入汇总。"""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    s = str(value).strip()
    return s == "" or s.lower() in {"nan", "nat", "none"}


def has_any_value(values):
    return any(not is_blank_value(v) for v in values)


def get_used_rows(ws):
    try:
        return int(ws.UsedRange.Rows.Count)
    except Exception:
        return 1


def is_auto_spacer_row(ws, row):
    try:
        return str(ws.Cells(row, AUTO_SPACER_HELPER_COL).Value).strip() == AUTO_SPACER_MARKER
    except Exception:
        return False


def get_row_height(ws, row, default=15.0):
    try:
        return safe_float(ws.Rows(row).RowHeight, default)
    except Exception:
        return default


def sum_row_heights(ws, start_row, end_row):
    total = 0.0
    for r in range(start_row, end_row + 1):
        total += get_row_height(ws, r, 15.0)
    return total


def is_backup_path(path):
    """自动备份目录中的文件只读保留，不参与任何预处理。"""
    parts = re.split(r"[\\/]+", os.path.abspath(path))
    return any(part.startswith("自动备份_") for part in parts)


def column_number_to_name(column):
    result = ""
    value = int(column)
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def is_merge_anchor(cell):
    try:
        if not cell.MergeCells:
            return True
        area = cell.MergeArea
        return int(cell.Row) == int(area.Row) and int(cell.Column) == int(area.Column)
    except Exception:
        return True


def get_work_content_column(mode, ws=None, max_col=None):
    """优先按真实表头识别工作内容列，模式列号只作为模板异常时的兜底。"""
    if ws is not None:
        try:
            column_limit = int(max_col or ws.UsedRange.Columns.Count)
            for row in range(1, 4):
                for col in range(1, column_limit + 1):
                    header = re.sub(r"\s+", "", str(ws.Cells(row, col).Value or ""))
                    if header == "工作内容":
                        return col
        except Exception:
            pass

    fallback_columns = {
        "risk": 6,
        "site": 4,
        "daogang": 3,
    }
    return fallback_columns.get(mode)


def find_header_column(ws, max_col, target_header):
    """在前3行按去空白后的表头文字动态查找列号。"""
    normalized_target = re.sub(r"\s+", "", str(target_header or ""))
    partial_matches = []
    for row in range(1, 4):
        for col in range(1, max_col + 1):
            try:
                header = re.sub(r"\s+", "", str(ws.Cells(row, col).Value or ""))
            except Exception:
                continue
            if header == normalized_target:
                return col
            if normalized_target and normalized_target in header:
                partial_matches.append(col)
    return partial_matches[0] if partial_matches else None


def is_yellow_interior(target):
    """识别标准黄及常见浅黄色填充。"""
    try:
        if int(target.Interior.ColorIndex) == 6:
            return True
    except Exception:
        pass
    try:
        color = int(target.Interior.Color)
    except Exception:
        return False
    if color < 0:
        return False
    color &= 0xFFFFFF
    red = color & 0xFF
    green = (color >> 8) & 0xFF
    blue = (color >> 16) & 0xFF
    return red >= 220 and green >= 200 and blue <= 170


def get_risk_level_number(text):
    """从“三级/四级/5级风险”等文本中提取风险等级数字。"""
    normalized = re.sub(r"\s+", "", str(text or ""))
    match = re.search(r"([1-9])级", normalized)
    if match:
        return int(match.group(1))
    chinese_levels = {
        "一级": 1, "二级": 2, "三级": 3, "四级": 4, "五级": 5,
        "六级": 6, "七级": 7, "八级": 8, "九级": 9,
    }
    for label, number in chinese_levels.items():
        if label in normalized:
            return number
    return None


def is_level_two_or_three_risk(text):
    """只有二级、三级风险需要黄色标识；四级、五级不标黄。"""
    level = get_risk_level_number(text)
    return level in {2, 3}


def enforce_level_three_risk_yellow(ws, used_rows, max_col, mode=None, filename=""):
    """
    作业风险等级为二级或三级时，单元格必须标黄；四级、五级不标黄。

    表头动态识别，兼容现场表、风险管控表以及删列/插列后的到岗到位版。
    该步骤必须放在最终白底和版式处理之后执行。
    """
    risk_col = find_header_column(ws, max_col, "作业风险等级")
    if not risk_col:
        log_red(
            f"  【二级/三级风险校验告警】{os.path.basename(filename) or '当前工作簿'}"
            "未识别到“作业风险等级”列，无法执行标黄检查"
        )
        return {
            "applied": False,
            "risk_column": None,
            "level_three_cells": 0,
            "fixed_cells": 0,
            "already_yellow": 0,
        }

    level_three_cells = 0
    fixed_cells = 0
    already_yellow = 0
    cleared_wrong_yellow = 0
    failed_cells = 0
    site_fill_source_col = None
    if mode == "site":
        site_fill_source_col = find_header_column(ws, max_col, "专业") or 2

    for row in range(DATA_START_ROW, used_rows + 1):
        if is_auto_spacer_row(ws, row):
            continue
        try:
            cell = ws.Cells(row, risk_col)
            if not is_merge_anchor(cell):
                continue
            risk_text = re.sub(r"\s+", "", str(cell.Value or ""))
            risk_level = get_risk_level_number(risk_text)
            target = cell.MergeArea if cell.MergeCells else cell
            address = str(target.Address)

            # 四级、五级明确不标黄；若旧版本曾误标黄，则主动纠正。
            if risk_level in {4, 5}:
                if is_yellow_interior(target):
                    if mode == "site":
                        # 原计划表恢复同一行专业底色，例如营销专业绿色。
                        copy_cell_fill(ws.Cells(row, site_fill_source_col), target)
                    else:
                        # 两张打印表正文应无营销绿色，四/五级恢复白底。
                        target.Interior.Pattern = XL_SOLID
                        target.Interior.Color = XL_COLOR_WHITE
                    cleared_wrong_yellow += 1
                    log(
                        f"  【四/五级风险去黄】{os.path.basename(filename) or '当前工作簿'} "
                        f"{ws.Name}!{address} 的作业风险等级为“{risk_text}”，已取消错误黄色"
                    )
                continue

            if risk_level not in {2, 3}:
                continue
            level_three_cells += 1
            if is_yellow_interior(target):
                already_yellow += 1
                continue
            target.Interior.Pattern = XL_SOLID
            target.Interior.Color = XL_COLOR_YELLOW
            fixed_cells += 1
            log_red(
                f"  【二级/三级风险标黄告警】{os.path.basename(filename) or '当前工作簿'} "
                f"{ws.Name}!{address} 的作业风险等级为“{risk_text}”，"
                "原单元格未标黄，程序已自动补为黄色"
            )
        except Exception as e:
            failed_cells += 1
            log_red(
                f"  【二级/三级风险标黄失败】{os.path.basename(filename) or '当前工作簿'} "
                f"第{row}行处理失败：{e}"
            )

    log(
        f"  二级/三级风险标黄校验：{column_number_to_name(risk_col)}列，"
        f"二级/三级风险 {level_three_cells} 个；原已标黄 {already_yellow} 个，"
        f"自动补黄 {fixed_cells} 个；四/五级错误去黄 {cleared_wrong_yellow} 个，失败 {failed_cells} 个"
    )
    return {
        "applied": failed_cells == 0,
        "risk_column": risk_col,
        "level_three_cells": level_three_cells,
        "fixed_cells": fixed_cells,
        "already_yellow": already_yellow,
        "cleared_wrong_yellow": cleared_wrong_yellow,
        "failed_cells": failed_cells,
    }


def is_red_font_color(color):
    """识别 Excel RGB/COLORREF 中的红色和常见深红、浅红字体。"""
    try:
        value = int(color)
    except Exception:
        return False
    if value < 0:
        return False
    value &= 0xFFFFFF
    red = value & 0xFF
    green = (value >> 8) & 0xFF
    blue = (value >> 16) & 0xFF
    return (
        red >= 128
        and green <= min(120, int(red * 0.55))
        and blue <= min(120, int(red * 0.55))
    )


def clear_red_font_outside_work_content(cell):
    """发现红字后仅把当前单元格字体颜色统一为黑色。"""
    try:
        color = cell.Font.Color
    except Exception:
        color = None

    if color is not None:
        if is_red_font_color(color):
            cell.Font.Color = XL_COLOR_BLACK
            return 1, 0
        return 0, 0

    text_length = len(str(cell.Value or ""))
    red_chars = 0
    for position in range(1, text_length + 1):
        try:
            char_font = cell.GetCharacters(position, 1).Font
            if is_red_font_color(char_font.Color):
                red_chars += 1
        except Exception:
            continue
    if red_chars:
        # 对 WPS 来源工作簿逐字符写字体可能生成 Excel 无法重开的富文本；
        # 整格只改 Color 属性仍会保留字号、字体、加粗等其他设置。
        cell.Font.Color = XL_COLOR_BLACK
        return 1, red_chars
    return 0, 0


def clear_cell_strikethrough(cell):
    """清除整格或字符级删除线，不改变字体颜色等其他属性。"""
    try:
        strike_value = cell.Font.Strikethrough
    except Exception:
        strike_value = False

    if strike_value not in (False, 0, None):
        try:
            cell.Font.Strikethrough = False
            return True
        except Exception:
            pass

    if strike_value is not None:
        return False

    text_length = len(str(cell.Value or ""))
    has_strikethrough = False
    for position in range(1, text_length + 1):
        try:
            char_font = cell.GetCharacters(position, 1).Font
            if char_font.Strikethrough not in (False, 0, None):
                has_strikethrough = True
                break
        except Exception:
            continue
    if has_strikethrough:
        cell.Font.Strikethrough = False
        return True
    return False


def normalize_font_effects_for_print(ws, used_rows, max_col, mode=None):
    """
    打印前字体效果约束：
    - 原始字体颜色全部保留，不再把任何红字或其他颜色改黑；
    - 标题、表头和数据区的删除线仍清除。
    """
    strike_cells_changed = 0

    for row in range(1, used_rows + 1):
        if is_auto_spacer_row(ws, row):
            continue
        for col in range(1, max_col + 1):
            try:
                cell = ws.Cells(row, col)
                if is_blank_value(cell.Value) or not is_merge_anchor(cell):
                    continue
                if clear_cell_strikethrough(cell):
                    strike_cells_changed += 1
            except Exception:
                continue

    log(
        "  字体效果预处理：原始字体颜色全部保留，不执行红字转黑；"
        f"清除删除线 {strike_cells_changed} 个单元格"
    )
    return {
        "red_cells_changed": 0,
        "red_chars_changed": 0,
        "strike_cells_changed": strike_cells_changed,
    }


def get_majority_format_value(cell, property_name):
    try:
        if property_name == "horizontal":
            return cell.HorizontalAlignment
        if property_name == "vertical":
            return cell.VerticalAlignment
        if property_name == "font_name":
            value = get_font_name(cell)
            return value or None
        if property_name == "font_size":
            value = safe_float(cell.Font.Size, 0.0)
            return round(value, 1) if value > 0 else None
    except Exception:
        return None
    return None


def set_majority_format_value(cell, property_name, value):
    if property_name == "horizontal":
        cell.HorizontalAlignment = value
    elif property_name == "vertical":
        cell.VerticalAlignment = value
    elif property_name == "font_name":
        cell.Font.Name = value
        try:
            cell.Font.NameFarEast = value
        except Exception:
            pass
    elif property_name == "font_size":
        cell.Font.Size = value


def normalize_data_region_majority_format(ws, used_rows, max_col, mode=None):
    """
    按列用严格多数派统一数据区排版。

    只处理对齐、字体名称和字号；字体颜色、加粗、斜体、填充、数字格式、
    边框和业务内容均保留，避免覆盖原表红字及其他人工强调。
    """
    if not ENABLE_MAJORITY_FORMAT_NORMALIZATION or used_rows < DATA_START_ROW:
        return 0

    property_labels = {
        "horizontal": "水平对齐",
        "vertical": "垂直对齐",
        "font_name": "字体",
        "font_size": "字号",
    }
    total_changed_cells = set()
    normalized_columns = []

    for col in range(1, max_col + 1):
        cells = []
        for row in range(DATA_START_ROW, used_rows + 1):
            if is_auto_spacer_row(ws, row):
                continue
            try:
                cell = ws.Cells(row, col)
                if is_blank_value(cell.Value) or not is_merge_anchor(cell):
                    continue
                cells.append((row, cell))
            except Exception:
                continue

        if len(cells) < 2:
            continue

        changed_by_property = []
        for property_name in property_labels:
            observed = []
            for row, cell in cells:
                value = get_majority_format_value(cell, property_name)
                if value is not None:
                    observed.append((row, cell, value))
            if len(observed) < 2:
                continue

            counts = Counter(value for _, _, value in observed)
            dominant_value, dominant_count = counts.most_common(1)[0]
            if dominant_count * 2 <= len(observed):
                continue

            changed = 0
            for row, cell, current_value in observed:
                if current_value == dominant_value:
                    continue
                try:
                    set_majority_format_value(cell, property_name, dominant_value)
                    total_changed_cells.add((row, col))
                    changed += 1
                except Exception:
                    continue
            if changed:
                changed_by_property.append(f"{property_labels[property_name]} {changed}处")

        if changed_by_property:
            normalized_columns.append(
                f"{column_number_to_name(col)}列（{'、'.join(changed_by_property)}）"
            )

    if normalized_columns:
        log(
            f"  多数派格式预处理：按列修正 {len(total_changed_cells)} 个非空单元格；"
            + "；".join(normalized_columns)
        )
    else:
        log("  多数派格式预处理：各列非空单元格格式已一致，或不存在严格多数派，无需修正")
    return len(total_changed_cells)


# ================= 备份 =================
def has_existing_backup_dir(cwd):
    try:
        return any(
            name.startswith("自动备份_") and os.path.isdir(os.path.join(cwd, name))
            for name in os.listdir(cwd)
        )
    except Exception:
        return False


def prepare_backup_dir(cwd):
    global CURRENT_BACKUP_DIR, BACKUP_SKIP_LOGGED
    if not ENABLE_BACKUP:
        return None
    if CURRENT_BACKUP_DIR:
        return CURRENT_BACKUP_DIR
    if SKIP_BACKUP_WHEN_BACKUP_DIR_EXISTS and has_existing_backup_dir(cwd):
        if not BACKUP_SKIP_LOGGED:
            log("  检测到当前目录已有自动备份，判断为已完成首轮处理，本轮不再创建备份。")
            BACKUP_SKIP_LOGGED = True
        return None
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    CURRENT_BACKUP_DIR = os.path.join(cwd, f"自动备份_{ts}")
    os.makedirs(CURRENT_BACKUP_DIR, exist_ok=True)
    return CURRENT_BACKUP_DIR


def backup_file(path, cwd):
    if not ENABLE_BACKUP:
        return None
    try:
        backup_dir = prepare_backup_dir(cwd)
        if not backup_dir:
            return None
        name = os.path.basename(path)
        backup_path = os.path.join(backup_dir, name)
        shutil.copy2(path, backup_path)
        log(f"  已备份原文件：{backup_path}")
        return backup_path
    except Exception as e:
        log(f"  警告：备份失败：{e}")
        return None


# ================= Excel 排版辅助函数 =================
def remove_old_auto_spacer_rows(ws):
    """删除上一次运行自动插入的分页空白行，保证脚本可反复测试。"""
    deleted = 0
    used_rows = get_used_rows(ws)
    for r in range(used_rows, 2, -1):
        if is_auto_spacer_row(ws, r):
            try:
                ws.Rows(r).Delete()
                deleted += 1
            except Exception as e:
                log(f"  删除旧分页空白行失败：row={r}, err={e}")
    if deleted:
        log(f"  已删除上次自动分页插入的空白行 {deleted} 行")
    else:
        log("  未发现上次自动分页插入的空白行")
    return deleted


def hide_helper_column(ws):
    try:
        ws.Columns(AUTO_SPACER_HELPER_COL).Hidden = True
    except Exception:
        pass


def copy_cell_fill(source_cell, target_cell):
    """只复制底色/填充图案，不动字体颜色、边框和其他格式。"""
    try:
        pattern = source_cell.Interior.Pattern
        target_cell.Interior.Pattern = pattern
        if pattern != XL_NONE:
            try:
                target_cell.Interior.Color = source_cell.Interior.Color
            except Exception:
                pass
            try:
                target_cell.Interior.PatternColor = source_cell.Interior.PatternColor
            except Exception:
                pass
        return True
    except Exception:
        return False


def sync_site_sequence_fill(ws, used_rows, source_col=2):
    """原计划表重排序号后，让A列序号底色跟随同一行专业列底色。"""
    changed = 0
    for row in range(DATA_START_ROW, used_rows + 1):
        if is_auto_spacer_row(ws, row):
            continue
        try:
            if copy_cell_fill(ws.Cells(row, source_col), ws.Cells(row, 1)):
                changed += 1
        except Exception:
            continue
    log(
        f"  原计划表序号底色同步：A列已按同一行"
        f"{column_number_to_name(source_col)}列同步 {changed} 行"
    )
    return changed


def clear_print_data_backgrounds(ws, used_rows, max_col, mode=None):
    """风险管控/到岗到位打印表正文统一白底；风险黄色在最终校验中补回。"""
    if mode not in {"risk", "daogang"} or used_rows < DATA_START_ROW:
        return False
    try:
        data_range = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(used_rows, max_col))
        data_range.Interior.Pattern = XL_SOLID
        data_range.Interior.Color = XL_COLOR_WHITE
        log(
            f"  打印表正文底色清理：R{DATA_START_ROW}C1:R{used_rows}C{max_col} 已统一白底；"
            "营销绿色取消，二级/三级风险黄色将在最终校验中保留/补齐；四级/五级保持白底"
        )
        return True
    except Exception as e:
        log(f"  打印表正文底色清理失败：{e}")
        return False


def renumber_sequence(ws, used_rows):
    seq = 1
    changed = 0
    for r in range(DATA_START_ROW, used_rows + 1):
        if is_auto_spacer_row(ws, r):
            try:
                ws.Cells(r, 1).ClearContents()
            except Exception:
                pass
            continue
        try:
            ws.Cells(r, 1).Value = seq
            changed += 1
        except Exception:
            pass
        seq += 1
    log(f"  已重排序号：有效数据行 {changed} 行，范围第{DATA_START_ROW}行到第{used_rows}行")
    return changed


def get_font_name(cell):
    font_sources = []
    try:
        font_sources.append(cell.Font)
    except Exception:
        pass
    try:
        font_sources.append(cell.GetCharacters(1, 1).Font)
    except Exception:
        pass
    for font in font_sources:
        for property_name in ("Name", "NameFarEast"):
            try:
                name = str(getattr(font, property_name) or "").strip()
                if name:
                    return name
            except Exception:
                continue
    return ""


def get_font_width_ratio(font_name):
    if font_name in AUTOFITROWEX_FONT_SETTINGS:
        return AUTOFITROWEX_FONT_SETTINGS[font_name][1]
    if font_name.startswith("MS") or font_name.startswith("HG"):
        return AUTOFITROWEX_MS_WIDTH_RATIO
    return 0.88 if AUTOFITROWEX_PRINT_MODE else 1.0


def get_font_height_ratio(font_name):
    if font_name in AUTOFITROWEX_FONT_SETTINGS:
        return AUTOFITROWEX_FONT_SETTINGS[font_name][0]
    if font_name.startswith("HG") or font_name.startswith("MS"):
        return 1.06
    if font_name == "Times New Roman":
        return 1.20
    if font_name == "Tahoma":
        return 1.05
    if font_name == "Verdana":
        return 1.15
    if font_name == "Meiryo":
        return 1.11
    if font_name == "Meiryo UI":
        return 1.20
    if font_name == "Arial Unicode MS":
        return 1.15
    if "Microsoft JhengHei" in font_name:
        return 1.10
    if font_name == "DFKai-SB":
        return 1.20
    if font_name == "Dotum":
        return 1.20
    if font_name == "FangSong":
        return 1.15
    if font_name == "Gulim":
        return 1.15
    if font_name == "KaiTi":
        return 1.12
    if font_name == "Malgun Gothic":
        return 1.25
    if font_name == "Microsoft YaHei Light":
        return 1.05
    if "MingLiU" in font_name or "SimSun" in font_name:
        return 1.08
    return AUTOFITROWEX_DEFAULT_EXTRA_HEIGHT


def delete_autofitrowex_temp_sheet(wb):
    try:
        app = wb.Application
        for sheet in list(wb.Worksheets):
            try:
                if str(sheet.Name) == AUTOFITROWEX_TEMP_SHEET:
                    previous_alerts = app.DisplayAlerts
                    app.DisplayAlerts = False
                    sheet.Delete()
                    app.DisplayAlerts = previous_alerts
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def create_autofitrowex_temp_sheet(wb):
    delete_autofitrowex_temp_sheet(wb)
    app = wb.Application
    temp = wb.Worksheets.Add()
    temp.Name = AUTOFITROWEX_TEMP_SHEET
    temp.Rows(1).Font.Size = 1
    try:
        temp.Visible = 0
    except Exception:
        pass
    try:
        app.CutCopyMode = False
    except Exception:
        pass
    return temp


def get_cell_merge_bounds(cell):
    try:
        ma = cell.MergeArea
        start_row = int(ma.Row)
        start_col = int(ma.Column)
        row_count = int(ma.Rows.Count)
        col_count = int(ma.Columns.Count)
        return ma, start_row, start_col, row_count, col_count
    except Exception:
        return cell, int(cell.Row), int(cell.Column), 1, 1


def choose_autofit_plus_height(ws, used_rows, mode="normal", compact=False):
    """根据最终字号自动选择行高余量；风险管控/到岗到位不再额外 +1。"""
    if mode in {"risk", "daogang"}:
        return 0.0
    if mode in {"site", "normal"} and not compact:
        return SITE_AUTOFIT_PLUS_HEIGHT

    sample_rows = [r for r in range(DATA_START_ROW, min(used_rows, DATA_START_ROW + 8) + 1)]
    font_sizes = []
    for r in sample_rows:
        for c in (1, 3, 4, 20):
            try:
                value = ws.Cells(r, c).Value
                if is_blank_value(value):
                    continue
                size = safe_float(ws.Cells(r, c).Font.Size, 0.0)
                if size > 0:
                    font_sizes.append(size)
            except Exception:
                continue

    base_size = max(font_sizes) if font_sizes else (DAOGANG_FONT_SIZE if mode == "daogang" else 11)
    if mode == "daogang" or base_size >= 28:
        return max(MIN_AUTOFIT_PLUS_HEIGHT, min(MAX_COMPACT_AUTOFIT_PLUS_HEIGHT, round(base_size * 0.025, 1)))
    return max(MIN_AUTOFIT_PLUS_HEIGHT, min(DEFAULT_AUTOFIT_PLUS_HEIGHT, round(base_size * 0.18, 1)))


def get_autofit_inserted_rows(result):
    try:
        return int(result[3])
    except Exception:
        return 0


def get_merged_width(ws, row, start_col, col_count):
    width = 0.0
    for c in range(start_col, start_col + col_count):
        try:
            width += safe_float(ws.Cells(row, c).ColumnWidth, 0.0)
        except Exception:
            pass
    if col_count > 1:
        width += (col_count - 1) * 0.5
    return width


def get_merged_height(ws, start_row, row_count, col):
    height = 0.0
    for r in range(start_row, start_row + row_count):
        height += get_row_height(ws, r, 15.0)
    return height


def prepare_temp_cell_from_source(temp_ws, source_cell, width, text=None):
    tmp = temp_ws.Cells(1, 1)
    try:
        tmp.Clear()
    except Exception:
        pass
    try:
        tmp.Value = source_cell.Value if text is None else text
        tmp.Font.Name = source_cell.Font.Name
        tmp.Font.Size = source_cell.Font.Size
        tmp.Font.Bold = source_cell.Font.Bold
        tmp.Font.Italic = source_cell.Font.Italic
        tmp.Font.Underline = source_cell.Font.Underline
        tmp.Font.Color = source_cell.Font.Color
    except Exception:
        try:
            tmp.Value = source_cell.Value if text is None else text
        except Exception:
            pass
    try:
        tmp.WrapText = True
    except Exception:
        pass
    try:
        ratio = get_font_width_ratio(get_font_name(source_cell)) if AUTOFITROWEX_PRINT_MODE else 1.0
        tmp.ColumnWidth = max(0.1, width * ratio)
    except Exception:
        pass
    return tmp


def fit_temp_cell_height(temp_ws, source_cell, width, text=None):
    tmp = prepare_temp_cell_from_source(temp_ws, source_cell, width, text=text)
    try:
        temp_ws.Rows(1).AutoFit()
    except Exception:
        pass
    height = get_row_height(temp_ws, 1, 15.0)
    if AUTOFITROWEX_PRINT_MODE:
        height *= get_font_height_ratio(get_font_name(source_cell))
    return height


def estimate_required_height_and_overflow_rows(temp_ws, source_cell, width):
    value = source_cell.Value
    height = fit_temp_cell_height(temp_ws, source_cell, width)
    overflow_chunks = 0
    measured_chunk_height = 0.0

    if height <= EXCEL_MAX_ROW_HEIGHT:
        return height, 0, height

    text = "" if value is None else str(value)
    remaining = text
    guard = 0
    while remaining and height > EXCEL_MAX_ROW_HEIGHT and guard < 200:
        guard += 1
        low = 1
        high = len(remaining)
        best = 0
        best_height = 0.0
        while low <= high:
            mid = (low + high) // 2
            test_height = fit_temp_cell_height(temp_ws, source_cell, width, text=remaining[:mid])
            if test_height <= EXCEL_MAX_ROW_HEIGHT:
                best = mid
                best_height = test_height
                low = mid + 1
            else:
                high = mid - 1

        if best <= 0:
            break
        overflow_chunks += 1
        measured_chunk_height += best_height
        remaining = remaining[best:]
        height = fit_temp_cell_height(temp_ws, source_cell, width, text=remaining)

    measured_required_height = measured_chunk_height + height
    # 不再用“完整块数 + 1”的保守高度；按真实测量总高度直接排版。
    return measured_required_height, overflow_chunks, measured_required_height


def merge_inserted_row_with_upper(ws, insert_row, max_col):
    merge_inserted_rows_with_upper(ws, insert_row, 1, max_col)


def merge_inserted_rows_with_upper(ws, insert_row, rows_to_insert, max_col):
    if rows_to_insert <= 0:
        return
    end_insert_row = insert_row + rows_to_insert - 1
    if rows_to_insert == 1:
        ws.Rows(insert_row).Insert()
    else:
        ws.Rows(f"{insert_row}:{end_insert_row}").Insert()
    for r in range(insert_row, end_insert_row + 1):
        try:
            ws.Rows(r).RowHeight = DEFAULT_INSERTED_CONTINUATION_ROW_HEIGHT
        except Exception:
            pass
    for c in range(1, max_col + 1):
        try:
            above = ws.Cells(insert_row - 1, c)
            ma, start_row, start_col, row_count, col_count = get_cell_merge_bounds(above)
            if c != start_col:
                continue
            end_col = min(start_col + col_count - 1, max_col)
            target = ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_insert_row, end_col))
            if ws.Cells(insert_row, end_col).MergeArea.Count == 1:
                target.Merge()
        except Exception:
            continue


def apply_height_to_merge_area(ws, top_row, col, required_height):
    _, start_row, _, row_count, _ = get_cell_merge_bounds(ws.Cells(top_row, col))
    current_height = get_merged_height(ws, start_row, row_count, col)
    if required_height <= current_height:
        return 0
    diff = required_height - current_height
    leftover = 0.0
    changed = 0
    for offset in range(row_count):
        row = start_row + offset
        current = get_row_height(ws, row, 15.0)
        add_height = round(diff / row_count + leftover / (row_count - offset), 1)
        if current + add_height > EXCEL_MAX_ROW_HEIGHT:
            actual_add = EXCEL_MAX_ROW_HEIGHT - current
            leftover += add_height - actual_add
        else:
            actual_add = add_height
        if actual_add > 0:
            ws.Rows(row).RowHeight = current + actual_add
            changed += 1
    return changed


def simple_autofit_rows(ws, used_rows, plus_height=5, start_row=1):
    adjusted = 0
    failed = 0
    for r in range(start_row, used_rows + 1):
        if is_auto_spacer_row(ws, r):
            continue
        try:
            ws.Rows(r).AutoFit()
            try:
                ws.Rows(r).RowHeight = min(EXCEL_MAX_ROW_HEIGHT, ws.Rows(r).RowHeight + plus_height)
            except Exception:
                pass
            adjusted += 1
        except Exception:
            failed += 1
    log(f"  已按原始流程自动调整行高并 +{plus_height}：成功 {adjusted} 行，失败 {failed} 行")
    return adjusted, failed


def autofit_rows(
    ws,
    used_rows,
    plus_height=5,
    max_col=None,
    start_row=AUTOFITROWEX_START_ROW,
    skip_columns=None,
    measure_columns=None,
    overflow_row_capacity_factor=1.0,
):
    skip_columns = set(skip_columns or [])
    measure_columns = set(measure_columns or [])
    overflow_row_capacity_factor = max(1.0, safe_float(overflow_row_capacity_factor, 1.0))
    if not ENABLE_AUTOFITROWEX_HEIGHT:
        adjusted = 0
        skipped_spacer = 0
        failed = 0
        for r in range(start_row, used_rows + 1):
            if is_auto_spacer_row(ws, r):
                skipped_spacer += 1
                continue
            try:
                ws.Rows(r).AutoFit()
                ws.Rows(r).RowHeight = min(EXCEL_MAX_ROW_HEIGHT, ws.Rows(r).RowHeight + plus_height)
                adjusted += 1
            except Exception:
                failed += 1
        log(f"  已自动调整行高并 +{plus_height}：成功 {adjusted} 行，跳过分页空白行 {skipped_spacer} 行，失败 {failed} 行")
        return adjusted, skipped_spacer, failed, 0

    adjusted = 0
    skipped_spacer = 0
    failed = 0
    overflow_inserted = 0
    measured_cells = 0
    planned_blocks = {}
    wb = ws.Parent
    temp_ws = None
    target_max_col = max_col if max_col else 60

    try:
        temp_ws = create_autofitrowex_temp_sheet(wb)
        ws.Activate()
        ws.Rows(f"{start_row}:{used_rows}").AutoFit()
        if skip_columns:
            log(f"  AutoFitRowEx：跳过列 {sorted(skip_columns)} 的高度测量")
        if measure_columns:
            log(f"  AutoFitRowEx：仅重点测量列 {sorted(measure_columns)}")

        r = start_row
        while r <= used_rows:
            if is_auto_spacer_row(ws, r):
                skipped_spacer += 1
                r += 1
                continue
            if r == start_row or (r - start_row) % 5 == 0:
                log(f"  AutoFitRowEx进度：正在测量第 {r}/{used_rows} 行")

            try:
                last_col = int(ws.Cells(r, ws.Columns.Count).End(-4159).Column)  # xlToLeft
            except Exception:
                last_col = 1
            target_max_col = max_col if max_col else last_col
            last_col = max(1, min(last_col, target_max_col, 60))

            for c in range(1, last_col + 1):
                if measure_columns and c not in measure_columns:
                    continue
                if c in skip_columns:
                    continue
                try:
                    cell = ws.Cells(r, c)
                    if is_blank_value(cell.Value):
                        continue
                    ma, merge_start_row, start_col, row_count, col_count = get_cell_merge_bounds(cell)
                    if r != merge_start_row or c != start_col:
                        continue
                    if not AUTOFITROWEX_PRINT_MODE and row_count == 1 and col_count == 1:
                        continue

                    width = get_merged_width(ws, r, start_col, col_count)
                    if width <= 0:
                        continue
                    required_height, overflow_chunks, measured_required_height = estimate_required_height_and_overflow_rows(
                        temp_ws, cell, width
                    )
                    measured_cells += 1

                    key = (merge_start_row, row_count)
                    plan = planned_blocks.get(key)
                    if plan is None:
                        plan = {
                            "start_row": merge_start_row,
                            "row_count": row_count,
                            "rows_needed": row_count,
                            "required_height": 0.0,
                            "representative_col": start_col,
                            "trigger_cols": set(),
                        }
                        planned_blocks[key] = plan

                    if overflow_chunks > 0 and ENABLE_AUTOFITROWEX_OVERFLOW_ROWS:
                        # 风险管控和到岗到位都不再机械使用 overflow_chunks + 1。
                        # 直接按真实测量总高度 / 单个物理行可承载高度计算需要的行数：需要几行就几行。
                        rows_needed_by_height = max(
                            1,
                            int(math.ceil(
                                measured_required_height
                                / (EXCEL_MAX_ROW_HEIGHT * overflow_row_capacity_factor)
                            )),
                        )
                        plan["rows_needed"] = max(plan["rows_needed"], rows_needed_by_height)
                        plan["trigger_cols"].add(start_col)
                    plan["required_height"] = max(plan["required_height"], required_height + plus_height)
                except Exception:
                    failed += 1
            r += 1

        plans = sorted(planned_blocks.values(), key=lambda p: p["start_row"], reverse=True)
        if plans:
            log(f"  AutoFitRowEx：扫描完成，需处理记录块 {len(plans)} 个，开始批量插行/调高")

        for plan in plans:
            merge_start_row = plan["start_row"]
            row_count = plan["row_count"]
            rows_needed = int(plan["rows_needed"])
            rows_to_insert = max(0, rows_needed - row_count)
            if rows_to_insert > 0:
                insert_at = merge_start_row + row_count
                try:
                    merge_inserted_rows_with_upper(ws, insert_at, rows_to_insert, target_max_col)
                    overflow_inserted += rows_to_insert
                    used_rows += rows_to_insert
                    trigger_cols = sorted(plan["trigger_cols"])
                    log(f"  AutoFitRowEx：R{merge_start_row} 文本超过单行高度上限，触发列 {trigger_cols}，已批量插入续行 {rows_to_insert} 行")
                except Exception as e:
                    failed += 1
                    log(f"  AutoFitRowEx：R{merge_start_row} 批量插入续行失败：{e}")
                    continue

            try:
                changed = apply_height_to_merge_area(
                    ws,
                    merge_start_row,
                    plan["representative_col"],
                    plan["required_height"],
                )
                adjusted += changed
            except Exception:
                failed += 1
    finally:
        try:
            ws.Activate()
        except Exception:
            pass
        if temp_ws is not None:
            delete_autofitrowex_temp_sheet(wb)

    log(
        f"  AutoFitRowEx风格行高调整：测量单元格 {measured_cells} 个，"
        f"调整行 {adjusted} 行，行高余量 +{plus_height}，插入超高续行 {overflow_inserted} 行，"
        f"跳过分页空白行 {skipped_spacer} 行，失败 {failed} 次"
    )
    return adjusted, skipped_spacer, failed, overflow_inserted


def round_up_height(value, step=5.0, maximum=EXCEL_MAX_ROW_HEIGHT):
    if value <= 0:
        return 0.0
    rounded = math.ceil(value / step) * step
    return min(maximum, round(rounded, 1))


def collect_sequence_record_blocks(ws, used_rows):
    """按序号列收集正文记录块，兼容到岗到位的跨行合并记录。"""
    blocks = []
    seen = set()
    row = DATA_START_ROW
    while row <= used_rows:
        if is_auto_spacer_row(ws, row):
            row += 1
            continue
        try:
            cell = ws.Cells(row, 1)
            area = cell.MergeArea
            start = int(area.Row)
            end = int(area.Row + area.Rows.Count - 1)
            value = area.Cells(1, 1).Value
        except Exception:
            start = end = row
            value = ws.Cells(row, 1).Value

        if start < DATA_START_ROW:
            row += 1
            continue
        key = (start, end)
        if key not in seen and not is_blank_value(value):
            seen.add(key)
            blocks.append({"start": start, "end": min(end, used_rows)})
        row = max(row + 1, end + 1)
    return blocks


def normalize_print_record_heights(ws, used_rows, max_col, mode=None):
    """
    将两个打印表的正文记录规整为统一打印高度。

    risk：每条记录统一为本表最大所需行高；
    daogang：保留长文本所需续行数，统一所有正文物理行的高度。
    """
    if mode not in {"risk", "daogang"}:
        return {"used_rows": used_rows, "inserted_rows": 0, "blocks": 0}

    blocks = collect_sequence_record_blocks(ws, used_rows)
    if not blocks:
        log("  打印正文行高规整：未识别到有效记录，跳过")
        return {"used_rows": used_rows, "inserted_rows": 0, "blocks": 0}

    if mode != "daogang":
        original_heights = [
            get_row_height(ws, block["start"], 15.0)
            for block in blocks
        ]
        target_height = round_up_height(max(original_heights), step=1.0)
        for block in blocks:
            ws.Rows(block["start"]).RowHeight = target_height
        log(
            f"  打印正文行高规整：{len(blocks)} 条记录统一为 {target_height:.1f}pt；"
            f"原行高范围 {min(original_heights):.1f}-{max(original_heights):.1f}pt"
        )
        return {
            "used_rows": used_rows,
            "inserted_rows": 0,
            "blocks": len(blocks),
            "row_height": target_height,
        }

    average_row_heights = []
    physical_rows = 0
    for block in blocks:
        row_count = block["end"] - block["start"] + 1
        physical_rows += row_count
        average_row_heights.append(
            sum_row_heights(ws, block["start"], block["end"]) / row_count
        )
    target_row_height = round_up_height(
        max(average_row_heights),
        step=2.5,
    )

    for block in blocks:
        ws.Rows(f"{block['start']}:{block['end']}").RowHeight = target_row_height

    log(
        f"  打印正文行高规整：{len(blocks)} 条到岗到位记录共 {physical_rows} 个正文行，"
        f"全部统一为 {target_row_height:.1f}pt；保留长文本原续行数，不额外补空行"
    )
    return {
        "used_rows": used_rows,
        "inserted_rows": 0,
        "blocks": len(blocks),
        "physical_rows": physical_rows,
        "row_height": target_row_height,
    }


def compact_print_rows_to_save_pages(ws, excel, used_rows, mode=None):
    """
    仅在小幅压缩行高即可少打一页时收紧正文，不修改字号、列宽或内容。
    """
    if mode not in {"risk", "daogang"} or used_rows < DATA_START_ROW:
        return {"applied": False}

    scale_percent, scale_source = get_effective_print_scale_percent(
        ws,
        excel,
        19 if mode == "daogang" else 12,
    )
    scale = max(0.01, scale_percent / 100.0)
    page_height = get_print_body_height_points(excel) * PAGE_HEIGHT_SAFETY_FACTOR
    title_height = (
        sum_row_heights(ws, 1, DATA_START_ROW - 1) * scale
        if ENABLE_REPEAT_TITLE_ROWS
        else 0.0
    )
    page_data_capacity = max(1.0, page_height - title_height)
    data_height = sum_row_heights(ws, DATA_START_ROW, used_rows) * scale
    current_pages = max(1, int(math.ceil(data_height / page_data_capacity)))

    if current_pages <= 1:
        log(
            f"  节纸行高检查：当前估算 1 页（缩放 {scale_percent:g}%，{scale_source}），"
            "保留现有阅读留白，不再压缩"
        )
        return {"applied": False, "pages": 1}

    target_pages = current_pages - 1
    exact_factor = target_pages * page_data_capacity / data_height
    if exact_factor < MIN_PRINT_ROW_COMPACT_FACTOR:
        log(
            f"  节纸行高检查：预计 {current_pages} 页；若减少到 {target_pages} 页需压缩"
            f"至 {exact_factor * 100:.1f}%，超过可读性限制，保持现有行高"
        )
        return {"applied": False, "pages": current_pages}

    compact_factor = max(
        MIN_PRINT_ROW_COMPACT_FACTOR,
        min(0.995, exact_factor * 0.995),
    )
    changed_rows = 0
    for row in range(DATA_START_ROW, used_rows + 1):
        if is_auto_spacer_row(ws, row):
            continue
        current_height = get_row_height(ws, row, 15.0)
        ws.Rows(row).RowHeight = round(current_height * compact_factor, 1)
        changed_rows += 1

    log(
        f"  节纸行高优化：不改字号，正文 {changed_rows} 行整体保留"
        f" {compact_factor * 100:.1f}% 行高；估算由 {current_pages} 页降至 {target_pages} 页"
    )
    return {
        "applied": True,
        "factor": compact_factor,
        "before_pages": current_pages,
        "after_pages": target_pages,
    }


def apply_full_borders(ws, used_rows, max_col):
    try:
        rng = ws.Range(ws.Cells(BORDER_START_ROW, 1), ws.Cells(used_rows, max_col))
        rng.Borders.LineStyle = XL_CONTINUOUS
        rng.Borders.Weight = XL_THIN
        rng.Borders.Color = XL_COLOR_BLACK
        log(f"  已添加/重刷全区域边框：R{BORDER_START_ROW}C1:R{used_rows}C{max_col}")
        return True
    except Exception as e:
        log(f"  添加边框失败: {e}")
        return False


def apply_print_border_hierarchy(ws, used_rows, max_col):
    """所有框线统一细实线；不再加粗外框或表头下沿。"""
    if used_rows < BORDER_START_ROW or max_col < 1:
        return False
    try:
        table_range = ws.Range(
            ws.Cells(BORDER_START_ROW, 1),
            ws.Cells(used_rows, max_col),
        )
        for border_id in (XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT):
            border = table_range.Borders(border_id)
            border.LineStyle = XL_CONTINUOUS
            border.Weight = XL_THIN
            border.Color = XL_COLOR_BLACK

        header_range = ws.Range(
            ws.Cells(BORDER_START_ROW, 1),
            ws.Cells(BORDER_START_ROW, max_col),
        )
        header_border = header_range.Borders(XL_EDGE_BOTTOM)
        header_border.LineStyle = XL_CONTINUOUS
        header_border.Weight = XL_THIN
        header_border.Color = XL_COLOR_BLACK
        log("  打印框线层次：内部、外框及表头下沿全部统一细线")
        return True
    except Exception as e:
        log(f"  打印框线层次设置失败：{e}")
        return False


def get_print_max_col(max_col, mode=None):
    """
    数据处理可以到更右侧列；打印区/框线统一按用户最终要求截止到 S 列。
    到岗到位 T 列仍保留并参与合并/行高处理，但不纳入打印框线。
    """
    if max_col >= PRINT_MAX_COLUMN_S:
        return PRINT_MAX_COLUMN_S
    return max_col


def setup_print_page(ws, excel, used_rows, max_col):
    """A3 横向 + 窄页边距 + 打印区域。"""
    if not ENABLE_PAGE_SETUP:
        log("  已跳过页面设置：ENABLE_PAGE_SETUP=False")
        return False
    print_comm_disabled = False
    try:
        try:
            excel.PrintCommunication = False
            print_comm_disabled = True
        except Exception:
            pass
        ps = ws.PageSetup
        ps.PaperSize = PAGE_PAPER_SIZE_A3
        ps.Orientation = PAGE_ORIENTATION_LANDSCAPE
        ps.LeftMargin = excel.InchesToPoints(MARGIN_LEFT_INCH)
        ps.RightMargin = excel.InchesToPoints(MARGIN_RIGHT_INCH)
        ps.TopMargin = excel.InchesToPoints(MARGIN_TOP_INCH)
        ps.BottomMargin = excel.InchesToPoints(MARGIN_BOTTOM_INCH)
        ps.HeaderMargin = excel.InchesToPoints(MARGIN_HEADER_INCH)
        ps.FooterMargin = excel.InchesToPoints(MARGIN_FOOTER_INCH)
        ps.PrintGridlines = False
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        if ENABLE_REPEAT_TITLE_ROWS:
            ps.PrintTitleRows = REPEAT_TITLE_ROWS_TEXT
        else:
            ps.PrintTitleRows = ""
        ps.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(used_rows, max_col)).Address
        if print_comm_disabled:
            try:
                excel.PrintCommunication = True
            except Exception:
                pass
            print_comm_disabled = False
        log("  已设置页面：A3 横向、窄页边距、适配为1页宽")
        log(f"  打印区域：R1C1:R{used_rows}C{max_col}")
        log(f"  每页重复标题行：{'开启 ' + REPEAT_TITLE_ROWS_TEXT if ENABLE_REPEAT_TITLE_ROWS else '关闭'}")
        return True
    except Exception as e:
        if print_comm_disabled:
            try:
                excel.PrintCommunication = True
            except Exception:
                pass
        log(f"  设置页面/打印区域失败: {e}")
        return False


def get_available_page_height_points(excel):
    """计算 A3 横向时每页可打印高度，单位 points。"""
    try:
        # A3 横向时，页面高度为短边 297mm。1 inch = 25.4mm，1 inch = 72 points。
        a3_landscape_page_height_points = 297.0 / 25.4 * 72.0
        available = (
            a3_landscape_page_height_points
            - excel.InchesToPoints(MARGIN_TOP_INCH)
            - excel.InchesToPoints(MARGIN_BOTTOM_INCH)
            - excel.InchesToPoints(MARGIN_HEADER_INCH)
            - excel.InchesToPoints(MARGIN_FOOTER_INCH)
        )
        return max(100.0, available * PAGE_HEIGHT_SAFETY_FACTOR)
    except Exception:
        # 保守兜底值
        return 650.0


def get_print_body_height_points(excel):
    """A3 横向打印正文高度，单位 points；页眉页脚在页边距内，不额外扣减。"""
    try:
        a3_landscape_page_height_points = 297.0 / 25.4 * 72.0
        return max(
            100.0,
            a3_landscape_page_height_points
            - excel.InchesToPoints(MARGIN_TOP_INCH)
            - excel.InchesToPoints(MARGIN_BOTTOM_INCH),
        )
    except Exception:
        return 733.0


def get_print_body_width_points(excel):
    """A3 横向打印正文宽度，单位 points。"""
    try:
        a3_landscape_page_width_points = 420.0 / 25.4 * 72.0
        return max(
            100.0,
            a3_landscape_page_width_points
            - excel.InchesToPoints(MARGIN_LEFT_INCH)
            - excel.InchesToPoints(MARGIN_RIGHT_INCH),
        )
    except Exception:
        return 1154.0


def get_range_width_points(ws, max_col):
    """读取当前打印列的实际宽度，单位 points。"""
    try:
        width = safe_float(ws.Range(ws.Cells(1, 1), ws.Cells(1, max_col)).Width, 0.0)
        if width > 0:
            return width
    except Exception:
        pass

    total = 0.0
    for c in range(1, max_col + 1):
        try:
            total += safe_float(ws.Range(ws.Cells(1, c), ws.Cells(1, c)).Width, 0.0)
        except Exception:
            pass
    return total


def get_effective_print_scale_percent(ws, excel, max_col):
    """按当前页面设置推导实际打印缩放，避免按某个样板写死比例。"""
    try:
        zoom = ws.PageSetup.Zoom
        zoom_percent = safe_float(zoom, 0.0)
        if zoom not in (False, None, "") and zoom_percent > 0:
            return zoom_percent, "PageSetup.Zoom"
    except Exception:
        pass

    try:
        fit_wide = safe_float(ws.PageSetup.FitToPagesWide, 0.0)
    except Exception:
        fit_wide = 0.0

    if fit_wide == 1:
        body_width = get_print_body_width_points(excel)
        table_width = get_range_width_points(ws, max_col)
        if body_width > 0 and table_width > 0:
            scale_percent = math.floor((body_width / table_width) * 100.0)
            scale_percent = max(MIN_EXCEL_PRINT_SCALE_PERCENT, scale_percent)
            scale_percent = min(MAX_EXCEL_PRINT_SCALE_PERCENT, scale_percent)
            return scale_percent, f"FitToPagesWide=1，打印宽 {body_width:.1f}pt / 表宽 {table_width:.1f}pt"

    return MAX_EXCEL_PRINT_SCALE_PERCENT, "默认 100%"


def compute_scaled_record_page_break_rows(
    ws,
    excel,
    used_rows,
    max_col,
    scale_percent=None,
    forced_break_rows=None,
):
    """
    按当前行高和打印缩放估算分页，并且只在完整记录块前分页。
    用于到岗到位版：Excel COM 有时读不到自动分页线，但实际打印仍会分页。
    """
    if used_rows < DATA_START_ROW:
        return []

    if scale_percent is None:
        scale_percent, scale_source = get_effective_print_scale_percent(ws, excel, max_col)
    else:
        scale_source = "指定缩放"

    scale = max(0.01, safe_float(scale_percent, MAX_EXCEL_PRINT_SCALE_PERCENT) / 100.0)
    page_height = (
        get_print_body_height_points(excel)
        * min(PAGE_HEIGHT_SAFETY_FACTOR, DAOGANG_PAGE_HEIGHT_SAFETY_FACTOR)
    )
    title_height = 0.0
    if ENABLE_REPEAT_TITLE_ROWS:
        title_height = sum_row_heights(ws, 1, DATA_START_ROW - 1) * scale

    blocks, merged_blocks, merge_area_count = build_no_split_blocks(ws, used_rows, max_col)
    data_blocks = [
        block for block in blocks
        if block["end"] >= DATA_START_ROW and block["start"] <= used_rows
    ]
    if not data_blocks:
        return []

    break_rows = []
    forced_break_rows = set(forced_break_rows or [])
    current_height = title_height
    page_has_data = False
    too_tall_blocks = []

    for block in data_blocks:
        start = max(DATA_START_ROW, block["start"])
        end = min(used_rows, block["end"])
        if start > end:
            continue
        block_height = sum_row_heights(ws, start, end) * scale

        if start in forced_break_rows and page_has_data:
            break_rows.append(start)
            current_height = title_height
            page_has_data = False

        if page_has_data and current_height + block_height > page_height:
            break_rows.append(start)
            current_height = title_height
            page_has_data = False

        current_height += block_height
        page_has_data = True

        if title_height + block_height > page_height:
            too_tall_blocks.append((start, end, title_height + block_height))

    log(
        "  到岗到位分页兜底估算："
        f"缩放 {scale_percent:g}%（{scale_source}），正文高度 {page_height:.1f}pt，"
        f"标题高度 {title_height:.1f}pt，合并区域 {merge_area_count} 个，"
        f"跨行记录块 {len(merged_blocks)} 个，"
        + (f"强制前移点 {sorted(forced_break_rows)}，" if forced_break_rows else "")
        + f"分页行 {break_rows or '无'}"
    )
    for start, end, height in too_tall_blocks[:5]:
        log(
            f"  警告：到岗到位记录 R{start}:R{end} 缩放后约 {height:.1f}pt，"
            "单条记录已超过一页正文高度"
        )
    if len(too_tall_blocks) > 5:
        log(f"  警告：另有 {len(too_tall_blocks) - 5} 个超高记录块未逐条列出")

    return break_rows


def apply_manual_page_break_rows(ws, break_rows, used_rows):
    added = 0
    for row in sorted(set(break_rows)):
        if not (DATA_START_ROW < row <= used_rows):
            continue
        try:
            ws.HPageBreaks.Add(Before=ws.Rows(row))
            added += 1
            log(f"  已添加手动分页符：第 {row} 行前")
        except Exception as e:
            log(f"  添加手动分页符失败：before_row={row}, err={e}")
    return added


def collect_merged_row_blocks(ws, used_rows, max_col):
    """
    收集所有跨多行的合并单元格，并合并成不可拆分页块。
    返回：blocks, merge_area_count
    block格式：{"start": int, "end": int, "addresses": [str]}
    """
    seen = set()
    intervals = []
    merge_area_count = 0

    for r in range(1, used_rows + 1):
        for c in range(1, max_col + 1):
            try:
                cell = ws.Cells(r, c)
                if not cell.MergeCells:
                    continue
                ma = cell.MergeArea
                addr = str(ma.Address)
                if addr in seen:
                    continue
                seen.add(addr)
                merge_area_count += 1
                start = int(ma.Row)
                end = int(ma.Row + ma.Rows.Count - 1)
                start_col = int(ma.Column)
                end_col = int(ma.Column + ma.Columns.Count - 1)
                # 只处理目标打印列范围内的合并区域
                if end_col < 1 or start_col > max_col:
                    continue
                if end > start:
                    intervals.append([start, end, [addr]])
            except Exception:
                continue

    if not intervals:
        return [], merge_area_count

    intervals.sort(key=lambda x: (x[0], x[1]))
    merged = []
    for start, end, addresses in intervals:
        if not merged or start > merged[-1]["end"]:
            merged.append({"start": start, "end": end, "addresses": list(addresses)})
        else:
            merged[-1]["end"] = max(merged[-1]["end"], end)
            merged[-1]["addresses"].extend(addresses)

    return merged, merge_area_count


def iter_contiguous_intervals(values):
    sorted_values = sorted(set(values))
    if not sorted_values:
        return
    start = prev = sorted_values[0]
    for value in sorted_values[1:]:
        if value == prev + 1:
            prev = value
            continue
        yield start, prev
        start = prev = value
    yield start, prev


def collect_merged_area_infos(ws, used_rows, max_col):
    """收集目标打印区域内的唯一合并区，避免边框阶段重复 COM 操作。"""
    seen = set()
    areas = []
    failed = 0

    for r in range(BORDER_START_ROW, used_rows + 1):
        for c in range(1, max_col + 1):
            try:
                cell = ws.Cells(r, c)
                if not cell.MergeCells:
                    continue
                ma = cell.MergeArea
                sr = int(ma.Row)
                sc = int(ma.Column)
                er = int(ma.Row + ma.Rows.Count - 1)
                ec = int(ma.Column + ma.Columns.Count - 1)
                if sr != r or sc != c:
                    continue
                if ec < 1 or sc > max_col or er < BORDER_START_ROW or sr > used_rows:
                    continue
                sr = max(BORDER_START_ROW, sr)
                er = min(used_rows, er)
                ec = min(max_col, ec)
                key = (sr, er, sc, ec)
                if key in seen:
                    continue
                seen.add(key)
                areas.append({"range": ma, "sr": sr, "er": er, "sc": sc, "ec": ec})
            except Exception:
                failed += 1

    return areas, failed


def set_black_thin_borders(rng):
    rng.Borders.LineStyle = XL_CONTINUOUS
    rng.Borders.Weight = XL_THIN
    rng.Borders.Color = XL_COLOR_BLACK


def fix_merged_borders(ws, used_rows, max_col):
    """重刷合并区域边框；先去重再批量处理，避免 Excel COM 长时间卡住。"""
    areas, failed = collect_merged_area_infos(ws, used_rows, max_col)
    fixed_rows = set()
    for item in areas:
        fixed_rows.update(range(item["sr"], item["er"] + 1))

    row_range_count = 0
    for start, end in iter_contiguous_intervals(fixed_rows):
        try:
            row_rng = ws.Range(ws.Cells(start, 1), ws.Cells(end, max_col))
            set_black_thin_borders(row_rng)
            row_range_count += 1
        except Exception:
            failed += 1

    fixed_areas = 0
    for item in areas:
        try:
            set_black_thin_borders(item["range"])
            fixed_areas += 1
        except Exception:
            failed += 1

    log(
        f"  合并单元格边框修复：合并区域 {fixed_areas} 个，占用行 {len(fixed_rows)} 行，"
        f"整行去重刷框区间 {row_range_count} 个，失败次数 {failed}"
    )
    if fixed_rows:
        rows_preview = sorted(fixed_rows)
        if len(rows_preview) > 20:
            log(f"  被重刷框线的合并区域行示例：{rows_preview[:20]} ... 共{len(rows_preview)}行")
        else:
            log(f"  被重刷框线的合并区域行：{rows_preview}")
    return fixed_areas, len(fixed_rows), failed


def set_range_border(rng, border_id, line_style=XL_CONTINUOUS, weight=XL_THIN):
    border = rng.Borders(border_id)
    border.LineStyle = line_style
    if line_style != XL_NONE:
        border.Weight = weight
        border.Color = XL_COLOR_BLACK


def get_record_page_break_misalignments(ws, used_rows):
    """返回落在序号记录块内部的分页行，以及应前移到的记录起始行。"""
    blocks = collect_sequence_record_blocks(ws, used_rows)
    break_rows = get_page_break_rows(ws, used_rows)
    misaligned = []
    for break_row in break_rows:
        for block in blocks:
            if block["start"] < break_row <= block["end"]:
                misaligned.append({
                    "break_row": break_row,
                    "record_start": block["start"],
                    "record_end": block["end"],
                })
                break
    return break_rows, misaligned


def apply_record_aware_borders(ws, used_rows, max_col, mode=None):
    """
    按序号列识别最终记录块后重建框线。

    AutoFitRowEx 插入的续行属于上一条记录，续行内部不应出现横线；只在每条
    记录的顶部和底部画横线。到岗到位 L/M 侧栏由专用函数管理，不参与记录横线。
    """
    blocks = collect_sequence_record_blocks(ws, used_rows)
    if not blocks:
        log("  记录边界框线校准：未识别到序号记录块，回退为全区域边框")
        return {
            "applied": apply_full_borders(ws, used_rows, max_col),
            "blocks": 0,
            "segments": 0,
            "misaligned_breaks": 0,
        }

    data_start = min(block["start"] for block in blocks)
    data_end = max(block["end"] for block in blocks)
    segments = [(1, max_col)]
    if mode == "daogang":
        segments = [(1, min(11, max_col))]
        if max_col >= 14:
            segments.append((14, max_col))

    failed = 0
    operations = 0
    try:
        header_range = ws.Range(
            ws.Cells(BORDER_START_ROW, 1),
            ws.Cells(BORDER_START_ROW, max_col),
        )
        set_black_thin_borders(header_range)
        operations += 1
    except Exception:
        failed += 1

    for start_col, end_col in segments:
        if start_col > end_col:
            continue
        try:
            data_range = ws.Range(
                ws.Cells(data_start, start_col),
                ws.Cells(data_end, end_col),
            )
            # 先一次性清掉插行继承的物理行横线，再按最终记录边界补回。
            set_range_border(data_range, XL_INSIDE_HORIZONTAL, line_style=XL_NONE)
            for border_id in (XL_EDGE_LEFT, XL_EDGE_RIGHT, XL_INSIDE_VERTICAL):
                set_range_border(data_range, border_id)
            operations += 1
        except Exception:
            failed += 1

        for block in blocks:
            try:
                block_range = ws.Range(
                    ws.Cells(block["start"], start_col),
                    ws.Cells(block["end"], end_col),
                )
                if block["end"] > block["start"]:
                    set_range_border(block_range, XL_INSIDE_HORIZONTAL, line_style=XL_NONE)
                set_range_border(block_range, XL_EDGE_TOP)
                set_range_border(block_range, XL_EDGE_BOTTOM)
                operations += 1
            except Exception:
                failed += 1

    if mode == "daogang":
        # L/M 侧栏只保留竖线、首页顶线和整表末行封底。
        clear_daogang_sidebar_horizontal_borders(ws, used_rows)

    break_rows, misaligned = get_record_page_break_misalignments(ws, used_rows)
    if misaligned:
        details = "；".join(
            f"分页R{item['break_row']}落在记录R{item['record_start']}:R{item['record_end']}内"
            for item in misaligned
        )
        log_red(f"  框线/分页边界检查发现错位：{details}")
    else:
        log(
            f"  记录边界框线校准：{len(blocks)} 条记录、{len(segments)} 个列区段，"
            f"续行内部横线已清除；分页符 {break_rows or '无'} 均位于记录边界"
        )
    if failed:
        log_red(f"  记录边界框线校准有 {failed} 次设置失败，请检查打印预览")
    return {
        "applied": failed == 0,
        "blocks": len(blocks),
        "segments": len(segments),
        "operations": operations,
        "failed": failed,
        "misaligned_breaks": len(misaligned),
    }


def apply_daogang_page_end_closure_borders(ws, used_rows, max_col):
    """
    把每个打印页末行的底边框真正写入物理单元格。

    对跨多行合并记录直接设置整个记录块的 EdgeBottom 时，Excel 屏幕上会显示
    边框，但保存后末行单元格可能没有实体底框，打印到分页边缘时就会漏线。
    因此分页稳定后，单独选中每个页末物理行重画底线。
    L/M 侧栏按既定版式跨页保持开放，只在整表末行由外框统一封底。
    """
    if used_rows < DATA_START_ROW:
        return {
            "applied": True,
            "page_end_rows": [],
            "failed": 0,
            "misaligned_page_ends": 0,
        }

    break_rows = get_page_break_rows(ws, used_rows)
    page_end_rows = sorted({
        row - 1
        for row in break_rows
        if DATA_START_ROW < row <= used_rows
    })
    page_end_rows.append(used_rows)
    page_end_rows = sorted(set(page_end_rows))

    record_end_rows = {
        block["end"]
        for block in collect_sequence_record_blocks(ws, used_rows)
    }
    misaligned_page_ends = [
        row for row in page_end_rows
        if row != used_rows and row not in record_end_rows
    ]
    if misaligned_page_ends:
        log_red(
            "  到岗到位页末封框发现分页仍落在记录内部："
            f"页末行 {misaligned_page_ends}"
        )

    segments = [(1, min(11, max_col))]
    if max_col >= 14:
        segments.append((14, max_col))

    failed = 0
    applied_ranges = 0
    for row in page_end_rows:
        for start_col, end_col in segments:
            if start_col > end_col:
                continue
            try:
                bottom_row_range = ws.Range(
                    ws.Cells(row, start_col),
                    ws.Cells(row, end_col),
                )
                set_range_border(
                    bottom_row_range,
                    XL_EDGE_BOTTOM,
                    line_style=XL_CONTINUOUS,
                    weight=XL_THIN,
                )
                applied_ranges += 1
            except Exception as e:
                failed += 1
                log_red(
                    f"  到岗到位页末封框失败：R{row}C{start_col}:C{end_col}，{e}"
                )

    log(
        "  到岗到位页末实体封框："
        f"分页符 {break_rows or '无'}，页末行 {page_end_rows}，"
        f"A:K/N:S 共落实 {applied_ranges} 段细底线；L/M 分页处保持开放"
    )
    return {
        "applied": failed == 0 and not misaligned_page_ends,
        "page_end_rows": page_end_rows,
        "failed": failed,
        "misaligned_page_ends": len(misaligned_page_ends),
    }


def build_no_split_blocks(ws, used_rows, max_col):
    """
    建立不可拆分页块。
    - 每一行本身是一个不可拆分页块；
    - 跨多行合并区域会合并成更大的不可拆分页块。
    """
    merged_blocks, merge_area_count = collect_merged_row_blocks(ws, used_rows, max_col)
    start_to_end = {b["start"]: b["end"] for b in merged_blocks}
    block_map = []
    r = 1
    while r <= used_rows:
        if r in start_to_end:
            end = start_to_end[r]
            block_map.append({"start": r, "end": end, "kind": "merged"})
            r = end + 1
        else:
            # 如果当前行落在某个合并块内部，跳到该块后面；正常来说不会走到这里
            inside = None
            for b in merged_blocks:
                if b["start"] < r <= b["end"]:
                    inside = b
                    break
            if inside:
                block_map.append({"start": inside["start"], "end": inside["end"], "kind": "merged"})
                r = inside["end"] + 1
            else:
                block_map.append({"start": r, "end": r, "kind": "row"})
                r += 1
    return block_map, merged_blocks, merge_area_count


def compute_page_layout_insertions(ws, excel, used_rows, max_col):
    available_height = get_available_page_height_points(excel)
    repeat_header_height = 0.0
    if ENABLE_REPEAT_TITLE_ROWS:
        repeat_header_height = sum_row_heights(ws, 1, 2)

    blocks, merged_blocks, merge_area_count = build_no_split_blocks(ws, used_rows, max_col)

    log("  分页估算信息：")
    log(f"    A3横向每页可用高度估算：{available_height:.1f} points，安全系数：{PAGE_HEIGHT_SAFETY_FACTOR}")
    log(f"    检测到合并区域总数：{merge_area_count} 个，跨多行合并块：{len(merged_blocks)} 个")
    if ENABLE_REPEAT_TITLE_ROWS:
        log(f"    重复标题行高度估算：{repeat_header_height:.1f} points")
    log(f"    分页模式：{'插入空白行 + 手动分页符' if ENABLE_INSERT_SPACER_ROWS else '只添加手动分页符，不插入可见空白行'}")

    insertions = []
    warnings_too_tall = []
    current_height = 0.0
    page_no = 1
    shift = 0

    for block in blocks:
        s = block["start"]
        e = block["end"]
        kind = block["kind"]
        block_height = sum_row_heights(ws, s, e)
        effective_page_limit = available_height
        block_too_tall = (
            block_height + (repeat_header_height if page_no > 1 and ENABLE_REPEAT_TITLE_ROWS else 0)
            > effective_page_limit
        )

        # 单个块已经超过一页，无法通过分页符让它变小：它只能单独占用纸面。
        # 旧逻辑仍把它的超额高度带入下一块，造成后续每行都被强制分页。
        if block_too_tall:
            if e > s or block_height > effective_page_limit:
                warnings_too_tall.append((s, e, block_height, kind))

        # 如果当前页放不下，则在该块前插入分页空白行，并让该块到下一页
        if current_height > 0 and current_height + block_height > effective_page_limit:
            remaining = max(0.0, effective_page_limit - current_height)
            if ENABLE_INSERT_SPACER_ROWS:
                if remaining <= MIN_SPACER_ROW_HEIGHT:
                    n_rows = 1
                    each_height = DEFAULT_SPACER_ROW_HEIGHT
                else:
                    n_rows = max(1, int(math.ceil(remaining / DEFAULT_SPACER_ROW_HEIGHT)))
                    each_height = max(MIN_SPACER_ROW_HEIGHT, min(MAX_SPACER_ROW_HEIGHT, remaining / n_rows))
            else:
                n_rows = 0
                each_height = 0

            insert_at = s + shift
            insertions.append({
                "insert_at": insert_at,
                "rows": n_rows,
                "height": each_height,
                "before_original_start": s,
                "before_original_end": e,
                "block_kind": kind,
                "old_page": page_no,
                "remaining": remaining,
            })
            shift += n_rows
            page_no += 1
            current_height = repeat_header_height if ENABLE_REPEAT_TITLE_ROWS else 0.0

        if block_too_tall:
            log(f"    超页块 R{s}:R{e} 高度 {block_height:.1f}，按独占页处理，不向后续记录累计高度")
            current_height = 0.0
            page_no += 1
        else:
            current_height += block_height

    return insertions, warnings_too_tall, available_height, merged_blocks


def apply_page_layout_no_split(ws, excel, used_rows, max_col, mode=None):
    """
    按 Excel 真实打印分页排版。
    先让 Excel 根据 A3 横向、页边距、缩放和打印区域计算分页线；只有分页线
    落在合并记录内部时，才添加手动分页符并将整条记录移到下一页。
    """
    if not ENABLE_AUTO_PAGE_LAYOUT:
        log("  已跳过自动分页排版：ENABLE_AUTO_PAGE_LAYOUT=False")
        return {"inserted_rows": 0, "insertions": 0, "warnings": 0}

    try:
        ws.ResetAllPageBreaks()
        log("  已清除旧手动分页符，正在读取 Excel 的 A3 横向真实分页线")
    except Exception as e:
        log(f"  无法刷新 Excel 打印分页线：{e}")

    if mode == "daogang":
        unmerge_daogang_sidebar(ws, used_rows)
        # 长表不能只依赖 Excel COM 当前返回的部分自动分页线。插入大量续行后，
        # Excel 有时只暴露前几处分隔，后半段仍会在打印时临时分页并切开记录。
        # 因此到岗到位版统一按最终行高、缩放和序号记录块计算完整分页方案。
        forced_break_rows = set()
        planned_break_rows = []
        planned_added = 0
        final_break_rows = []
        misaligned = []
        for repair_round in range(1, 8):
            planned_break_rows = compute_scaled_record_page_break_rows(
                ws,
                excel,
                used_rows,
                max_col,
                forced_break_rows=forced_break_rows,
            )
            planned_break_rows = sorted(set(planned_break_rows))
            try:
                ws.ResetAllPageBreaks()
            except Exception:
                pass
            planned_added = apply_manual_page_break_rows(ws, planned_break_rows, used_rows)
            final_break_rows, misaligned = get_record_page_break_misalignments(ws, used_rows)
            if not misaligned:
                break

            repair_logs = []
            for item in misaligned:
                # Excel 比估算提前分页时，固定当前记录起始行为新页开头；
                # 下一轮从头重新装页，后续分页点也随之重新计算。
                forced_break_rows.add(item["record_start"])
                repair_logs.append(
                    f"自动分页R{item['break_row']}切入记录R{item['record_start']}:R{item['record_end']}，"
                    f"固定在R{item['record_start']}前分页并重排后续页面"
                )

            log(
                f"  到岗到位分页第 {repair_round} 轮校准："
                + "；".join(repair_logs)
            )
            if not repair_logs:
                break

        if misaligned:
            details = "；".join(
                f"R{item['break_row']}位于记录R{item['record_start']}:R{item['record_end']}内"
                for item in misaligned
            )
            log_red(f"  到岗到位完整分页方案仍有错位：{details}")
        log(
            f"  到岗到位完整分页排版：计划分页行 {planned_break_rows or '无'}，"
            f"成功写入 {planned_added} 个，最终分页符 {final_break_rows or '无'}"
        )
        return {
            "inserted_rows": 0,
            "insertions": planned_added,
            "warnings": len(misaligned),
            "planned_break_rows": planned_break_rows,
        }

    blocks, _, _ = build_no_split_blocks(ws, used_rows, max_col)
    manual_rows = set()
    warnings = []

    for _ in range(7):
        actual_break_rows = get_page_break_rows(ws, used_rows)
        move_before_rows = set()
        for break_row in actual_break_rows:
            for block in blocks:
                if block["start"] < break_row <= block["end"]:
                    if block["start"] <= DATA_START_ROW:
                        warnings.append((break_row, block["start"], block["end"]))
                    else:
                        move_before_rows.add(block["start"])
                    break

        new_rows = sorted(move_before_rows - manual_rows)
        if not new_rows:
            break

        for row in new_rows:
            try:
                ws.HPageBreaks.Add(Before=ws.Rows(row))
                manual_rows.add(row)
                log(f"  Excel 真实分页线落在合并记录内，已将分页前移到第 {row} 行前")
            except Exception as e:
                log(f"  添加手动分页符失败：before_row={row}, err={e}")
    final_break_rows = get_page_break_rows(ws, used_rows)
    for break_row, start, end in warnings:
        log(f"  警告：Excel 在第 {break_row} 行切开首条合并记录 {start}-{end}，该记录无法与首页标题同时容纳")
    log(
        f"  自动分页排版完成：基于 Excel A3 横向真实分页线 {len(final_break_rows)} 条，"
        f"为保护合并记录新增手动分页符 {len(manual_rows)} 个"
    )
    return {"inserted_rows": 0, "insertions": len(manual_rows), "warnings": len(warnings)}


def get_page_break_rows(ws, used_rows):
    """读取脚本已添加的水平分页符，返回每个新页的起始行。"""
    break_rows = []
    try:
        for index in range(1, ws.HPageBreaks.Count + 1):
            row = int(ws.HPageBreaks(index).Location.Row)
            if DATA_START_ROW < row <= used_rows and row not in break_rows:
                break_rows.append(row)
    except Exception as e:
        log(f"  读取分页符位置失败，将按单页处理：{e}")
    return sorted(break_rows)


def unmerge_daogang_sidebar(ws, used_rows):
    """只解除 L/M 数据区的合并，避免重复运行时旧首页合并干扰新的分页。"""
    seen = set()
    for r in range(DATA_START_ROW, used_rows + 1):
        for c in (12, 13):
            try:
                cell = ws.Cells(r, c)
                if not cell.MergeCells:
                    continue
                area = cell.MergeArea
                address = str(area.Address)
                if address in seen:
                    continue
                seen.add(address)
                area.UnMerge()
            except Exception:
                pass


def apply_daogang_font_layout(ws, used_rows, max_col=20):
    """到岗到位专属字体版式：第2行起 A:T 字号45，第2行行高204。"""
    if used_rows < 2:
        return
    try:
        rng = ws.Range(ws.Cells(2, 1), ws.Cells(used_rows, max_col))
        rng.Font.Size = DAOGANG_FONT_SIZE
        ws.Rows(2).RowHeight = DAOGANG_HEADER_ROW_HEIGHT
        log(
            f"  到岗到位字体版式：R2C1:R{used_rows}C{max_col} 字号={DAOGANG_FONT_SIZE}，"
            f"第2行行高={DAOGANG_HEADER_ROW_HEIGHT}"
        )
    except Exception as e:
        log(f"  到岗到位字体版式设置失败：{e}")


def get_text_max_line_length(text):
    lines = str(text or "").replace("\r", "\n").split("\n")
    return max((len(line.strip()) for line in lines if line.strip()), default=0)


def count_nonempty_text_lines(text):
    lines = str(text or "").replace("\r", "\n").split("\n")
    return sum(1 for line in lines if line.strip())


def normalize_personnel_multiline_text(text):
    """
    人员/职务类单元格统一为单元格内换行。
    保留顿号连接的人名组合，只把斜杠、分号、手机号后的分隔符等人工格式差异拉齐。
    """
    if is_blank_value(text):
        return ""
    s = str(text).replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("／", "/").replace("\\", "/")
    s = re.sub(r"\s*/\s*", "\n", s)
    s = re.sub(r"\s*[；;]\s*", "\n", s)
    s = re.sub(r"(\d{7,})\s*[、,，]\s*", r"\1\n", s)
    lines = []
    for line in s.split("\n"):
        cleaned = re.sub(r"\s+", "", line).strip("、,，；; ")
        if cleaned:
            lines.append(cleaned)
    return "\n".join(lines)


PHONE_TOKEN_PATTERN = re.compile(
    r"(?<!\d)(?:\+?86[- ]?)?(?:1[3-9]\d{9}|\d{7,12})(?!\d)"
)


def is_phone_token(text):
    normalized = re.sub(r"[\s-]+", "", str(text or ""))
    normalized = re.sub(r"^\+?86", "", normalized)
    return bool(re.fullmatch(r"(?:1[3-9]\d{9}|\d{7,12})", normalized))


def normalize_name_phone_multiline_text(text):
    """N/O列自动识别姓名和手机号，统一为“姓名在上、手机号在下”。"""
    if is_blank_value(text):
        return ""

    raw = str(text).replace("\r\n", "\n").replace("\r", "\n")
    raw = raw.replace("／", "/").replace("\\", "/")
    raw = re.sub(r"[\n/；;、,，]+", "\n", raw)

    tokens = []
    for chunk in raw.split("\n"):
        chunk = chunk.strip()
        if not chunk:
            continue
        pos = 0
        for match in PHONE_TOKEN_PATTERN.finditer(chunk):
            before = re.sub(r"\s+", "", chunk[pos:match.start()]).strip("、,，；;:： ")
            if before:
                tokens.append(before)
            phone = re.sub(r"\s+", "", match.group(0)).strip()
            if phone:
                tokens.append(phone)
            pos = match.end()
        tail = re.sub(r"\s+", "", chunk[pos:]).strip("、,，；;:： ")
        if tail:
            tokens.append(tail)

    # 若原文写成“手机号在前、姓名在后”，自动交换为“姓名在上、手机号在下”。
    ordered = []
    i = 0
    while i < len(tokens):
        current = tokens[i]
        if (
            is_phone_token(current)
            and i + 1 < len(tokens)
            and not is_phone_token(tokens[i + 1])
        ):
            ordered.append(tokens[i + 1])
            ordered.append(current)
            i += 2
            continue
        ordered.append(current)
        i += 1

    result = []
    for token in ordered:
        token = token.strip()
        if token and (not result or token != result[-1]):
            result.append(token)
    return "\n".join(result)


def apply_personnel_multiline_layout(ws, used_rows, mode=None):
    """到岗到位N/O整理姓名+手机号为上下换行；P列保持原结构并允许换行。"""
    if used_rows < 2 or mode != "daogang":
        return

    changed = 0
    adjusted_cols = []

    # N/O：自动判断姓名、手机号，统一姓名在上、手机号在下。
    for col in (14, 15):
        col_changed = 0
        longest_line = 0
        try:
            for row in range(DATA_START_ROW, used_rows + 1):
                cell = ws.Cells(row, col)
                value = cell.Value
                if is_blank_value(value) or not is_merge_anchor(cell):
                    continue
                normalized = normalize_name_phone_multiline_text(value)
                current = str(value).replace("\r\n", "\n").replace("\r", "\n")
                if normalized and current != normalized:
                    cell.Value = normalized
                    col_changed += 1
                longest_line = max(
                    longest_line,
                    get_text_max_line_length(normalized or value),
                )

            rng = ws.Range(ws.Cells(BORDER_START_ROW, col), ws.Cells(used_rows, col))
            rng.HorizontalAlignment = XL_CENTER
            rng.VerticalAlignment = XL_CENTER_VERTICAL
            rng.WrapText = True
            ws.Columns(col).AutoFit()
            autofit_width = safe_float(ws.Columns(col).ColumnWidth, 0.0)
            padded_width = max(
                autofit_width * PERSONNEL_COLUMN_WIDTH_PADDING_RATIO,
                autofit_width + PERSONNEL_COLUMN_WIDTH_PADDING_MIN,
                PERSONNEL_DAOGANG_MIN_COLUMN_WIDTH,
            )
            final_width = round(
                min(PERSONNEL_DAOGANG_MAX_COLUMN_WIDTH, padded_width),
                1,
            )
            ws.Columns(col).ColumnWidth = final_width
            adjusted_cols.append(
                f"{column_number_to_name(col)}列={final_width}"
                f"（最长显示行 {longest_line} 字）"
            )
            changed += col_changed
        except Exception as e:
            log(f"  N/O姓名手机号换行整理失败：列 {col}，{e}")

    # P：不改任何单元格文字，不把原换行合并成顿号，也不强制拉成单行。
    if used_rows >= BORDER_START_ROW:
        try:
            p_range = ws.Range(ws.Cells(BORDER_START_ROW, 16), ws.Cells(used_rows, 16))
            p_range.HorizontalAlignment = XL_CENTER
            p_range.VerticalAlignment = XL_CENTER_VERTICAL
            p_range.WrapText = True
            log("  P列保持原内容结构，启用自动换行，不执行单行压缩或内容重写")
        except Exception as e:
            log(f"  P列保持结构/自动换行设置失败：{e}")

    if adjusted_cols:
        log(
            f"  N/O姓名手机号上下换行：更新单元格 {changed} 个；"
            f"列宽 {'，'.join(adjusted_cols)}"
        )


def apply_daogang_sequence_column_layout(ws, used_rows):
    """到岗到位45号字下为序号列保留足够宽度，避免数字和表头贴框。"""
    if used_rows < BORDER_START_ROW:
        return
    try:
        if used_rows >= DATA_START_ROW:
            data_range = ws.Range(
                ws.Cells(DATA_START_ROW, 1),
                ws.Cells(used_rows, 1),
            )
            data_range.WrapText = False
            data_range.HorizontalAlignment = XL_CENTER
            data_range.VerticalAlignment = XL_CENTER_VERTICAL
        ws.Cells(BORDER_START_ROW, 1).WrapText = False
        ws.Columns(1).AutoFit()
        autofit_width = safe_float(ws.Columns(1).ColumnWidth, 0.0)
        final_width = round(min(
            DAOGANG_SEQUENCE_MAX_COLUMN_WIDTH,
            max(
                DAOGANG_SEQUENCE_MIN_COLUMN_WIDTH,
                autofit_width * DAOGANG_SEQUENCE_WIDTH_PADDING_RATIO,
                autofit_width + 2.0,
            ),
        ), 1)
        ws.Columns(1).ColumnWidth = final_width
        log(
            f"  到岗到位序号列宽：A列={final_width}"
            f"（AutoFit {round(autofit_width, 1)}，正文单行显示）"
        )
    except Exception as e:
        log(f"  到岗到位序号列宽设置失败：{e}")


def align_daogang_sidebar_display_text(operations_text, leader_text):
    """
    当 L 列为两行专业、M 列为一行领导时，给 L 列前面补一个换行。
    这等价于用户在单元格开头按 Alt+Enter，让两列文字视觉高度看齐。
    """
    operations_lines = count_nonempty_text_lines(operations_text)
    leader_lines = count_nonempty_text_lines(leader_text)
    if operations_text and leader_text and operations_lines >= 2 and leader_lines == 1:
        return "\n" + operations_text, True
    return operations_text, False


def compute_daogang_sidebar_column_width(ws, sidebar=None):
    texts = []
    for col in (12, 13):
        try:
            texts.append(ws.Cells(2, col).Value)
        except Exception:
            pass
    if sidebar:
        texts.extend([
            sidebar.get("operations", ""),
            sidebar.get("leader", ""),
        ])

    longest = max((get_text_max_line_length(text) for text in texts), default=0)
    width = longest * DAOGANG_SIDEBAR_WIDTH_PER_CHAR + DAOGANG_SIDEBAR_WIDTH_PADDING
    width = max(DAOGANG_SIDEBAR_MIN_COLUMN_WIDTH, width)
    width = min(DAOGANG_SIDEBAR_MAX_COLUMN_WIDTH, width)
    return round(width, 1), longest


def apply_daogang_sidebar_column_layout(ws, used_rows, sidebar=None):
    """L/M 新增列按内容估算同宽，并统一为居中、垂直居中、自动换行。"""
    if used_rows < 2:
        return
    try:
        width, longest = compute_daogang_sidebar_column_width(ws, sidebar)
        ws.Columns(12).ColumnWidth = width
        ws.Columns(13).ColumnWidth = width
        rng = ws.Range(ws.Cells(2, 12), ws.Cells(used_rows, 13))
        rng.WrapText = True
        rng.HorizontalAlignment = XL_CENTER
        rng.VerticalAlignment = XL_CENTER_VERTICAL
        rng.Font.Size = DAOGANG_FONT_SIZE
        log(
            f"  到岗到位 L/M 列版式：两列同宽 {width}，"
            f"最长文本行 {longest} 字，居中/垂直居中/自动换行"
        )
    except Exception as e:
        log(f"  到岗到位 L/M 列版式设置失败：{e}")


def apply_daogang_final_print_style(ws, used_rows, max_col=20):
    """到岗到位最终打印版：正文白底；风险黄色保留；原字体颜色不改。"""
    if used_rows < DATA_START_ROW:
        return

    preserved_risk_yellow = []
    risk_col = find_header_column(ws, max_col, "作业风险等级")
    if risk_col:
        for row in range(DATA_START_ROW, used_rows + 1):
            try:
                cell = ws.Cells(row, risk_col)
                if not is_merge_anchor(cell):
                    continue
                risk_text = re.sub(r"\s+", "", str(cell.Value or ""))
                if not is_level_two_or_three_risk(risk_text):
                    continue
                target = cell.MergeArea if cell.MergeCells else cell
                if is_yellow_interior(target):
                    preserved_risk_yellow.append(target)
            except Exception:
                continue

    try:
        data_range = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(used_rows, max_col))
        data_range.Interior.Pattern = XL_SOLID
        data_range.Interior.Color = XL_COLOR_WHITE
        data_range.Font.Size = DAOGANG_FONT_SIZE
        data_range.WrapText = True
        data_range.VerticalAlignment = XL_CENTER_VERTICAL

        # N/O 保持“姓名在上、手机号在下”；P 保持原结构，三列都允许换行。
        if max_col >= 14:
            personnel_end_col = min(16, max_col)
            ws.Range(
                ws.Cells(DATA_START_ROW, 14),
                ws.Cells(used_rows, personnel_end_col),
            ).WrapText = True

        try:
            base_font_name = str(ws.Cells(DATA_START_ROW, 1).Font.Name or "").strip()
            if base_font_name:
                data_range.Font.Name = base_font_name
        except Exception:
            pass

        for target in preserved_risk_yellow:
            try:
                target.Interior.Pattern = XL_SOLID
                target.Interior.Color = XL_COLOR_YELLOW
            except Exception:
                pass

        log(
            f"  到岗到位最终打印样式：有效数据 R{DATA_START_ROW}C1:R{used_rows}C{max_col} "
            "正文已统一白底，营销绿色取消；原字体颜色保留；N/O/P允许换行"
            + (
                f"；已保留 {len(preserved_risk_yellow)} 个原有二级/三级风险黄色单元格"
                if preserved_risk_yellow
                else ""
            )
        )
    except Exception as e:
        log(f"  到岗到位最终打印样式设置失败：{e}")


def clear_daogang_sidebar_horizontal_borders(ws, used_rows):
    """L/M 侧栏只保留顶部连接和整表末行封底，分页处及中间不留横线。"""
    if used_rows < DATA_START_ROW:
        return
    cleared = 0
    for r in range(DATA_START_ROW, used_rows + 1):
        try:
            row_range = ws.Range(ws.Cells(r, 12), ws.Cells(r, 13))
            row_range.Borders(XL_EDGE_TOP).LineStyle = XL_NONE
            row_range.Borders(XL_EDGE_BOTTOM).LineStyle = XL_NONE
            cleared += 1
        except Exception:
            pass

    try:
        top_range = ws.Range(ws.Cells(DATA_START_ROW, 12), ws.Cells(DATA_START_ROW, 13))
        top_range.Borders(XL_EDGE_TOP).LineStyle = XL_CONTINUOUS
        top_range.Borders(XL_EDGE_TOP).Weight = XL_THIN
        top_range.Borders(XL_EDGE_TOP).Color = XL_COLOR_BLACK
    except Exception:
        pass
    try:
        tail_range = ws.Range(ws.Cells(used_rows, 12), ws.Cells(used_rows, 13))
        tail_range.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        tail_range.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
        tail_range.Borders(XL_EDGE_BOTTOM).Color = XL_COLOR_BLACK
    except Exception:
        pass
    log(f"  到岗到位侧栏横线收口：已清理 L/M 数据区横线 {cleared} 行，仅保留顶部和末行封底")


def apply_daogang_sidebar(ws, used_rows, sidebar):
    """
    处理到岗到位 L/M 两列的纸面版式：
    - 首页数据区合并并居中写入；
    - 后续页面保持空白；
    - 数据区不画内部横线，只保留左右竖线；
    - 仅首页顶部和整份表末行绘制封口横线。
    """
    if used_rows < DATA_START_ROW:
        return

    operations_text = str(sidebar.get("operations", "") or "").strip()
    leader_text = str(sidebar.get("leader", "") or "").strip()
    operations_display_text, added_visual_blank = align_daogang_sidebar_display_text(operations_text, leader_text)
    page_break_rows = get_page_break_rows(ws, used_rows)
    first_page_end = (page_break_rows[0] - 1) if page_break_rows else used_rows
    first_page_end = max(DATA_START_ROW, min(first_page_end, used_rows))

    unmerge_daogang_sidebar(ws, used_rows)
    data_range = ws.Range(ws.Cells(DATA_START_ROW, 12), ws.Cells(used_rows, 13))
    try:
        data_range.ClearContents()
        data_range.Borders.LineStyle = XL_NONE
        data_range.Font.Size = DAOGANG_FONT_SIZE
        data_range.WrapText = True
        data_range.HorizontalAlignment = XL_CENTER
        data_range.VerticalAlignment = XL_CENTER_VERTICAL
    except Exception:
        pass

    # 全程只保留侧栏的三条竖线：L 左边、L/M 中线、M 右边。
    for border_id in (XL_EDGE_LEFT, XL_EDGE_RIGHT, XL_INSIDE_VERTICAL):
        try:
            data_range.Borders(border_id).LineStyle = XL_CONTINUOUS
            data_range.Borders(border_id).Weight = XL_THIN
            data_range.Borders(border_id).Color = XL_COLOR_BLACK
        except Exception:
            pass

    # 首页与表头正常连接；整份表只在最后一行封底，不在分页处封口。
    for row, border_id in ((DATA_START_ROW, XL_EDGE_TOP), (used_rows, XL_EDGE_BOTTOM)):
        try:
            border_range = ws.Range(ws.Cells(row, 12), ws.Cells(row, 13))
            border_range.Borders(border_id).LineStyle = XL_CONTINUOUS
            border_range.Borders(border_id).Weight = XL_THIN
            border_range.Borders(border_id).Color = XL_COLOR_BLACK
        except Exception:
            pass

    # 首页的两个侧栏分别合并；之后各页不合并、不重复文字。
    for col, value in ((12, operations_display_text), (13, leader_text)):
        try:
            page_one_cell = ws.Range(ws.Cells(DATA_START_ROW, col), ws.Cells(first_page_end, col))
            page_one_cell.Merge()
            page_one_cell = ws.Cells(DATA_START_ROW, col).MergeArea
            page_one_cell.Cells(1, 1).Value = value
            page_one_cell.WrapText = True
            page_one_cell.HorizontalAlignment = XL_CENTER
            page_one_cell.VerticalAlignment = XL_CENTER_VERTICAL
            page_one_cell.Font.Size = DAOGANG_FONT_SIZE
            page_one_cell.Borders(XL_EDGE_LEFT).LineStyle = XL_CONTINUOUS
            page_one_cell.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
            page_one_cell.Borders(XL_EDGE_TOP).LineStyle = XL_CONTINUOUS
            page_one_cell.Borders(XL_EDGE_LEFT).Weight = XL_THIN
            page_one_cell.Borders(XL_EDGE_RIGHT).Weight = XL_THIN
            page_one_cell.Borders(XL_EDGE_TOP).Weight = XL_THIN
            page_one_cell.Borders(XL_EDGE_LEFT).Color = XL_COLOR_BLACK
            page_one_cell.Borders(XL_EDGE_RIGHT).Color = XL_COLOR_BLACK
            page_one_cell.Borders(XL_EDGE_TOP).Color = XL_COLOR_BLACK
            if first_page_end == used_rows:
                page_one_cell.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
                page_one_cell.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
                page_one_cell.Borders(XL_EDGE_BOTTOM).Color = XL_COLOR_BLACK
            else:
                page_one_cell.Borders(XL_EDGE_BOTTOM).LineStyle = XL_NONE
        except Exception as e:
            log(f"  到岗到位首页侧栏合并失败：列 {col}，{e}")

    # 合并后再次确保后续页不出现横向封口，整份表只在最后一行封底。
    try:
        data_range.Borders(XL_INSIDE_HORIZONTAL).LineStyle = XL_NONE
        tail_range = ws.Range(ws.Cells(used_rows, 12), ws.Cells(used_rows, 13))
        tail_range.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        tail_range.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
        tail_range.Borders(XL_EDGE_BOTTOM).Color = XL_COLOR_BLACK
    except Exception:
        pass
    clear_daogang_sidebar_horizontal_borders(ws, used_rows)

    log(
        f"  到岗到位侧栏已完成：首页合并行 {DATA_START_ROW}-{first_page_end}，"
        f"后续页留空，分页处不封口，末行 {used_rows} 已封底；分页符 {page_break_rows or '无'}"
        + ("；L列已自动补首行空行用于视觉对齐" if added_visual_blank else "")
    )


def apply_daogang_sidebar_with_stable_page_breaks(ws, excel, used_rows, max_col, sidebar):
    """最终阶段稳定分页线后再写 L/M 侧栏。"""
    setup_print_page(ws, excel, used_rows, max_col)
    before_rows = get_page_break_rows(ws, used_rows)
    log(f"  到岗到位侧栏分页稳定检查：分页符 {before_rows or '无'}")
    apply_daogang_sidebar(ws, used_rows, sidebar)

    after_rows = get_page_break_rows(ws, used_rows)
    if after_rows != before_rows:
        log(f"  到岗到位侧栏写入后分页符变化为 {after_rows or '无'}，按新分页符重写侧栏")
        apply_daogang_sidebar(ws, used_rows, sidebar)


# ================= Excel 预处理 =================
def detect_processing_mode(title, filename=""):
    text = f"{title} {filename}"
    if "到岗到位" in text:
        return "daogang"
    if "风险管控" in text:
        return "risk"
    if "现场作业计划" in text:
        return "site"
    return "normal"


def get_business_mode_label(mode, daogang_sidebar=None):
    """业务阶段命名：模式A=首轮处理，模式B=到岗到位生成后处理。"""
    if mode == "daogang" or daogang_sidebar is not None:
        return "模式B（到岗到位文件生成后处理）"
    return "模式A（原始文件首轮处理）"


PLAN_DATE_PATTERN = re.compile(r"(?<!\d)(\d{1,2})月(\d{1,2})日")
PLAN_DATE_ACCEPTED_STATUSES = {"tomorrow", "weekend_advance"}


def extract_plan_month_day(text):
    match = PLAN_DATE_PATTERN.search(str(text or ""))
    if not match:
        return None
    month = int(match.group(1))
    day = int(match.group(2))
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return month, day


def resolve_plan_date(month_day, today):
    """把不含年份的“X月X日”解析为最接近本机日期的实际日期。"""
    if not month_day:
        return None
    month, day = month_day
    candidates = []
    for year in (today.year - 1, today.year, today.year + 1):
        try:
            candidates.append(today.replace(year=year, month=month, day=day))
        except ValueError:
            continue
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda item: (abs((item - today).days), (item - today).days < 0),
    )


def classify_plan_date(plan_date, today):
    if plan_date is None:
        return "missing"
    delta_days = (plan_date - today).days
    if delta_days == 1:
        return "tomorrow"
    # 周末批量计划例外：允许提前最多4天制作星期六、星期日或星期一计划。
    if 1 <= delta_days <= 4 and plan_date.weekday() in {5, 6, 0}:
        return "weekend_advance"
    if delta_days == 0:
        return "same_day"
    if delta_days < 0:
        return "past"
    return "future_nonstandard"


def format_plan_date_cn(plan_date):
    if plan_date is None:
        return "未识别"
    weekday_names = "一二三四五六日"
    return (
        f"{plan_date.month}月{plan_date.day}日"
        f"（星期{weekday_names[plan_date.weekday()]}）"
    )


def replace_title_plan_date(title, plan_date):
    replacement = f"{plan_date.month}月{plan_date.day}日"
    if PLAN_DATE_PATTERN.search(str(title or "")):
        return PLAN_DATE_PATTERN.sub(replacement, str(title), count=1)
    return str(title or "")


def evaluate_plan_date_consistency(filename, title, today=None):
    """
    评估文件名、表内标题和本机日期的关系。

    常规计划日期应为本机日期+1；未来4天内的星期六、星期日、星期一允许作为
    周末提前计划。函数只给出建议，不直接修改文件。
    """
    today = today or datetime.now().date()
    file_date = resolve_plan_date(extract_plan_month_day(filename), today)
    title_date = resolve_plan_date(extract_plan_month_day(title), today)
    file_status = classify_plan_date(file_date, today)
    title_status = classify_plan_date(title_date, today)
    dates_match = bool(file_date and title_date and file_date == title_date)
    suggested_date = None
    reason = ""

    if file_date and title_date and not dates_match:
        if file_status in PLAN_DATE_ACCEPTED_STATUSES:
            suggested_date = file_date
            reason = "文件名日期符合计划日期规则，表内标题日期不一致"
        elif title_status in PLAN_DATE_ACCEPTED_STATUSES:
            reason = "表内标题日期合理，但文件名日期不一致"
        elif file_date > today and title_date <= today:
            suggested_date = file_date
            reason = "文件名为未来日期，表内标题为当天或历史日期"
        else:
            reason = "文件名日期与表内标题日期不一致，且无法仅凭本机日期唯一判断"
    elif title_date:
        if title_status == "same_day":
            suggested_date = today + timedelta(days=1)
            reason = "表内标题为本机当天，不符合次日计划规则"
        elif title_status == "future_nonstandard":
            suggested_date = today + timedelta(days=1)
            reason = "表内标题不是次日计划，也不属于周末提前计划例外"
        elif title_status == "past":
            reason = "表内标题早于本机日期，按历史文件提示，不自动建议修改"
    elif file_date and file_status in PLAN_DATE_ACCEPTED_STATUSES:
        suggested_date = file_date
        reason = "表内标题未识别到日期，文件名日期符合计划日期规则"
    else:
        reason = "文件名和表内标题均未识别到可校验的计划日期"

    if dates_match and title_status == "tomorrow":
        reason = "文件名与表内标题一致，符合本机次日计划规则"
    elif dates_match and title_status == "weekend_advance":
        reason = "文件名与表内标题一致，符合周末/周一提前计划例外"

    needs_title_change = bool(
        suggested_date
        and title_date != suggested_date
        and extract_plan_month_day(title)
    )
    return {
        "today": today,
        "file_date": file_date,
        "title_date": title_date,
        "file_status": file_status,
        "title_status": title_status,
        "dates_match": dates_match,
        "suggested_date": suggested_date,
        "needs_title_change": needs_title_change,
        "reason": reason,
    }


def audit_plan_date_and_prompt_title_correction(ws, path, title):
    result = evaluate_plan_date_consistency(os.path.basename(path), title)
    today = result["today"]
    file_text = format_plan_date_cn(result["file_date"])
    title_text = format_plan_date_cn(result["title_date"])
    suggested = result["suggested_date"]
    status = result["title_status"]

    summary = (
        f"本机日期 {today:%Y-%m-%d}；文件名日期 {file_text}；"
        f"表内标题日期 {title_text}；{result['reason']}"
    )
    if result["dates_match"] and status in PLAN_DATE_ACCEPTED_STATUSES:
        log(f"  计划日期校验：{summary}")
        return title

    if status == "past" and not result["needs_title_change"]:
        log(f"  计划日期提示：{summary}")
        return title

    log_red(f"  【计划日期提醒】{summary}")
    if not result["needs_title_change"]:
        log("  未形成唯一且可执行的标题修改建议，已保留原日期")
        return title

    suggested_text = format_plan_date_cn(suggested)
    prompt = (
        "发现现场作业计划日期可能有误。\n\n"
        f"{summary}\n\n"
        f"建议把表内标题修改为：{suggested_text}\n"
        "是否立即修改表内标题？"
    )
    if not ask_yes_no_dialog("现场作业计划日期校验", prompt, default=False):
        log("  用户选择不修改表内标题日期，已保留原内容")
        return title

    corrected_title = replace_title_plan_date(title, suggested)
    try:
        ws.Cells(1, 1).Value = corrected_title
        log_red(f"  【计划日期已修改】{title} -> {corrected_title}")
        return corrected_title
    except Exception as e:
        log_red(f"  修改表内标题日期失败，已保留原内容：{e}")
        return title


def preprocess_excel(path, cwd, mode=None, daogang_sidebar=None):
    log_section(f"开始处理 Excel 文件：{os.path.basename(path)}")
    log(f"完整路径：{path}")
    abs_path = os.path.abspath(path)

    if not os.path.exists(abs_path):
        log("  文件不存在！")
        return "文件不存在"

    if is_backup_path(abs_path):
        log("  检测到自动备份目录中的文件：只读保留，跳过全部预处理")
        return "备份文件已跳过"

    if is_file_locked(abs_path):
        log("  警告：文件正被其他程序占用，跳过处理！")
        return "文件被占用"

    backup_file(abs_path, cwd)

    try:
        os.chmod(abs_path, 0o666)
        log("  已移除文件只读属性")
    except Exception as e:
        log(f"  警告：无法移除只读属性，跳过处理: {e}")
        return "只读无法修改"

    excel = None
    wb = None
    title = "未知标题"
    try:
        timer_open = step_timer("启动并打开工作簿")
        log("  正在启动独立 Excel 实例...")
        excel = create_excel_application()
        configure_excel_silent(excel)

        log("  正在打开工作簿...")
        wb = excel.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )
        configure_excel_silent(excel, wb)
        log("  工作簿已打开")
        step_done("启动并打开工作簿", timer_open)
        ws = wb.Worksheets(1)

        # ===== 唯一新增前置任务：新版现场作业计划 B 列新增“填报人”，进入原处理流程前先删除 =====
        # 必须在 hide_helper_column() 之前执行，避免删除 B 列后辅助列整体左移。
        try:
            raw_title = str(ws.Cells(1, 1).Value).strip() if ws.Cells(1, 1).Value else ""
            is_site_plan = (
                "现场作业计划" in raw_title
                or "现场作业计划" in os.path.basename(path)
            )
            if is_site_plan:
                hv = ws.Cells(2, 2).Value
                if hv is not None and "填报人" in str(hv):
                    ws.Columns(2).Delete()
                    log(f"  前置任务：现场作业计划 B 列检测到“{str(hv).strip()}”，已删除填报人列")
                else:
                    log("  前置任务：现场作业计划 B 列未检测到“填报人”，按原流程继续处理")
        except Exception as e:
            log(f"  警告：删除现场作业计划“填报人”列时出错: {e}")
        # ===== 前置任务结束 =====

        hide_helper_column(ws)

        title = str(ws.Cells(1, 1).Value).strip() if ws.Cells(1, 1).Value else "未知标题"
        if (
            mode != "daogang"
            and (
                "现场作业计划" in title
                or "现场作业计划" in os.path.basename(path)
                or "风险管控" in title
                or "风险管控" in os.path.basename(path)
            )
        ):
            title = audit_plan_date_and_prompt_title_correction(ws, abs_path, title)
        if mode is None:
            mode = detect_processing_mode(title, os.path.basename(path))
        max_col = 20 if mode == "daogang" else (19 if "现场作业计划" in title else 12)
        print_max_col = get_print_max_col(max_col, mode=mode)
        autofit_max_col = print_max_col if mode == "daogang" else max_col
        autofit_skip_columns = {12, 13} if mode == "daogang" else set()
        autofit_measure_columns = DAOGANG_AUTOFIT_MEASURE_COLUMNS if mode == "daogang" else set()
        autofit_overflow_row_capacity_factor = (
            DAOGANG_OVERFLOW_ROW_CAPACITY_FACTOR if mode in {"risk", "daogang"} else 1.0
        )
        business_mode_label = get_business_mode_label(mode, daogang_sidebar)
        log(f"  标题：{title}")
        log(f"  识别最大处理列：{max_col} 列；打印/框线截止列：{print_max_col} 列")
        log(f"  业务模式：{business_mode_label}")
        log(f"  技术处理模式：{mode}（site=首轮轻处理，risk/daogang=允许插行排版）")

        # 删除上次自动分页空白行，保证重复运行不会越插越多
        remove_old_auto_spacer_rows(ws)

        # 删除“例”行（只删除第一处，从第1行开始）
        deleted_example = False
        used_rows = get_used_rows(ws)
        for r in range(1, used_rows + 1):
            try:
                cellv = ws.Cells(r, 1).Value
                if cellv and "例" in str(cellv):
                    ws.Rows(r).Delete()
                    deleted_example = True
                    break
            except Exception:
                continue
        if deleted_example:
            log("  已删除含 '例' 的行")
        else:
            log("  未发现含 '例' 的行")

        # 删除空行（从底向上）
        deleted_rows = 0
        used_rows = get_used_rows(ws)
        for r in range(used_rows, 2, -1):
            if is_auto_spacer_row(ws, r):
                continue
            vals = []
            for c in range(1, max_col + 1):
                try:
                    vals.append(ws.Cells(r, c).Value)
                except Exception:
                    vals.append(None)
            if all(v in [None, ""] for v in vals):
                ws.Rows(r).Delete()
                deleted_rows += 1
        if deleted_rows:
            log(f"  已删除空行 {deleted_rows} 行")
        else:
            log("  未发现需删除的空行")

        # 先约束红字范围、清除删除线，再按每列严格多数派规整基础排版。
        used_rows = get_used_rows(ws)
        timer_font_effects = step_timer("红字范围和删除线预处理")
        normalize_font_effects_for_print(ws, used_rows, max_col, mode=mode)
        step_done("红字范围和删除线预处理", timer_font_effects)

        timer_majority_format = step_timer("按列多数派格式预处理")
        normalize_data_region_majority_format(ws, used_rows, max_col, mode=mode)
        step_done("按列多数派格式预处理", timer_majority_format)

        # 第一次重排序号、行高、边框、页面设置
        apply_personnel_multiline_layout(ws, used_rows, mode=mode)
        if mode == "daogang":
            apply_daogang_font_layout(ws, used_rows, max_col=max_col)
            apply_daogang_sequence_column_layout(ws, used_rows)
            apply_daogang_sidebar_column_layout(ws, used_rows, daogang_sidebar)
        renumber_sequence(ws, used_rows)
        if mode == "site":
            sync_site_sequence_fill(ws, used_rows, source_col=2)
        timer_autofit = step_timer("行高调整/AutoFitRowEx")
        autofit_result = None
        if mode in {"risk", "daogang"}:
            row_extra_height = choose_autofit_plus_height(ws, used_rows, mode=mode)
            log(f"  自动行高余量：+{row_extra_height} points（按最终字号自动选择）")
            autofit_result = autofit_rows(
                ws,
                used_rows,
                plus_height=row_extra_height,
                max_col=autofit_max_col,
                start_row=AUTOFITROWEX_START_ROW,
                skip_columns=autofit_skip_columns,
                measure_columns=autofit_measure_columns,
                overflow_row_capacity_factor=autofit_overflow_row_capacity_factor,
            )
        else:
            log("  普通现场表无需打印：保留原行高，不执行 AutoFit 或打印行高调整")
        step_done("行高调整/AutoFitRowEx", timer_autofit)
        used_rows = get_used_rows(ws)

        inserted_by_autofit = get_autofit_inserted_rows(autofit_result)
        if inserted_by_autofit > 0:
            if mode == "daogang" and not ENABLE_DAOGANG_POST_INSERT_COMPACT_AUTOFIT:
                log(
                    f"  到岗到位已插入续行 {inserted_by_autofit} 行："
                    "跳过插行后全量二次 AutoFitRowEx 测量，沿用首轮测量高度"
                )
            else:
                compact_extra_height = choose_autofit_plus_height(ws, used_rows, mode=mode, compact=True)
                timer_post_insert_autofit = step_timer("插行后紧凑行高复算")
                log(f"  检测到 AutoFitRowEx 已插入续行 {inserted_by_autofit} 行，打印区域设置前复算行高，余量 +{compact_extra_height}")
                autofit_rows(
                    ws,
                    used_rows,
                    plus_height=compact_extra_height,
                    max_col=autofit_max_col,
                    start_row=AUTOFITROWEX_START_ROW,
                    skip_columns=autofit_skip_columns,
                    measure_columns=autofit_measure_columns,
                    overflow_row_capacity_factor=autofit_overflow_row_capacity_factor,
                )
                used_rows = get_used_rows(ws)
                step_done("插行后紧凑行高复算", timer_post_insert_autofit)

        if mode in {"risk", "daogang"}:
            timer_uniform_heights = step_timer("打印表正文行高规整")
            normalize_print_record_heights(ws, used_rows, max_col, mode=mode)
            used_rows = get_used_rows(ws)
            step_done("打印表正文行高规整", timer_uniform_heights)
        else:
            log("  普通现场表无需打印：跳过正文行高和人员列宽规整")

        timer_page_setup_1 = step_timer("首次边框和页面设置")
        if mode == "daogang":
            log("  到岗到位首次框线：延后到分页和侧栏稳定后按记录边界一次性重建")
        else:
            apply_full_borders(ws, used_rows, print_max_col)
        setup_print_page(ws, excel, used_rows, print_max_col)
        step_done("首次边框和页面设置", timer_page_setup_1)

        if mode in {"risk", "daogang"}:
            timer_page_compact = step_timer("节纸行高检查")
            page_compact_result = compact_print_rows_to_save_pages(
                ws,
                excel,
                used_rows,
                mode=mode,
            )
            if page_compact_result.get("applied"):
                setup_print_page(ws, excel, used_rows, print_max_col)
            step_done("节纸行高检查", timer_page_compact)

        # 自动分页排版：先根据真实行高估算，再插入空白行/分页符
        timer_layout = step_timer("读取 Excel 真实分页线并保护合并记录")
        if mode == "risk":
            layout_result = apply_page_layout_no_split(ws, excel, used_rows, print_max_col, mode=mode)
        elif mode == "daogang":
            layout_result = {"inserted_rows": 0, "insertions": 0, "warnings": 0}
            log("  到岗到位完整分页：延后到人员列宽和最终行结构稳定后执行")
        else:
            layout_result = {"inserted_rows": 0, "insertions": 0, "warnings": 0}
            log("  首轮现场作业计划/普通文件：跳过自动分页符排版")
        step_done("读取 Excel 真实分页线并保护合并记录", timer_layout)

        # 插行后重新计算、重排序号、行高、边框、页面设置
        used_rows = get_used_rows(ws)
        if layout_result.get("inserted_rows", 0) > 0:
            log_sub("插行后重新整理")
            renumber_sequence(ws, used_rows)
            if mode == "site":
                sync_site_sequence_fill(ws, used_rows, source_col=2)
            if mode in {"risk", "daogang"}:
                compact_extra_height = choose_autofit_plus_height(ws, used_rows, mode=mode, compact=True)
                log(f"  插行后重新整理采用紧凑行高余量：+{compact_extra_height} points")
                autofit_rows(
                    ws,
                    used_rows,
                    plus_height=compact_extra_height,
                    max_col=autofit_max_col,
                    start_row=AUTOFITROWEX_START_ROW,
                    skip_columns=autofit_skip_columns,
                    measure_columns=autofit_measure_columns,
                    overflow_row_capacity_factor=autofit_overflow_row_capacity_factor,
                )
            else:
                compact_extra_height = choose_autofit_plus_height(ws, used_rows, mode=mode, compact=True)
                simple_autofit_rows(ws, used_rows, plus_height=compact_extra_height, start_row=1)
            used_rows = get_used_rows(ws)

        apply_personnel_multiline_layout(ws, used_rows, mode=mode)

        if mode == "daogang":
            timer_final_layout = step_timer("到岗到位最终列宽后的完整分页")
            setup_print_page(ws, excel, used_rows, print_max_col)
            layout_result = apply_page_layout_no_split(
                ws,
                excel,
                used_rows,
                print_max_col,
                mode=mode,
            )
            step_done("到岗到位最终列宽后的完整分页", timer_final_layout)

        timer_borders = step_timer("最终边框和合并单元格边框修复")
        if mode == "daogang":
            log("  到岗到位最终框线：等待侧栏合并和分页稳定后再按记录边界校准")
        else:
            apply_full_borders(ws, used_rows, print_max_col)
            fix_merged_borders(ws, used_rows, print_max_col)
        step_done("最终边框和合并单元格边框修复", timer_borders)

        # 风险管控打印表正文取消营销绿色等底色；仅二级/三级风险黄色在最终校验中补回。
        if mode == "risk":
            clear_print_data_backgrounds(ws, used_rows, print_max_col, mode=mode)

        if mode == "daogang":
            timer_print_style = step_timer("到岗到位最终打印样式收口")
            apply_daogang_final_print_style(ws, used_rows, max_col=print_max_col)
            step_done("到岗到位最终打印样式收口", timer_print_style)

        if mode == "daogang" and daogang_sidebar is not None:
            timer_sidebar = step_timer("最终页面设置和到岗到位 L/M 侧栏写入")
            apply_daogang_sidebar_with_stable_page_breaks(ws, excel, used_rows, print_max_col, daogang_sidebar)
            step_done("最终页面设置和到岗到位 L/M 侧栏写入", timer_sidebar)
        else:
            timer_page_setup_2 = step_timer("最终页面设置")
            setup_print_page(ws, excel, used_rows, print_max_col)
            step_done("最终页面设置", timer_page_setup_2)

        if mode == "daogang":
            timer_record_borders = step_timer("到岗到位记录边界框线校准")
            apply_record_aware_borders(ws, used_rows, print_max_col, mode=mode)
            step_done("到岗到位记录边界框线校准", timer_record_borders)

        # 风险管控和到岗到位按用户要求全部使用细线。
        if mode in {"risk", "daogang"}:
            log(
                f"  {get_business_mode_label(mode, daogang_sidebar)}框线层次："
                "表头、外框、记录边界和页末统一使用细线"
            )
        else:
            apply_print_border_hierarchy(ws, used_rows, print_max_col)
        if mode == "daogang":
            timer_page_end_borders = step_timer("到岗到位分页末行实体封框")
            apply_daogang_page_end_closure_borders(ws, used_rows, print_max_col)
            step_done("到岗到位分页末行实体封框", timer_page_end_borders)

        if mode in {"site", "risk", "daogang"}:
            timer_level_three_yellow = step_timer("二级/三级风险强制标黄校验")
            enforce_level_three_risk_yellow(
                ws,
                used_rows,
                max_col,
                mode=mode,
                filename=abs_path,
            )
            step_done("二级/三级风险强制标黄校验", timer_level_three_yellow)

        # 汇总报告：分页符数量
        try:
            hp_count = ws.HPageBreaks.Count
        except Exception:
            hp_count = "无法读取"
        log(f"  最终 UsedRows：{used_rows}，最终手动/自动水平分页符数量：{hp_count}")

        # 已经以可写方式打开原文件，优先 Save，避免同路径 SaveAs 触发 Excel
        # 对合并单元格和手动分页符进行额外的格式/打印重算而长时间卡住。
        timer_save = step_timer("保存工作簿")
        try:
            wb.Save()
            log("  已保存文件（Save）")
        except Exception as e:
            log(f"  Save 失败，尝试 SaveAs: {e}")
            try:
                if abs_path.lower().endswith('.xlsm'):
                    wb.SaveAs(abs_path, FileFormat=52)
                else:
                    wb.SaveAs(abs_path, FileFormat=51)
                log("  已保存文件（SaveAs 回退）")
            except Exception as e2:
                log(f"  保存失败: {e2}")
                raise e
        step_done("保存工作簿", timer_save)

        return title

    except Exception as e:
        err_text = str(e)
        if 'CLSIDToPackageMap' in err_text or ('module' in err_text and 'has no attribute' in err_text):
            log(f"检测到win32com缓存错误: {e}")
            return "缓存错误"
        else:
            log(f"预处理 Excel 出错：{e}")
            log("详细错误：")
            log(traceback.format_exc())
            return "处理失败"
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass


def check_n_column_ready_for_daogang(path):
    """
    到岗到位生成前置条件：
    现场作业计划的有效数据行中，N列必须全部非空。
    只要有一个有效行 N 列为空，就不弹出生成询问，只继续汇总。
    """
    try:
        df = pd.read_excel(path, header=None, sheet_name=0, engine="openpyxl", dtype=object)
    except Exception as e:
        log(f"  检测 N 列失败：{e}")
        return False, 0, []

    valid_rows = []
    blank_n_rows = []
    max_check_col = min(df.shape[1], 18)
    for zero_idx, row in df.iloc[2:].iterrows():
        excel_row = int(zero_idx) + 1
        row_values = [row.iloc[i] if i < len(row) else None for i in range(max_check_col)]
        if all(is_blank_value(v) for v in row_values):
            continue
        first_col = row.iloc[0] if len(row) > 0 else None
        if not is_blank_value(first_col) and "例" in str(first_col):
            continue

        valid_rows.append(excel_row)
        n_value = row.iloc[13] if df.shape[1] > 13 and len(row) > 13 else None
        if is_blank_value(n_value):
            blank_n_rows.append(excel_row)

    return bool(valid_rows) and not blank_n_rows, len(valid_rows), blank_n_rows


def ask_yes_no(prompt, default=False):
    if not ENABLE_DAOGANG_VERSION_PROMPT:
        return default
    suffix = "Y/n" if default else "y/N"
    try:
        answer = input(f"{prompt} ({suffix})：").lstrip("\ufeff").strip().lower()
    except Exception:
        log("  当前环境无法交互输入，默认不生成到岗到位版本")
        return default
    if answer == "":
        return default
    return answer in {"y", "yes", "是", "生成", "1"}


def ask_yes_no_dialog(title, prompt, default=False):
    """优先弹出 Windows 是/否对话框；无图形环境时回退到控制台询问。"""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        answer = messagebox.askyesno(
            title,
            prompt,
            parent=root,
            default="yes" if default else "no",
        )
        root.destroy()
        return bool(answer)
    except Exception as e:
        log(f"  Windows确认框不可用，回退到控制台询问：{e}")
        return ask_yes_no(prompt.replace("\n", " "), default=default)


def ask_text_dialog(title, prompt):
    """优先使用 Windows 输入对话框；无法显示窗口时回退到控制台输入。"""
    try:
        if not sys.stdin.isatty():
            return input(f"{prompt}（直接回车表示无人值班）：").strip()
    except Exception:
        pass
    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        value = simpledialog.askstring(title, prompt, parent=root)
        root.destroy()
        return (value or "").strip()
    except Exception as e:
        log(f"  无法显示输入对话框，改用控制台输入：{e}")
        try:
            return input(f"{prompt}（直接回车表示无人值班）：").strip()
        except Exception:
            return ""


def format_company_leader_text(text):
    cleaned = str(text or "").strip()
    if not cleaned:
        return ""
    if cleaned.startswith("公司领导"):
        return cleaned
    return f"公司领导-{cleaned}"


def collect_daogang_sidebar_input():
    """收集首页 L/M 两列内容；空输入不会生成对应专业文字。"""
    operations = normalize_personnel_multiline_text(
        ask_text_dialog("到岗到位 - 运检专业", "请输入运检专业到岗人员：")
    )
    marketing = normalize_personnel_multiline_text(
        ask_text_dialog("到岗到位 - 营销专业", "请输入营销专业到岗人员：")
    )
    leader = normalize_personnel_multiline_text(
        ask_text_dialog("到岗到位 - 公司领导", "请输入公司领导到岗信息：")
    )

    operation_lines = []
    if operations:
        operation_lines.append(f"运检专业：{operations}")
    if marketing:
        operation_lines.append(f"营销专业：{marketing}")

    leader_text = format_company_leader_text(leader)
    result = {"operations": "\n".join(operation_lines), "leader": leader_text}
    log(
        "  已收集到岗到位首页信息："
        f"运检专业={'已填写' if operations else '无人值班'}，"
        f"营销专业={'已填写' if marketing else '无人值班'}，"
        f"公司领导={'已填写：' + leader_text if leader_text else '未填写'}"
    )
    return result


def build_daogang_path(source_path):
    root, ext = os.path.splitext(source_path)
    if root.endswith("（到岗到位）"):
        return source_path
    return f"{root}（到岗到位）{ext}"


def build_daogang_work_path(target_path):
    return f"{target_path}.处理中.xlsx"


def transform_daogang_columns(path):
    log_section(f"生成到岗到位列结构：{os.path.basename(path)}")
    abs_path = os.path.abspath(path)
    excel = None
    wb = None
    try:
        timer_transform = step_timer("到岗到位列结构转换")
        log("  正在启动独立 Excel 实例（到岗到位结构转换）...")
        excel = create_excel_application()
        configure_excel_silent(excel)
        log("  正在打开到岗到位工作簿...")
        wb = excel.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
        configure_excel_silent(excel, wb)
        log("  到岗到位工作簿已打开")
        ws = wb.Worksheets(1)

        # 到岗到位成品结构：删除“公司名称”，再在原“到岗到位人员”前插入两列。
        try:
            if str(ws.Cells(2, 3).Value or "").strip() == "公司名称":
                ws.Columns(3).Delete()
                log("  已删除 C 列：公司名称")
            else:
                log("  C 列不是“公司名称”，未删除，避免误伤列结构")
        except Exception as e:
            log(f"  删除 C 列失败：{e}")

        try:
            header_l = str(ws.Cells(2, 12).Value or "").strip()
            header_m = str(ws.Cells(2, 13).Value or "").strip()
            if "保证体系履责" not in header_l and "公司领导" not in header_m:
                ws.Columns(12).Insert()
                ws.Columns(12).Insert()
                ws.Cells(2, 12).Value = "保证体系履责\n（视频监控中心）"
                ws.Cells(2, 13).Value = "公司领导\n监控中心到岗到位"
                log("  已在 L/M 列插入：保证体系履责、公司领导监控中心到岗到位")
            else:
                log("  已存在到岗到位新增列，未重复插入")
        except Exception as e:
            log(f"  插入到岗到位新增列失败：{e}")

        try:
            if ws.Cells(1, 1).MergeCells:
                ws.Cells(1, 1).MergeArea.UnMerge()
            ws.Range(ws.Cells(1, 1), ws.Cells(1, 20)).Merge()
            ws.Cells(1, 1).HorizontalAlignment = -4108  # xlCenter
        except Exception as e:
            log(f"  标题行合并到 A:T 失败：{e}")
        try:
            used_rows = get_used_rows(ws)
            if ENABLE_REPEAT_TITLE_ROWS:
                ws.PageSetup.PrintTitleRows = REPEAT_TITLE_ROWS_TEXT
            ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(used_rows, PRINT_MAX_COLUMN_S)).Address
        except Exception:
            pass

        wb.Save()
        log("  到岗到位列结构转换已保存")
        step_done("到岗到位列结构转换", timer_transform)
        return True
    except Exception as e:
        log(f"  到岗到位列结构转换失败：{e}")
        log(traceback.format_exc())
        return False
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass


def maybe_generate_daogang_version(source_path, cwd):
    n_ready, valid_count, blank_n_rows = check_n_column_ready_for_daogang(source_path)
    if not n_ready:
        if blank_n_rows:
            preview = blank_n_rows[:20]
            suffix = " ..." if len(blank_n_rows) > 20 else ""
            log(
                f"  N列有效数据行存在空值，不弹出到岗到位生成询问；"
                f"有效行 {valid_count} 行，N列空值行：{preview}{suffix}"
            )
        else:
            log("  未识别到有效数据行，不弹出到岗到位生成询问")
        return None

    log(f"  N列有效数据行全部非空：有效行 {valid_count} 行，可生成到岗到位版本")
    # V2.8：恢复真正的 Windows 弹窗。V2.7 这里仍走控制台 input()，
    # 在双击运行/无控制台环境下会直接回退为“不生成”，表现为弹窗和到岗到位文件同时消失。
    if not ask_yes_no_dialog(
        "生成到岗到位版本",
        "检测到 N 列有效数据行均已填写到岗/同进同出信息。\n是否生成“到岗到位”版本？",
        default=False,
    ):
        log("  用户选择不生成到岗到位版本")
        return None

    sidebar = collect_daogang_sidebar_input()

    target_path = build_daogang_path(source_path)
    work_path = build_daogang_work_path(target_path)
    failure_titles = {"文件被占用", "只读无法修改", "处理失败", "文件不存在", "缓存错误"}
    try:
        if os.path.exists(work_path):
            try:
                os.remove(work_path)
                log(f"  已删除上次残留的到岗到位处理中临时文件：{work_path}")
            except Exception as e:
                log(f"  删除到岗到位处理中临时文件失败：{e}")
                return None
        shutil.copy2(source_path, work_path)
        log(f"  已复制生成到岗到位处理中临时文件：{work_path}")
    except Exception as e:
        log(f"  复制到岗到位处理中临时文件失败：{e}")
        return None

    if not transform_daogang_columns(work_path):
        try:
            if os.path.exists(work_path):
                os.remove(work_path)
        except Exception:
            pass
        return None

    title = preprocess_excel(work_path, cwd, mode="daogang", daogang_sidebar=sidebar)
    if title in failure_titles:
        log(f"  到岗到位处理中临时文件处理失败，未替换最终文件：{title}")
        try:
            if os.path.exists(work_path):
                os.remove(work_path)
        except Exception:
            pass
        return None

    try:
        if os.path.exists(target_path) and is_file_locked(target_path):
            log(f"  到岗到位最终文件正被占用，无法替换：{target_path}")
            return None
        os.replace(work_path, target_path)
        log(f"  到岗到位最终文件已生成/替换：{target_path}")
    except Exception as e:
        log(f"  替换到岗到位最终文件失败，处理中临时文件保留供检查：{e}")
        return None

    log(f"  到岗到位版本处理完成：{os.path.basename(target_path)} -> {title}")
    return target_path


# ================= 内容提取 =================
LIVE_WORK_LEADERS = {"陈承鲁", "杨文波", "毛武田"}
MODEB_PROFESSION_ORDER = ["带电作业", "输电", "变电", "配电", "配网工程", "营销"]
RISK_LEVELS = ["二级", "三级", "四级", "五级"]
MARKETING_ACTION_ORDER = [
    "高压计量更换与新装电能表",
    "更换高压电流互感器及二次回路接线检查",
    "电流互感器及二次回路接线检查",
    "新装采集终端",
    "新装电能表、采集终端、电流互感器及二次回路接线检查",
    "电流互感器及二次回路接线",
    "高压增容工程竣工验收",
    "高压业扩工程竣工验收",
    "新装关口电能表",
    "低压计量新装计量箱及电能表、批量更换电能表",
    "低压计量新装及批量更换电能表",
    "低压计量新装计量箱及电能表",
    "新装计量箱与电能表",
    "新装与更换电能表",
    "低压计量批量更换电能表",
]


def clean_person_name(text):
    s = remove_phone_and_noise(text)
    s = re.sub(r'^(联系人|负责人|姓名)[:：\s]*', '', s).strip()
    if "/" in s:
        s = s.split("/")[0].strip()
    if "\n" in s:
        s = s.split("\n")[0].strip()
    if "-" in s:
        tail = s.split("-")[-1].strip()
        if tail:
            s = tail
    return s


def classify_profession(profession, leader):
    prof = str(profession).strip() if not is_blank_value(profession) else ""
    leader_name = clean_person_name(leader)
    if prof == "配电" and leader_name in LIVE_WORK_LEADERS:
        return "带电作业"
    return prof


def normalize_work_text(text, keep_newlines=False):
    if is_blank_value(text):
        return ""
    s = str(text).replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\u3000", " ")
    if keep_newlines:
        lines = [re.sub(r"[ \t]+", " ", x).strip() for x in s.split("\n")]
        return "\n".join(x for x in lines if x)
    s = re.sub(r"\s+", "", s)
    return s.strip()


def remove_parentheses_notes(text):
    """删除中英文圆括号及其全部内容，支持混合和嵌套括号。

    如果原文只有左括号而没有右括号，保留该段原文，避免因录入错误
    把后续真实工作内容全部删除。
    """
    source = str(text or "")
    output = []
    pending = []
    depth = 0
    for char in source:
        if char in "（(":
            if depth == 0:
                pending = [char]
            else:
                pending.append(char)
            depth += 1
            continue
        if char in "）)":
            if depth > 0:
                depth -= 1
                if depth == 0:
                    pending = []
                else:
                    pending.append(char)
            continue
        if depth > 0:
            pending.append(char)
        else:
            output.append(char)
    if depth > 0:
        output.extend(pending)
    return "".join(output)


def count_work_items(raw_text):
    s = normalize_work_text(raw_text, keep_newlines=True)
    matches = re.findall(r"(?m)(?:^|\n)\s*\d+\s*[.．、]", s)
    if matches:
        return len(matches)
    if "；" in s or ";" in s:
        parts = [p for p in re.split(r"[；;]", s) if p.strip()]
        return len(parts) if len(parts) >= 2 else 1
    return 1


def cleanup_work_for_summary(raw_text, remove_branch=True):
    s = normalize_work_text(raw_text, keep_newlines=True)
    s = re.sub(r"(?<=\d)[kK][vV]?", "kV", s)
    s = remove_parentheses_notes(s)
    s = re.sub(r"(?m)(^|\n)\s*\d+\s*[.．、]\s*", r"\1", s)
    s = s.replace("变电站", "站")
    s = s.replace("济宁泗水", "")
    s = s.replace("及相关附件", "及附件").replace("相关附件", "附件")
    s = s.replace("相关配备设备", "配套设备").replace("柜内附件", "柜附件")
    if remove_branch:
        for branch_word in ("联络支线", "过驾峪支线", "庠厂5号变支线"):
            s = s.replace(branch_word, "")
    s = s.replace("\n", "，")
    s = re.sub(r"[；;。]+", "，", s)
    s = s.replace("：，", "").replace(":，", "")
    s = re.sub(r"，{2,}", "，", s)
    s = re.sub(r"\s+", "", s)
    return s.strip("，、；;。. ")


def split_summary_clauses(text):
    parts = [p.strip("，、；;。. ") for p in re.split(r"[，,；;。]", text) if p.strip("，、；;。. ")]
    merged = []
    for p in parts:
        if p and p not in merged:
            merged.append(p)
    return merged


def shorten_clause(text, max_len):
    text = text.strip("，、；;。. ")
    if len(text) <= max_len:
        return text
    cut = text[:max_len].rstrip("，、；;:：/ ")
    return cut


def finish_work_sentence(text, item_count):
    s = text.strip("，、；;。. ")
    if not s:
        return ""
    if item_count >= 2 and not s.endswith("等作业内容"):
        s = f"{s}等作业内容"
    return s + "。"


def summarize_marketing_work_for_mode_a(raw_text, compact=False):
    raw = normalize_work_text(raw_text)
    if is_high_voltage_metering_marketing_work(raw):
        return summarize_high_voltage_metering_list(raw_text, compact=compact)
    text = cleanup_work_for_summary(raw_text, remove_branch=False)
    company_or_area = ""
    m = re.search(r"线(.+?)(更换|新装|业扩|批量更换)", text)
    if m:
        company_or_area = m.group(1).strip("，、；;:：。")
    if not company_or_area:
        m = re.search(r"(?:台区|计量箱)(.+?)(新装|更换|批量更换)", text)
        if m:
            company_or_area = m.group(0).split(m.group(2))[0].strip("，、；;:：。")

    if "竣工验收" in raw:
        action = "工程竣工验收"
        if "高压增容" in raw or "增容" in raw:
            action = "高压增容工程竣工验收"
        elif "业扩" in raw:
            action = "高压业扩工程竣工验收"
    elif "批量" in raw and "电能表" in raw:
        action = "批量更换电能表"
    elif "电流互感器" in raw and "二次回路" in raw:
        action = "更换高压电流互感器及二次回路接线检查"
    elif "电能表" in raw and "采集终端" in raw:
        action = "新装电能表、采集终端、电流互感器及二次回路接线检查"
    elif "电能表" in raw and "新装" in raw:
        action = "新装电能表"
    elif "电能表" in raw and "更换" in raw:
        action = "更换电能表"
    else:
        action = shorten_clause(text, 45 if compact else 80)

    if compact and company_or_area:
        company_or_area = re.sub(r"泗水县|济宁|有限公司|有限责任公司", "", company_or_area)
    prefix = f"{company_or_area}" if company_or_area else ""
    return finish_work_sentence(f"{prefix}{action}", 1)



# ================= V2.4 真实语料压缩核心（营销专业保持5.2原逻辑） =================
# V2.4边界：营销专业不进入V2语义压缩；模式A和模式B均沿用5.2营销处理。
# 其他专业继续使用V2.3真实语料压缩与语义守恒。
# 设计原则：模式A“缩写版”领导式压缩并保主要事实；模式B只从模式A继续生成领导总览。
# 压缩主要依靠：安全说明删除、工程套话瘦身、杆号并列、动作共享、设备列表共享谓词。
# 不以“动作重要/不重要”词典直接删动作；守恒失败立即回退。

V23_SAFETY_WORDS = (
    "高低压同杆架设", "高压线路带电", "低压线路带电", "低压线路不带电",
    "配电二票", "有限空间", "同塔架设", "上方线路带电", "临近线路",
    "邻近线路", "保持安全距离", "专责监护", "作业范围有0.4kV", "通讯线",
)
V23_BUSINESS_NOTE_WORDS = (
    "号变", "台区", "配电室", "有限公司", "公司", "配合更换", "废弃线路",
    "迁改线路", "联络工程", "高压计量", "低压计量", "箱变", "主变", "环网柜",
)
V23_STRONG_DEVICES = (
    "变压器", "主变", "JP柜", "母线", "配电盘", "低压柜", "环网柜", "环网箱",
    "分支柜", "接头柜", "箱变", "断路器", "PT", "FTU", "电缆终端", "电缆",
)
V23_DEVICE_LIST_WORDS = (
    "一二次融合断路器", "融合断路器", "变压器", "JP柜", "母线", "配电盘", "低压柜",
    "低压出线电缆", "出线电缆", "台架", "配电设备", "环网柜", "环网箱", "分支柜",
    "接头柜", "箱变", "断路器", "PT", "FTU", "电缆终端", "电缆", "耐张金具",
)
V23_ACTION_PATTERNS = {
    "EXCAVATE": (r"基础开挖", r"杆坑开挖", r"挖坑"),
    "ERECT": (r"电杆组立", r"管塔组立", r"钢管杆(?:吊装)?组立", r"塔(?:吊装)?组立", r"立杆", r"挖坑立杆"),
    "INSTALL": (r"安装", r"装设", r"新装", r"就位", r"装驱鸟器", r"装占位器", r"装耐张金具", r"装横担"),
    "REPLACE": (r"更换", r"换线"),
    "REMOVE": (r"拆除", r"拆旧", r"拆隔离开关", r"拆跌落开关"),
    "DISCONNECT": (r"断引流线", r"断开[^，；。]{0,16}引流线", r"弓子线断开", r"断、接引流线", r"断接引流线"),
    "CONNECT": (r"接引流线", r"接[^，；。]{0,16}引流线", r"弓子线搭接", r"断、接引流线", r"断接引流线", r"搭接"),
    "SHORT": (r"短接",),
    "RECONNECT": (r"改接",),
    "LAY": (r"敷设",),
    "STRING": (r"展放", r"架设"),
    "MAKE": (r"制作",),
    "TEST": (r"试验", r"耐压", r"回路电阻"),
    "CRIMP": (r"压接",),
    "BIND": (r"绑扎", r"绑导线"),
    "FIX": (r"固定",),
    "MOVE": (r"迁移", r"改移", r"迁改"),
    "CLEAN": (r"清理", r"清扫", r"除尘"),
    "INSPECT": (r"巡视", r"检查", r"检测"),
    "TIGHTEN": (r"紧固", r"复紧"),
    "DEBUG": (r"调试",),
    "DRIVE": (r"传动",),
    "VERIFY": (r"检验", r"校验"),
    "CURRENT": (r"升流",),
    "WRAP": (r"绝缘包裹", r"包裹"),
    "SHIELD_INSTALL": (r"装绝缘遮蔽", r"装、拆绝缘遮蔽", r"装拆绝缘遮蔽"),
    "SHIELD_REMOVE": (r"拆绝缘遮蔽", r"装、拆绝缘遮蔽", r"装拆绝缘遮蔽"),
    "ACCEPT": (r"验收",),
    "SAMPLE": (r"取油样",),
    "JOINT_DEBUG": (r"联调",),
}


def v23_norm(text, keep_newlines=False):
    import unicodedata
    if is_blank_value(text):
        return ""
    s = unicodedata.normalize("NFC", str(text)).replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("＃", "#").replace("﹟", "#").replace("\u3000", " ")
    s = re.sub(r"(?<=\d)[kK][vV]", "kV", s)
    s = s.replace("装装绝缘遮蔽", "装绝缘遮蔽").replace("包裹包裹", "包裹")
    s = s.replace("相关配备设备", "配套设备")
    if keep_newlines:
        lines = [re.sub(r"[ \t]+", " ", x).strip() for x in s.split("\n")]
        return "\n".join(x for x in lines if x)
    return re.sub(r"\s+", "", s).strip()


def v23_is_safety_note(note):
    n = v23_norm(note)
    if not n:
        return False
    # 强安全状态优先：即使句中带台区/杆号，也仍属于安全说明。
    flat = re.sub(r"[、，,\s]", "", n)
    if any(re.sub(r"[、，,\s]", "", w) in flat for w in V23_SAFETY_WORDS):
        return True
    if any(w in n for w in V23_BUSINESS_NOTE_WORDS):
        return False
    return False


def v23_classify_parentheses(text):
    """V3.2统一口径：所有圆括号说明都不进入模式A和模式B。"""
    return remove_parentheses_notes(text)


def v23_slim_title(title, mode="A"):
    s = v23_norm(v23_classify_parentheses(title))
    s = s.replace("济宁市泗水县", "").replace("济宁泗水", "").replace("济宁市", "").replace("泗水县", "")
    for p, rpl in (
        (r"优化营商项目", ""), (r"重过载治理工程", ""), (r"电压越限治理工程", ""),
        (r"新建电网配套工程", ""), (r"电网配套工程", ""), (r"配套工程项目", "配套"),
        (r"供电改造工程", "供电改造"), (r"低压线路维修工程", "低压线路维修"),
        (r"线路维修工程", "线路维修"), (r"新建联络工程", "联络工程"),
    ):
        s = re.sub(p, rpl, s)
    s = re.sub(r"工程(?=$)", "", s)
    s = re.sub(r"(台区)新建台区$", r"\1", s)
    s = s.replace("台区台区", "台区")
    s = s.strip("，、；;。. :：")
    # 只有上级站名时通常不是最小充分定位。
    if re.fullmatch(r"(?:35|110|220)kV[^，；。:：]{1,28}(?:变电)?站", s):
        return ""
    return s


def v23_split_title_items(raw_text):
    """按项目级编号切分；编号可挤在同一行，分号本身不当作项目号。"""
    s = v23_norm(v23_classify_parentheses(raw_text), keep_newlines=True)
    # 只在行首/分号/句号后的数字识别项目号；兼容“1.10kV”，不会把0.4kV当编号。
    marked = re.sub(r"(^|[\n；;。:：])\s*([1-9]\d?)[.．]", lambda m: (m.group(1) if m.group(1) not in ("；", ";", "。", ":", "：") else "\n") + f"@@V23ITEM{m.group(2)}@@", s, flags=re.M)
    lines = [x.strip() for x in marked.split("\n") if x.strip()]
    title_parts, items = [], []
    current = None
    for line in lines:
        m = re.match(r"@@V23ITEM\d+@@(.*)$", line)
        if m:
            if current:
                items.append(current.strip(" ，、；;。"))
            current = m.group(1).strip()
            continue
        if current is None:
            title_parts.append(line)
        else:
            # 中途出现纯站名标题，不硬拼到上一项；下项本身通常已有线路定位。
            if re.fullmatch(r"(?:35|110|220)kV[^，；。:：]{1,28}(?:变电)?站[:：]?", line):
                continue
            current += "；" + line
    if current:
        items.append(current.strip(" ，、；;。"))
    if not items:
        return "", [s.strip(" ，、；;。")]
    title = "".join(title_parts).strip(" ，、；;。:：")
    return title, [x for x in items if x]


def v23_compact_phrases(text):
    s = v23_norm(text)
    s = s.replace("变电站", "站")
    s = s.replace("及相关附件", "及附件").replace("相关附件", "附件")
    s = s.replace("相关配电设备", "配电设备").replace("相关配套设备", "配套设备")
    repls = (
        (r"电杆基础开挖[、，]?(?:电杆)?组立", "挖坑立杆"),
        (r"杆坑开挖[、，]?(?:电杆)?组立", "挖坑立杆"),
        (r"基础开挖[、，]?(?:电杆)?组立", "挖坑立杆"),
        (r"下户线制作[、，]?安装", "下户线制作安装"),
        (r"电缆终端制作[、，]?试验[、，]?压接", "电缆终端制作试验压接"),
        (r"终端制作[、，]?试验[、，]?压接", "终端制作试验压接"),
        (r"断引流线[、，]?接引流线", "断、接引流线"),
        (r"弓子线断开[、，]?弓子线搭接", "弓子线断开、搭接"),
        (r"安装耐张线夹[、，]?耐张绝缘子[、，]?耐张金具(?:等附件)?", "装耐张金具"),
        (r"装设绝缘遮蔽|安装绝缘遮蔽|设置绝缘遮蔽", "装绝缘遮蔽"),
        (r"拆除绝缘遮蔽", "拆绝缘遮蔽"),
        (r"安装直线横担", "装横担"),
        (r"绑扎固定导线", "绑导线"),
        (r"安装驱鸟器", "装驱鸟器"),
        (r"安装占位器", "装占位器"),
        (r"一二次融合断路器[、，]PT[、，]FTU及附件安装", "融合断路器/PT/FTU安装"),
        (r"一二次融合断路器[、，]PT[、，]FTU安装", "融合断路器/PT/FTU安装"),
        (r"箱变就位及附件安装", "箱变就位、附件安装"),
        (r"导、地线", "导地线"),
    )
    for p, rpl in repls:
        s = re.sub(p, rpl, s)
    # 设备列表共享谓词：只对明确设备名列表做，避免普通名词被误并。
    words = sorted(V23_DEVICE_LIST_WORDS, key=len, reverse=True)
    token = "(?:" + "|".join(re.escape(w) for w in words) + ")"
    pattern = re.compile(rf"({token}(?:(?:、|，|及){token}){{1,7}})(更换|安装|拆除)")
    def list_repl(m):
        head = re.sub(r"(?:、|，|及)", "/", m.group(1))
        return re.sub(r"/+", "/", head) + m.group(2)
    s = pattern.sub(list_repl, s)
    s = re.sub(r"[；;](（[^）]+）)$", r"\1", s)
    s = re.sub(r"[，,]{2,}", "，", s)
    return s.strip("，、；;。. ")


def v23_strip_station_if_line(text):
    s = str(text or "")
    m = re.match(r"^((?:35|110|220)kV[^，；;。:：]{1,25}站)[:：]?(.*)$", s)
    if m and re.search(r"(?:10|35|110|220)kV[^，；;。:：]{1,45}?线", m.group(2)):
        return m.group(2).lstrip("，、:：")
    return s


def v23_parse_pole_item(item):
    s = str(item or "")
    # 已是范围/多个杆的一句话不重构，防止破坏“#03-#08杆”等杆号范围。
    if re.search(r"#\d+\s*-\s*#?\d+(?=(?:号)?杆)", s):
        return None
    if len(re.findall(r"(?:号)?杆", s)) != 1:
        return None
    m = re.search(r"(?P<pole>(?:(?:新建|新|原|旧)?#?\d+(?:[-+]\d+)*|[AB]))(?:号)?杆", s)
    if not m:
        return None
    prefix = s[:m.start()].strip("，、:：")
    pole = m.group("pole").replace("新建", "新").lstrip("#")
    tail = s[m.end():].strip("，、:：")
    if not tail:
        return None
    return prefix, pole, tail


def v23_merge_same_pole_actions(items):
    groups, order = {}, []
    for idx, item in enumerate(items):
        parsed = v23_parse_pole_item(item)
        if not parsed:
            key=("raw", idx); groups[key]=item; order.append(key); continue
        prefix,pole,tail=parsed
        key=("pole",prefix,tail)
        if key not in groups:
            groups[key]=[]; order.append(key)
        if pole not in groups[key]: groups[key].append(pole)
    out=[]
    for key in order:
        if key[0]=="raw": out.append(groups[key]); continue
        _,prefix,tail=key
        out.append(f"{prefix}{'、'.join(groups[key])}杆{tail}")
    return out


def v23_merge_paired_actions(items):
    out=[]; used=[False]*len(items)
    for i,cur in enumerate(items):
        if used[i]: continue
        m1=re.match(r"^(.*?杆)(装绝缘遮蔽|拆绝缘遮蔽)(（[^）]+）)?$",cur)
        if m1:
            for j in range(i+1,len(items)):
                if used[j]: continue
                m2=re.match(r"^(.*?杆)(装绝缘遮蔽|拆绝缘遮蔽)(（[^）]+）)?$",items[j])
                if m2 and m1.group(1)==m2.group(1) and (m1.group(3) or "")== (m2.group(3) or "") and {m1.group(2),m2.group(2)}=={"装绝缘遮蔽","拆绝缘遮蔽"}:
                    out.append(f"{m1.group(1)}装、拆绝缘遮蔽{m1.group(3) or ''}")
                    used[i]=used[j]=True; break
        if used[i]: continue
        m1=re.match(r"^(.*?杆)(断引流线|接引流线)$",cur)
        if m1:
            for j in range(i+1,len(items)):
                if used[j]: continue
                m2=re.match(r"^(.*?杆)(断引流线|接引流线)$",items[j])
                if m2 and m1.group(1)==m2.group(1) and {m1.group(2),m2.group(2)}=={"断引流线","接引流线"}:
                    out.append(f"{m1.group(1)}断、接引流线"); used[i]=used[j]=True; break
        if not used[i]: out.append(cur); used[i]=True
    return out


def v23_action_signature(text):
    s=v23_norm(text)
    found=set()
    for key,patterns in V23_ACTION_PATTERNS.items():
        if any(re.search(p,s) for p in patterns): found.add(key)
    # “制作安装”必须同时算两动作。
    if "制作安装" in s: found.update(("MAKE","INSTALL"))
    if "挖坑立杆" in s: found.update(("EXCAVATE","ERECT"))
    return found


def v23_device_signature(text):
    s=v23_norm(text)
    return {d for d in V23_STRONG_DEVICES if d in s}


def v23_protected_ids(text, mode="A"):
    s=v23_norm(text); ids=set()
    for m in re.finditer(r"#?\d+(?=(?:号)?(?:箱变|主变|配电室|环网柜|环网箱|分支柜|接头柜))",s): ids.add(m.group(0).lstrip("#"))
    for m in re.finditer(r"\bF\d{1,3}\b",s,re.I): ids.add(m.group(0).upper())
    for m in re.finditer(r"(?<!\d)(?:1001|1002|1003|1004)(?!\d)",s): ids.add(m.group(0))
    for m in re.finditer(r"(?:新建|新|原|旧)#?\d+(?:[-+]\d+)*(?=(?:号)?杆)",s): ids.add(m.group(0).replace("新建","新"))
    if mode=="A":
        for m in re.finditer(r"(?<![A-Za-z0-9])#?\d+(?:[-+]\d+)*(?=(?:号)?(?:杆|塔))",s): ids.add(m.group(0).lstrip("#"))
    return ids


def v23_has_id(text, token):
    s=v23_norm(text); t=str(token)
    if t.startswith("F"): return t.upper() in s.upper()
    if t.startswith(("新","原","旧")): return t in s.replace("新建","新")
    return bool(re.search(rf"(?<!\d)#?{re.escape(t.lstrip('#'))}(?!\d)",s))


def v23_line_names(text):
    s = v23_norm(text)
    # 防止“110kV某站110kV某线”被整体误识别成一条线路；取离“线”最近的电压前缀。
    pat = r"(?<!\d)(?:10|35|110|220)kV(?:(?!(?:10|35|110|220)kV)[^，、；;。:：]){1,42}?线"
    return set(x.strip("，、；;。. :：") for x in re.findall(pat, s))


def v23_guard(source,candidate,mode="A",profession=""):
    if not candidate: return False
    clean_source = v23_classify_parentheses(source)
    clean_candidate = v23_classify_parentheses(candidate)
    sa,ca=v23_action_signature(clean_source),v23_action_signature(clean_candidate)
    missing=sa-ca
    if "FIX" in missing and "BIND" in sa and "BIND" in ca: missing.discard("FIX")
    if missing: return False
    if not v23_device_signature(clean_source).issubset(v23_device_signature(clean_candidate)): return False
    for token in v23_protected_ids(clean_source,mode=mode):
        if not v23_has_id(clean_candidate,token): return False
    src_lines=v23_line_names(clean_source); cand_lines=v23_line_names(clean_candidate)
    if mode=="B" and ("0.4kV" in v23_norm(clean_source) or "低压线路" in clean_source) and "台区" in clean_candidate:
        pass
    elif src_lines and not (src_lines & cand_lines):
        return False
    # 新旧/改接/迁改属于关系，不可被普通“施工”替代。
    for rel in ("改接","废弃线路"):
        if rel in clean_source and rel not in clean_candidate: return False
    if re.search(r"(?:原|旧).{0,24}(?:新建|新)|(?:新建|新).{0,24}(?:原|旧)",clean_source):
        if not (re.search(r"原|旧",clean_candidate) and re.search(r"新建|新",clean_candidate)): return False
    if any(w in clean_source for w in ("迁改","迁移","改移")) and not any(w in clean_candidate for w in ("迁改","迁移","改移")): return False
    return True


def v23_low_voltage_mode_b(source):
    s=v23_norm(source)
    if not ("0.4kV" in s or "低压线路" in s or ("低压" in s and any(x in s for x in ("导线","电缆","下户线")))):
        return ""
    # 台区优先，保留上级10kV线路只在标题中容易取得时。
    taiqus=re.findall(r"([^，、；;。:：]{1,34}?台区)",s)
    anchor=taiqus[0] if taiqus else ""
    if anchor:
        # 去掉台区前重复的工程套话，但保留可能存在的10kV线路。
        m=re.search(r"((?:10kV[^，、；;。:：]{1,35}?线)?[^，、；;。:：]{0,24}?台区)",anchor)
        if m: anchor=m.group(1)
    actions=[]
    def add(x):
        if x and x not in actions: actions.append(x)
    if re.search(r"T接点拆除",s): add("T接点拆除")
    if re.search(r"低压出线电缆[^；。]{0,20}(?:搭接|接入)|搭接新放低压出线电缆",s): add("出线电缆搭接")
    if re.search(r"低压电缆[^；。]{0,16}敷设",s):
        add("电缆敷设搭接" if re.search(r"低压电缆[^；。]{0,24}敷设[^；。]{0,10}搭接",s) else "电缆敷设")
    if re.search(r"低压导线[^；。]{0,16}更换",s): add("导线更换")
    if re.search(r"低压导线[^；。]{0,16}展放",s): add("导线展放")
    if re.search(r"原低压电缆[^；。]{0,16}拆除",s): add("原电缆拆除")
    elif re.search(r"低压电缆[^；。]{0,16}拆除",s): add("电缆拆除")
    if re.search(r"下户线[^；。]{0,16}改接",s): add("下户改接")
    if re.search(r"下户线[^；。]{0,16}制作安装",s): add("下户线制作安装")
    if re.search(r"弓子线[^；。]{0,16}(?:断开|断接).{0,12}(?:搭接|接)",s): add("弓子线断接")
    if not actions: return ""
    return f"{anchor + '低压线路' if anchor else '低压线路'}{'、'.join(actions)}"


def v23_compact_work(raw_text, profession="", mode="A"):
    prof=str(profession or "").strip()
    title,items=v23_split_title_items(raw_text)
    slim_title=v23_slim_title(title,mode=mode)
    compact=[]
    for item in items:
        c=v23_strip_station_if_line(v23_compact_phrases(item))
        if c: compact.append(c)
    compact=v23_merge_same_pole_actions(compact)
    compact=v23_merge_paired_actions(compact)
    body="；".join(compact).strip("；。")
    # 标题只在承担地点/台区/场所定位时保留；避免纯站名重复。
    if slim_title and slim_title not in body:
        candidate=f"{slim_title}：{body}" if body else slim_title
    else:
        candidate=body or slim_title
    candidate=re.sub(r"[:：]；", "：", candidate)
    candidate=re.sub(r"；{2,}", "；", candidate).strip("，、；;。. ")
    if not v23_guard(raw_text,candidate,mode="A",profession=prof):
        # 保守回退仍删除纯安全括号，但不截断动作。
        fallback=v23_norm(v23_classify_parentheses(raw_text),keep_newlines=True)
        fallback=re.sub(r"(?m)(^|\n)\s*\d+\s*[.．、]\s*",r"\1",fallback)
        fallback=v23_compact_phrases(fallback.replace("\n","；"))
        if v23_guard(raw_text,fallback,mode="A",profession=prof): candidate=fallback
        else: candidate=v23_norm(v23_classify_parentheses(raw_text))
    return candidate.rstrip("。") + "。" if candidate else ""


def v23_mode_b_from_a(mode_a_text, profession=""):
    source=str(mode_a_text or "").rstrip("。")
    prof=str(profession or "").strip()
    if not source: return ""
    candidates=[]
    lv=v23_low_voltage_mode_b(source)
    if lv: candidates.append(lv)
    # 继续沿用5.2各专业的候选套路，但输入明确是模式A。
    try:
        if prof=="输电": candidates.append(summarize_transmission_mode_b(source))
        elif prof=="带电作业": candidates.append(summarize_live_work_mode_b(source))
        elif prof=="变电": candidates.append(summarize_substation_mode_b(source))
        elif prof=="配网工程": candidates.append(summarize_peiwang_project_mode_b(source))
        elif prof=="配电": candidates.append(summarize_peidian_mode_b(source))
        else: candidates.append(summarize_general_mode_b_candidate_a(source,max_clauses=8,clause_len=180))
    except Exception:
        pass
    # 通用候选：模式A再次瘦工程标题和上级站，不删除动作。
    generic=v23_compact_phrases(v23_strip_station_if_line(source))
    generic=re.sub(r"(?:重过载治理|电网配套|配套工程项目)","",generic)
    candidates.extend([generic,source])
    valid=[]
    for idx,c in enumerate(candidates):
        c=mode_b_summary_plain_text(c)
        if c and v23_guard(source,c,mode="B",profession=prof):
            # 以保真为前提选择较短候选；长度接近时优先人工专用候选。
            valid.append((len(c),idx,c))
    chosen=min(valid,key=lambda x:(x[0],x[1]))[2] if valid else source
    chosen=chosen.replace("台区台区", "台区")
    chosen=re.sub(r"[:：]{2,}", "：", chosen)
    return chosen.rstrip("。")+"。"


# ================= V2.6 领导视角摘要层 =================
# 设计依据：
# 1) 模式A“缩写版”也是领导阅读版：一条Excel计划仍对应一条，保留全部主要动作/设备/关系，
#    主要通过去重复定位、合并杆号、共享设备/动作、低压范围概括来变短；完整版紧随其后保留原文。
# 2) 模式B是领导视角总览：按专业汇总，以“重要性 + 覆盖度 + 去冗余 + 句子融合”为原则，
#    允许省略不影响判断的次要施工细节，但不得凭空增加设备、线路、编号或拓扑关系。
# 3) 营销专业完全绕过本层，继续使用5.2原版模式A/模式B逻辑。

V26_LV_ACTION_RULES = (
    (r"T接点拆除", "T接点拆除", 5.0),
    (r"原低压出线电缆拆除|原低压电缆拆除", "原低压电缆拆除", 6.0),
    (r"低压出线电缆拆除|低压电缆拆除", "低压电缆拆除", 5.5),
    (r"低压出线电缆更换", "低压出线电缆更换", 5.7),
    (r"低压电缆更换", "低压电缆更换", 5.4),
    (r"裸导线更换绝缘导线", "裸导线更换绝缘导线", 6.0),
    (r"低压导线更换|(?<!高压)导线更换", "低压导线更换", 5.6),
    (r"低压导线迁改|低压导线迁移|低压导线改移|(?<!高压)导线迁改|(?<!高压)导线迁移|(?<!高压)导线改移", "低压导线迁改", 5.8),
    (r"低压导线展放|(?<!高压)导线展放", "低压导线展放", 5.1),
    (r"低压导线架设|(?<!高压)导线架设", "低压导线架设", 5.1),
    (r"搭接新放低压出线电缆|新放低压出线电缆(?:搭接)?", "新放低压出线电缆", 5.9),
    (r"低压出线电缆敷设(?:搭接)?", "低压出线电缆敷设", 4.6),
    (r"低压电缆敷设搭接", "低压电缆敷设搭接", 4.4),
    (r"低压电缆敷设", "低压电缆敷设", 4.1),
    (r"低压下户电缆改接|下户电缆改接", "下户电缆改接", 5.2),
    (r"低压下户线改接|下户线改接", "下户线改接", 5.2),
    (r"下户线制作安装", "下户线制作安装", 5.2),
    (r"下户线更换", "下户线更换", 5.2),
    (r"导线瓷瓶绑扎", "导线瓷瓶绑扎", 3.9),
    (r"低压线路搭接", "低压线路搭接", 3.7),
    (r"弓子线断开、搭接|弓子线断接", "弓子线断接", 4.7),
    (r"金具安装", "金具安装", 3.9),
)


def v26_unique(values):
    out = []
    for value in values:
        value = str(value or "").strip("，、；;。. :：")
        if value and value not in out:
            out.append(value)
    return out


def v26_line_names_ordered(text):
    s = v23_norm(text)
    pattern = r"(?<!\d)(?:10|35|110|220)kV(?:(?!(?:10|35|110|220)kV)[^，、；;。:：]){1,42}?线"
    return v26_unique(re.findall(pattern, s))


def v26_taiqu_names(text):
    s = v23_norm(text)
    values = []
    for match in re.finditer(r"([^，、；;。:：\n]{1,40}?台区)", s):
        value = match.group(1)
        value = re.sub(r"^[（(]?\d+[.．、]", "", value).lstrip("（(")
        if "kV" in value:
            value = value.split("kV")[-1]
        if "线" in value:
            value = value.split("线")[-1]
        value = re.sub(r"^(?:新建|原)", "", value)
        for word in ("工程", "治理", "改造", "项目", "镇", "街道"):
            if word in value:
                value = value.split(word)[-1]
        value = value.strip()
        if value and value != "台区" and value not in values:
            values.append(value)
    return values


def v26_place_anchor(text):
    s = v23_norm(text)
    taiqus = v26_taiqu_names(s)
    if taiqus:
        return taiqus[0]
    suffixes = (
        "服务区南区", "服务区北区", "家属院", "小区", "服务区",
        "配电室", "台变", "公变", "村",
    )
    for suffix in suffixes:
        match = re.search(rf"([^，、；;。:：\n]{{2,32}}?{re.escape(suffix)})", s)
        if not match:
            continue
        value = match.group(1)
        value = re.sub(r"^[（(]?\d+[.．、]", "", value).lstrip("（(")
        if "kV" in value and suffix not in ("服务区南区", "服务区北区"):
            value = value.split("kV")[-1]
        if "线" in value and suffix not in ("配电室",):
            value = value.split("线")[-1]
        for word in ("工程", "项目", "镇", "街道"):
            if word in value:
                value = value.split(word)[-1]
        return value.strip("，、；;。. :：")
    return ""


def v26_strip_pole_refs(text):
    """模式B删除普通杆塔定位；箱变/主变/柜间隔等设备编号不在此处理。"""
    s = str(text or "")
    s = re.sub(
        r"#?\d+(?:[-+]\d+)*(?:、#?\d+(?:[-+]\d+)*)*(?:号)?杆(?:至#?\d+(?:[-+]\d+)*(?:号)?杆)?",
        "",
        s,
    )
    s = re.sub(
        r"#?\d+(?:[-+]\d+)*(?:号)?塔(?:至#?\d+(?:[-+]\d+)*(?:号)?塔)?",
        "",
        s,
    )
    s = s.replace("杆至", "")
    s = re.sub(r"[，、]{2,}", "、", s)
    return s


def v26_extract_lv_actions(text):
    s = v23_norm(text)
    found = []
    for pattern, label, score in V26_LV_ACTION_RULES:
        match = re.search(pattern, s)
        if match:
            found.append([label, score, match.start()])
    found.sort(key=lambda x: x[2])
    unique = []
    labels = set()
    for item in found:
        if item[0] not in labels:
            unique.append(item)
            labels.add(item[0])
    suppress = {
        "低压电缆敷设搭接": {"低压电缆敷设"},
        "原低压电缆拆除": {"低压电缆拆除"},
        "低压出线电缆更换": {"低压电缆更换"},
        "新放低压出线电缆": {"低压出线电缆敷设", "低压电缆敷设", "低压电缆敷设搭接"},
        "下户电缆改接": {"下户线改接"},
    }
    removed = set()
    for key, values in suppress.items():
        if key in labels:
            removed.update(values & labels)
    return [tuple(item) for item in unique if item[0] not in removed]


def v26_mode_a_low_voltage(source):
    s = v23_norm(source)
    if not (
        "0.4kV" in s
        or "低压线路" in s
        or ("低压" in s and any(word in s for word in ("导线", "电缆", "下户")))
    ):
        return ""
    actions = v26_extract_lv_actions(s)
    if not actions:
        return ""
    labels = []
    for label, _score, _pos in sorted(actions, key=lambda x: x[2]):
        if label == "新放低压出线电缆" and re.search(
            r"搭接新放低压出线电缆|新放低压出线电缆.{0,6}搭接", s
        ):
            label = "新放低压出线电缆搭接"
        if label == "导线瓷瓶绑扎" and re.search(r"(?:新建|新)#?\d+", s):
            label = "新建杆导线瓷瓶绑扎"
        if label not in labels:
            labels.append(label)
    lines = v26_line_names_ordered(s)
    taiqus = v26_taiqu_names(s)
    if lines and taiqus:
        anchor = f"{lines[0]}{taiqus[0]}"
    elif taiqus:
        anchor = taiqus[0]
    elif lines:
        anchor = lines[0]
    else:
        anchor = v26_place_anchor(s)
    return f"{anchor}低压线路：{'、'.join(labels)}" if anchor else f"低压线路：{'、'.join(labels)}"


def v26_mode_a_newbuild(source):
    s = v23_norm(source)
    if not (
        re.search(r"(?:基础开挖|杆坑开挖|挖坑)", s)
        and re.search(r"(?:组立|立杆)", s)
    ):
        return ""
    # 多线路/多台区的同一Excel项不强行揉成一个对象，交给原V2保守结果。
    lines = v26_line_names_ordered(s)
    taiqus = v26_taiqu_names(s)
    if len(lines) > 1 or len(taiqus) > 1:
        return ""
    parts = ["电杆基础开挖、电杆组立"]
    if "金具安装" in s:
        parts.append("金具安装")
    devices = []
    for raw, label in (
        ("台架", "台架"), ("变压器", "变压器"), ("JP柜", "JP柜"),
        ("两路出线电缆", "两路出线电缆"), ("配电设备", "配电设备"),
    ):
        if raw in s and label not in devices:
            devices.append(label)
    if devices:
        parts.append("/".join(devices) + "安装")
    for label, _score, _pos in v26_extract_lv_actions(s):
        if label == "新放低压出线电缆" and "搭接" in s:
            label = "新放低压出线电缆搭接"
        if label not in parts and label != "金具安装":
            parts.append(label)
    if taiqus:
        anchor = taiqus[0] + "："
        if lines:
            anchor += lines[0]
    else:
        anchor = lines[0] if lines else ""
    return anchor + "、".join(parts) if anchor else "、".join(parts)


def v26_mode_a_box(source):
    s = v23_norm(source)
    if "箱变" not in s:
        return ""
    place = v26_place_anchor(s)
    if place:
        s = re.sub(rf"^{re.escape(place)}10kV供电改造[:：]?", "", s)
        s = re.sub(rf"^{re.escape(place)}供电改造[:：]?", "", s)
    clauses = [
        item.strip("，、；;。. ")
        for item in re.split(r"[；;。]", s)
        if item.strip("，、；;。. ")
    ]
    outputs = []
    for clause in clauses:
        clause = v23_strip_station_if_line(v23_compact_phrases(clause))
        if place and clause.startswith(place):
            clause = clause[len(place):].lstrip(":：")
        clause = clause.replace("13杆线", "13杆")
        clause = clause.replace("一二次融合开关", "融合开关")
        clause = clause.replace("低压电缆分支箱", "低压分支箱")
        if clause and clause not in outputs:
            outputs.append(clause)
    if not outputs:
        return ""
    return f"{place}：{'；'.join(outputs)}" if place else "；".join(outputs)


def v26_mode_a_parallel_terminal(source):
    s = v23_norm(source)
    if s.count("电缆终端制作") < 2 or "出线柜" not in s:
        return ""
    place = v26_place_anchor(s)
    match = re.search(r"([^，；。]{0,30}#\d+中心配电室)", s)
    base = match.group(1) if match else place
    if place and place in base:
        base = base[base.find(place):]
    refs = []
    for match in re.finditer(r"([ⅠⅡⅢIV]+)段母线#?(\d+)出线柜", s):
        ref = f"{match.group(1)}段母线#{match.group(2)}"
        if ref not in refs:
            refs.append(ref)
    if refs:
        return f"{base}{'、'.join(refs)}出线柜电缆终端制作、压接"
    return ""


def v26_mode_a_formalize(text):
    s = v23_norm(text)
    s = s.replace("绑导线", "绑扎固定导线")
    s = re.sub(r"(?<!安)装驱鸟器", "安装驱鸟器", s)
    return s


def v26_mode_a_guard(source, candidate, profession=""):
    """A层守恒：守动作、强设备、线路和业务关系；普通杆号由完整版承担追溯。"""
    if not candidate:
        return False
    src = v23_classify_parentheses(source)
    cand = v23_classify_parentheses(candidate)
    src_actions = v23_action_signature(src)
    cand_actions = v23_action_signature(cand)
    missing = src_actions - cand_actions
    if "FIX" in missing and "BIND" in src_actions and "BIND" in cand_actions:
        missing.discard("FIX")
    if missing:
        return False
    if not v23_device_signature(src).issubset(v23_device_signature(cand)):
        return False
    src_lines = set(v26_line_names_ordered(src))
    cand_lines = set(v26_line_names_ordered(cand))
    if src_lines and not (src_lines & cand_lines):
        # 低压台区允许以台区替代上级线路定位。
        if not (("0.4kV" in src or "低压线路" in src) and "台区" in cand):
            return False
    # 真正承担设备/拓扑区分的编号仍保护；普通杆塔号不保护。
    for token in re.findall(r"#?\d+(?=(?:号)?(?:箱变|主变|配电室|环网柜|环网箱|分支柜|接头柜))", src):
        if not v23_has_id(cand, token.lstrip("#")):
            return False
    for token in re.findall(r"\bF\d{1,3}\b|(?<!\d)(?:1001|1002|1003|1004)(?!\d)", src, re.I):
        if token.upper() not in cand.upper():
            return False
    for rel in ("改接", "废弃线路"):
        if rel in src and rel not in cand:
            return False
    if any(word in src for word in ("迁改", "迁移", "改移")) and not any(word in cand for word in ("迁改", "迁移", "改移")):
        return False
    if re.search(r"(?:原|旧).{0,28}(?:新建|新)|(?:新建|新).{0,28}(?:原|旧)", src):
        if not (re.search(r"原|旧", cand) and re.search(r"新建|新", cand)):
            return False
    return True


def v26_mode_a_abbrev(raw_text, profession=""):
    """模式A缩写版：保主要事实但允许删除普通杆号/重复定位；完整版保留原文。"""
    prof = str(profession or "").strip()
    if prof == "营销":
        return summarize_marketing_work_for_mode_a(raw_text, compact=False)
    base = v23_compact_work(raw_text, profession=prof, mode="A").rstrip("。")
    candidates = []
    for function in (
        v26_mode_a_parallel_terminal,
        v26_mode_a_box,
        v26_mode_a_newbuild,
        v26_mode_a_low_voltage,
    ):
        try:
            candidate = function(base)
        except Exception:
            candidate = ""
        if candidate:
            candidates.append(candidate)
    candidates.append(base)

    valid = []
    for index, candidate in enumerate(candidates):
        candidate = v26_mode_a_formalize(candidate).strip("，、；;。. ")
        # 借用V2的模式B守恒口径：动作/强设备/线路/新旧关系仍守恒，
        # 但普通杆号不再作为A层硬保留项，因为下方完整版已经承担逐杆追溯。
        if candidate and v26_mode_a_guard(base, candidate, profession=prof):
            valid.append((len(candidate), index, candidate))
    chosen = min(valid, key=lambda item: (item[0], item[1]))[2] if valid else v26_mode_a_formalize(base)
    return chosen.rstrip("。") + "。"


def v26_leader_strip_station(text):
    s = v23_norm(text)
    return re.sub(r"^(?:35|110|220)kV[^，；。:：]{1,25}(?:变电)?站[:：]?", "", s)


def v26_leader_old_new_line(source):
    s = v23_norm(source)
    if not (
        ("拆除旧导线" in s or "拆旧导线" in s)
        and "耐张金具" in s
        and "接引流线" in s
    ):
        return ""
    lines = v26_line_names_ordered(s)
    line = lines[0] if lines else ""
    return f"{line}断引流线、拆旧导线，装耐张金具、挂接导线、接引流线".lstrip()


def v26_live_action_key(clause):
    c = v23_norm(clause)
    if "安装驱鸟器" in c or "装驱鸟器" in c:
        return "安装驱鸟器"
    if "断引流线" in c and "废弃线路" in c:
        return "断废弃线路引流线"
    if "断引流线" in c:
        return "断引流线"
    if "接引流线" in c:
        return "接引流线"
    if "更换避雷器" in c:
        return "更换避雷器"
    if "更换跌落开关" in c:
        return "更换跌落开关"
    if "绝缘包裹" in c:
        return "绝缘包裹"
    return ""


def v26_leader_live(source):
    """带电作业自身动作摘要。V2.9起不在本函数内臆造“配合对象”；跨专业配合关系由全日关联层判定。"""
    special = v26_leader_old_new_line(source)
    if special:
        return special
    s = v23_norm(source).rstrip("。")
    clauses = [item.strip("，、；;。. ") for item in re.split(r"[；;。]", s) if item.strip("，、；;。. ")]
    insulation = {}
    grouped = {}
    group_order = []
    other = []
    for clause in clauses:
        lines = v26_line_names_ordered(clause)
        line = lines[0] if lines else ""
        if "绝缘遮蔽" in clause and line:
            poles = v29_extract_pole_ids(clause)
            bucket = insulation.setdefault(line, [])
            for pole in poles:
                if pole not in bucket:
                    bucket.append(pole)
            continue
        if ("绑扎固定导线" in clause or "绑导线" in clause) and line:
            key = "绑扎固定导线"
        elif re.search(r"断(?:、|,|，)?接引流线|断、接引流线|断引流线[^；;。]{0,20}接引流线", clause):
            key = "断接引流线"
        else:
            key = v26_live_action_key(clause)
        complex_chain = any(word in clause for word in ("短接", "带负荷", "拆除隔离开关", "拆除跌落开关", "断隔离开关", "断跌落开关"))
        if key and line and not complex_chain:
            if key not in grouped:
                grouped[key] = []
                group_order.append(key)
            if line not in grouped[key]:
                grouped[key].append(line)
            continue
        compact = v26_strip_pole_refs(clause)
        compact = re.sub(r"（[^）]*）", "", compact)
        compact = re.sub(r"(?<!安)装驱鸟器", "安装驱鸟器", compact)
        if compact:
            other.append(compact)

    outputs = []
    for line, poles in insulation.items():
        pole_text = f"{'、'.join(poles)}杆" if poles else ""
        outputs.append(f"{line}{pole_text}装、拆绝缘遮蔽")
    for key in group_order:
        outputs.append(f"{'、'.join(grouped[key])}{key}")
    outputs.extend(other)
    return "；".join(v26_unique(outputs))




# ================= V2.9 同日跨专业强关联层 =================
# 只服务于文字提炼：不参与Excel处理、到岗到位、弹窗、打印或文件生成流程。
# 从模式A缩写版建立“线路/杆位/台区或配变/配合目的/动作”签名，再在当天正式记录之间匹配。
# 核心原则：同线路本身绝不够；必须叠加业务对象、杆位或“配合目的↔主体动作”中的至少一项强证据。

V29_SUPPORT_ACTION_WORDS = (
    "绝缘遮蔽", "断引流线", "接引流线", "断接引流线", "电缆引流线",
    "中压发电车", "带负荷", "短接", "直线杆改耐张", "挂接紧固导线",
)


def v29_extract_pole_ids(text):
    s = v23_norm(text)
    values = []
    # 兼容“123+01、126+01杆”和“#01杆至#05杆”。
    for m in re.finditer(r"(#?\d+(?:[-+]\d+)*(?:、#?\d+(?:[-+]\d+)*)*)(?:号)?杆", s):
        for token in m.group(1).split("、"):
            token = token.lstrip("#")
            if token and token not in values:
                values.append(token)
    return values


def v29_normalize_business_key(value):
    s = v23_norm(value)
    s = re.sub(r"^(?:新建|原|旧)", "", s)
    s = s.replace("＃", "#")
    # 2号变 / #2台区 / 2号台变视为同一业务对象编号，但保留前面的地名。
    s = re.sub(r"#?(\d+)号?(?:台架变|台变|公变|配变|变|台区)$", r"#\1", s)
    s = re.sub(r"#(\d+)台区$", r"#\1", s)
    s = re.sub(r"(?:台架变|台变|公变|配变|台区)$", "", s)
    s = re.sub(r"\s+", "", s)
    return s.strip("，、；;。. :：")


def v29_business_objects(text):
    s = v23_norm(text)
    values = []
    patterns = (
        r"([\u4e00-\u9fffA-Za-z]{2,18}#?\d+号?(?:台架变|台变|公变|配变|变|台区))",
        r"([\u4e00-\u9fffA-Za-z]{2,18}#\d+(?:台区)?)",
        r"([\u4e00-\u9fffA-Za-z]{2,18}(?:台区|台架变|台变|公变|配变))",
    )
    for pattern in patterns:
        for m in re.finditer(pattern, s):
            raw = m.group(1)
            # 去掉线路/支线前缀，只保留业务对象尾部。
            if "线" in raw:
                raw = raw.split("线")[-1]
            raw = re.sub(r"^#?\d+(?:[-+]\d+)*(?:号)?杆", "", raw)
            key = v29_normalize_business_key(raw)
            if len(key) >= 2 and key not in values:
                values.append(key)
    return values


def v29_explicit_support_purpose(text):
    s = v23_norm(text)
    purposes = []
    for m in re.finditer(r"配合([^，、；;。）（()]{1,32})", s):
        value = m.group(1).strip("工作 ：:")
        if value and value not in purposes:
            purposes.append(value)
    return purposes


def v29_purpose_matches_target(purposes, target_text):
    t = v23_norm(target_text)
    for p in purposes:
        p = v23_norm(p)
        if not p:
            continue
        if "更换变压器" in p and "变压器" in t and "更换" in t:
            return True
        if "配变改造" in p and any(x in t for x in ("变压器", "JP柜", "配电盘", "母线")) and any(x in t for x in ("更换", "改造", "迁改")):
            return True
        if any(x in p for x in ("更换开关", "换开关")) and any(x in t for x in ("开关", "断路器")) and "更换" in t:
            return True
        if "更换JP柜" in p and "JP柜" in t and "更换" in t:
            return True
        if "更换环网柜" in p and "环网柜" in t and "更换" in t:
            return True
        if any(x in p for x in ("迁改台架", "台架迁改")) and any(x in t for x in ("台架", "台变", "配变")) and any(x in t for x in ("迁改", "迁移", "改移")):
            return True
        if any(x in p for x in ("迁改线路", "线路迁改")) and any(x in t for x in ("迁改", "迁移", "改移", "线路改接")):
            return True
        if "立杆" in p and any(x in t for x in ("电杆组立", "组立", "立杆", "基础开挖")):
            return True
        if "展放导线" in p and any(x in t for x in ("导线展放", "导线架设")):
            return True
        if "联络" in p and "联络" in t:
            return True
        if "检修" in p and any(x in t for x in (
            "导线更换", "开关更换", "断路器更换", "电缆终端", "旧导线拆除",
            "新导线", "耐张", "变压器更换", "JP柜更换", "母线更换", "检修",
        )):
            return True
    return False


def v29_record_signature(mode_a_text, profession=""):
    s = v23_norm(mode_a_text)
    return {
        "profession": str(profession or "").strip(),
        "text": s,
        "lines": set(v26_line_names_ordered(s)),
        "poles": set(v29_extract_pole_ids(s)),
        "objects": set(v29_business_objects(s)),
        "purposes": v29_explicit_support_purpose(s),
        "support_like": any(word in s for word in V29_SUPPORT_ACTION_WORDS) or "配合" in s,
    }


def v29_association_score(source_sig, target_sig):
    if source_sig.get("profession") == target_sig.get("profession"):
        return 0, []
    if not source_sig.get("support_like"):
        return 0, []
    shared_lines = source_sig["lines"] & target_sig["lines"]
    shared_poles = source_sig["poles"] & target_sig["poles"]
    shared_objects = source_sig["objects"] & target_sig["objects"]
    purpose_match = v29_purpose_matches_target(source_sig.get("purposes") or [], target_sig.get("text") or "")
    score = 0
    reasons = []
    if shared_lines:
        score += 6; reasons.append("同线路")
    if shared_objects:
        score += 8; reasons.append("同业务对象")
    if shared_poles:
        score += 7; reasons.append("同杆位")
    if purpose_match:
        score += 6; reasons.append("配合目的匹配主体动作")
    # 仅同线路不得关联；强关联至少要有第二个证据。
    strong_second = bool(shared_objects or shared_poles or purpose_match)
    if not strong_second:
        return 0, []
    # 若连线路都不同，只允许“同业务对象+同杆位”这种极强身份关系。
    if not shared_lines and not (shared_objects and shared_poles):
        return 0, []
    return score, reasons


def v29_target_project_label(mode_a_text, profession=""):
    s = v23_norm(mode_a_text).rstrip("。")
    lines = v26_line_names_ordered(s)
    line = lines[0] if lines else ""
    taiqus = v26_taiqu_names(s)
    obj = taiqus[0] if taiqus else ""
    if not obj:
        # 优先取可读的配变/公变/台变对象。
        m = re.search(r"([^，、；;。:：]{2,24}?(?:#?\d+号)?(?:台架变|台变|公变|配变|变))", s)
        if m:
            obj = m.group(1)
            if "线" in obj:
                obj = obj.split("线")[-1]
            obj = re.sub(r"^#?\d+(?:[-+]\d+)*(?:号)?杆", "", obj)
    prefix = line
    if obj and obj not in prefix:
        prefix += obj

    if ("基础开挖" in s or "挖坑" in s) and ("组立" in s or "立杆" in s):
        purpose = "新建"
    elif "变压器" in s and "更换" in s:
        purpose = "变压器更换"
    elif "更换" in s and any(x in s for x in ("JP柜", "母线", "配电盘", "低压柜")):
        purpose = "配变改造"
    elif any(x in s for x in ("开关更换", "断路器更换", "更换跌落开关", "更换隔离开关")):
        purpose = "开关更换"
    elif any(x in s for x in ("迁改", "迁移", "改移")):
        purpose = "迁改"
    elif "联络" in s:
        purpose = "联络工程"
    elif any(x in s for x in ("导线更换", "旧导线拆除", "耐张金具", "电缆终端拆除")):
        purpose = "线路检修"
    elif "检修" in s:
        purpose = "检修"
    else:
        purpose = ""
    if prefix and purpose:
        return prefix + purpose
    if prefix:
        return prefix + "工作"
    # 保守兜底：只取主体摘要开头，不凭空造项目名。
    compact = v26_leader_summary_from_a(s, profession).rstrip("。") if profession != "带电作业" else s
    return compact.rstrip("，、；;：:")


def v29_build_day_associations(records, mode_a_cache):
    prepared = []
    for idx, record in enumerate(records):
        prof = record.get("profession") or "其他"
        mode_a = mode_a_cache[idx]
        prepared.append({"index": idx, "record": record, "mode_a": mode_a, "sig": v29_record_signature(mode_a, prof)})

    associations = {}
    for src in prepared:
        sig = src["sig"]
        if not sig.get("support_like"):
            continue
        candidates = []
        for dst in prepared:
            if dst["index"] == src["index"]:
                continue
            if dst["record"].get("profession") == "营销":
                continue
            score, reasons = v29_association_score(sig, dst["sig"])
            if score >= 12:
                candidates.append((score, dst["index"], reasons))
        candidates.sort(key=lambda x: (-x[0], x[1]))
        if not candidates:
            continue
        # 两个不同主体分数非常接近时不猜，宁可保持原摘要。
        if len(candidates) >= 2 and candidates[0][0] - candidates[1][0] < 2:
            continue
        score, target_idx, reasons = candidates[0]
        associations[src["index"]] = {
            "target_idx": target_idx,
            "score": score,
            "reasons": reasons,
            "label": v30_target_project_label(mode_a_cache[target_idx], records[target_idx].get("profession") or "其他"),
        }
    return associations


def v34_live_multibranch_defect_summary(raw_text, profession=""):
    """超长带电缺陷治理：同站多支线提升为一句主题，其他作业性质严格退出。"""
    if str(profession or "").strip() != "带电作业":
        return ""
    source = remove_parentheses_notes(raw_text)
    clean = v23_norm(source, keep_newlines=True)
    clauses = v31_split_work_clauses(clean)
    if len(re.sub(r"\s+", "", clean)) < 300 and len(clauses) < 8:
        return ""

    signatures = v23_action_signature(clean)
    allowed_signatures = {"INSTALL", "REPLACE", "WRAP", "BIND", "FIX", "TIGHTEN", "CLEAN"}
    if not signatures or not signatures.issubset(allowed_signatures):
        return ""
    forbidden = re.compile(
        r"断、接引流线|断接引流线|断引流线|接引流线|短接|拆除|"
        r"发电车|导线(?:展放|更换|拆除|改接)|线路搭接|基础开挖|电杆组立|"
        r"电缆(?:敷设|终端制作)|变压器更换|更换(?:跌落开关|隔离开关|真空开关)|"
        r"装绝缘遮蔽|拆绝缘遮蔽|装拆绝缘遮蔽"
    )
    if forbidden.search(clean):
        return ""
    defect_cues = (
        "绝缘包裹", "护套", "驱鸟器", "占位器", "驱鸟占位器",
        "更换避雷器", "更换直线绝缘子", "紧固直线绝缘子", "清除鸟巢",
    )
    if sum(any(cue in clause for cue in defect_cues) for clause in clauses) < 4:
        return ""

    station_match = re.search(
        r"(?<!\d)((?:35|110|220)kV[^，、；;。\n:：]{1,28}?(?:变电站|站))",
        clean,
    )
    if not station_match:
        return ""
    station = station_match.group(1)
    if station.endswith("站") and not station.endswith("变电站"):
        station = station[:-1] + "变电站"

    lines = v26_unique(v26_line_names_ordered(clean))
    if not lines or len(lines) > 3:
        return ""
    line_labels = [re.sub(r"^(?:10|35|110|220)kV", "", line) for line in lines]

    branches = []
    for clause in clauses:
        anchor = v30_line_context(clause)
        if anchor and re.search(r"分支线|支线|联络支", anchor) and anchor not in branches:
            branches.append(anchor)
    if len(branches) < 2:
        return ""
    return f"{station}：{'、'.join(line_labels)}多条支线进行绝缘化缺陷治理工作"


def v34_distribution_old_new_rebuild_summary(raw_text, profession=""):
    """同一支线的新杆/原杆、新线/旧线成套改造，模式B提升为新旧线路改造主题。"""
    if str(profession or "").strip() not in ("配电", "配网工程", "省管产业", "其他"):
        return ""
    clean = v23_norm(remove_parentheses_notes(raw_text), keep_newlines=True)
    if not (re.search(r"新建?#?\d+|新#?\d+", clean) and re.search(r"原#?\d+|旧导线|旧线", clean)):
        return ""
    action_groups = (
        r"基坑开挖|基础开挖|挖坑",
        r"电杆组立|立杆",
        r"导线开断|断开导线",
        r"导线拆除|旧导线拆除",
        r"导线展放|导线架设",
        r"改接|改移|迁改|迁移",
        r"搭接",
    )
    if sum(bool(re.search(pattern, clean)) for pattern in action_groups) < 4:
        return ""
    if any(word in clean for word in (
        "变压器", "JP柜", "环网柜", "环网箱", "箱变", "分支箱",
        "出线柜", "电缆终端", "融合断路器", "PT", "FTU",
    )):
        return ""

    lines = v26_unique(v26_line_names_ordered(clean))
    if len(lines) != 1:
        return ""
    main_line = lines[0]
    branches = []
    for clause in v31_split_work_clauses(clean):
        match = re.search(
            rf"{re.escape(main_line)}([^，、；;。\n:：]{{1,36}}?)(?=(?:新建|新|原|旧)?#?\d+(?:[-+]\d+)*(?:号)?杆)",
            clause,
        )
        if not match:
            continue
        branch = match.group(1).strip("，、；;。. :：")
        branch = re.sub(r"分支$", "分支线", branch)
        branch = re.sub(r"(?<!分)支$", "支线", branch)
        if branch and branch not in branches:
            branches.append(branch)
    if len(branches) != 1:
        return ""
    return f"{main_line}{branches[0]}线路新旧线路改造"


def v34_raw_mode_b_summary(raw_text, profession=""):
    return (
        v34_live_multibranch_defect_summary(raw_text, profession)
        or v34_distribution_old_new_rebuild_summary(raw_text, profession)
    )


def v35_source_grounded_summary(raw_text, profession, candidate, fallback=""):
    """审核案例可精确复用；其他候选若出现无来源性质，则保守回退。"""
    source_for_summary = audit_source_entry_text(raw_text)
    result = str(candidate or "").replace("提接", "T接").strip("，、；;。. ")
    if not raw_text or not V35_REVIEW_LIBRARY_AVAILABLE:
        return result
    try:
        approved = v35_exact_approved_summary(raw_text, profession)
        if approved:
            return approved.replace("提接", "T接").rstrip("。")
        # V3.4主题层本身已经包含长度、专业、动作白名单、排除动作、
        # 多支线/新旧线路身份等严格证据，不再由简化词表重复否决。
        validated_theme = v34_raw_mode_b_summary(source_for_summary, profession)
        if validated_theme and result == validated_theme.rstrip("。"):
            return result
        violations = v35_provenance_violations(source_for_summary, result)
        if violations:
            log_red(
                f"  【模式B来源证明拒绝】专业={profession}，"
                f"无当前原文来源：{'、'.join(violations)}；已回退为完整摘要"
            )
            safe_fallback = audit_source_entry_text(
                fallback or source_for_summary
            ).strip("，、；;。. ")
            if v35_provenance_violations(source_for_summary, safe_fallback):
                safe_fallback = remove_parentheses_notes(source_for_summary).strip("，、；;。. ")
            return safe_fallback
    except Exception as error:
        log(f"  模式B审核语料库暂不可用，继续沿用V3.4规则：{error}")
    return result


def v29_leader_summary_with_context(mode_a_text, profession, association=None, raw_text=None):
    special = v34_raw_mode_b_summary(raw_text, profession) if raw_text else ""
    if special:
        return v35_source_grounded_summary(raw_text, profession, special, mode_a_text)
    base = v30_leader_summary_from_a(mode_a_text, profession).rstrip("。")
    if not association or not association.get("label"):
        return v35_source_grounded_summary(raw_text, profession, base, mode_a_text)
    label = association["label"]
    if label.endswith("工作"):
        label = label[:-2]
    if base.startswith("配合"):
        base = re.sub(r"^配合[^；;。]+[；;]?", "", base).strip("，、；; ")

    # 同线路关联时，把“配合谁”与“本专业干什么”串成一句，避免重复线路名。
    src_lines = v26_line_names_ordered(mode_a_text)
    first_line = src_lines[0] if src_lines else ""
    parts = [p.strip("，、；;。. ") for p in re.split(r"[；;]", base) if p.strip("，、；;。. ")]
    if parts and first_line and first_line in label and parts[0].startswith(first_line):
        tail = parts[0][len(first_line):].strip("，、；; ")
        head = f"配合{label}"
        if tail:
            head += f"，{tail}"
        parts[0] = head
        return v35_source_grounded_summary(raw_text, profession, "；".join(parts), mode_a_text)
    combined = f"配合{label}；{base}" if base else f"配合{label}"
    return v35_source_grounded_summary(raw_text, profession, combined, mode_a_text)


def v26_leader_low_voltage(source, profession=""):
    s = v23_norm(source)
    actions = v26_extract_lv_actions(s)
    if not actions:
        return ""
    taiqus = v26_taiqu_names(s)
    anchor = taiqus[0] if taiqus else (v26_place_anchor(s) or (v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""))

    # 配网工程里“新放低压出线电缆”反复出现时，说明它是当天建设结果/主题，
    # 领导摘要直接提炼这一建设成果，不再罗列T接点和每段杆号。
    if profession == "配网工程" and "新放低压出线电缆" in s:
        return f"{anchor}新放低压出线电缆" if anchor else "新放低压出线电缆"

    # 普通低压维修按重要性选前2个主要动作；模式A已经完整保留全部动作。
    ranked = sorted(actions, key=lambda item: (-item[1], item[2]))
    take = 2 if len(ranked) >= 3 else len(ranked)
    selected = {item[0] for item in ranked[:take]}
    labels = [item[0] for item in sorted(actions, key=lambda x: x[2]) if item[0] in selected]
    return f"{anchor}{'、'.join(labels)}" if anchor else "、".join(labels)


def v26_leader_old_new_transform(source):
    s = v23_norm(source)
    if not (
        "改接" in s
        and "拆除" in s
        and any(word in s for word in ("台变", "变压器", "JP柜", "配电盘"))
        and any(word in s for word in ("新建", "挖坑立杆", "基础开挖"))
    ):
        return ""
    lines = v26_line_names_ordered(s)
    line = lines[0] if lines else ""
    place_match = re.search(r"新建([^，；。]{2,18}?台变)", s)
    place = place_match.group(1) if place_match else v26_place_anchor(s)
    install_devices = [word for word in ("台架", "变压器", "JP柜") if word in s]
    remove_devices = [word for word in ("变压器", "JP柜", "配电盘", "低压柜") if re.search(rf"{re.escape(word)}[^；。]{{0,20}}拆除|{re.escape(word)}(?:/|、)[^；。]{{0,20}}拆除", s)]
    if not remove_devices:
        if "配电盘拆除" in s:
            remove_devices.append("配电盘")
        if re.search(r"变压器[/、]配电盘拆除|变压器[^；。]{0,12}拆除", s):
            remove_devices.insert(0, "变压器")
    prefix = f"{line}{place}" if place and place not in line else (line or place)
    first = ("新建：电杆基础开挖、组立" if place and place.endswith("台变") else "新建台变：电杆基础开挖、组立")
    if install_devices:
        first += "，" + "/".join(v26_unique(install_devices)) + "安装"
    second_parts = []
    if "高压导线" in s and "低压电缆" in s and "改接" in s:
        second_parts.append("高低压线路改接")
    elif "改接" in s:
        second_parts.append("线路改接")
    if remove_devices:
        second_parts.append("原" + "/".join(v26_unique(remove_devices)) + "拆除")
    return f"{prefix}{first}；{'、'.join(second_parts)}".strip("；")


def v26_leader_newbuild(source):
    s = v23_norm(source)
    if "联络" in s:
        return ""
    if not (
        ("基础开挖" in s or "挖坑" in s)
        and ("组立" in s or "立杆" in s)
    ):
        return ""
    lines = v26_line_names_ordered(s)
    line = lines[0] if lines else ""
    parts = ["新建电杆基础开挖、组立"]
    if "金具安装" in s:
        parts.append("金具安装")
    devices = [word for word in ("台架", "变压器", "JP柜") if word in s]
    if devices:
        parts.append("、".join(devices))
    if "低压导线展放" in s:
        parts.append("低压导线展放")
    elif "低压导线架设" in s:
        parts.append("低压导线架设")
    elif "导线展放" in s:
        parts.append("导线展放")
    elif "导线架设" in s:
        parts.append("导线架设")
    return f"{line}{'、'.join(parts)}" if line else "、".join(parts)


def v26_leader_fusion_install(source):
    s = v23_norm(source)
    if not ("融合断路器" in s or ("PT" in s and "FTU" in s)):
        return ""
    clauses = [item.strip("，、；;。. ") for item in re.split(r"[；;。]", s) if item.strip("，、；;。. ")]
    outputs = []
    for clause in clauses:
        if not ("融合断路器" in clause or ("PT" in clause and "FTU" in clause)):
            continue
        lines = v26_line_names_ordered(clause)
        line = lines[0] if lines else ""
        branch = ""
        if line:
            tail = clause[clause.find(line) + len(line):]
            match = re.match(r"([^，、；;。]{1,24}?(?:支线|分支线|支))", tail)
            if match:
                branch = match.group(1)
        action = "融合开关/PT/FTU安装"
        if "电杆组立" in clause or "挖坑立杆" in clause:
            action = "电杆组立、" + action
        label = f"{line}{branch}{action}" if (line or branch) else action
        if label not in outputs:
            outputs.append(label)
    return "；".join(outputs)


def v26_leader_linkage(source):
    s = v23_norm(source)
    if "联络" not in s:
        return ""
    lines = v26_line_names_ordered(s)
    anchor = "、".join(lines[:2]) if lines else "联络工程"
    actions = []
    if "管塔组立" in s:
        actions.append("管塔组立")
    if any(word in s for word in ("电杆组立", "挖坑立杆", "基础开挖")):
        actions.append("电杆基础开挖、组立")
    if "融合断路器" in s or ("PT" in s and "FTU" in s):
        actions.append("融合开关/PT/FTU安装")
    if "导线架设" in s:
        actions.append("导线架设")
    elif "导线展放" in s:
        actions.append("导线展放")
    if "电缆敷设" in s:
        actions.append("电缆敷设")
    if "环网柜" in s and "安装" in s:
        actions.append("环网柜安装")
    return f"{anchor}联络工程：{'、'.join(v26_unique(actions))}" if actions else ""


def v26_leader_box_relation(source):
    s = v23_norm(source)
    if not ("#1箱变" in s and "#2箱变" in s):
        return ""
    place = v26_place_anchor(s)
    prefix = f"{place}：" if place else ""
    parts = []
    if "#2箱变" in s and ("就位" in s or "吊装" in s):
        parts.append("#2箱变就位")
    if "改接" in s:
        relation = "#1箱变进线电缆改接#2箱变"
        if "终端制作" in s or "电缆终端制作" in s:
            relation += "、终端制作压接"
        parts.append(relation)
    if "电缆敷设" in s:
        relation = "#2箱变至#1箱变电缆敷设"
        if "终端制作" in s or "两侧终端制作" in s:
            relation += "、终端制作压接"
        parts.append(relation)
    return prefix + "，".join(v26_unique(parts)) if parts else ""


def v26_leader_box_general(source):
    s = v23_norm(source)
    if "箱变" not in s:
        return ""
    outputs = []
    lines = v26_line_names_ordered(s)
    if ("融合断路器" in s or "PT" in s or "FTU" in s) and lines:
        outputs.append(f"{lines[0]}融合开关及附件安装")
    site = v26_place_anchor(s)
    if site:
        phrase = "箱变就位"
        if "附件安装" in s or "附件" in s:
            phrase += "、相关附件安装"
        if "低压电缆敷设" in s:
            phrase += "及低压电缆敷设"
        elif "电缆敷设" in s:
            phrase += "及电缆敷设"
        outputs.append(f"{site}{phrase}")
    return "；".join(outputs)


def v26_leader_parallel_terminal(source):
    s = v23_norm(source)
    if s.count("电缆终端制作") < 2:
        return ""
    # 小区/项目 + 多配电室/多变压器共同动作 -> 对象类别共享一次动作。
    match = re.search(r"([^，；。]{2,28}?四期)", s)
    site = match.group(1) if match else v26_place_anchor(s)
    room_ids = v26_unique(re.findall(r"([12一二]号配电室)", s))
    if room_ids and "高压出线柜" in s:
        return f"{site}{'、'.join(room_ids)}变压器高压出线柜/高压侧电缆终端制作、压接"
    if "出线柜" in s:
        return f"{site}出线柜电缆终端制作、压接" if site else "出线柜电缆终端制作、压接"
    return ""


def v26_leader_line_overhaul(source):
    s = v23_norm(source)
    if not ("原" in s and "电缆终端拆除" in s and "导线更换" in s):
        return ""
    lines = v26_line_names_ordered(s)
    anchor = "、".join(lines[:2]) if lines else ""
    parts = ["原电缆终端拆除、新终端制作试验压接"]
    if "隔离开关拆除" in s:
        parts.append("隔离开关拆除")
    if "导线更换" in s:
        parts.append("导线更换")
    if "新电缆固定" in s:
        parts.append("新电缆固定压接")
    return f"{anchor}：{'、'.join(parts)}" if anchor else "、".join(parts)


def v26_leader_peidian(source):
    for function in (
        v26_leader_box_relation,
        v26_leader_line_overhaul,
        v26_leader_parallel_terminal,
        v26_leader_box_general,
    ):
        result = function(source)
        if result:
            return result
    low = v26_leader_low_voltage(source, "配电")
    if low:
        return low
    s = v23_norm(source)
    if len(s) <= 65:
        return v26_leader_strip_station(s)
    compact = v26_strip_pole_refs(v26_leader_strip_station(s))
    return compact.rstrip("，、；;")


def v26_leader_peiwang(source):
    for function in (
        v26_leader_old_new_transform,
        v26_leader_linkage,
        v26_leader_fusion_install,
        v26_leader_newbuild,
    ):
        result = function(source)
        if result:
            return result
    low = v26_leader_low_voltage(source, "配网工程")
    if low:
        return low
    s = v23_norm(source)
    if len(s) <= 70:
        return v26_leader_strip_station(s)
    return v26_strip_pole_refs(v26_leader_strip_station(s)).rstrip("，、；;")


def v26_leader_industry(source):
    s = v23_norm(source)
    if "电缆终端制作" in s and "出线柜" in s:
        site = v26_place_anchor(s)
        return f"{site}出线柜电缆终端制作、压接" if site else "出线柜电缆终端制作、压接"
    return ""


def v26_leader_transmission(source):
    s = v26_leader_strip_station(v23_norm(source))
    # 原文已经短时不为了“缩”而破坏塔号/区间。
    if len(s) <= 70:
        return s
    return v26_strip_pole_refs(s).rstrip("，、；;")


def v26_leader_substation(source):
    s = v23_norm(source)
    s = s.replace("保护屏内#1主变保护装置", "#1主变保护装置")
    s = s.replace("保护屏内保护装置", "保护装置")
    s = re.sub(r"本体清扫[、，]检查", "本体清扫检查", s)
    return s.rstrip("，、；;")


def v26_leader_generic(source):
    s = v23_norm(source)
    if len(s) <= 70:
        return v26_leader_strip_station(s)
    return v26_strip_pole_refs(v26_leader_strip_station(s)).rstrip("，、；;")


def v26_relation_guard(source, candidate):
    """模式B允许删次要动作，但关键拓扑/编号不能被摘要破坏。"""
    s = v23_norm(source)
    c = v23_norm(candidate)
    if not c:
        return False
    # 不得生成源文没有的线路名。
    src_lines = set(v26_line_names_ordered(s))
    cand_lines = set(v26_line_names_ordered(c))
    if cand_lines and not cand_lines.issubset(src_lines):
        return False
    # #1/#2箱变改接属于强拓扑关系，必须完整保留。
    if "#1箱变" in s and "#2箱变" in s and "改接" in s:
        if not ("#1箱变" in c and "#2箱变" in c and "改接" in c):
            return False
    # 若摘要仍提到改接，至少不能把新旧关系写反/写没。
    if "改接" in c and "改接" not in s:
        return False
    # F01/F02、1001/1002等间隔号：摘要一旦描述接入/改接，源文中的关系编号应保留。
    strong_ids = re.findall(r"\bF\d{1,3}\b|(?<!\d)(?:1001|1002|1003|1004)(?!\d)", s, re.I)
    if strong_ids and any(word in c for word in ("接入", "改接")):
        if not all(token.upper() in c.upper() for token in strong_ids):
            return False
    # 任何候选都不应比A层更长。
    if len(c) > len(s) + 4:
        return False
    return True


def v26_leader_summary_from_a(mode_a_text, profession=""):
    """模式B领导总览：输入必须是模式A缩写版，不重新读取原文。"""
    source = str(mode_a_text or "").rstrip("。")
    prof = str(profession or "").strip()
    if not source:
        return ""
    if prof == "营销":
        # 实际营销模式B由build_mode_b_summary_lines走5.2原汇总分支；这里仅作兼容兜底。
        return summarize_marketing_work_for_mode_a(source, compact=True)
    if prof == "带电作业":
        candidate = v26_leader_live(source)
    elif prof == "配电":
        candidate = v26_leader_peidian(source)
    elif prof == "配网工程":
        candidate = v26_leader_peiwang(source)
    elif prof == "省管产业":
        candidate = v26_leader_industry(source) or v26_leader_generic(source)
    elif prof == "输电":
        candidate = v26_leader_transmission(source)
    elif prof == "变电":
        candidate = v26_leader_substation(source)
    else:
        candidate = v26_leader_generic(source)
    candidate = str(candidate or "").strip("，、；;。. ")
    if not v26_relation_guard(source, candidate):
        # 关系守恒失败时回退到A层，宁可稍长也不改事实。
        candidate = source
    return candidate.rstrip("。") + "。"




# ================= V3.0 全语料语义校准层 =================
# 仅修改“文字缩写/提炼/跨专业关联”逻辑。
# Excel预处理、风险表、到岗到位、弹窗、打印、分页、文件生成顺序均不在本层处理。
# 设计依据：2029行真实工作内容 + V2总规范。
# A层：领导可读的技术缩写，仍保动作/强设备/新旧/源目的关系；完整版承担逐杆追溯。
# B层：从A层建立事件框架后做领导总览；允许降定位粒度，但禁止新增动作、改设备、改施工阶段。

V30_ACTION_HINTS = (
    "安装", "更换", "拆除", "制作", "试验", "压接", "敷设", "搭接", "展放", "架设",
    "组立", "开挖", "浇筑", "清理", "检查", "调试", "传动", "升流", "紧固", "改接",
    "迁改", "迁移", "改移", "接入", "断引流线", "接引流线", "遮蔽", "包裹", "绑扎",
)


def v30_has_low_voltage(text):
    s = v23_norm(text)
    return bool(
        "0.4kV" in s
        or "低压" in s
        or "下户线" in s
        or "下户电缆" in s
    )


def v30_line_context(text):
    """主线路 + 必要支线；不带普通杆号。"""
    s = v23_norm(text)
    lines = v26_line_names_ordered(s)
    if not lines:
        return ""
    line = lines[0]
    start = s.find(line) + len(line)
    tail = s[start:start + 45]
    m = re.match(r"([^，、；;。:：]{1,25}?(?:分支线|支线|联络支线|支))(?=#?\d|新|原|旧|$)", tail)
    branch = m.group(1) if m else ""
    return line + branch


def v30_project_suffix_candidate(raw_text, profession=""):
    """工程标题只作定位；若冒号后已有完整施工事实，优先用事实段作为A层候选。"""
    s = v23_norm(v23_classify_parentheses(raw_text)).rstrip("。")
    m = re.search(r"(?:工程|治理|改造|项目)\s*[:：]", s)
    if not m:
        return ""
    suffix = s[m.end():].strip("，、；;。. :：")
    if len(suffix) < 8 or not any(k in suffix for k in V30_ACTION_HINTS):
        return ""
    try:
        cand = v23_compact_work(suffix, profession=str(profession or "").strip(), mode="A").rstrip("。")
    except Exception:
        return ""
    return cand


V31_ADMIN_PREFIX_RE = re.compile(
    r"(^|[，、；;。:：\n（(]|配合|位于|在)"
    r"(?:山东省济宁市泗水县|山东省济宁市|山东济宁泗水|山东济宁|"
    r"济宁市泗水县|济宁泗水|济宁市|泗水县|济宁|泗水)"
    r"(?=[^，、；;。:：\n]{0,24}(?:镇|乡|街道|村|社区|庄|路|台区|台架|配电室|"
    r"箱变|环网柜|(?:10|35|110|220)kV))"
)

V31_FORMAL_PROJECT_CUES = (
    "农村电网", "配电网", "农网", "专项", "改造", "建设", "治理",
    "大修", "技改", "业扩", "增容", "迁改", "改扩建", "投资", "储备",
)

V31_LV_ACTION_RULES = (
    (r"下户线敷设[、，,]?搭接", "下户线敷设、搭接"),
    (r"低压出线电缆敷设[、，,]?搭接", "低压出线电缆敷设、搭接"),
    (r"低压电缆敷设[、，,]?搭接", "低压电缆敷设搭接"),
    (r"裸导线更换(?:为)?绝缘导线", "裸导线更换绝缘导线"),
    (r"下户线制作安装", "下户线制作安装"),
    (r"下户线改接", "下户线改接"),
    (r"下户线更换", "下户线更换"),
    (r"下户线敷设", "下户线敷设"),
    (r"下户线搭接", "下户线搭接"),
    (r"低压出线电缆更换", "低压出线电缆更换"),
    (r"低压出线电缆拆除", "低压出线电缆拆除"),
    (r"低压出线电缆敷设", "低压出线电缆敷设"),
    (r"低压电缆更换", "低压电缆更换"),
    (r"低压电缆拆除", "低压电缆拆除"),
    (r"低压电缆敷设", "低压电缆敷设"),
    (r"低压导线更换|0\.4kV[^，；。]{0,30}?导线更换", "低压导线更换"),
    (r"低压导线展放|0\.4kV[^，；。]{0,30}?导线展放", "低压导线展放"),
    (r"低压导线架设|0\.4kV[^，；。]{0,30}?导线架设", "低压导线架设"),
    (r"低压导线迁改|低压导线迁移|低压导线改移", "低压导线迁改"),
    (r"低压导线搭接", "低压导线搭接"),
)

V31_SIMPLE_SAME_ACTION_RULES = (
    (r"(?:一二次)?融合断路器(?:[/、]PT[/、]FTU|、PT、FTU)?(?:及(?:相关)?附件)?安装", "融合断路器/PT/FTU安装"),
    (r"绑扎固定导线|耐张导线绑扎|绑扎导线", "绑扎固定导线"),
    (r"安装驱鸟器", "安装驱鸟器"),
    (r"安装占位器", "安装占位器"),
    (r"装拆绝缘遮蔽|装、拆绝缘遮蔽", "装拆绝缘遮蔽"),
    (r"装绝缘遮蔽", "装绝缘遮蔽"),
    (r"拆绝缘遮蔽", "拆绝缘遮蔽"),
    (r"绝缘包裹", "绝缘包裹"),
    (r"导线更换", "导线更换"),
    (r"导线展放", "导线展放"),
    (r"电缆敷设", "电缆敷设"),
)


def v31_extract_formal_project_names(text):
    """提取含行政地域字样的正式工程/项目名，供摘要清洗前建立保护区。"""
    s = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    names = []
    for match in re.finditer(r"([^，,；;。\n:：]{1,120}?(?:工程|项目))(?=[:：，,；;。\n]|$)", s):
        candidate = re.sub(r"^\s*\d+\s*[.．、]\s*", "", match.group(1)).strip(" （(，、")
        if not any(word in candidate for word in ("济宁", "泗水")):
            continue
        is_formal = bool(re.search(r"20\d{2}年", candidate)) or any(
            cue in candidate for cue in V31_FORMAL_PROJECT_CUES
        )
        if is_formal and candidate not in names:
            names.append(candidate)
    return names


def v31_strip_admin_regions(text):
    """只删摘要中的行政地域前缀；线路、设备名中的同名文字不做全局替换。"""
    s = str(text or "")
    protected = v31_extract_formal_project_names(s)
    placeholders = {}
    for index, name in enumerate(sorted(protected, key=len, reverse=True)):
        token = f"__V31_PROJECT_{index}__"
        if name in s:
            s = s.replace(name, token)
            placeholders[token] = name
    for _ in range(3):
        cleaned = V31_ADMIN_PREFIX_RE.sub(lambda m: m.group(1), s)
        if cleaned == s:
            break
        s = cleaned
    for token, name in placeholders.items():
        s = s.replace(token, name)
    s = re.sub(r"([，、；;。:：])\1+", r"\1", s)
    return s.strip("，、；;。. ")


def v31_finalize_leader_text(text, original=""):
    """领导摘要统一收口：清行政前缀、恢复正式工程名、清理标点。"""
    candidate = v31_strip_admin_regions(text)
    project_names = v31_extract_formal_project_names(original)
    missing = [name for name in project_names if name not in candidate]
    if missing:
        candidate = f"{'、'.join(missing)}：{candidate}" if candidate else "、".join(missing)
    candidate = re.sub(r"[:：][；;]", "：", candidate)
    candidate = re.sub(r"[，、]{2,}", "、", candidate)
    return candidate.strip("，、；;。. ")


def v31_split_work_clauses(text):
    s = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[（(]\s*新\s*[）)]", "", s)
    try:
        s = strip_mode_b_risk_notes(s)
    except Exception:
        pass
    clauses = []
    for part in re.split(r"[；;。\n]+", s):
        item = re.sub(r"^\s*\d+\s*[.．、]\s*", "", part).strip("，、；;。. ")
        if item:
            clauses.append(item)
    return clauses


def v31_extract_lv_actions(clause):
    s = v23_norm(clause)
    found = []
    residual = s
    for pattern, label in V31_LV_ACTION_RULES:
        match = re.search(pattern, s)
        if match:
            found.append((match.start(), label))
            residual = re.sub(pattern, "", residual)
    labels = [label for _pos, label in sorted(found, key=lambda item: item[0])]
    labels = v26_unique(labels)
    suppress = {
        "下户线敷设、搭接": {"下户线敷设", "下户线搭接"},
        "低压出线电缆敷设、搭接": {"低压出线电缆敷设"},
        "低压电缆敷设搭接": {"低压电缆敷设"},
        "裸导线更换绝缘导线": {"低压导线更换"},
    }
    removed = set()
    for composite, children in suppress.items():
        if composite in labels:
            removed.update(children)
    labels = [label for label in labels if label not in removed]
    # 已识别一部分、却仍残留其他施工动词时不冒险合并，交回旧摘要链路。
    if re.search(r"更换|敷设|搭接|改接|制作|安装|展放|架设|迁改|迁移|改移|拆除", residual):
        return []
    return labels


def v33_normalize_lv_device_name(text):
    """只归一低压设备名称的类别尾词，地点、村名、方位和设备编号完整保留。"""
    device = v23_norm(text).strip("，、；;。. :：")
    device = re.sub(r"(?:台区){2,}$", "台区", device)
    device = re.sub(
        r"台架(?:变压器|配变|台变|变)?(?:台区)+$",
        "台架",
        device,
    )
    device = re.sub(r"(?:配电)?变压器(?:台区)+$", "台区", device)
    device = re.sub(r"(?:公变|配变|台变)(?:台区)+$", "台区", device)
    return device


def v31_extract_lv_device(clause):
    s = v31_strip_admin_regions(v23_norm(clause))
    s = re.sub(r"[（(]\s*新\s*[）)]", "", s)
    # 先在0.4kV标识前取定位段，防止把“0.4kVⅠ支线”误当成10kV主线。
    scope = re.split(r"0\.4kV|低压", s, maxsplit=1)[0]
    scope = re.sub(
        r"^(?:(?:35|110|220)kV[^，、；;。:：]{1,32}?(?:变电)?站)[:：]?",
        "",
        scope,
    )
    line_matches = list(re.finditer(
        r"(?<!\d)(?:10|35|110|220)kV(?:(?!(?:0\.4|10|35|110|220)kV)[^，、；;。:：]){1,42}?线",
        scope,
    ))
    if line_matches:
        scope = scope[line_matches[-1].end():]
    else:
        scope = re.sub(r"^(?:10|35|110|220)kV", "", scope)
    scope = re.sub(r"^(?:[^，、；;。:：]{1,28}?(?:分支线|支线))", "", scope)
    scope = re.sub(r"^#?\d+(?:[-+]\d+)*(?:号)?杆", "", scope)
    scope = scope.split("：")[-1].strip("，、；;。. :：")
    match = re.search(
        r"([^，、；;。:：]{1,64}?(?:台架台区|台架|变压器台区|变台区|台区|台变|公变|配变))$",
        scope,
    )
    if not match:
        return ""
    device = v33_normalize_lv_device_name(match.group(1))
    return v31_strip_admin_regions(device)


def v31_extract_lv_branch(clause):
    s = v23_norm(clause)
    match = re.search(r"0\.4kV\s*([ⅠⅡⅢⅣⅤIVX0-9A-Za-z#-]+)\s*支线", s)
    if match:
        return f"{match.group(1)}支线"
    if "0.4kV" in s and "支线" in s:
        return "支线"
    return ""


def v31_try_merge_low_voltage_same_action(raw_text, profession=""):
    """按设备+支线身份去重，再将完全相同的低压动作后置一次。"""
    prof = str(profession or "").strip()
    if prof not in ("配电", "配网工程", "带电作业", "省管产业", "其他"):
        return ""
    clauses = v31_split_work_clauses(raw_text)
    lv_clauses = [
        clause for clause in clauses
        if "0.4kV" in clause or "低压" in clause or "下户线" in clause
    ]
    if not lv_clauses:
        return ""
    # 同一行若还含高压施工事实，不用低压专用模板吞掉它。
    for clause in clauses:
        if clause in lv_clauses or any(name in clause for name in v31_extract_formal_project_names(raw_text)):
            continue
        if any(hint in clause for hint in V30_ACTION_HINTS):
            return ""

    entries = []
    for clause in lv_clauses:
        device = v31_extract_lv_device(clause)
        actions = tuple(v31_extract_lv_actions(clause))
        branch = v31_extract_lv_branch(clause)
        if not device or not actions:
            return ""
        entry = (device, branch, actions)
        if entry not in entries:
            entries.append(entry)
    if not entries:
        return ""

    groups = {}
    order = []
    for device, branch, actions in entries:
        if actions not in groups:
            groups[actions] = {"devices": [], "branch_ids": [], "entry_count": 0}
            order.append(actions)
        groups[actions]["entry_count"] += 1
        if device not in groups[actions]["devices"]:
            groups[actions]["devices"].append(device)
        branch_id = (device, branch)
        if branch and branch_id not in groups[actions]["branch_ids"]:
            groups[actions]["branch_ids"].append(branch_id)

    outputs = []
    for actions in order:
        group = groups[actions]
        branch_ids = group["branch_ids"]
        if len(branch_ids) >= 2 and len(branch_ids) == group["entry_count"]:
            voltage_scope = "0.4kV多条支线"
        elif len(branch_ids) == 1 and group["entry_count"] == 1:
            voltage_scope = f"0.4kV{branch_ids[0][1]}"
        else:
            voltage_scope = "0.4kV"
        outputs.append(f"{'、'.join(group['devices'])}{voltage_scope}{'、'.join(actions)}")
    candidate = "；".join(outputs)
    return v31_finalize_leader_text(candidate, original=raw_text)


def v31_try_merge_simple_same_action(raw_text, profession=""):
    """10kV/带电同动作合并；存在新旧、先后或强拓扑关系时立即退出。"""
    clauses = v31_split_work_clauses(raw_text)
    parsed = []
    action_clause_count = 0
    forbidden = re.compile(
        r"原|旧|新设备|断引流线|接引流线|断开|接通|拆除|改接|迁改|迁移|"
        r"至|环网柜|环网箱|箱变|出线柜|电缆头|终端|\bF\d+\b|#1|#2"
    )
    for clause in clauses:
        matches = []
        for pattern, label in V31_SIMPLE_SAME_ACTION_RULES:
            match = re.search(pattern, clause)
            if match:
                matches.append((match.start(), match.end(), label))
        labels = v26_unique(label for _start, _end, label in matches)
        if not labels:
            if any(hint in clause for hint in V30_ACTION_HINTS):
                return ""
            continue
        action_clause_count += 1
        if len(labels) != 1 or forbidden.search(clause):
            return ""
        first = min(matches, key=lambda item: item[0])
        action = labels[0]
        target = clause[:first[0]].strip("，、；;。. :：")
        tail = clause[first[1]:].strip("，、；;。. :：")
        if tail and any(hint in tail for hint in V30_ACTION_HINTS):
            return ""
        target = v31_strip_admin_regions(v26_strip_pole_refs(target)).strip("，、；;。. :：")
        if not target:
            return ""
        voltages = tuple(v26_unique(re.findall(r"(?:10|35|110|220)kV", clause)))
        parsed.append((target, action, voltages))
    if action_clause_count < 2 or len(parsed) < 2:
        return ""
    actions = {item[1] for item in parsed}
    voltage_sets = {item[2] for item in parsed}
    if len(actions) != 1 or len(voltage_sets) != 1:
        return ""
    action = parsed[0][1]
    if str(profession or "").strip() == "带电作业" and action not in {
        "绑扎固定导线", "安装驱鸟器", "安装占位器", "装绝缘遮蔽",
        "拆绝缘遮蔽", "装拆绝缘遮蔽", "绝缘包裹",
    }:
        return ""
    targets = v26_unique(item[0] for item in parsed)
    if len(targets) < 2:
        return ""
    return v31_finalize_leader_text(f"{'、'.join(targets)}{action}", original=raw_text)


def v31_semantically_complete(source, candidate):
    c = str(candidate or "").strip("，、；;。. ")
    if not c:
        return False
    dangling = (
        "下户", "低压导", "电缆敷", "断路器安", "搭接新", "导线更",
        "电缆更", "终端制", "引流", "绝缘遮", "台架台",
    )
    if c.endswith(dangling):
        return False
    s = str(source or "")
    if any(hint in s for hint in V30_ACTION_HINTS):
        has_action = any(hint in c for hint in V30_ACTION_HINTS)
        has_action = has_action or bool(v31_extract_lv_actions(c))
        if not has_action:
            return False
    return True


def v30_mode_a_abbrev(raw_text, profession=""):
    """A层沿用V2.9为主，仅修复全语料验证出的“凭空加低压/丢管塔”语义错误。"""
    prof = str(profession or "").strip()
    if prof == "营销":
        marketing = summarize_marketing_work_for_mode_a(raw_text, compact=False).rstrip("。")
        return v31_finalize_leader_text(marketing, original=raw_text).rstrip("。") + "。"
    special = v31_try_merge_low_voltage_same_action(raw_text, prof)
    if not special:
        special = v31_try_merge_simple_same_action(raw_text, prof)
    if special and v31_semantically_complete(raw_text, special):
        return special.rstrip("。") + "。"
    base = v26_mode_a_abbrev(raw_text, profession=prof).rstrip("。")
    raw_base = v23_compact_work(raw_text, profession=prof, mode="A").rstrip("。")
    # 原文没有低压语义时，禁止把普通导线展放/架设升级成“低压导线”。
    if not v30_has_low_voltage(raw_text):
        base = base.replace("低压导线展放", "导线展放").replace("低压导线架设", "导线架设")
        base = base.replace("低压导线更换", "导线更换")
    # 规范化一次到位，避免二次运行继续缩短。
    base = base.replace("电杆基础开挖、电杆组立", "挖坑立杆")
    if ("新建杆导线瓷瓶绑扎" in v23_norm(raw_text) or re.search(r"(?:新建|新)#?\d+[^，；。]{0,12}导线瓷瓶绑扎", v23_norm(raw_text))):
        if "新建杆导线瓷瓶绑扎" not in base and "导线瓷瓶绑扎" in base:
            base = base.replace("导线瓷瓶绑扎", "新建杆导线瓷瓶绑扎")
    # 管塔是强设备；旧A层的“电杆新建”模板不能吞掉管塔组立。
    if "管塔组立" in raw_base and "管塔组立" not in base:
        base = raw_base
    if "钢管塔" in raw_base and "组立" in raw_base and "钢管塔" not in base:
        base = raw_base
    base = v31_finalize_leader_text(v26_mode_a_formalize(base), original=raw_text)
    if not v31_semantically_complete(raw_text, base):
        base = v31_finalize_leader_text(raw_base, original=raw_text)
    return base.rstrip("。") + "。"


def v30_action_events(text):
    """按原文位置提取领导摘要需要的业务动作；只归一表达，不新增动作。"""
    s = v23_norm(text)
    events = []
    def add(pos, label, key=None):
        label = str(label or "").strip("，、；;。. ")
        if not label:
            return
        k = key or label
        if any(x[2] == k for x in events):
            return
        events.append((max(0, int(pos)), label, k))

    # 绝缘遮蔽必须精确区分“装 / 拆 / 装拆”，不能凭空补动作。
    for m in re.finditer(r"装拆绝缘遮蔽|装、拆绝缘遮蔽|装绝缘遮蔽|拆绝缘遮蔽", s):
        raw = m.group(0)
        if raw in ("装拆绝缘遮蔽", "装、拆绝缘遮蔽"):
            add(m.start(), "装拆绝缘遮蔽", "insulation_both")
        elif raw.startswith("装"):
            add(m.start(), "装绝缘遮蔽", "insulation_install")
        else:
            add(m.start(), "拆绝缘遮蔽", "insulation_remove")

    patterns = [
        (r"断、接引流线|断接引流线", lambda m:"断、接引流线", "disconnect_connect_jumper"),
        (r"带负荷(?:短接)?(?:隔离开关|跌落开关|真空开关)(?:、真空开关)?", lambda m:m.group(0), "live_short"),
        (r"短接(?:隔离开关|跌落开关|真空开关)(?:、真空开关)?", lambda m:m.group(0), "short"),
        (r"带负荷(?:直线杆|直线)改耐张杆", lambda m:m.group(0).replace("直线改", "直线杆改"), "tension"),
        (r"直线杆改耐张杆", lambda m:m.group(0), "tension"),
        (r"断(?:隔离开关|跌落开关|真空开关)(?:、真空开关)?两侧引流线", lambda m:m.group(0), "disconnect_switch_jumper"),
        (r"断(?:旧PT|原PT|PT)引流线", lambda m:m.group(0), "disconnect_pt_jumper"),
        (r"断电缆引流线", lambda m:m.group(0), "disconnect_cable_jumper"),
        (r"断(?:[^，、；;。]{0,8})?引流线", lambda m:m.group(0), "disconnect_jumper"),
        (r"拆除(?:隔离开关|跌落开关|真空开关)", lambda m:m.group(0), "remove_switch"),
        (r"更换(?:跌落开关|隔离开关|真空开关)", lambda m:m.group(0), "replace_switch"),
        (r"安装(?:跌落开关|隔离开关|真空开关)(?:及附件)?", lambda m:m.group(0), "install_switch"),
        (r"(?:接入|接)(?:一二次|一二)?融合断路器", lambda m:re.sub(r"一二次|一二", "", m.group(0)), "connect_fusion"),
        (r"接(?:跌落开关|隔离开关|真空开关)(?:、真空开关)?两侧引流线", lambda m:m.group(0), "connect_switch_jumper"),
        (r"接(?:[^，、；;。]{0,8})?引流线", lambda m:m.group(0), "connect_jumper"),
        (r"绑扎固定导线|耐张导线绑扎|绑扎导线", lambda m:"绑扎固定导线", "bind"),
        (r"更换避雷器引流线", lambda m:m.group(0), "arrester_jumper"),
        (r"更换避雷器", lambda m:m.group(0), "replace_arrester"),
        (r"绝缘包裹(?:并沟线夹|电缆头|真空开关|跌落开关|隔离开关)?", lambda m:m.group(0), "wrap"),
        (r"安装(?:隔离开关护套|真空开关护套|跌落开关护套|并沟线夹护套)(?:、(?:隔离开关护套|真空开关护套|跌落开关护套|并沟线夹护套))*", lambda m:m.group(0), "cover"),
        (r"安装驱鸟器", lambda m:m.group(0), "bird"),
        (r"(?:安装|装)占位器", lambda m:"安装占位器", "occupier"),
    ]
    for pat, fn, key in patterns:
        for m in re.finditer(pat, s):
            add(m.start(), fn(m), key)

    events.sort(key=lambda x:x[0])
    labels = [x[1] for x in events]
    # 只有完全泛化的“断引流线 + 接引流线”才合并；对象不同则保留迁移关系。
    if "断、接引流线" in labels:
        labels = [label for label in labels if label not in ("断引流线", "接引流线")]
    if "更换避雷器引流线" in labels and "更换避雷器" in labels:
        labels.remove("更换避雷器")
    if "断引流线" in labels and "接引流线" in labels:
        di = labels.index("断引流线"); ci = labels.index("接引流线")
        if not any(w in s for w in ("旧PT", "原PT", "跌落开关两侧", "隔离开关两侧", "真空开关两侧")):
            first = min(di,ci); second=max(di,ci)
            labels[first] = "断接引流线"
            labels.pop(second)
    return v26_unique(labels)


def v30_leader_live(source):
    """带电：精确保留自己的核心动作；可省普通杆号，但不省“干什么”。"""
    s = v23_norm(source).rstrip("。")
    special = v26_leader_old_new_line(s)
    if special:
        return special
    clauses = [x.strip("，、；;。. ") for x in re.split(r"[；;。]", s) if x.strip("，、；;。. ")]
    simple_groups = {}
    order = []
    complex_out = []
    for clause in clauses:
        actions = v30_action_events(clause)
        if not actions:
            compact = v26_strip_pole_refs(v26_leader_strip_station(clause)).strip("，、；;。. ")
            if compact:
                complex_out.append(compact)
            continue
        anchor = v30_line_context(clause) or (v26_line_names_ordered(clause)[0] if v26_line_names_ordered(clause) else v26_place_anchor(clause))
        # 单一重复动作允许跨线路共享一次动作；复杂链必须整条保留。
        if len(actions) == 1 and actions[0] in (
            "绑扎固定导线", "安装驱鸟器", "安装占位器", "装绝缘遮蔽", "拆绝缘遮蔽", "装拆绝缘遮蔽",
        ):
            key = actions[0]
            if key not in simple_groups:
                simple_groups[key] = []
                order.append(key)
            if anchor and anchor not in simple_groups[key]:
                simple_groups[key].append(anchor)
        else:
            prefix = anchor
            action_text = "、".join(actions)
            compact = f"{prefix}{action_text}" if prefix else action_text
            if compact and compact not in complex_out:
                complex_out.append(compact)
    outputs = []
    for key in order:
        anchors = simple_groups[key]
        if anchors:
            outputs.append(f"{'、'.join(anchors)}{key}")
    outputs.extend(complex_out)
    return "；".join(v26_unique(outputs))


def v30_leader_low_voltage(source, profession=""):
    s = v23_norm(source)
    if not v30_has_low_voltage(s):
        return ""
    return v26_leader_low_voltage(s, profession)


def v30_leader_conductor(source):
    """非低压导线作业，不得被误写成低压。"""
    s = v23_norm(source)
    if v30_has_low_voltage(s):
        return ""
    if not any(w in s for w in ("导线展放", "导线架设", "导线更换", "旧导线拆除", "裸导线", "挂接紧固导线", "金具安装")):
        return ""
    lines = v26_line_names_ordered(s)
    anchor = "、".join(lines[:2]) if lines else v26_place_anchor(s)
    acts=[]
    if "裸导线" in s and "绝缘导线" in s and "更换" in s:
        acts.append("裸导线更换绝缘导线")
    elif "导线更换" in s:
        acts.append("导线更换")
    if "旧导线拆除" in s or re.search(r"拆除[^，；。]{0,8}旧导线", s):
        acts.append("旧导线拆除")
    if "金具安装" in s:
        acts.append("金具安装")
    if "导线展放" in s:
        acts.append("导线展放")
    elif "导线架设" in s:
        acts.append("导线架设")
    if "挂接紧固导线" in s:
        acts.append("挂接紧固导线")
    if "接引流线" in s:
        acts.append("接引流线")
    return f"{anchor}{'、'.join(v26_unique(acts))}" if anchor else "、".join(v26_unique(acts))


def v30_leader_tower_construction(source):
    """配网工程中的管塔/钢管塔施工独立于普通电杆和导线规则。"""
    s=v23_norm(source)
    if not any(x in s for x in ("管塔组立","钢管塔组立","钢管塔安装")):
        return ""
    line=v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""
    acts=[]
    if "管塔组立" in s: acts.append("管塔组立")
    elif "钢管塔" in s and ("组立" in s or "安装" in s): acts.append("钢管塔组立")
    if "金具" in s and "安装" in s: acts.append("金具及附件安装" if "附件" in s else "金具安装")
    if "导线展放" in s: acts.append("导线展放")
    elif "导线架设" in s: acts.append("导线架设")
    return f"{line}{'、'.join(v26_unique(acts))}" if line else "、".join(v26_unique(acts))


def v30_leader_fusion(source):
    """区分：融合设备施工安装 / 带负荷接入 / 原开关更换三个阶段。"""
    s = v23_norm(source).rstrip("。")
    if "融合断路器" not in s and "一二融合断路器" not in s and not ("PT" in s and "FTU" in s):
        return ""
    anchor = v30_line_context(s) or (v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else "")
    # 带电接入阶段：绝不能补PT/FTU“安装”。
    if "接入" in s and "融合断路器" in s:
        acts=[]
        if "带负荷" in s:
            if "直线杆改耐张" in s or "直线改耐张" in s:
                acts.append("带负荷直线杆改耐张杆")
            else:
                acts.append("带负荷")
        elif "直线杆改耐张" in s or "直线改耐张" in s:
            acts.append("直线杆改耐张杆")
        acts.append("接入融合断路器")
        return f"{anchor}{'、'.join(v26_unique(acts))}" if anchor else "、".join(v26_unique(acts))
    # 旧开关替换为融合设备：保“原→新”的状态。
    if "更换" in s and any(w in s for w in ("原真空开关", "原隔离开关", "原开关", "真空开关更换", "隔离开关更换")):
        old = "原真空开关" if "真空开关" in s else ("原隔离开关" if "隔离开关" in s else "原开关")
        new = "融合断路器/PT/FTU" if ("PT" in s and "FTU" in s) else "融合断路器"
        return f"{anchor}{old}更换{new}" if anchor else f"{old}更换{new}"
    # 施工安装阶段要求原文确有“安装”。
    if "安装" in s and ("PT" in s or "FTU" in s or "融合断路器" in s):
        acts=[]
        if ("基础开挖" in s or "挖坑" in s) and ("组立" in s or "立杆" in s):
            acts.append("挖坑立杆")
        elif "电杆组立" in s:
            acts.append("电杆组立")
        device = "融合断路器/PT/FTU安装" if ("PT" in s and "FTU" in s) else "融合断路器安装"
        acts.append(device)
        return f"{anchor}{'、'.join(acts)}" if anchor else "、".join(acts)
    return ""


def v30_leader_newbuild(source):
    s = v23_norm(source)
    if "联络" in s:
        return ""
    if not (("基础开挖" in s or "杆坑开挖" in s or "挖坑" in s) and ("组立" in s or "立杆" in s)):
        return ""
    line = v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""
    obj = v26_taiqu_names(s)[0] if v26_taiqu_names(s) else v26_place_anchor(s)
    # 过长/明显工程标题型地点不作为对象。
    if obj and len(obj) > 20:
        obj = ""
    anchor = line + (obj if obj and obj not in line else "")
    acts=["新建电杆基础开挖、组立"]
    if "金具安装" in s: acts.append("金具安装")
    devices=[]
    for w in ("台架", "变压器", "JP柜"):
        if w in s: devices.append(w)
    if not devices and "配电设备" in s:
        devices.append("配电设备")
    if devices: acts.append("、".join(v26_unique(devices)) + "安装")
    if "低压电缆" in s and "敷设" in s: acts.append("低压电缆敷设")
    if "低压导线展放" in s: acts.append("低压导线展放")
    elif "低压导线架设" in s: acts.append("低压导线架设")
    elif "导线展放" in s: acts.append("导线展放")
    elif "导线架设" in s: acts.append("导线架设")
    return f"{anchor}{'，'.join(v26_unique(acts))}" if anchor else "，".join(v26_unique(acts))


def v30_leader_linkage(source):
    """联络工程按实际施工阶段提炼；基础施工绝不能臆造成组立。"""
    s = v23_norm(source)
    if "联络" not in s:
        return ""
    lines = v26_line_names_ordered(s)
    anchor = "、".join(lines[:2]) if lines else "联络工程"
    acts=[]
    # 管塔/铁塔基础与组立严格分开。
    if "管塔基础开挖" in s:
        acts.append("管塔基础开挖")
    if re.search(r"管塔[^，；。]{0,12}基础制作|管塔基础制作", s) or ("管塔" in s and "基础制作" in s):
        acts.append("管塔基础制作")
    if "管塔基础浇筑" in s or ("管塔" in s and "基础浇筑" in s):
        acts.append("管塔基础浇筑")
    if "管塔组立" in s:
        acts.append("管塔组立")
    if "钢管塔" in s and "组立" in s:
        acts.append("钢管塔组立")
    if "电杆基础开挖" in s or ("基础开挖" in s and "管塔" not in s and "塔基础" not in s):
        if "组立" in s or "立杆" in s:
            acts.append("挖坑立杆")
        else:
            acts.append("电杆基础开挖")
    if "金具安装" in s:
        acts.append("金具安装")
    if "导线展放" in s:
        acts.append("导线展放")
    elif "导线架设" in s:
        acts.append("导线架设")
    if "电缆沟" in s and "检查" in s:
        acts.append("电缆沟检查")
    if "电缆沟" in s and "清理" in s:
        acts.append("电缆沟清理")
    if "电缆接头柜" in s and "安装" in s:
        acts.append("电缆接头柜安装")
    if ("电缆头制作" in s or "电缆终端制作" in s):
        if "环网柜" in s and "接头柜" in s:
            acts.append("接头柜/环网柜电缆头制作")
        else:
            acts.append("电缆终端制作")
    fusion = v30_leader_fusion(s)
    if fusion:
        # 只取融合动作尾部，避免重复线路。
        tail=fusion
        for line in lines:
            tail=tail.replace(line, "", 1)
        tail=tail.strip("，、；;。. ")
        if tail: acts.append(tail)
    if "环网柜" in s and "安装" in s and "电缆接头柜" not in s:
        acts.append("环网柜安装")
    if not acts:
        return ""
    return f"{anchor}联络工程：{'、'.join(v26_unique(acts))}"


def v30_leader_transformer(source):
    s = v23_norm(source)
    if not any(w in s for w in ("变压器", "JP柜", "配电盘", "母线", "台架")):
        return ""
    # 新建多阶段优先交新建/旧新迁改专用逻辑。
    oldnew = v26_leader_old_new_transform(s)
    if oldnew:
        return oldnew
    newbuild = v30_leader_newbuild(s)
    if newbuild:
        return newbuild
    line = v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""
    obj = v26_taiqu_names(s)[0] if v26_taiqu_names(s) else v26_place_anchor(s)
    if obj and len(obj)>22: obj=""
    anchor = line + (obj if obj and obj not in line else "")
    acts=[]
    # 同动作设备共享一次动作。
    for verb in ("更换", "安装", "拆除"):
        devs=[]
        for d in ("变压器", "JP柜", "母线", "配电盘", "低压柜", "出线电缆", "台架"):
            if d in s:
                # A层大多已把同动作设备并列；领导层保守地只在全文含该动词时归组。
                if verb in s and d not in devs:
                    devs.append(d)
        if devs:
            prefix = "原" if verb == "拆除" and ("原" in s or "旧" in s) else ""
            acts.append(prefix + "/".join(devs) + verb)
            break
    if "附件安装" in s and "安装" not in "".join(acts):
        acts.append("附件安装")
    if not acts:
        return ""
    return f"{anchor}{'、'.join(acts)}" if anchor else "、".join(acts)


def v30_compact_cable_relation(source):
    s = v23_norm(source)
    # 环网箱/柜 F间隔 -> 新柜/配电室 的明确源目的关系。
    m = re.search(
        r"(?P<src>(?:10kV[^，；。]{0,55}?线)?[^，；。]{0,40}?(?:环网箱|环网柜)F\d{1,3}间隔)至(?:新建)?(?P<dst>(?:10kV[^，；。]{0,55}?线)?[^，；。]{0,45}?(?:环网箱|环网柜|配电室)(?:#?\d+进线柜|基础)?)",
        s,
    )
    if m:
        src=m.group('src'); dst=m.group('dst')
        lines=v26_line_names_ordered(src)
        line=lines[0] if lines else ""
        if line and line in dst:
            dst=dst.replace(line,"",1)
        acts=[]
        if "电缆通道清理" in s: acts.append("电缆通道清理")
        if "电缆敷设" in s: acts.append("电缆敷设")
        if "电缆终端制作" in s or "终端制作" in s: acts.append("终端制作压接" if "压接" in s else "终端制作")
        return f"{src}至{dst}：{'、'.join(v26_unique(acts))}" if acts else f"{src}至{dst}"
    return ""


def v30_leader_cable_rmu(source):
    s = v23_norm(source)
    box = v26_leader_box_relation(s)
    if box:
        return box
    rel = v30_compact_cable_relation(s)
    if rel:
        return rel
    if "环网柜" in s and "更换为高标准环网柜" in s:
        line=v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""
        m=re.search(r"([^，；。]{2,28}?环网柜)更换为高标准环网柜", s)
        obj=m.group(1) if m else "环网柜"
        if line and obj.startswith(line): obj=obj[len(line):]
        out=f"{line}{obj}更换为高标准环网柜" if line else f"{obj}更换为高标准环网柜"
        if "原顺序接入新环网柜" in s or ("原环网柜" in s and "接入新环网柜" in s):
            out += "，原各间隔电缆顺序接入新柜"
        if "电缆终端制作" in s:
            out += "、终端制作试验压接" if "试验" in s and "压接" in s else "、终端制作压接"
        return out
    # 有强编号间隔时保留原关系，仅压缩“制作、试验、压接”。
    if any(w in s for w in ("环网柜", "环网箱", "接头柜", "电缆终端", "电缆敷设", "电缆改接", "配网电缆接入")):
        c=s.replace("电缆终端制作、试验、压接", "电缆终端制作试验压接")
        c=c.replace("终端制作、试验、压接", "终端制作试验压接")
        c=v26_leader_strip_station(c)
        if len(c)<=110:
            return c
    return ""


def v30_leader_peidian(source):
    s=v23_norm(source)
    if "#1箱变" in s and "#2箱变" in s:
        x=v30_leader_cable_rmu(s)
        if x: return x
    if "箱变" in s:
        x=v26_leader_box_general(s)
        if x: return x
    x=v30_leader_cable_rmu(s)
    if x: return x
    tr=v30_leader_transformer(s)
    if tr: return tr
    lv=v30_leader_low_voltage(s,"配电")
    if lv: return lv
    cond=v30_leader_conductor(s)
    if cond: return cond
    # 短作业不硬缩；长作业只去站名/普通杆号。
    if len(s)<=72: return v26_leader_strip_station(s)
    return v26_strip_pole_refs(v26_leader_strip_station(s)).rstrip("，、；;")


def v30_leader_peiwang(source):
    s=v23_norm(source)
    link=v30_leader_linkage(s)
    if link: return link
    tower=v30_leader_tower_construction(s)
    if tower: return tower
    fusion=v30_leader_fusion(s)
    if fusion: return fusion
    tr=v30_leader_transformer(s)
    if tr: return tr
    new=v30_leader_newbuild(s)
    if new: return new
    lv=v30_leader_low_voltage(s,"配网工程")
    if lv: return lv
    cable=v30_leader_cable_rmu(s)
    if cable: return cable
    cond=v30_leader_conductor(s)
    if cond: return cond
    if len(s)<=78: return v26_leader_strip_station(s)
    return v26_strip_pole_refs(v26_leader_strip_station(s)).rstrip("，、；;")


def v30_leader_substation(source):
    s=v23_norm(source).rstrip("。")
    # 保护/主变类保留动作族但去重复“保护装置”等词。
    if any(w in s for w in ("保护装置", "保护试验", "保护调试", "开关传动", "CT一次升流", "回路电阻试验")):
        prefix=""
        m=re.match(r"((?:35|110|220)kV[^：:]{1,28}(?:站|变电站))[:：]?", s)
        if m:
            prefix=m.group(1).replace("变电站","站")+"："
            body=s[m.end():].lstrip("：:")
        else:
            body=s
        body=body.replace("保护屏内保护装置","保护装置").replace("保护屏内#1主变保护装置","#1主变保护装置")
        body=body.replace("本体清扫、检查","本体清扫检查")
        return (prefix+body).rstrip("，、；;")
    return v26_leader_substation(s)


def v30_customer_name(source):
    s=v23_norm(source)
    # 业扩/验收对象：从线路之后、业扩之前提取，保公司/单位名。
    m=re.search(r"(?:10kV[^，；。:：]{1,42}?线)([^，；。:：]{2,48}?)(?:业扩(?:增容|新装)?工程|业扩|竣工验收)", s)
    if m:
        return m.group(1).strip("，、；;。. ")
    return ""


def v30_leader_industry(source):
    s=v23_norm(source)
    if "出线柜" in s and "电缆终端制作" in s:
        site=v26_place_anchor(s)
        if site:
            return f"{site}出线柜电缆终端制作、压接"
        return "出线柜电缆终端制作、压接"
    cable=v30_leader_cable_rmu(s)
    if cable: return cable
    if "业扩" in s or "竣工验收" in s:
        line=v26_line_names_ordered(s)[0] if v26_line_names_ordered(s) else ""
        customer=v30_customer_name(s)
        if "更换" in s and "电流互感器" in s:
            action="电流互感器更换、二次回路检查" if "二次回路" in s else "电流互感器更换"
        elif "增容" in s and "竣工验收" in s:
            action="业扩增容竣工验收"
        elif "新装" in s and "竣工验收" in s:
            action="业扩新装竣工验收"
        elif "竣工验收" in s:
            action="业扩竣工验收"
        else:
            action="业扩工作"
        return f"{line}{customer}{action}" if (line or customer) else action
    if "电缆终端制作" in s and "出线柜" in s:
        site=v26_place_anchor(s)
        return f"{site}出线柜电缆终端制作、压接" if site else "出线柜电缆终端制作、压接"
    return v26_leader_generic(s)


def v30_leader_transmission(source):
    """输电领导摘要保塔号/区间/源目的关系；A层已做站名清理，这里不再删杆塔定位。"""
    s=v23_norm(source).rstrip("。")
    return s.rstrip("，、；;")


def v30_relation_guard(source, candidate):
    s=v23_norm(source); c=v23_norm(candidate)
    if not c: return False
    src_lines=set(v26_line_names_ordered(s)); cand_lines=set(v26_line_names_ordered(c))
    if cand_lines and not cand_lines.issubset(src_lines): return False
    # 禁止凭空增加低压属性。
    if "低压导线" in c and not v30_has_low_voltage(s): return False
    # 绝缘遮蔽装/拆状态精确守恒。
    if "装拆绝缘遮蔽" in c and not ("装拆绝缘遮蔽" in s or "装、拆绝缘遮蔽" in s or ("装绝缘遮蔽" in s and "拆绝缘遮蔽" in s)):
        return False
    if "装绝缘遮蔽" in c and "装绝缘遮蔽" not in s and "装拆绝缘遮蔽" not in s and "装、拆绝缘遮蔽" not in s:
        return False
    if "拆绝缘遮蔽" in c and "拆绝缘遮蔽" not in s and "装拆绝缘遮蔽" not in s and "装、拆绝缘遮蔽" not in s:
        return False
    # 融合设备施工阶段守恒：接入不能改成安装，原开关更换不能改成新装。
    if "融合断路器/PT/FTU安装" in c and not ("安装" in s and "PT" in s and "FTU" in s): return False
    if "接入融合断路器" in c and "接入" not in s: return False
    if "直线杆改耐张杆" in s and "接入" in s and not ("直线杆改耐张杆" in c and "接入融合断路器" in c): return False
    # 基础/组立不可互相臆造。
    if "管塔组立" in c and "管塔组立" not in s: return False
    if "挖坑立杆" in c and not (("基础开挖" in s or "挖坑" in s) and ("组立" in s or "立杆" in s)): return False
    # 强拓扑关系。
    if "#1箱变" in s and "#2箱变" in s and "改接" in s and not ("#1箱变" in c and "#2箱变" in c and "改接" in c): return False
    for token in re.findall(r"\bF\d{1,3}\b|(?<!\d)(?:1001|1002|1003|1004)(?!\d)", s, re.I):
        if any(w in c for w in ("接入","改接","至")) and token.upper() not in c.upper(): return False
    if "改接" in c and "改接" not in s: return False
    # 新旧迁改关系出现时不能反转/消失到只剩一个状态。
    if re.search(r"(?:原|旧).{0,30}(?:新建|新)|(?:新建|新).{0,30}(?:原|旧)", s):
        if any(w in c for w in ("改接","迁改","迁移","更换","拆除")) and not (re.search(r"原|旧", c) and re.search(r"新", c)):
            # 箱变专用已由上面强关系保护；普通领导摘要若无法表达新旧则回A。
            if not ("#1箱变" in s and "#2箱变" in s): return False
    if len(c)>len(s)+6: return False
    return True


def v30_leader_summary_from_a(mode_a_text, profession=""):
    source=remove_parentheses_notes(mode_a_text).rstrip("。")
    prof=str(profession or "").strip()
    if not source: return ""
    if prof=="营销":
        marketing=summarize_marketing_work_for_mode_a(source, compact=True).rstrip("。")
        return v31_finalize_leader_text(marketing, original=source).rstrip("。")+"。"
    cand = ""
    if "0.4kV多条支线" in source:
        cand = source
    if not cand:
        cand = v31_try_merge_low_voltage_same_action(source, prof)
    if not cand:
        cand = v31_try_merge_simple_same_action(source, prof)
    if not cand:
        if prof=="带电作业": cand=v30_leader_live(source)
        elif prof=="配电": cand=v30_leader_peidian(source)
        elif prof=="配网工程": cand=v30_leader_peiwang(source)
        elif prof=="省管产业": cand=v30_leader_industry(source)
        elif prof=="输电": cand=v30_leader_transmission(source)
        elif prof=="变电": cand=v30_leader_substation(source)
        else:
            cand=v30_leader_cable_rmu(source) or v30_leader_conductor(source) or v26_leader_generic(source)
    cand=str(cand or "").strip("，、；;。. ")
    cand=v31_finalize_leader_text(cand, original=source)
    if not v31_semantically_complete(source, cand):
        cand=v31_finalize_leader_text(source, original=source)
    if not v30_relation_guard(source,cand):
        cand=v31_finalize_leader_text(source, original=source)
    return cand.rstrip("。")+"。"


def v30_target_project_label(mode_a_text, profession=""):
    s=v23_norm(mode_a_text).rstrip("。")
    lines=v26_line_names_ordered(s); line=lines[0] if lines else ""
    taiqus=v26_taiqu_names(s); obj=taiqus[0] if taiqus else ""
    if not obj:
        m=re.search(r"([^，、；;。:：]{2,22}?(?:#?\d+号)?(?:台架变|台变|公变|配变|变))",s)
        if m:
            obj=m.group(1)
            if "线" in obj: obj=obj.split("线")[-1]
            obj=re.sub(r"^#?\d+(?:[-+]\d+)*(?:号)?杆","",obj)
    prefix=line+(obj if obj and obj not in line else "")
    if ("基础开挖" in s or "挖坑" in s) and ("组立" in s or "立杆" in s): purpose="新建"
    elif "变压器" in s and "更换" in s: purpose="变压器更换"
    elif "更换" in s and any(x in s for x in ("JP柜","母线","配电盘","低压柜")): purpose="配变改造"
    elif "更换" in s and any(x in s for x in ("开关","断路器")): purpose="开关更换"
    elif any(x in s for x in ("迁改","迁移","改移")): purpose="迁改"
    elif "联络" in s: purpose="联络工程"
    elif any(x in s for x in ("导线更换","旧导线拆除","耐张金具","电缆终端拆除")): purpose="线路检修"
    elif "检修" in s: purpose="检修"
    else: purpose=""
    if prefix and purpose: return prefix+purpose
    if prefix: return prefix+"工作"
    cand=v30_leader_summary_from_a(s,profession).rstrip("。") if profession!="带电作业" else s
    return cand.rstrip("，、；;：:")


def summarize_work_clear(raw_text, profession=""):
    """模式A【缩写版】：先排除所有圆括号说明，再按专业缩写。"""
    prof = str(profession or "").strip()
    source = remove_parentheses_notes(audit_source_entry_text(raw_text))
    if prof == "营销":
        return summarize_marketing_work_for_mode_a(source, compact=False)
    return v30_mode_a_abbrev(source, profession=prof)


def summarize_work_compact(raw_text, profession=""):
    """兼容旧调用名；先排除所有圆括号说明。"""
    prof = str(profession or "").strip()
    source = remove_parentheses_notes(audit_source_entry_text(raw_text))
    if prof == "营销":
        return summarize_marketing_work_for_mode_a(source, compact=True)
    return v30_mode_a_abbrev(source, profession=prof)


MODEB_KEY_ACTION_WORDS = (
    "变压器", "JP柜", "环网柜", "开关", "刀闸", "电缆", "导线", "下户线",
    "电能表", "互感器", "计量箱", "采集终端", "附件", "更换", "安装",
    "改接", "敷设", "压接", "拆除", "搭接", "展放", "试验", "制作",
    "检查", "验收", "组立", "断引流线", "接引流线", "绝缘", "防护", "包裹",
)

MODEB_COMPACT_ACTION_ORDER = [
    "线路拆线防护",
    "绝缘防护",
    "绝缘包裹",
    "断、接引流线",
    "断引流线",
    "接引流线",
    "拆除隔离开关",
    "开关更换",
    "安装跌落开关",
    "更换避雷器",
    "联络开关、PT、FTU及附件更换",
    "电缆通道清理",
    "电缆敷设",
    "低压出线电缆敷设",
    "低压出线电缆更换",
    "低压导线更换",
    "下户线改接",
    "终端制作",
    "试验",
    "塔组立",
    "导地线展放",
    "压接",
]


def strip_mode_b_risk_notes(text):
    """V3.2口径：模式B不保留任何圆括号说明。"""
    return remove_parentheses_notes(text)


def cleanup_mode_b_work_text(raw_text):
    """模式B小结专用清洗：删除全部圆括号说明并统一标点。"""
    s = normalize_work_text(raw_text, keep_newlines=True)
    s = re.sub(r"(?<=\d)[kK][vV]?", "kV", s)
    s = strip_mode_b_risk_notes(s)
    s = s.replace("（", "").replace("）", "").replace("(", "").replace(")", "")
    s = re.sub(r"(?m)(^|\n)\s*\d+\s*[.．、]\s*", r"\1", s)
    s = s.replace("变电站", "站")
    s = s.replace("济宁泗水", "")
    s = s.replace("及相关附件", "及附件").replace("相关附件", "附件")
    s = s.replace("相关配备设备", "配套设备").replace("柜内附件", "柜附件")
    s = s.replace("\n", "，")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"，{2,}", "，", s)
    return s.strip("，、；;。. ")


def cleanup_mode_b_item_text(text):
    s = remove_parentheses_notes(text).strip()
    s = re.sub(r"(?<=\d)[kK][vV]?", "kV", s)
    s = re.sub(r"^\s*\d+\s*[.．、]\s*", "", s)
    s = s.replace("变电站", "站")
    s = s.replace("（", "").replace("）", "").replace("(", "").replace(")", "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("；", "，").replace(";", "，")
    s = re.sub(r"[:：][，,]+", "，", s)
    s = re.sub(r"[，,]{2,}", "，", s)
    s = s.replace("配合断引流线、配合接引流线", "配合断、接引流线")
    s = s.replace("配合断引流线，配合接引流线", "配合断、接引流线")
    s = s.replace("断引流线、接引流线", "断、接引流线")
    s = s.replace("断引流线，接引流线", "断、接引流线")
    inline_pairs = (
        ("装设绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
        ("安装绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
        ("设置绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
        ("断开引流线", "接通引流线", "断开、接通引流线"),
        ("解开引流线", "恢复引流线", "解开、恢复引流线"),
        ("拆除引流线", "安装引流线", "拆、装引流线"),
        ("拆除引流线", "搭接引流线", "拆除、搭接引流线"),
        ("断开导线", "接通导线", "断开、接通导线"),
        ("解开导线", "恢复导线", "解开、恢复导线"),
    )
    for left, right, combined in inline_pairs:
        for separator in ("、", "，"):
            s = s.replace(f"{left}{separator}{right}", combined)
            s = s.replace(f"{right}{separator}{left}", combined)
    return s.strip("，、；;。. ")


def strip_station_context_prefix(text):
    return re.sub(r"^(?:(?:35|110|220)kV)?[^，,；;。:：\n]{1,18}(?:站)", "", str(text or ""))


def same_mode_b_object_prefix(left, right):
    if left == right:
        return True
    left_core = strip_station_context_prefix(left)
    right_core = strip_station_context_prefix(right)
    return bool(left_core) and left_core == right_core


LIVE_WORK_ACTION_PAIRS = (
    ("断引流线", "接引流线", "断、接引流线"),
    ("断开引流线", "接通引流线", "断开、接通引流线"),
    ("解开引流线", "恢复引流线", "解开、恢复引流线"),
    ("拆除引流线", "安装引流线", "拆、装引流线"),
    ("拆除引流线", "搭接引流线", "拆除、搭接引流线"),
    ("装设绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
    ("安装绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
    ("设置绝缘遮蔽", "拆除绝缘遮蔽", "装拆绝缘遮蔽"),
    ("断开导线", "接通导线", "断开、接通导线"),
    ("解开导线", "恢复导线", "解开、恢复导线"),
)


def split_live_work_action(text, action):
    idx = str(text or "").find(action)
    if idx <= 0:
        return None
    return text[:idx], cleanup_mode_b_item_text(text[idx + len(action):])


def merge_live_work_tails(left_tail, right_tail):
    left = cleanup_mode_b_item_text(left_tail)
    right = cleanup_mode_b_item_text(right_tail)
    if left == right:
        return left
    if not left:
        return right
    if not right:
        return left

    for marker in ("配合", "用于", "协助"):
        left_idx = left.rfind(marker)
        right_idx = right.rfind(marker)
        if left_idx < 0 or right_idx < 0:
            continue
        left_context = left[left_idx:]
        right_context = right[right_idx:]
        if left_context != right_context:
            continue
        extras = []
        for extra in (left[:left_idx], right[:right_idx]):
            extra = extra.strip("，、；;。. ")
            if extra and extra not in extras:
                extras.append(extra)
        return f"{'、'.join(extras)}，{left_context}" if extras else left_context
    if left.endswith(right):
        return left
    if right.endswith(left):
        return right
    return None


def try_merge_live_work_pair(current, nxt):
    for left_action, right_action, combined_action in LIVE_WORK_ACTION_PAIRS:
        for first_action, second_action in (
            (left_action, right_action),
            (right_action, left_action),
        ):
            first = split_live_work_action(current, first_action)
            second = split_live_work_action(nxt, second_action)
            if not first or not second:
                continue
            first_prefix, first_tail = first
            second_prefix, second_tail = second
            if not same_mode_b_object_prefix(first_prefix, second_prefix):
                continue
            merged_tail = merge_live_work_tails(first_tail, second_tail)
            if merged_tail is None:
                continue
            if merged_tail:
                separator = "" if merged_tail.startswith(("配合", "用于", "协助")) else "，"
                merged_tail = f"{separator}{merged_tail}"
            return f"{first_prefix}{combined_action}{merged_tail}"
    return ""


def merge_reverse_live_work_items(items):
    merged = []
    idx = 0
    while idx < len(items):
        current = cleanup_mode_b_item_text(items[idx])
        if idx + 1 < len(items):
            nxt = cleanup_mode_b_item_text(items[idx + 1])
            combined = try_merge_live_work_pair(current, nxt)
            if combined:
                merged.append(combined)
                idx += 2
                continue
        merged.append(current)
        idx += 1
    return merged


def split_mode_b_work_items_with_context(raw_text):
    s = normalize_work_text(raw_text, keep_newlines=True)
    s = strip_mode_b_risk_notes(s)
    s = s.replace("变电站", "站")
    s = re.sub(
        r"((?:35|110|220)kV[^，,；;。:：\n]{1,18})变(?=[:：])",
        r"\1站",
        s,
    )
    raw_parts = []
    for line in s.split("\n"):
        raw_parts.extend(part for part in re.split(r"[；;。]+", line) if part.strip())

    items = []
    station_context = ""
    first_item_in_station = True
    station_pattern = re.compile(r"^((?:35|110|220)kV[^，,；;。:：\n]{1,18}站)[:：]?(.*)$")
    for raw_part in raw_parts:
        part = cleanup_mode_b_item_text(raw_part)
        if not part:
            continue
        m = station_pattern.match(part)
        if m and not m.group(2):
            station_context = m.group(1)
            first_item_in_station = True
            continue
        if m and m.group(2):
            station_context = m.group(1)
            part = cleanup_mode_b_item_text(m.group(2))
            first_item_in_station = True
        has_own_station = bool(
            re.match(r"^(?:35|110|220)kV[^，,；;。:：\n]{1,18}站", part)
        )
        if station_context and not has_own_station:
            part = f"{station_context}{part}"
        if part:
            items.append(cleanup_mode_b_item_text(part))
            first_item_in_station = False
    return merge_reverse_live_work_items(items)


def summarize_mode_b_items(raw_text, max_items=3, clause_len=130, require_action=False):
    items = split_mode_b_work_items_with_context(raw_text)
    if require_action:
        action_items = [item for item in items if any(word in item for word in MODEB_KEY_ACTION_WORDS)]
        if action_items:
            items = action_items
    if not items:
        text = cleanup_mode_b_work_text(raw_text)
        clauses = split_summary_clauses(text)
        items = [normalize_mode_b_clause_text(c) for c in clauses if c]
    if not items:
        return finish_mode_b_work_sentence(normalize_work_text(raw_text, keep_newlines=False))
    kept = [shorten_clause(cleanup_mode_b_item_text(item), clause_len) for item in items[:max_items]]
    body = "，".join(kept)
    if len(items) > max_items:
        body += "等作业内容"
    return finish_mode_b_work_sentence(body)


def mode_b_summary_plain_text(text):
    return str(text or "").replace("提接", "T接").strip("，、；;。. ")


def score_mode_b_summary_candidate(raw_text, candidate, prefer_compact=False):
    cand = mode_b_summary_plain_text(candidate)
    if not cand:
        return -999

    raw = cleanup_mode_b_work_text(raw_text)
    raw_items = split_mode_b_work_items_with_context(raw_text)
    raw_item_count = max(len(raw_items), count_work_items(raw_text))
    raw_len = len(raw)
    cand_len = len(cand)

    score = 0
    if re.search(r"(?:10|35|110|220)kV[^，,；;。]{1,30}?线", cand):
        score += 30
    if any(word in cand for word in MODEB_KEY_ACTION_WORDS):
        score += 35
    if any(word in cand for word in ("台区", "站", "杆", "支线", "环网柜", "开关柜", "塔")):
        score += 12
    if "低压线路维修" in raw and "低压线路维修" in cand:
        score += 25
    if "断、接引流线" in cand:
        score += 10

    if "；" in cand or ";" in cand or ":，" in cand or "：，" in cand:
        score -= 80
    if cand.endswith("线") or cand.endswith("10kV") or cand.endswith("35kV"):
        score -= 80

    if raw_item_count >= 3:
        if cand_len <= 95:
            score += 18
        elif cand_len <= 130:
            score += 10
        elif cand_len > 180:
            score -= 25
    else:
        if cand_len > 160:
            score -= 15

    # 如果原文很长，候选却几乎没有压缩，降低优先级。
    if raw_len > 180 and cand_len > raw_len * 0.75:
        score -= 20
    if prefer_compact and cand_len <= 120:
        score += 10

    # 过度压缩成只有动作、没有对象，也不合格。
    if not re.search(r"(?:10|35|110|220)kV|台区|站|杆|塔|柜", cand):
        score -= 30
    return score


def choose_mode_b_summary_candidate(raw_text, candidates, prefer_compact=False):
    valid = [c for c in candidates if mode_b_summary_plain_text(c)]
    if not valid:
        return ""
    scored = [
        (score_mode_b_summary_candidate(raw_text, c, prefer_compact=prefer_compact), len(mode_b_summary_plain_text(c)), c)
        for c in valid
    ]
    scored.sort(key=lambda item: (item[0], -item[1]), reverse=True)
    return scored[0][2]


def remove_pole_range_noise(text):
    # 模式B小结里杆号、区间往往是有效作业位置，不能机械删除。
    return re.sub(r"[、，]{2,}", "，", str(text or "")).strip("，、；;。. ")


def finish_mode_b_work_sentence(text):
    s = str(text or "").strip("，、；;。. ")
    return f"{s}。" if s else ""


def normalize_mode_b_clause_text(text):
    c = str(text or "").strip("，、；;。. ")
    c = c.replace("新建", "")
    c = c.replace("塔吊装组立", "塔组立").replace("吊装组立", "组立")
    c = c.replace("导、地线", "导地线")
    c = c.replace("电缆通道清理，电缆敷设", "电缆通道清理、电缆敷设")
    c = c.replace("断引流线，接引流线", "断、接引流线")
    c = c.replace("断引流线、接引流线", "断、接引流线")
    c = re.sub(r"等\d+项工作$", "", c)
    return c.strip("，、；;。. ")


def extract_voltage_line_name(text):
    m = re.search(r"((?:10|35|110|220)kV[^，,；;。:：#号杆]*?线(?:[^，,；;。:：#号杆]*?支线)?)", str(text or ""))
    return m.group(1).strip("，、；;。. ") if m else ""


def extract_voltage_line_names(text):
    names = []
    for m in re.finditer(r"((?:10|35|110|220)kV[^，,；;。:：#号杆]*?线(?:[^，,；;。:：#号杆]*?支线)?)", str(text or "")):
        add_unique(names, m.group(1).strip("，、；;。. "))
    return names


def v31_target_count_unit(targets):
    """只对同一身份类别计数，主线路、支线、台区和变电站绝不混算。"""
    kinds = []
    for target in targets:
        value = cleanup_mode_b_target_text(target)
        if re.search(r"(?:分支线|支线)$", value):
            kind = "branch"
        elif re.search(r"(?:10|35|110|220)kV[^，、；;。:：]+线$", value):
            kind = "main_line"
        elif value.endswith("站"):
            kind = "station"
        elif value.endswith("台区"):
            kind = "taiqu"
        else:
            kind = "other"
        kinds.append(kind)
    if not kinds or len(set(kinds)) != 1:
        return ""
    return {
        "branch": "条支线",
        "main_line": "条线路",
        "station": "座变电站",
        "taiqu": "个台区",
    }.get(kinds[0], "")


def extract_mode_b_compact_targets(raw_text, max_targets=4):
    targets = []
    items = split_mode_b_work_items_with_context(raw_text)
    if not items:
        items = split_summary_clauses(cleanup_mode_b_work_text(raw_text))

    for item in items:
        c = cleanup_mode_b_item_text(item)
        if not c:
            continue
        m = re.search(
            r"((?:10|35|110|220)kV[^，,；;。:：#号杆]{1,28}?线)"
            r"([^，,；;。:：]{0,24}?台区)",
            c,
        )
        if m:
            add_unique(targets, f"{m.group(1)}{normalize_lv_taiqu_name(m.group(2))}")
            continue

        line_names = extract_voltage_line_names(c)
        if line_names:
            for line_name in line_names:
                add_mode_b_target_unique(targets, cleanup_mode_b_target_text(line_name))
            continue

        station = ""
        m_station = re.match(r"((?:(?:35|110|220)kV)?[^，,；;。:：\n]{1,18}站)", c)
        if m_station:
            station = m_station.group(1)
        if station:
            add_mode_b_target_unique(targets, station)

    if len(targets) > max_targets:
        count_unit = v31_target_count_unit(targets)
        if not count_unit:
            return targets
        shown = targets[:max_targets]
        shown[-1] = f"{shown[-1]}等{len(targets)}{count_unit}"
        return shown
    return targets


def add_mode_b_target_unique(targets, target):
    target = cleanup_mode_b_target_text(target)
    if not target:
        return
    for old in list(targets):
        if old == target:
            return
        if old.endswith(target):
            return
        if target.endswith(old):
            targets[targets.index(old)] = target
            return
    targets.append(target)


def cleanup_mode_b_target_text(text):
    target = cleanup_mode_b_item_text(text)
    target = re.sub(r"(?:新建|更换|安装|拆除|敷设|展放|压接|试验|制作|改接|清理).*$", "", target)
    target = re.sub(r"(?:联络开关|开关|PT|FTU|附件|导地线|导线|电缆|低压出线电缆).*$", "", target)
    return target.strip("，、；;。. ")


def extract_mode_b_compact_actions(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    actions = []

    if "断引流线" in text and "接引流线" in text:
        add_unique(actions, "断、接引流线")
    elif "断引流线" in text:
        add_unique(actions, "断引流线")
    elif "接引流线" in text:
        add_unique(actions, "接引流线")

    if "线路拆线防护" in text:
        add_unique(actions, "线路拆线防护")
    if "绝缘防护" in text:
        add_unique(actions, "绝缘防护")
    if "绝缘包裹" in text or "改造包裹" in text:
        add_unique(actions, "绝缘包裹")

    if "联络开关" in text and "PT" in text and "FTU" in text and "更换" in text:
        add_unique(actions, "联络开关、PT、FTU及附件更换")
    elif "开关" in text and "更换" in text:
        add_unique(actions, "开关更换")
    if "拆除隔离开关" in text:
        add_unique(actions, "拆除隔离开关")
    if "安装跌落开关" in text:
        add_unique(actions, "安装跌落开关")
    if "更换避雷器" in text:
        add_unique(actions, "更换避雷器")

    if "电缆通道清理" in text:
        add_unique(actions, "电缆通道清理")
    if "电缆敷设" in text:
        add_unique(actions, "电缆敷设")
    if "电缆终端制作" in text or "终端制作" in text:
        add_unique(actions, "终端制作")
    if "试验" in text:
        add_unique(actions, "试验")
    if "压接" in text:
        add_unique(actions, "压接")

    if "塔组立" in text or "塔吊装组立" in text or "吊装组立" in text:
        add_unique(actions, "塔组立")
    if "导地线展放" in text or "导、地线展放" in text or "导线展放" in text:
        add_unique(actions, "导地线展放")

    if "低压出线电缆" in text and "敷设" in text:
        add_unique(actions, "低压出线电缆敷设")
    if "低压出线电缆" in text and "更换" in text:
        add_unique(actions, "低压出线电缆更换")
    if "低压导线" in text and "更换" in text:
        add_unique(actions, "低压导线更换")
    if "下户线" in text and "改接" in text:
        add_unique(actions, "下户线改接")

    ordered = [action for action in MODEB_COMPACT_ACTION_ORDER if action in actions]
    ordered.extend(action for action in actions if action not in ordered)
    return ordered


def summarize_compact_targets_actions_mode_b(raw_text, max_targets=4):
    targets = extract_mode_b_compact_targets(raw_text, max_targets=max_targets)
    actions = extract_mode_b_compact_actions(raw_text)
    if not targets or not actions:
        return ""
    target_text = "、".join(targets)
    action_text = "、".join(actions)
    return finish_mode_b_work_sentence(f"{target_text}{action_text}")


def remove_leading_station_prefix(text):
    return re.sub(r"^(?:\d+kV[^，,；;。:：]{1,30}站[:：])", "", str(text or "")).strip("，、；;。. ")


def summarize_general_mode_b_candidate_a(raw_text, max_clauses=5, clause_len=120):
    candidates = [
        summarize_compact_targets_actions_mode_b(raw_text, max_targets=4),
        summarize_mode_b_items(raw_text, max_items=max_clauses, clause_len=clause_len, require_action=True),
    ]
    return choose_mode_b_summary_candidate(raw_text, candidates, prefer_compact=True)


def extract_production_context_prefix(text):
    """提取行首站、线路上下文；地点和设备留在正文中，不能随前缀一起省略。"""
    clean = cleanup_mode_b_item_text(text)
    station_match = re.match(
        r"^((?:35|110|220)kV[^，,；;。:：\n]{1,18}站)",
        clean,
    )
    station = station_match.group(1) if station_match else ""
    search_start = station_match.end() if station_match else 0
    line_match = re.search(
        r"((?<!\d)(?:10|35|110|220)kV[^，,；;。:：#号杆]{1,40}?(?<!母)线)",
        clean[search_start:],
    )
    line = line_match.group(1) if line_match else ""
    end = search_start
    if line_match:
        end = search_start + line_match.end()
    return station, line, clean[:end]


PRODUCTION_IDENTITY_ACTIONS = (
    "断、接引流线", "断引流线", "接引流线", "断开", "接通",
    "更换", "新装", "安装", "拆除", "改接", "敷设", "制作",
    "展放", "维修", "治理", "试验", "压接", "检查", "验收",
    "组立", "搭接", "绑扎", "固定", "复紧", "移至", "清理",
    "校验", "检测", "改造", "开挖", "防护", "包裹", "紧固",
)

PRODUCTION_ACTION_OBJECTS = (
    "低压出线电缆", "电缆终端", "基础保护帽", "冷却风机",
    "室内外消防设施", "消防设施", "辅控设施", "联络开关",
    "低压导线", "导、地线", "导地线", "出线电缆", "下户线",
    "绝缘子", "引流线", "变压器", "配电设备", "保护装置",
    "光缆附件", "导线附件", "附件", "光缆", "导线", "电缆",
    "线路拆线", "JP柜", "母线", "开关", "螺栓", "绝缘", "大门",
)


def normalize_production_action_object(action_object):
    value = str(action_object or "")
    replacements = {
        "低压导线": "导线",
        "导、地线": "导地线",
    }
    return replacements.get(value, value)


def split_production_scope_action(text):
    """拆成站线、设备身份、动作对象和动作，避免只按“更换”等动词归组。"""
    clean = cleanup_mode_b_item_text(text)
    station, line, prefix = extract_production_context_prefix(clean)
    body = clean[len(prefix):].lstrip("，、:：")
    matches = []
    for action in PRODUCTION_IDENTITY_ACTIONS:
        idx = body.find(action)
        if idx >= 0:
            matches.append((idx, action))
    action_idx, action_word = min(matches, default=(-1, ""), key=lambda pair: pair[0])
    if action_idx < 0:
        return station, line, body, "", "", clean, prefix
    scope = body[:action_idx].strip("，、:：")
    action_text = body[action_idx:].strip("，、:：")
    action_object = ""

    for candidate in PRODUCTION_ACTION_OBJECTS:
        if scope.endswith(candidate):
            action_object = candidate
            scope = scope[:-len(candidate)].strip("，、:：")
            break

    if not action_object and action_text.startswith(action_word):
        after_verb = action_text[len(action_word):]
        for candidate in PRODUCTION_ACTION_OBJECTS:
            if after_verb == candidate:
                action_object = candidate
                action_text = action_word
                break

    action_object = normalize_production_action_object(action_object)
    action_text = normalize_production_action_key(action_text)
    return station, line, scope, action_object, action_text, clean, prefix


def normalize_production_action_key(action_text):
    action = cleanup_mode_b_item_text(action_text).replace("检查工作", "检查")
    replacements = (
        (r"^更换低压导线$", "导线更换"),
        (r"^更换导线$", "导线更换"),
        (r"^更换低压出线电缆$", "出线电缆更换"),
        (r"^敷设低压出线电缆$", "出线电缆敷设"),
    )
    for pattern, replacement in replacements:
        action = re.sub(pattern, replacement, action)
    return action


def compact_contextual_identity_references(
    references,
    station_colon=False,
    initial_context=None,
    return_context=False,
):
    """按新身份逐层更新上下文，省略后续已确认未变化的站、线前缀。"""
    compacted = []
    if initial_context is None:
        previous_station = None
        previous_line = None
    else:
        previous_station, previous_line = initial_context
    for reference in references:
        ref = mode_b_summary_plain_text(reference)
        if not ref:
            continue
        station, line, prefix = extract_production_context_prefix(ref)
        if previous_station is None:
            display = ref
        elif station != previous_station:
            display = ref
        elif line != previous_line:
            station_prefix = station if station and ref.startswith(station) else ""
            display = ref[len(station_prefix):].lstrip("，、:：")
        elif prefix:
            display = ref[len(prefix):].lstrip("，、:：") or ref
        else:
            display = ref
        if station_colon and station and not line and display.startswith(station):
            rest = display[len(station):].lstrip("，、:：")
            display = f"{station}：{rest}"
        add_unique(compacted, display)
        previous_station = station
        previous_line = line
    if return_context:
        return compacted, (previous_station, previous_line)
    return compacted


def compact_parallel_identity_references(references):
    """合并同一上下文中的1号/2号主变、#5/#8塔、台架I/II等并列设备。"""
    refs = [mode_b_summary_plain_text(ref) for ref in references if ref]
    patterns = (
        r"^(.*?)(\d+号)(主变)(.*)$",
        r"^(.*?)(\d+(?:[-+]\d+)*)(号杆)(.*)$",
        r"^(.*?)(#\d+(?:[-+]\d+)*)(杆)(.*)$",
        r"^(.*?)(#?\d+)(塔)(.*)$",
        r"^(.*?10kV)([ⅠⅡⅢⅣIV]+)(段母线)(.*)$",
        r"^(.*台架(?:变)?)([ⅠⅡⅢⅣIV]+)(.*)$",
    )
    compacted = []
    group_indexes = {}
    group_labels = {}
    for ref in refs:
        matched = False
        for pattern_index, pattern in enumerate(patterns):
            match = re.match(pattern, ref)
            if not match:
                continue
            key = (pattern_index, match.group(1), match.group(3), match.group(4))
            if key not in group_indexes:
                group_indexes[key] = len(compacted)
                group_labels[key] = []
                compacted.append("")
            add_unique(group_labels[key], match.group(2))
            compacted[group_indexes[key]] = (
                f"{match.group(1)}{'、'.join(group_labels[key])}"
                f"{match.group(3)}{match.group(4)}"
            )
            matched = True
            break
        if not matched:
            add_unique(compacted, ref)
    return compacted


def summarize_identity_aware_production_mode_b(
    raw_text, station_colon=False, max_items=None
):
    """先逐行绑定身份与动作，再按相同动作归组并统一后置动作。"""
    items = split_mode_b_work_items_with_context(raw_text)
    if not items:
        text = cleanup_mode_b_work_text(raw_text)
        return finish_mode_b_work_sentence(text)

    entries = []
    selected_items = items if max_items is None else items[:max_items]
    for item_index, item in enumerate(selected_items):
        (
            station,
            line,
            scope,
            action_object,
            action_text,
            clean,
            prefix,
        ) = split_production_scope_action(item)
        clean = clean.replace("检查工作", "检查")
        if not clean:
            continue
        if action_text:
            entries.append(
                {
                    "index": item_index,
                    "identity": f"{prefix}{scope}".strip("，、:："),
                    "object": action_object,
                    "action": action_text,
                }
            )
        else:
            entries.append({"index": item_index, "standalone": clean})

    phrase_counts = Counter(
        (entry.get("object", ""), entry.get("action", ""))
        for entry in entries
        if "standalone" not in entry
    )
    groups = {}
    output_order = []
    for entry in entries:
        if "standalone" in entry:
            key = ("standalone", entry["index"])
            groups[key] = {"text": entry["standalone"]}
            output_order.append(key)
            continue
        phrase_key = (entry["object"], entry["action"])
        if phrase_counts[phrase_key] > 1:
            key = ("phrase",) + phrase_key
        else:
            key = ("identity", entry["identity"], entry["object"])
        if key not in groups:
            groups[key] = {
                "object": entry["object"],
                "actions": [],
                "references": [],
            }
            output_order.append(key)
        add_unique(groups[key]["actions"], entry["action"])
        add_unique(groups[key]["references"], entry["identity"])

    kept = []
    context = None
    for group_key in output_order:
        group = groups[group_key]
        if group_key[0] == "standalone":
            add_unique(kept, shorten_clause(group["text"], 160))
            station, line, _ = extract_production_context_prefix(group["text"])
            context = (station, line)
            continue
        references = compact_parallel_identity_references(group["references"])
        references, context = compact_contextual_identity_references(
            references,
            station_colon=station_colon,
            initial_context=context,
            return_context=True,
        )
        identity_text = "、".join(references)
        action_text = "、".join(group["actions"])
        add_unique(
            kept,
            shorten_clause(
                f"{identity_text}{group['object']}{action_text}",
                220,
            ),
        )

    body = "，".join(kept)
    if max_items is not None and len(items) > max_items:
        body += "等作业内容"
    return finish_mode_b_work_sentence(body)


def summarize_transmission_mode_b(raw_text):
    return summarize_identity_aware_production_mode_b(raw_text)


def summarize_live_work_mode_b(raw_text):
    return summarize_identity_aware_production_mode_b(raw_text, max_items=3)


def summarize_substation_mode_b(raw_text):
    return summarize_identity_aware_production_mode_b(raw_text)


def strip_peiwang_project_title(text):
    """配网工程模式B只保留实际线路、设备、动作，不输出工程包装名。"""
    s = str(text or "").strip("，、；;。. ")
    s = re.sub(r"^配网工程\d*项[，,、:：]*", "", s)
    s = re.sub(r"^[^，,；;。]{0,140}(?:工程|治理|改造)[^，,；;。:：]{0,20}[:：]", "", s)
    s = re.sub(r"^.*?(?:工程|治理|改造)(?=(?:10|35|110|220)kV)", "", s)
    return s.strip("，、；;。. :：")


def summarize_peiwang_project_mode_b(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    text = strip_peiwang_project_title(text)
    candidates = [summarize_compact_targets_actions_mode_b(text, max_targets=4)]
    # 去掉“某某工程/治理”这类包装话，保留真正线路、设备、动作。
    m = re.search(r"(?:治理|改造|工程)(?=((?:10|35|110)kV[^，。；;]{1,30}?线))", text)
    if m:
        text = text[m.end():]
    else:
        m = re.search(r"((?:10|35|110)kV[^，。；;]{1,35}?线.*)", text)
        if m:
            text = m.group(1)

    clauses = split_summary_clauses(text)
    kept = []
    for clause in clauses:
        c = normalize_mode_b_clause_text(remove_pole_range_noise(clause))
        if any(word in c for word in MODEB_KEY_ACTION_WORDS):
            add_unique(kept, shorten_clause(c, 120))
    if not kept and text:
        kept = [shorten_clause(text, 140)]
    body = "，".join(kept[:6])
    candidates.append(finish_mode_b_work_sentence(body))
    return choose_mode_b_summary_candidate(raw_text, candidates, prefer_compact=True)


def normalize_lv_taiqu_name(text):
    name = str(text or "").strip("，、；;。. :：")
    name = name.replace("#", "")
    name = re.sub(r"(?<=\d)号台区", "台区", name)
    name = re.sub(r"\s+", "", name)
    return name


def extract_lv_repair_targets(text):
    targets = []
    pattern = re.compile(
        r"((?:10|35|110)kV[^，,；;。:：]{1,30}?线)"
        r"([^，,；;。:：]{1,30}?台区)"
        r"(?:低压线路维修工程|低压线路维修|线路维修工程|线路维修)"
    )
    for match in pattern.finditer(text):
        line_name = match.group(1).strip("，、；;。. :：")
        taiqu = normalize_lv_taiqu_name(match.group(2))
        if line_name and taiqu:
            add_unique(targets, f"{line_name}{taiqu}")
    return targets


def summarize_low_voltage_repair_mode_b(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    targets = extract_lv_repair_targets(text)
    if not targets:
        return ""

    actions = []
    if "低压出线电缆" in text:
        if "低压出线电缆敷设" in text and "低压出线电缆更换" not in text:
            add_unique(actions, "敷设低压出线电缆")
        else:
            add_unique(actions, "更换低压出线电缆")
    if "低压导线" in text:
        add_unique(actions, "低压导线")
    if "下户线" in text and "改接" in text:
        add_unique(actions, "下户线改接")

    if not actions:
        return ""
    if actions and actions[0].startswith(("更换", "敷设")) and len(actions) >= 2:
        action_text = f"{actions[0]}、{'及'.join(actions[1:])}"
    else:
        action_text = "、".join(actions)
    return finish_mode_b_work_sentence(f"{'、'.join(targets)}低压线路维修，{action_text}")


def split_peidian_segments(text):
    parts = [p.strip("，、；;。. ") for p in re.split(r"[。；;]+", text) if p.strip("，、；;。. ")]
    if len(parts) >= 2:
        return parts
    # 兜底：没有句号时，按第二个及以后的 10kV 开头拆成多项。
    starts = [m.start() for m in re.finditer(r"(?<!^)(?=(?:10|35|110)kV)", text)]
    if not starts:
        return [text] if text else []
    segments = []
    last = 0
    for start in starts:
        if start > last:
            segments.append(text[last:start].strip("，、；;。. "))
        last = start
    segments.append(text[last:].strip("，、；;。. "))
    return [p for p in segments if p]


def simplify_peidian_action_clause(clause):
    c = remove_pole_range_noise(clause)
    c = re.sub(r"^\d+\s*[.．、]\s*", "", c).strip("，、；;。. ")
    c = re.sub(r"等\d+项工作$", "", c)
    if not c:
        return ""
    if "下户线" in c and "改接" in c:
        return "下户线改接"
    if "低压出线电缆" in c and "低压导线" in c and "更换" in c:
        m = re.search(r"([^，,；;。]{0,45}低压出线电缆及低压导线更换)", c)
        return m.group(1) if m else "低压出线电缆及低压导线更换"
    if "低压出线电缆" in c and ("更换" in c or "敷设" in c):
        action_word = "更换" if "更换" in c else "敷设"
        m = re.search(rf"([^，,；;。]{{0,45}}低压出线电缆{action_word})", c)
        return m.group(1) if m else f"低压出线电缆{action_word}"
    if "低压导线" in c and "更换" in c:
        m = re.search(r"([^，,；;。]{0,35}0\.4kV[ⅠIⅡⅢA-Za-z0-9#号]*支线)?[^，,；;。]{0,25}低压导线更换", c)
        if m:
            phrase = m.group(0).strip("，、；;。. ")
            if "0.4kV" in phrase and ("台区" in phrase or "支线" in phrase):
                return phrase
        return "低压导线更换"
    if any(word in c for word in MODEB_KEY_ACTION_WORDS):
        return shorten_clause(c, 90)
    return ""


def summarize_peidian_segment_mode_b(segment):
    seg = cleanup_mode_b_work_text(segment)
    title = ""
    m = re.match(r"((?:10|35|110)kV[^，,；;。]{0,90}?(?:低压线路维修工程|低压线路维修|线路维修工程|线路维修))", seg)
    if m:
        title = m.group(1).strip("，、；;。. ")
        rest = seg[m.end():]
    else:
        m = re.match(r"((?:10|35|110)kV[^，,；;。]{0,45})", seg)
        if m:
            title = m.group(1).strip("，、；;。. ")
            rest = seg[m.end():]
        else:
            rest = seg

    actions = []
    for clause in split_summary_clauses(rest):
        action = simplify_peidian_action_clause(clause)
        add_unique(actions, action)

    if not actions:
        return shorten_clause(seg, 140)
    joiner = "" if title.endswith("工程") else "："
    return f"{title}{joiner}{'、'.join(actions)}" if title else "、".join(actions)


def summarize_peidian_cable_mode_b(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    lines = []
    for match in re.finditer(r"((?<!\d)10kV[^，,；;。:：#号杆]{1,40}?线)", text):
        add_unique(lines, match.group(1).strip("，、；;。. "))
    if not lines:
        lines = extract_voltage_line_names(text)
    actions = []
    if "电缆通道清理" in text:
        add_unique(actions, "电缆通道清理")
    if "电缆敷设" in text:
        add_unique(actions, "电缆敷设")
    if "电缆终端制作" in text or "终端制作" in text:
        add_unique(actions, "终端制作")
    if "试验" in text:
        add_unique(actions, "试验")
    if "压接" in text:
        add_unique(actions, "压接")
    if lines and actions:
        return finish_mode_b_work_sentence(f"{'、'.join(lines)}{'、'.join(actions)}")
    return summarize_general_mode_b_candidate_a(raw_text, max_clauses=5, clause_len=130)


def summarize_peidian_mode_b(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    if "电缆通道清理" in text and "电缆敷设" in text:
        return summarize_peidian_cable_mode_b(raw_text)
    candidates = [
        summarize_low_voltage_repair_mode_b(raw_text),
        summarize_compact_targets_actions_mode_b(raw_text, max_targets=4),
    ]
    if "低压线路维修" not in text:
        candidates.append(summarize_general_mode_b_candidate_a(raw_text, max_clauses=5, clause_len=130))
        return choose_mode_b_summary_candidate(raw_text, candidates, prefer_compact=True)
    segments = split_peidian_segments(text)
    summarized = []
    for segment in segments:
        add_unique(summarized, summarize_peidian_segment_mode_b(segment))
    body = "。".join(s for s in summarized if s)
    candidates.append(finish_mode_b_work_sentence(body) if body else "")
    candidates.append(summarize_general_mode_b_candidate_a(raw_text, max_clauses=5, clause_len=130))
    return choose_mode_b_summary_candidate(raw_text, candidates, prefer_compact=True)


def summarize_work_mode_b(raw_text, profession=""):
    """模式B领导总览：调用者必须传入模式A【缩写版】，不重新读取原文。"""
    return v26_leader_summary_from_a(raw_text, profession=profession)


def make_record_block(record, index, work_text):
    def compact_value(value):
        text = str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()
        return re.sub(r"\n{2,}", "\n", text)

    lines = [
        f"{index}.专业:{record.get('profession', '')}",
        f"工作内容:{compact_value(work_text)}",
    ]
    if record.get("leader"):
        lines.append(f"负责人:{compact_value(record['leader'])}")
    if record.get("time_part"):
        lines.append(f"作业时间:{compact_value(record['time_part'])}")
    if record.get("risk"):
        lines.append(f"作业风险等级:{compact_value(record['risk'])}")
    if record.get("grid_risk"):
        lines.append(f"电网风险等级:{compact_value(record['grid_risk'])}")
    if record.get("same_inout"):
        lines.append(f"同进同出:{compact_value(record['same_inout'])}")
    return "\n".join(lines)


def sort_mode_a_records(records):
    need = [r for r in records if r.get("need_leader_no")]
    other = [r for r in records if not r.get("need_leader_no")]
    need.sort(key=lambda r: r.get("need_leader_no") or 999999)
    other.sort(key=lambda r: r.get("order") or 999999)
    return need + other


def company_plan_title(title):
    text = str(title or "现场作业计划").strip()
    if "公司" in text:
        text = re.sub(r"^.*?公司", "公司", text, count=1)
    elif not text.startswith("公司"):
        text = f"公司{text}"
    return text


def build_mode_a_summary_lines(title, records):
    ordered = sort_mode_a_records(records)
    display_title = company_plan_title(title)
    lines = ["【缩写版】", display_title, ""]
    for idx, record in enumerate(ordered, 1):
        work = summarize_work_clear(summary_source_text(record), record.get("profession"))
        lines.append(make_record_block(record, idx, work))
        lines.append("")

    lines.extend(["------------------------------------", "【完整版】", display_title, ""])
    for idx, record in enumerate(ordered, 1):
        full_work = normalize_work_text(record.get("work"), keep_newlines=True)
        lines.append(make_record_block(record, idx, full_work))
        lines.append("")
    return lines[:-1]


def add_unique(items, value):
    value = str(value).strip()
    if value and value not in items:
        items.append(value)


def is_high_voltage_marketing_work(raw_text):
    raw = normalize_work_text(raw_text)
    if not raw:
        return False
    if "低压" in raw and not any(word in raw for word in ("高压", "业扩", "增容", "关口", "竣工验收")):
        return False
    high_words = (
        "高压计量", "高压业扩", "高压增容", "业扩新装", "业扩增容",
        "增容工程竣工验收", "业扩工程竣工验收", "竣工验收",
        "电流互感器", "二次回路", "关口电能表",
    )
    return any(word in raw for word in high_words)


def is_high_voltage_metering_marketing_work(raw_text):
    raw = normalize_work_text(raw_text)
    return bool(raw) and ("高压计量" in raw or "关口电能表" in raw)


def strip_marketing_supply_prefix(text):
    s = str(text or "").strip("，、；;。. ")
    s = re.sub(r"^(?:35|110|220)kV[^，,；;。:：]{1,25}站[:：]?", "", s)
    s = re.sub(r"^(?:10|35|110|220)kV[^，,；;。:：]{0,45}?线", "", s)
    return s.strip("，、；;。. ")


def remove_marketing_quantity_words(text):
    s = str(text or "")
    s = re.sub(r"\d+(?:\.\d+)?\s*(?:只|块|台|套|个|组|支|根|条|面|具|户|处)", "", s)
    s = re.sub(r"[，,、]{2,}", "，", s)
    return s.strip("，、；;。. ")


def extract_high_marketing_object(body):
    text = remove_marketing_quantity_words(body)
    m = re.search(r"^(?:新装|更换)(.+?关口)电能表", text)
    if m:
        return shorten_clause(m.group(1), 45)

    action_starts = []
    for word in ("更换", "新装", "业扩", "增容工程", "高压增容", "高压业扩", "竣工验收"):
        idx = text.find(word)
        if idx > 0:
            action_starts.append(idx)
    if not action_starts:
        return ""
    obj = text[:min(action_starts)].strip("，、；;。. ")
    return shorten_clause(obj, 45)


def summarize_high_marketing_action(raw_text):
    raw = normalize_work_text(raw_text)
    if "电能表" in raw and ("校验" in raw or "检验" in raw):
        return "现场校验电能表"
    if "竣工验收" in raw:
        if "增容" in raw:
            return "高压增容工程竣工验收"
        if "业扩" in raw:
            return "高压业扩工程竣工验收"
        return "工程竣工验收"
    if "关口" in raw and "电能表" in raw:
        if "更换" in raw and "新装" not in raw:
            return "更换电能表"
        return "新装电能表"
    if "电能表" in raw and "采集终端" in raw and "电流互感器" in raw and "二次回路" in raw:
        return "新装电能表、采集终端、电流互感器及二次回路接线检查"
    if "电流互感器" in raw and "二次回路" in raw:
        if "更换" in raw:
            return "更换高压电流互感器及二次回路接线检查"
        return "电流互感器及二次回路接线检查"
    if "高压计量" in raw and "电能表" in raw:
        if "更换" in raw and "新装" in raw:
            return "高压计量更换与新装电能表"
        if "更换" in raw:
            return "更换高压计量电能表"
        return "新装高压计量电能表"
    if "采集终端" in raw:
        return "新装采集终端" if "新装" in raw else "采集终端处理"
    return ""


def summarize_high_voltage_metering_list(raw_text, compact=False):
    text = cleanup_mode_b_work_text(raw_text)
    line_names = extract_voltage_line_names(text)
    if not line_names:
        return finish_work_sentence(summarize_high_marketing_action(raw_text), 1)

    max_lines = 4 if compact else 6
    kept = line_names[:max_lines]
    first_voltage = ""
    compact_names = []
    for idx, name in enumerate(kept):
        match = re.match(r"((?:10|35|110|220)kV)(.+)", name)
        if idx == 0 and match:
            first_voltage = match.group(1)
            compact_names.append(name)
        elif match and first_voltage and match.group(1) == first_voltage:
            compact_names.append(match.group(2))
        else:
            compact_names.append(name)

    target_text = "、".join(compact_names)
    if len(line_names) > len(kept):
        target_text += f"等{len(line_names)}条线路"
    action = summarize_high_marketing_action(raw_text) or "高压计量作业"
    return finish_work_sentence(f"{target_text}{action}", 1)


def summarize_high_voltage_marketing_mode_b(raw_text):
    text = cleanup_mode_b_work_text(raw_text)
    body = strip_marketing_supply_prefix(text)
    obj = extract_high_marketing_object(body)
    action = summarize_high_marketing_action(raw_text)
    if obj and action:
        return f"{obj}{action}"
    if action:
        return action
    return summarize_marketing_work_for_mode_a(raw_text, compact=False).rstrip("。")


def summarize_high_voltage_metering_mode_b(raw_text):
    return summarize_high_voltage_metering_list(raw_text, compact=True).rstrip("。")


def add_low_voltage_marketing_actions(raw, actions):
    has_batch_meter = "批量" in raw and "更换" in raw and "电能表" in raw
    has_box_new_meter = "计量箱" in raw and "新装" in raw and "电能表" in raw
    if "采集终端" in raw:
        add_unique(actions, "新装采集终端" if "新装" in raw else "采集终端处理")
    if has_box_new_meter:
        add_unique(actions, "新装计量箱与电能表")
    if ("新装" in raw or "更换" in raw) and "电能表" in raw and "关口" not in raw and not has_batch_meter:
        if not (has_box_new_meter and "更换" not in raw):
            add_unique(actions, "新装与更换电能表")
    if has_batch_meter:
        add_unique(actions, "低压计量批量更换电能表")
    if "计量装置" in raw and "隐患" in raw and ("治理" in raw or "排查" in raw):
        add_unique(actions, "低压计量装置隐患治理")


def merge_low_voltage_meter_actions(actions):
    low_meter_actions = {
        "新装计量箱与电能表",
        "新装与更换电能表",
        "低压计量批量更换电能表",
    }
    has_box_new = "新装计量箱与电能表" in actions
    has_general = "新装与更换电能表" in actions
    has_batch = "低压计量批量更换电能表" in actions
    if not any((has_box_new, has_general, has_batch)):
        return actions

    merged = ""
    if has_box_new and has_batch:
        merged = "低压计量新装计量箱及电能表、批量更换电能表"
    elif has_general and has_batch:
        merged = "低压计量新装及批量更换电能表"
    elif has_box_new:
        merged = "低压计量新装计量箱及电能表"
    elif has_batch:
        merged = "低压计量批量更换电能表"
    else:
        merged = "新装与更换电能表"

    result = []
    inserted = False
    for action in actions:
        if action in low_meter_actions:
            if not inserted:
                result.append(merged)
                inserted = True
            continue
        result.append(action)
    return result


def extract_marketing_action_types(records):
    actions = []
    for record in records:
        if record.get("profession") != "营销":
            continue
        raw = normalize_work_text(record.get("work"))
        if is_high_voltage_metering_marketing_work(raw):
            continue
        if is_high_voltage_marketing_work(raw):
            add_unique(actions, summarize_high_marketing_action(raw))
            continue
        add_low_voltage_marketing_actions(raw, actions)
    ordered = [a for a in MARKETING_ACTION_ORDER if a in actions]
    ordered.extend([a for a in actions if a not in ordered])
    return merge_low_voltage_meter_actions(ordered)


def extract_high_voltage_metering_marketing_details(records):
    details = []
    for record in records:
        if record.get("profession") != "营销":
            continue
        raw = normalize_work_text(record.get("work"))
        if is_high_voltage_metering_marketing_work(raw):
            add_unique(details, summarize_high_voltage_metering_mode_b(raw))
    return details


def extract_marketing_units(records):
    units = []
    unit_keywords = [
        "营销部营业班", "营销部计量班", "营销部营销班", "城区供电中心",
        "星村", "泗张", "中册", "华村", "泉林", "柘沟", "苗馆",
        "圣水峪", "金庄", "高峪", "杨柳", "济河", "泗河", "城区",
    ]
    for record in records:
        if record.get("profession") != "营销":
            continue
        text = normalize_work_text(record.get("unit"), keep_newlines=True)
        if not text:
            continue
        for part in re.split(r"[、,，/；;\n]+", text):
            p = part.strip()
            if not p:
                continue
            for key in unit_keywords:
                if key in p:
                    add_unique(units, "城区" if key == "城区供电中心" else key)
                    break
    return units


def title_to_work_plan_prefix(title):
    base = company_plan_title(title)
    base = base.replace("现场作业计划", "").replace("作业计划", "").strip()
    return base or "公司"


def format_numbered_mode_b_descriptions(descs):
    parts = []
    for idx, desc in enumerate(descs, 1):
        clean = mode_b_summary_plain_text(desc)
        if clean:
            parts.append(f"{idx}.{clean}。")
    return "".join(parts)


def summarize_marketing_record_mode_b(record):
    raw = record.get("work")
    if is_high_voltage_marketing_work(normalize_work_text(raw)):
        return summarize_high_voltage_marketing_mode_b(raw)
    return summarize_marketing_work_for_mode_a(raw, compact=True).rstrip("。")


PEIDIAN_SUMMARY_ACTIONS = (
    "更换", "新装", "安装", "拆除", "改接", "敷设", "制作",
    "展放", "维修", "治理", "试验", "压接", "检查", "验收",
)


def first_peidian_action(text, start=0):
    matches = []
    for action in PEIDIAN_SUMMARY_ACTIONS:
        idx = text.find(action, start)
        if idx >= 0:
            matches.append((idx, action))
    return min(matches, default=(-1, ""), key=lambda item: item[0])


def summarize_high_voltage_peidian_item(item):
    """10kV及以上保留站线和定级关键对象，设备明细可归并为配套设备。"""
    seg = cleanup_mode_b_item_text(item)
    if not seg:
        return ""
    if "电缆通道清理" in seg and "电缆敷设" in seg:
        return summarize_peidian_cable_mode_b(seg).rstrip("。")

    marker = ""
    marker_end = -1
    if "台架" in seg:
        marker = "台架"
        marker_end = seg.find("台架") + len("台架")
    elif "配电室" in seg and "变压器" in seg:
        marker = "变压器"
        marker_end = seg.find("变压器") + len("变压器")
    elif "配电室" in seg:
        marker = "配电室"
        marker_end = seg.find("配电室") + len("配电室")
    elif "变压器" in seg:
        marker = "变压器"
        marker_end = seg.find("变压器") + len("变压器")

    if marker_end < 0:
        return shorten_clause(seg, 150)

    action_idx, action = first_peidian_action(seg, marker_end)
    if action_idx < 0:
        return shorten_clause(seg, 150)

    target = seg[:marker_end].strip("，、；;。. ")
    between = seg[marker_end:action_idx].strip("，、；;。. ")
    if re.search(r"0\.4kV|支线|杆|至|之间|通道|线路", between):
        return shorten_clause(seg, 150)
    has_extra_equipment = bool(between)
    if marker == "台架" and between.startswith("变压器"):
        has_extra_equipment = True
    middle = "及配套设备" if has_extra_equipment else ""
    return f"{target}{middle}{action}"


def summarize_low_voltage_peidian_item(item):
    """0.4kV省略支线和杆号，只保留定级关键对象、设备及动作。"""
    seg = cleanup_mode_b_item_text(item)
    if not seg:
        return ""

    actions = []
    if "低压出线电缆" in seg or "0.4kV出线电缆" in seg:
        if "更换" in seg:
            add_unique(actions, "出线电缆更换")
        elif "敷设" in seg:
            add_unique(actions, "出线电缆敷设")
    elif "电缆" in seg:
        if "更换" in seg:
            add_unique(actions, "电缆更换")
        elif "敷设" in seg:
            add_unique(actions, "电缆敷设")
    if "导线" in seg:
        if "更换" in seg:
            add_unique(actions, "导线更换")
        elif "展放" in seg:
            add_unique(actions, "导线展放")
        elif "维修" in seg:
            add_unique(actions, "导线维修")
    if "下户线" in seg:
        if "改接" in seg:
            add_unique(actions, "下户线改接")
        elif "更换" in seg:
            add_unique(actions, "下户线更换")

    critical = ""
    critical_end = -1
    for marker in ("台架", "配电室", "变压器"):
        if marker in seg:
            critical = marker
            critical_end = seg.find(marker) + len(marker)
            break
    if critical:
        action_idx, action = first_peidian_action(seg, critical_end)
        between = seg[critical_end:action_idx] if action_idx >= 0 else ""
        is_actual_target = action_idx >= 0 and not re.search(
            r"0\.4kV|低压|出线|支线|杆|至|之间|通道|线路|#|[ⅠⅡⅢⅣIV]",
            between,
        )
        if is_actual_target:
            other_equipment = any(
                word in between for word in ("JP柜", "母线", "电缆", "导线", "开关", "配电箱")
            )
            critical_text = f"{critical}{'及配套设备' if other_equipment else ''}{action}"
            add_unique(actions, critical_text)

    if not actions:
        _, action = first_peidian_action(seg)
        if action:
            return f"0.4kV设备{action}"
        return "0.4kV低压作业"
    return f"0.4kV{'、'.join(actions)}"


def extract_exact_10kv_line_name(text):
    names = []
    for match in re.finditer(r"((?<!\d)10kV[^，,；;。:：#号杆]{1,40}?线)", str(text or "")):
        add_unique(names, match.group(1).strip("，、；;。. "))
    return names[-1] if names else ""


def normalize_parent_location(prefix, line_name=""):
    text = cleanup_mode_b_item_text(prefix)
    if line_name and line_name in text:
        text = text[text.rfind(line_name) + len(line_name):]
    text = re.sub(r"^(?:(?:35|110|220)kV)?[^，,；;。:：]{1,18}站", "", text)
    text = text.replace("0.4kV", "")
    text = re.sub(r"(?:低压线路维修工程|线路维修工程|低压线路维修)[:：]?", "", text)
    return text.strip("，、；;。. :：")


def extract_parent_device_identity(text, default_line=""):
    """把双重名称映射为内部身份，不改写原始设备名称和杆号牌。"""
    clean = cleanup_mode_b_item_text(text)
    patterns = (
        r"(?P<num>\d+)号台架(?:变压器|变)?",
        r"(?P<num>\d+)号变压器",
        r"#(?P<num>\d+)变(?:[ⅠⅡⅢⅣIV]+)?(?:支线)?",
        r"(?P<num>\d+)号变(?!压器)",
    )
    for pattern in patterns:
        match = re.search(pattern, clean)
        if not match:
            continue
        line_name = extract_exact_10kv_line_name(clean) or default_line
        location = normalize_parent_location(clean[:match.start()], line_name)
        if not line_name or not location:
            return None
        return line_name, location, int(match.group("num"))
    return None


def extract_low_voltage_reference_names(text, default_line=""):
    names = []
    clean = cleanup_mode_b_item_text(text)
    line_name = extract_exact_10kv_line_name(clean) or default_line
    if "低压线路维修" in clean and re.search(r"[:：]", clean):
        work_part = re.split(r"[:：]", clean, maxsplit=1)[1]
        work_part = re.sub(r"^\s*\d+\s*[.．、]\s*", "", work_part)
        if line_name:
            clean = f"{line_name}{work_part}"
    action_idx, _ = first_peidian_action(clean)
    identity = clean[:action_idx] if action_idx >= 0 else clean

    # 动作前常带“低压导线/出线电缆”等动作对象，不属于地点或设备身份。
    identity = re.sub(
        r"(?:低压)?(?:出线)?(?:电缆|导线|下户线)$",
        "",
        identity,
    )
    # 0.4kV杆段只保留所属设备；杆号不参与小结，但不能删掉台架、变压器等身份。
    identity = re.sub(
        r"(0\.4kV(?:#?\d+变|\d+号变)?(?:[ⅠⅡⅢⅣIV]+)?(?:支线)?)"
        r"(?:#?\d+(?:[-+]#?\d+)*(?:号)?杆?(?:至#?\d+(?:[-+]#?\d+)*(?:号)?杆?)?)$",
        r"\1",
        identity,
    )
    # “台架I低压出线#1-#3”中的#1-#3是低压出线杆段，不是台架编号。
    identity = re.sub(
        r"((?:台架(?:变)?[ⅠⅡⅢⅣIV]+)?低压出线)"
        r"#?\d+(?:[-+]#?\d+)*(?:号)?(?:杆)?(?:至#?\d+(?:[-+]#?\d+)*(?:号)?杆?)?$",
        r"\1",
        identity,
    )
    identity = identity.strip("，、；;。. :：")

    has_specific_identity = bool(
        re.search(
            r"(?:35|110|220)kV[^，,；;。:：]{1,18}站"
            r"|10kV[^，,；;。:：]{1,40}?线"
            r"|台架|配电室|变压器|台区|0\.4kV(?:#?\d+变|\d+号变)",
            identity,
        )
    )
    if has_specific_identity:
        if line_name and line_name not in identity:
            identity = f"{line_name}{identity}"
        add_unique(names, identity)
    elif "0.4kV" in clean:
        add_unique(names, "0.4kV")
    return names


def merge_low_voltage_descriptions(descriptions):
    actions = []
    for description in descriptions:
        body = mode_b_summary_plain_text(description)
        body = re.sub(r"^0\.4kV", "", body)
        for action in body.split("、"):
            add_unique(actions, action)
    return f"0.4kV{'、'.join(actions)}" if actions else ""


def compact_parallel_low_voltage_references(references):
    """相同站、线、地点和设备主体下合并I/II，其他身份保持独立。"""
    refs = [mode_b_summary_plain_text(ref) for ref in references if ref]
    patterns = (
        r"^(.*台架(?:变)?)([ⅠⅡⅢⅣIV]+)(低压出线)$",
        r"^(.*0\.4kV#?\d+变)([ⅠⅡⅢⅣIV]+)(支线)$",
    )
    compacted = []
    group_indexes = {}
    group_labels = {}
    for ref in refs:
        matched = False
        for pattern_index, pattern in enumerate(patterns):
            match = re.match(pattern, ref)
            if not match:
                continue
            key = (pattern_index, match.group(1), match.group(3))
            if key not in group_indexes:
                group_indexes[key] = len(compacted)
                group_labels[key] = []
                compacted.append("")
            add_unique(group_labels[key], match.group(2))
            compacted[group_indexes[key]] = (
                f"{match.group(1)}{'、'.join(group_labels[key])}{match.group(3)}"
            )
            matched = True
            break
        if not matched:
            add_unique(compacted, ref)
    return compacted


def collect_peidian_mode_b_descriptions(raw_text):
    descriptions = []
    items = split_mode_b_work_items_with_context(raw_text)
    if not items:
        items = [raw_text]
    default_line = extract_exact_10kv_line_name(cleanup_mode_b_work_text(raw_text))
    high_parent_keys = set()
    low_groups = {}
    pending_low_outputs = []

    for item_index, item in enumerate(items):
        clean = cleanup_mode_b_item_text(item)
        if re.search(r"(?:低压线路维修)?工程[:：]?$", clean):
            continue
        is_low_voltage = (
            "0.4kV" in clean
            or "低压线路维修" in clean
            or (
                "低压" in clean
                and any(word in clean for word in ("导线", "下户线", "出线电缆", "低压电缆"))
            )
        )
        if is_low_voltage:
            desc = summarize_low_voltage_peidian_item(clean)
            parent_key = extract_parent_device_identity(clean, default_line=default_line)
            references = extract_low_voltage_reference_names(
                clean, default_line=default_line
            )
            group_key = parent_key or (
                ("reference", tuple(references))
                if references
                else ("unmatched", item_index)
            )
            group = low_groups.setdefault(
                group_key,
                {"key": parent_key, "descriptions": [], "references": []},
            )
            add_unique(group["descriptions"], desc)
            for reference in references:
                add_unique(group["references"], reference)
        else:
            desc = summarize_high_voltage_peidian_item(clean)
            parent_key = extract_parent_device_identity(clean, default_line=default_line)
            if parent_key:
                high_parent_keys.add(parent_key)
            add_unique(descriptions, desc)

    low_group_count = len(low_groups)
    for group in low_groups.values():
        merged = merge_low_voltage_descriptions(group["descriptions"])
        if not merged:
            continue
        safe_single_group = (
            low_group_count == 1
            and group["key"] is not None
            and group["key"] in high_parent_keys
        )
        if safe_single_group:
            add_unique(descriptions, merged)
            continue

        references = group["references"]
        action_text = re.sub(r"^0\.4kV", "", merged)
        if references:
            pending_low_outputs.append((references, action_text))
        else:
            add_unique(descriptions, merged)

    outputs_by_action = {}
    output_order = []
    for references, action_text in pending_low_outputs:
        if action_text not in outputs_by_action:
            outputs_by_action[action_text] = []
            output_order.append(action_text)
        for reference in references:
            add_unique(outputs_by_action[action_text], reference)
    for action_text in output_order:
        references = outputs_by_action[action_text]
        references = compact_parallel_low_voltage_references(references)
        references = compact_contextual_identity_references(references)
        add_unique(descriptions, f"{'、'.join(references)}{action_text}")
    return descriptions


def summarize_mode_b_record_description(record, profession):
    """一条Excel记录对应一项小结；特定超长主题使用原文判定，其余从模式A继续生成。"""
    raw_text = record.get("work")
    special = v34_raw_mode_b_summary(raw_text, profession)
    if special:
        return v35_source_grounded_summary(raw_text, profession, special, raw_text)
    mode_a = summarize_work_clear(raw_text, profession).rstrip("。")
    candidate = summarize_work_mode_b(mode_a, profession).rstrip("。")
    return v35_source_grounded_summary(raw_text, profession, candidate, mode_a)


def _recount_summary_risks(records):
    """模式B统计的唯一可信来源：最终现场计划正式记录。"""
    counts = {level: 0 for level in RISK_LEVELS}
    counts["其他"] = 0
    for record in records:
        risk = str(record.get("risk") or "").strip()
        if risk in counts:
            counts[risk] += 1
        else:
            counts["其他"] += 1
    return counts


def summary_count_guard(records, risk_counter=None, grouped=None):
    """
    领导摘要项数硬校验。

    统计口径：
    - 一条最终现场计划Excel正式记录 = 1项工作；
    - 内部1./2./3.施工步骤不增加工作项数；
    - 模式B跨项融合不减少工作项数；
    - 总项数 = 各专业项数之和 = 各风险等级项数之和。

    外部传入的risk_counter仅用于交叉核对；最终输出以records重新统计为准。
    """
    records = list(records or [])
    total = len(records)

    if grouped is None:
        grouped = {}
        for record in records:
            grouped.setdefault(record.get("profession") or "其他", []).append(record)

    profession_counts = {prof: len(items) for prof, items in grouped.items()}
    profession_total = sum(profession_counts.values())
    recounted_risks = _recount_summary_risks(records)
    risk_total = sum(recounted_risks.values())

    passed = total == profession_total == risk_total
    if risk_counter is not None:
        supplied_total = sum(int(risk_counter.get(k, 0) or 0) for k in list(RISK_LEVELS) + ["其他"])
        if supplied_total != total:
            log_red(
                f"  【模式B统计校验】传入风险统计合计{supplied_total}项，与最终正式记录{total}项不一致；"
                "已按最终现场计划重新统计。"
            )

    if passed:
        prof_text = " + ".join(f"{prof}{count}" for prof, count in profession_counts.items()) or "无专业记录"
        risk_text = " + ".join(
            f"{level}{recounted_risks[level]}"
            for level in list(RISK_LEVELS) + ["其他"]
            if recounted_risks.get(level, 0)
        ) or "风险0"
        log(f"  模式B统计校验通过：总项数{total} = 专业合计{profession_total}（{prof_text}） = 风险合计{risk_total}（{risk_text}）")
    else:
        log_red(
            f"  【模式B统计校验失败】总项数={total}，专业合计={profession_total}，风险合计={risk_total}。"
            "摘要仍按最终正式记录统计，禁止按摘要句子数量改写项数。"
        )

    return {
        "total": total,
        "profession_counts": profession_counts,
        "risk_counts": recounted_risks,
        "profession_total": profession_total,
        "risk_total": risk_total,
        "passed": passed,
        "grouped": grouped,
    }


def build_mode_b_summary_lines(title, records, risk_counter):
    # 先按最终处理后的现场计划正式记录建立统计，再生成领导总览。
    grouped = {}
    for record in records:
        grouped.setdefault(record.get("profession") or "其他", []).append(record)

    count_info = summary_count_guard(records, risk_counter=risk_counter, grouped=grouped)
    total = count_info["total"]
    risk_counts = count_info["risk_counts"]

    risk_parts = [
        f"{level}风险等级{risk_counts.get(level, 0)}项"
        for level in RISK_LEVELS
        if risk_counts.get(level, 0)
    ]
    if risk_counts.get("其他", 0):
        risk_parts.append(f"其他风险等级{risk_counts['其他']}项")
    if not risk_parts:
        risk_parts = ["风险等级0项"]

    lines = [f"{title_to_work_plan_prefix(title)}工作计划{total}项，其中{'，'.join(risk_parts)}。", ""]
    ordered_professions = [p for p in MODEB_PROFESSION_ORDER if p in grouped]
    ordered_professions.extend([p for p in grouped if p not in ordered_professions])

    # V2.9：模式B先对当天全部模式A缩写版做一次跨专业强关联，再按专业输出。
    # 这里仍然只从模式A继续生成，不重新读取原始长文本。
    mode_a_cache = [
        summarize_work_clear(summary_source_text(record), record.get("profession") or "其他").rstrip("。")
        for record in records
    ]
    associations = v29_build_day_associations(records, mode_a_cache)
    index_by_id = {id(record): idx for idx, record in enumerate(records)}
    if associations:
        for src_idx, rel in sorted(associations.items()):
            log(
                f"  模式B强关联：#{src_idx + 1} -> #{rel['target_idx'] + 1}，"
                f"得分{rel['score']}（{'/'.join(rel['reasons'])}），主体={rel['label']}"
            )

    first = True
    for prof in ordered_professions:
        prof_records = grouped[prof]
        if prof == "营销":
            # 营销专业严格沿用5.2原版模式B：
            # 高压计量明细 + 营销动作类型 + 涉及单位，不套V2/V2.9通用压缩或关联逻辑。
            actions = extract_marketing_action_types(prof_records)
            metering_details = extract_high_voltage_metering_marketing_details(prof_records)
            non_metering_records = [
                record for record in prof_records
                if not is_high_voltage_metering_marketing_work(record.get("work"))
            ]
            units = extract_marketing_units(prof_records)
            unit_text = "、".join(units) if units else "未提取到涉及单位"
            prefix = "其中" if first else ""
            if actions:
                action_text = "、".join(actions)
            elif non_metering_records:
                action_text = "、".join(
                    summarize_work_compact(r.get("work"), prof).rstrip("。")
                    for r in non_metering_records
                )
            else:
                action_text = ""
            marketing_parts = []
            if metering_details:
                marketing_parts.append(f"高压计量：{'、'.join(metering_details)}")
            if action_text:
                marketing_parts.append(action_text)
            marketing_parts.append(f"涉及单位：{unit_text}")
            marketing_text = mode_b_summary_plain_text("；".join(marketing_parts))
            lines.append(f"{prefix}营销{len(prof_records)}项。（{marketing_text}。）")
        else:
            descs = []
            for record in prof_records:
                idx = index_by_id[id(record)]
                mode_a = mode_a_cache[idx]
                desc = v29_leader_summary_with_context(
                    mode_a,
                    prof,
                    associations.get(idx),
                    raw_text=record.get("work"),
                )
                if V35_REVIEW_LIBRARY_AVAILABLE:
                    try:
                        collected = v35_collect_pending_case(
                            record.get("work"),
                            desc,
                            profession=prof,
                            risk_level=record.get("risk"),
                            source_ref=f"现场计划正式记录#{idx + 1}",
                        )
                        if collected:
                            hits = v35_search_reviewed_cases(
                                record.get("work"), profession=prof, top_k=1
                            )
                            hint = ""
                            if hits:
                                hit = hits[0]
                                hint = (
                                    f"；最相近已审核案例得分{hit['score']:.3f}，"
                                    f"来源={hit['case']['source_ref']}"
                                )
                            log(f"  模式B待审核语料已收集：正式记录#{idx + 1}{hint}")
                    except Exception as error:
                        log(f"  模式B待审核语料收集失败，不影响正式输出：{error}")
                desc = mode_b_summary_plain_text(desc)
                if desc:
                    descs.append(desc)
            if len(descs) >= 2:
                desc_text = format_numbered_mode_b_descriptions(descs)
            elif descs:
                desc_text = descs[0].rstrip("。") + "。"
            else:
                desc_text = ""
            prefix = "其中" if first else ""
            lines.append(f"{prefix}{prof}{len(prof_records)}项，（{desc_text}）")
        first = False
    return lines



def process_excel(path, sheet_name=0):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            df = pd.read_excel(path, header=None, sheet_name=sheet_name, engine='openpyxl', dtype=object)
        except Exception as e:
            log(f"读取 Excel 失败: {path} -> {e}")
            return {"title": "", "records": []}, {"二级": 0, "三级": 0, "四级": 0, "五级": 0, "其他": 0}

    data = df.values.tolist()
    if len(data) < 2:
        return {"title": "", "records": []}, {"二级": 0, "三级": 0, "四级": 0, "五级": 0, "其他": 0}

    title_row = data[0]
    title = next((str(x).strip() for x in title_row if x not in [None, ''] and str(x).strip() != ''), '现场作业计划')

    records = []
    empty_n_counter = 0

    for ridx, row in enumerate(data[2:], start=3):
        def cell(cidx):
            return row[cidx] if cidx < len(row) else None

        if all(is_blank_value(cell(i)) for i in range(min(len(row), 18))):
            continue

        # 保持原来 18 列映射
        A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R = [cell(i) for i in range(18)]

        # TXT最终提取只认“工作内容”非空的真实作业行。
        # Excel处理过程中产生的空白续行、分页辅助行、只有格式/序号但无工作内容的行均不进入汇总。
        if is_blank_value(D):
            continue

        seq = str(A).strip() if not is_blank_value(A) else str(ridx - 2)
        raw_leader = J

        leader = ''
        if not is_blank_value(raw_leader):
            leader = clean_person_name(raw_leader)
            if leader == "":
                leader = None

        dt_k = parse_datetime_from_cell(K)
        dt_l = parse_datetime_from_cell(L)
        start_hm = format_HHMM(dt_k)
        end_hm = format_HHMM(dt_l)
        time_part = f"{start_hm} - {end_hm}" if start_hm or end_hm else ""

        risk = str(Q).strip() if not is_blank_value(Q) else ''
        grid_risk = ''
        if not is_blank_value(R) and str(R).strip() not in ['无', '无 ']:
            grid_risk = str(R).strip()

        need_no = None
        if not is_blank_value(N):
            nclean = remove_phone_and_noise(N)
            same_inout = nclean if nclean else ''
        else:
            empty_n_counter += 1
            need_no = empty_n_counter
            same_inout = f"需要领导{empty_n_counter}"

        profession = classify_profession(B, raw_leader)
        records.append({
            "order": len(records) + 1,
            "source_row": ridx,
            "seq": seq,
            "profession": profession,
            "raw_profession": str(B).strip() if not is_blank_value(B) else "",
            "work": D,
            "unit": H,
            "leader": leader,
            "time_part": time_part,
            "risk": risk,
            "grid_risk": grid_risk,
            "same_inout": same_inout,
            "need_leader_no": need_no,
        })

    risk_counter = {"二级": 0, "三级": 0, "四级": 0, "五级": 0, "其他": 0}
    for record in records:
        r = record.get("risk", "").strip()
        if r in risk_counter:
            risk_counter[r] += 1
        else:
            risk_counter["其他"] += 1

    log(f"\n处理并排序输出（来自文件：{path}）")
    for idx, record in enumerate(sort_mode_a_records(records), 1):
        first_work = normalize_work_text(record.get("work")).split("。")[0]
        need_mark = f"；{record['same_inout']}" if record.get("need_leader_no") else ""
        log(f"  {idx}. 专业:{record.get('profession')} 工作内容:{first_work[:60]} -> 风险等级: {record.get('risk') or '其他'}{need_mark}")
    log("========== 文件统计 ==========")
    log(f"总条目={sum(risk_counter.values())}  二级={risk_counter['二级']}  三级={risk_counter['三级']}  四级={risk_counter['四级']}  五级={risk_counter['五级']}  其他={risk_counter['其他']}\n")

    return {"title": title, "records": records}, risk_counter


def write_text_file(path, lines):
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines).rstrip() + "\n")


def flatten_site_summaries(site_summaries):
    all_records = []
    title = "现场作业计划"
    for summary in site_summaries:
        if summary.get("title") and title == "现场作业计划":
            title = summary.get("title")
        all_records.extend(summary.get("records") or [])
    return title, all_records


def write_summary_txt(cwd, site_summaries, site_files_processed, generated_daogang_files=None):
    title, records = flatten_site_summaries(site_summaries)
    legacy_path = os.path.join(cwd, "汇总.txt")

    if not records:
        lines = ["未提取到有效作业条目，请检查现场作业计划的数据行和筛选条件。"]
        write_text_file(legacy_path, lines)
        log(f"\n汇总完成，输出文件：{legacy_path}，提取条目 0 条")
        return [legacy_path]

    prepare_summary_source_entries(records)

    need_count = sum(1 for r in records if r.get("need_leader_no"))
    mode_a_name = "汇总-模式A-需要领导.txt" if need_count else "汇总-模式A.txt"
    mode_a_path = os.path.join(cwd, mode_a_name)
    mode_a_lines = build_mode_a_summary_lines(title, records)
    write_text_file(mode_a_path, mode_a_lines)
    write_text_file(legacy_path, mode_a_lines)

    total_risk_counter = {"二级": 0, "三级": 0, "四级": 0, "五级": 0, "其他": 0}
    for record in records:
        risk = record.get("risk") or "其他"
        total_risk_counter[risk if risk in total_risk_counter else "其他"] += 1
    mode_b_path = os.path.join(cwd, "模式B-输出小结.txt")
    mode_b_lines = build_mode_b_summary_lines(title, records, total_risk_counter)
    write_text_file(mode_b_path, mode_b_lines)

    log(f"\n模式A汇总完成：{mode_a_path}，记录 {len(records)} 条，需要领导 {need_count} 条")
    log(f"兼容输出已同步：{legacy_path}")
    log(f"模式B输出小结完成：{mode_b_path}")
    return [mode_a_path, legacy_path, mode_b_path]


def clear_previous_summary_txt(cwd):
    """
    删除上一次运行留下的汇总TXT，避免本轮Excel处理失败时把旧文件误认为新结果。
    只删除本程序固定生成的TXT，不碰用户其他文本文件。
    """
    names = (
        "汇总.txt",
        "汇总-模式A.txt",
        "汇总-模式A-需要领导.txt",
        "模式B-输出小结.txt",
        "原始填报错误提醒报告.txt",
    )
    removed = []
    for name in names:
        path = os.path.join(cwd, name)
        if os.path.isfile(path):
            try:
                os.remove(path)
                removed.append(name)
            except Exception as e:
                log_red(f"【旧TXT清理失败】{name}: {e}")
    if removed:
        log("已清理上次运行的汇总TXT：" + "、".join(removed))
    return removed


# ================= 主入口 =================
def main():
    SOURCE_ENTRY_ISSUES.clear()
    SOURCE_ENTRY_ISSUE_KEYS.clear()
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    try:
        os.chdir(exe_dir)
    except Exception:
        pass
    cwd = os.getcwd()

    log_section("程序启动")
    log(f"工作目录：{cwd}")
    log("5.2-V2.8：以V2.7为主干；恢复到岗到位Windows生成弹窗及5.2生成链路；模式A=缩写版+完整版；模式B=领导视角总览且保留核心动作；统计硬校验；营销沿用5.2；TXT最后生成；打印余量+2pt。")

    if os.path.basename(cwd).startswith("自动备份_"):
        log("检测到当前位于自动备份目录：为保护备份原貌，本次不处理任何文件。")
        return

    # 先清除固定名称的旧汇总，防止旧结果/旧样例被误认为本轮新生成。
    clear_previous_summary_txt(cwd)

    all_excel_files = [
        f for f in os.listdir(cwd)
        if (f.endswith(".xlsx") or f.endswith(".xlsm"))
        and not f.startswith("~$")
        and not is_backup_path(os.path.join(cwd, f))
    ]
    files = [f for f in all_excel_files if "（到岗到位）" not in f]
    if not files:
        log("未找到需处理的xlsx/xlsm文件")
        # 没有Excel时也只在处理流程结束处给出明确空结果，不生成模式A/B旧内容。
        write_summary_txt(cwd, [], 0)
        save_report(cwd)
        return

    # ---------- 第1阶段：先完成并保存全部Excel处理 ----------
    kill_excel_processes()
    title_map = {}
    success_files = []
    generated_daogang_files = []
    cache_cleared = False
    for idx, file in enumerate(files, 1):
        path = os.path.join(cwd, file)
        title = "处理失败"
        try:
            for attempt in range(2):
                run_mode = detect_processing_mode("", file)
                title = preprocess_excel(path, cwd, mode=run_mode if run_mode != "normal" else None)
                if title not in ["缓存错误", "处理失败"]:
                    break
                if title == "缓存错误" and not cache_cleared:
                    log("  检测到缓存错误，尝试清除缓存并重试...")
                    kill_excel_processes()
                    if clear_win32com_cache():
                        cache_cleared = True
                        continue
                    break
                break
        except Exception as e:
            log_red(f"【Excel处理异常，不终止其他文件】{file}: {e}")
            log(traceback.format_exc())
            title = "处理失败"
        title_map[file] = title
        if title not in ["文件被占用", "只读无法修改", "处理失败", "文件不存在", "缓存错误"]:
            success_files.append(file)
        log(f"[Excel预处理 {idx}/{len(files)}] 完成：{file} -> 标题：{title}")

        if file in success_files and "现场作业计划" in str(title) and "风险管控" not in str(title):
            try:
                generated_path = maybe_generate_daogang_version(path, cwd)
                if generated_path:
                    generated_daogang_files.append(generated_path)
            except Exception as e:
                log_red(f"【到岗到位生成异常，不影响TXT】{file}: {e}")
                log(traceback.format_exc())

    # ---------- 第2阶段：Excel全部处理完后，再从最终正式现场计划生成TXT ----------
    # 只读取“本轮处理成功 + 标题确认为现场作业计划”的原始源文件；
    # 不使用处理前缓存、不读取到岗到位派生表、不读取风险管控表，也不读取程序内测试案例。
    final_site_files = [
        file for file in success_files
        if "现场作业计划" in str(title_map.get(file, ""))
        and "风险管控" not in str(title_map.get(file, ""))
        and "（到岗到位）" not in file
    ]

    site_summaries = []
    site_files_processed = 0
    total_risk_counter = {"二级": 0, "三级": 0, "四级": 0, "五级": 0, "其他": 0}

    for idx, file in enumerate(final_site_files, 1):
        path = os.path.join(cwd, file)
        try:
            summary, file_risk_count = process_excel(path)
            # 再次用最终文件自身标题做一道校验，避免仅凭预处理返回标题误收其他表。
            final_title = str(summary.get("title") or "")
            if "现场作业计划" not in final_title or "风险管控" in final_title:
                log(f"[最终TXT提取 {idx}/{len(final_site_files)}] 跳过：{file} -> {final_title}")
                continue
            records = summary.get("records") or []
            if not records:
                log(f"[最终TXT提取 {idx}/{len(final_site_files)}] {file}：0条有效工作内容，跳过空表")
                continue
            site_files_processed += 1
            site_summaries.append(summary)
            for k in total_risk_counter:
                total_risk_counter[k] += file_risk_count.get(k, 0)
            log(f"[最终TXT提取 {idx}/{len(final_site_files)}] {file}：{len(records)} 条有效作业")
        except Exception as e:
            log_red(f"【最终TXT提取异常】{file}: {e}")
            log(traceback.format_exc())

    # 所有表格处理、保存及到岗到位生成完成后，才真正写TXT。
    write_summary_txt(cwd, site_summaries, site_files_processed, generated_daogang_files)
    log("TXT已在全部表格处理完成后生成，数据来源为本轮处理成功的最终现场作业计划。")
    write_source_entry_issue_report(cwd)

    total_entries = sum(total_risk_counter.values())
    log("========== 全部文件汇总统计 ==========")
    log(f"最终TXT有效条目数: {total_entries}")
    for key in ("二级", "三级", "四级", "五级", "其他"):
        log(f"{key}: {total_risk_counter[key]}")
    log("====================================")
    save_report(cwd)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            log_red(f"【程序未捕获异常】{e}")
            log(traceback.format_exc())
            save_report(os.getcwd())
        except Exception:
            print(traceback.format_exc())
        try:
            input("程序发生异常，错误已写入处理报告。按回车键退出...")
        except Exception:
            pass
