# -*- coding: utf-8 -*-
"""
邮件附件编号检测工具 V2.7
============================================================
本版本重点优化
1. 结果预览：
   - 底部结果区改为“统计卡片 + 搜索定位 + 四个结果页”
   - 结果区默认占更大空间，运行日志缩小到右侧
   - 编号总览改为“可调列数”的 Excel 式网格，支持 6~36 列
   - 缺失 / 异常 / 重复分别独立列表，可直接滚动查看
   - 搜索可正确定位“正常 / 缺失 / 重复”编号
   - 双击问题编号可定位到总览
2. 老旧/新电脑性能与稳定性：
   - 扫描线程可选 1/2/4/8/16/32/64
   - 根据 CPU、内存、32/64 位 Python 自动给出推荐值，并在高线程数风险场景自动降级，避免老旧电脑卡死
   - 使用限量任务池，不一次性提交全部 EML
   - 扫描期间锁定全部非停止功能，避免并发操作破坏状态
   - 不使用全局 WaitCursor，不让弹窗继承沙漏状态
   - 关闭窗口时等待扫描任务安全退出，避免 QThread 未结束导致崩溃
   - 结果网格列数可配置 6~36，真正按用户设置生效
   - 优化表格批量刷新，减少老旧电脑 UI 重绘
3. 稳定性 / Bug 修复：
   - 停止扫描不再一次性提交所有 EML，减少停止等待时间
   - 使用 threading.Event 作为停止信号
   - 文件打开使用 QDesktopServices，避免 shell 命令拼接
   - 所有临时文件写入后再原子替换
   - 分析结果、无效行、XLSX 均有异常保护
   - 预览搜索修复了“缺失:75 / 重复:75(2次)”无法定位的问题
   - Excel 输出使用 openpyxl，保证彩色 XLSX 功能正常
4. 保留 V2.2 核心功能：
   - 扫描指定目录及子目录全部 .eml
   - 仅提取 Content-Disposition: attachment 的附件
   - RFC2047 / UTF-8 / GBK 等附件名解码
   - 生成 附件名称列表.csv
   - 检测缺失 / 异常 / 重复编号
   - 输出分析结果 CSV
   - 输出无法识别编号 CSV
   - 可选彩色 XLSX
   - 自动查找程序目录中的 EML
   - 后台线程扫描
   - 可停止扫描
   - 全局异常捕获与错误日志

依赖：
    pip install PyQt5 openpyxl

推荐打包（Win7 SP1 老电脑建议使用 onedir）：
    pyinstaller --noconfirm --clean --onedir --windowed ^
      --name "Email_Number_Check_V2.7" ^
      --exclude-module matplotlib ^
      --exclude-module IPython ^
      --exclude-module tkinter ^
      --exclude-module PySide ^
      --exclude-module PySide2 ^
      Email_Number_Check_V2.7.py

说明：
- 彩色 XLSX 重新使用 openpyxl，保证 Excel 兼容性。CSV 默认不依赖 openpyxl。
- 建议在干净 venv 中打包，只安装 PyQt5、openpyxl、PyInstaller。
- Win7 老电脑优先推荐 --onedir：比 --onefile 启动更快，避免每次启动都解压整个程序。
"""

import csv
import os
import re
import sys
import time
import traceback
import threading
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from collections import Counter
from email.header import decode_header
from email.parser import BytesParser
from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import QObject, QThread, pyqtSignal, Qt, QUrl
from PyQt5.QtGui import QFont, QTextCursor, QColor, QStandardItemModel
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QPushButton, QLabel, QLineEdit, QSpinBox, QProgressBar,
    QPlainTextEdit, QGridLayout, QVBoxLayout, QHBoxLayout,
    QGroupBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QSplitter, QAbstractItemView, QTabWidget,
    QFrame, QComboBox, QMenu, QAction
)
from PyQt5.QtGui import QDesktopServices


APP_VERSION = "V2.7"
APP_NAME = "Email附件检测分析工具 V2.7"
ATTACHMENT_CSV = "附件名称列表.csv"
RESULT_CSV = "编号分析结果.csv"
INVALID_CSV = "无法识别编号.csv"
XLSX_RESULT = "编号分析结果.xlsx"
HELP_DOCX = "Email附件数据分析v2.0使用说明.docx"
ERROR_LOG = "程序错误日志.txt"

# CSV / Excel 文件保持与旧版接近的矩阵尺寸，避免影响原有使用习惯。
MATRIX_COLUMNS = 36
MATRIX_ROWS = 37

# GUI 预览列数范围。
# 与截图类似，允许很多列；用户设置多少就显示多少，不再被视口宽度强制压缩。
UI_MIN_COLUMNS = 6
UI_MAX_COLUMNS = 36
UI_DEFAULT_COLUMNS = 18
UI_CELL_WIDTH = 88

# 防止误输入超大编号范围导致 list/set 占用大量内存、界面长时间无响应。
MAX_EXPECTED_NUMBERS = 2_000_000


# ============================================================
# 基础工具
# ============================================================

def get_app_dir() -> str:
    """兼容普通 Python 运行和 PyInstaller 打包运行。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_default_mail_dir() -> str:
    """默认邮件目录：程序目录的父级目录。启动时不递归扫描磁盘。"""
    app_dir = get_app_dir()
    parent_dir = os.path.dirname(app_dir)
    return parent_dir if os.path.isdir(parent_dir) else app_dir


def write_error_log(message: str) -> None:
    try:
        path = os.path.join(get_app_dir(), ERROR_LOG)
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80 + "\n")
            f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
            f.write(message + "\n")
    except Exception:
        pass


def safe_remove(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def format_int(value: int) -> str:
    return "{:,}".format(int(value))


def get_total_memory_bytes() -> int:
    """尽量不依赖第三方库获取物理内存；Windows 通过 GlobalMemoryStatusEx。"""
    try:
        if os.name == "nt":
            import ctypes
            class MEMORYSTATUSEX(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]
            status = MEMORYSTATUSEX()
            status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return int(status.ullTotalPhys)
    except Exception:
        pass
    return 0


def get_runtime_capabilities() -> Dict[str, int]:
    """返回线程配置所需的轻量级系统信息。"""
    cpu = os.cpu_count() or 2
    memory_bytes = get_total_memory_bytes()
    memory_gb = int(memory_bytes / (1024 ** 3)) if memory_bytes else 0
    pointer_bits = 64 if sys.maxsize > 2 ** 32 else 32
    return {"cpu": cpu, "memory_gb": memory_gb, "pointer_bits": pointer_bits}


def recommend_scan_threads(cpu: int, memory_gb: int, pointer_bits: int) -> Tuple[int, int, str]:
    """返回 (推荐线程, 安全上限, 说明)。EML 扫描偏 I/O，因此默认不盲目等于 CPU 核数。"""
    if cpu <= 2:
        safe_cap = 8
        rec = 4
        reason = "CPU 核心较少，建议 4 线程；超过 8 线程可能增加磁盘与上下文切换负担。"
    elif cpu <= 4:
        safe_cap = 16
        rec = 8
        reason = "中低端电脑建议 8 线程；16 可尝试，32/64 不建议。"
    elif cpu <= 8:
        safe_cap = 32 if memory_gb >= 8 else 16
        rec = 8 if memory_gb < 8 else 12
        reason = "中高端电脑建议 8~12 线程；SSD 可尝试 16/32。"
    elif cpu <= 16:
        safe_cap = 64 if memory_gb >= 16 else 32
        rec = 16 if memory_gb >= 8 else 12
        reason = "多核电脑可尝试 16~32 线程；SSD + 大内存可尝试 64。"
    else:
        safe_cap = 64 if memory_gb >= 16 else 32
        rec = 32 if memory_gb >= 16 else 16
        reason = "高核心电脑可尝试 32 线程；具备 SSD + 16GB 以上内存时可尝试 64。"
    if pointer_bits == 32:
        safe_cap = min(safe_cap, 16)
        rec = min(rec, 8)
        reason += " 当前为 32 位 Python，已将安全上限限制为 16 线程。"
    return rec, safe_cap, reason


def decode_filename(filename: Optional[str]) -> Optional[str]:
    """RFC2047 / 多编码附件名解码。"""
    if not filename:
        return None

    try:
        result = []
        for text, charset in decode_header(filename):
            if isinstance(text, bytes):
                encodings = []
                if charset:
                    encodings.append(charset)
                encodings.extend(["utf-8", "gb18030", "gbk", "big5", "latin1"])

                decoded = None
                for enc in encodings:
                    try:
                        decoded = text.decode(enc, errors="strict")
                        break
                    except Exception:
                        continue

                if decoded is None:
                    decoded = text.decode("utf-8", errors="replace")
                text = decoded

            result.append(str(text))

        value = "".join(result).strip()
        return value or None
    except Exception:
        return str(filename)


def extract_number(filename: str) -> Optional[int]:
    """
    与旧版本逻辑保持一致：取附件名称中的第一个连续数字。
    0038.jpg -> 38
    ABC0038.jpg -> 38
    """
    if not filename:
        return None

    match = re.search(r"\d+", str(filename).strip())
    if not match:
        return None

    try:
        return int(match.group())
    except Exception:
        return None


def find_eml_folders(base_path: str) -> List[str]:
    """查找所有包含 .eml 的目录。"""
    found = []
    try:
        for root, _dirs, files in os.walk(base_path):
            if any(name.lower().endswith(".eml") for name in files):
                found.append(root)
    except Exception:
        pass
    return sorted(set(found))


# ============================================================
# EML 解析
# ============================================================

def get_attachments_from_eml(
    eml_path: str
) -> Tuple[List[Dict[str, str]], Optional[str]]:
    """
    返回：
        attachments, error_message

    只保留 Content-Disposition: attachment。
    单封邮件损坏不会中断整个扫描。
    """
    attachments = []

    try:
        with open(eml_path, "rb") as f:
            msg = BytesParser().parse(f)

        for part in msg.walk():
            if part.is_multipart():
                continue

            disposition = part.get_content_disposition()
            if disposition != "attachment":
                continue

            filename = decode_filename(part.get_filename())
            if not filename:
                continue

            attachments.append({
                "filename": filename,
                "disposition": disposition or "",
                "content_type": part.get_content_type() or "",
            })

        return attachments, None
    except Exception as e:
        return [], "{} | {}".format(eml_path, e)


# ============================================================
# 编号分析核心
# ============================================================

def generate_expected_numbers(start: int, end: int, step: int) -> List[int]:
    if step <= 0 or end < start:
        return []
    return list(range(start, end + 1, step))


def analyze_numbers(
    actual_numbers: List[int],
    start: int,
    end: int,
    step: int
):
    expected_numbers = generate_expected_numbers(start, end, step)
    actual_set = set(actual_numbers)
    expected_set = set(expected_numbers)

    missing_numbers = sorted(expected_set - actual_set)
    abnormal_numbers = sorted(actual_set - expected_set)

    duplicates = {
        number: count
        for number, count in Counter(actual_numbers).items()
        if count > 1
    }

    return expected_numbers, missing_numbers, abnormal_numbers, duplicates


def read_numbers_from_csv(csv_path: str):
    numbers = []
    invalid_rows = []
    total_rows = 0

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        next(reader, None)

        for row_index, row in enumerate(reader, start=2):
            total_rows += 1
            if not row:
                continue

            filename = row[0].strip() if row else ""
            if not filename:
                continue

            number = extract_number(filename)
            if number is None:
                invalid_rows.append((row_index, filename))
            else:
                numbers.append(number)

    return numbers, invalid_rows, total_rows


# ============================================================
# CSV / XLSX 保存
# ============================================================

def save_attachment_csv(folder_path: str, rows: List[List[str]]) -> str:
    output_csv = os.path.join(folder_path, ATTACHMENT_CSV)
    temp_csv = output_csv + ".tmp"

    safe_remove(temp_csv)

    try:
        with open(temp_csv, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([
                "附件名称", "邮件文件名", "邮件完整路径",
                "Content-Disposition", "Content-Type"
            ])
            writer.writerows(rows)

        os.replace(temp_csv, output_csv)
        return output_csv
    except Exception:
        safe_remove(temp_csv)
        raise


def save_invalid_rows(
    folder_path: str,
    invalid_rows: List[Tuple[int, str]]
) -> Optional[str]:
    if not invalid_rows:
        # 避免上一次分析留下的旧文件被误认为本次仍存在无法识别编号。
        safe_remove(os.path.join(folder_path, INVALID_CSV))
        return None

    path = os.path.join(folder_path, INVALID_CSV)
    temp = path + ".tmp"

    safe_remove(temp)

    try:
        with open(temp, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["CSV行号", "附件名称"])
            writer.writerows(invalid_rows)

        os.replace(temp, path)
        return path
    except Exception:
        safe_remove(temp)
        raise


def build_number_matrix(
    expected_numbers,
    actual_numbers,
    missing_numbers,
    duplicate_numbers,
    columns=MATRIX_COLUMNS,
    rows=MATRIX_ROWS
):
    """按旧版 CSV / Excel 的竖向填充方式构造矩阵。"""
    missing_set = set(missing_numbers)
    actual_counter = Counter(actual_numbers)

    matrix = []
    total = len(expected_numbers)

    for row_index in range(rows):
        row = []
        for col_index in range(columns):
            index = col_index * rows + row_index

            if index >= total:
                row.append(("", "empty"))
                continue

            number = expected_numbers[index]

            if number in missing_set:
                row.append(("缺失:{}".format(number), "missing"))
            elif actual_counter[number] > 1:
                row.append(
                    ("重复:{}({}次)".format(number, actual_counter[number]), "duplicate")
                )
            else:
                row.append((str(number), "normal"))

        matrix.append(row)

    return matrix


def save_matrix_csv(
    folder_path: str,
    start: int,
    end: int,
    step: int,
    expected_numbers: List[int],
    actual_numbers: List[int],
    missing_numbers: List[int],
    abnormal_numbers: List[int],
    duplicate_numbers: Dict[int, int],
) -> str:
    """生成兼容旧版本的横向矩阵 CSV。"""
    path = os.path.join(folder_path, RESULT_CSV)
    temp = path + ".tmp"

    safe_remove(temp)

    export_rows = MATRIX_ROWS
    export_columns = max(1, (len(expected_numbers) + export_rows - 1) // export_rows)
    matrix = build_number_matrix(
        expected_numbers,
        actual_numbers,
        missing_numbers,
        duplicate_numbers,
        columns=export_columns,
        rows=export_rows
    )

    try:
        with open(temp, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)

            writer.writerow([APP_NAME])
            writer.writerow([
                "编号规则", "开始编号", start,
                "结束编号", end,
                "步长", step
            ])
            writer.writerow([
                "附件记录数", len(actual_numbers),
                "理论编号数", len(expected_numbers),
                "缺失", len(missing_numbers),
                "异常", len(abnormal_numbers),
                "重复", len(duplicate_numbers)
            ])
            writer.writerow([])
            writer.writerow([
                "竖向编号矩阵（每列从上到下递增；缺失/重复带文字标记）"
            ])

            for row in matrix:
                writer.writerow([value for value, _status in row])

            writer.writerow([])
            writer.writerow(["异常编号"])

            if abnormal_numbers:
                abnormal_columns = max(1, (len(abnormal_numbers) + MATRIX_ROWS - 1) // MATRIX_ROWS)
                abnormal_matrix = []
                for row_index in range(MATRIX_ROWS):
                    row = []
                    for col_index in range(abnormal_columns):
                        index = col_index * MATRIX_ROWS + row_index
                        row.append(
                            abnormal_numbers[index]
                            if index < len(abnormal_numbers)
                            else ""
                        )
                    abnormal_matrix.append(row)
                writer.writerows(abnormal_matrix)
            else:
                writer.writerow(["无"])

            writer.writerow([])
            writer.writerow(["重复编号明细", "出现次数"])

            if duplicate_numbers:
                for number, count in sorted(duplicate_numbers.items()):
                    writer.writerow([number, count])
            else:
                writer.writerow(["无", ""])

        os.replace(temp, path)
        return path
    except Exception:
        safe_remove(temp)
        raise


# ---------- 彩色 XLSX（可选 openpyxl） ----------

def try_save_styled_xlsx(
    folder_path: str,
    start: int,
    end: int,
    step: int,
    expected_numbers: List[int],
    actual_numbers: List[int],
    missing_numbers: List[int],
    abnormal_numbers: List[int],
    duplicate_numbers: Dict[int, int],
) -> Optional[str]:
    """生成彩色 Excel，使用命名参数兼容不同 openpyxl 版本。"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except Exception:
        return None

    path = os.path.join(folder_path, "编号分析结果.xlsx")
    temp = path + ".tmp.xlsx"
    safe_remove(temp)

    wb = Workbook()
    ws = wb.active
    ws.title = "编号分析结果"

    green = PatternFill(fill_type="solid", fgColor="92D050")
    red = PatternFill(fill_type="solid", fgColor="FF0000")
    yellow = PatternFill(fill_type="solid", fgColor="FFD966")
    blue = PatternFill(fill_type="solid", fgColor="5B9BD5")
    title_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    white_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    export_rows = MATRIX_ROWS
    export_columns = max(1, (len(expected_numbers) + export_rows - 1) // export_rows)
    # Excel 单工作表最多 16384 列，超出则拒绝而不是生成截断结果。
    if export_columns > 16384:
        safe_remove(temp)
        return None

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=export_columns)
    c = ws.cell(row=1, column=1, value="邮件附件编号自动序列号（绿色正常 / 红色缺失 / 黄色异常 / 蓝色重复）")
    c.fill = title_fill
    c.font = Font(bold=True, size=14)
    c.alignment = center

    headers = [("开始", start), ("结束", end), ("步长", step)]
    col = 1
    for label, value in headers:
        ws.cell(row=2, column=col, value=label).font = Font(bold=True)
        ws.cell(row=2, column=col + 1, value=value)
        col += 2

    matrix = build_number_matrix(
        expected_numbers, actual_numbers, missing_numbers, duplicate_numbers,
        columns=export_columns, rows=export_rows
    )
    start_row = 3
    for r_index, row in enumerate(matrix):
        for c_index, (value, status) in enumerate(row, start=1):
            cell = ws.cell(row=start_row + r_index, column=c_index, value=value)
            cell.alignment = center
            cell.border = border
            if status == "normal":
                cell.fill = green
            elif status == "missing":
                cell.fill = red
                cell.font = white_font
            elif status == "duplicate":
                cell.fill = blue
                cell.font = white_font

    abnormal_start = start_row + MATRIX_ROWS + 2
    ws.cell(row=abnormal_start, column=1, value="异常编号").font = Font(bold=True)
    if abnormal_numbers:
        for i, number in enumerate(abnormal_numbers):
            row = abnormal_start + 1 + (i % MATRIX_ROWS)
            col = 1 + (i // MATRIX_ROWS)
            if col > 16384:
                safe_remove(temp)
                return None
            cell = ws.cell(row=row, column=col, value=number)
            cell.fill = yellow
            cell.border = border
            cell.alignment = center
    else:
        ws.cell(row=abnormal_start + 1, column=1, value="无")

    abnormal_columns = max(1, (len(abnormal_numbers) + MATRIX_ROWS - 1) // MATRIX_ROWS)
    duplicate_start = abnormal_start + max(3, abnormal_columns + 2) + 2
    ws.cell(row=duplicate_start, column=1, value="重复编号明细").font = Font(bold=True)
    ws.cell(row=duplicate_start, column=2, value="出现次数").font = Font(bold=True)
    if duplicate_numbers:
        for i, (number, count) in enumerate(sorted(duplicate_numbers.items()), start=1):
            ws.cell(row=duplicate_start + i, column=1, value=number)
            ws.cell(row=duplicate_start + i, column=2, value=count)

    for col_index in range(1, export_columns + 1):
        ws.column_dimensions[get_column_letter(col_index)].width = 12
    ws.freeze_panes = "A3"

    try:
        wb.save(temp)
        os.replace(temp, path)
        return path
    except Exception:
        safe_remove(temp)
        return None

# ============================================================
# 扫描线程
# ============================================================

class ScanWorker(QObject):
    """
    工作线程负责调度：
    - 不一次性提交全部 EML，停止响应更快
    - 最多同时运行 max_workers 个任务
    - 主线程只处理 UI
    """

    progress = pyqtSignal(int, int, int, str)
    log = pyqtSignal(str)
    finished = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, folder_path: str, max_workers: int = 8):
        super().__init__()
        self.folder_path = folder_path
        self._stop_event = threading.Event()
        cpu = os.cpu_count() or 2
        # I/O 为主的 EML 扫描通常 8 个线程是较稳妥的上限；16 不一定更快，
        # 尤其是机械硬盘/老旧 CPU，过多线程反而增加磁盘寻道与上下文切换。
        self.max_workers = max(1, min(64, int(max_workers)))
        self.cpu_count = cpu

    def request_stop(self):
        self._stop_event.set()

    @staticmethod
    def _parse_one(eml_path: str):
        attachments, error = get_attachments_from_eml(eml_path)
        return eml_path, attachments, error

    def run(self):
        try:
            if not os.path.isdir(self.folder_path):
                raise RuntimeError("邮件目录不存在：{}".format(self.folder_path))

            eml_files = []
            for root, _dirs, files in os.walk(self.folder_path):
                if self._stop_event.is_set():
                    self.finished.emit({
                        "stopped": True,
                        "processed": 0,
                        "total": 0,
                        "attachments": 0,
                        "output": None,
                        "errors": 0,
                    })
                    return

                for name in files:
                    if name.lower().endswith(".eml"):
                        eml_files.append(os.path.join(root, name))

            eml_files.sort()
            total = len(eml_files)

            if total == 0:
                raise RuntimeError("没有找到任何 .eml 文件。")

            self.log.emit(
                "发现 {} 封 EML，使用 {} 个并发工作线程开始扫描（CPU {} 核）。".format(
                    total, self.max_workers, self.cpu_count
                )
            )

            rows = []
            processed = 0
            attachment_count = 0
            error_count = 0

            worker_count = max(1, min(self.max_workers, total))
            executor = ThreadPoolExecutor(max_workers=worker_count)
            futures = set()
            self.log.emit(
                "线程结构：1 个 Qt 工作线程负责调度 + {} 个线程池任务线程解析 EML。".format(
                    worker_count
                )
            )
            next_index = 0

            def submit_more():
                nonlocal next_index
                while (
                    next_index < total
                    and len(futures) < self.max_workers
                    and not self._stop_event.is_set()
                ):
                    email_path = eml_files[next_index]
                    future = executor.submit(self._parse_one, email_path)
                    try:
                        future._email_path = email_path
                    except Exception:
                        pass
                    futures.add(future)
                    next_index += 1

            submit_more()

            try:
                while futures:
                    done, _pending = wait(
                        futures,
                        timeout=0.25,
                        return_when=FIRST_COMPLETED
                    )

                    if self._stop_event.is_set() and not done:
                        continue

                    if not done:
                        continue

                    for future in done:
                        futures.discard(future)

                        try:
                            eml_path, attachments, error = future.result()
                        except Exception as e:
                            eml_path = future_map_path = getattr(future, "_email_path", self.folder_path)
                            attachments = []
                            error = "{} | {}".format(eml_path, e)

                        processed += 1

                        if error:
                            error_count += 1
                            if error_count <= 20:
                                self.log.emit(
                                    "读取失败（已跳过）：{}".format(error)
                                )
                            elif error_count == 21:
                                self.log.emit(
                                    "读取失败邮件较多，后续错误不再逐条显示。"
                                )

                        mail_name = os.path.basename(eml_path)

                        for attachment in attachments:
                            rows.append([
                                attachment["filename"],
                                mail_name,
                                eml_path,
                                attachment["disposition"],
                                attachment["content_type"],
                            ])

                        attachment_count += len(attachments)

                        if (
                            processed == 1
                            or processed % 5 == 0
                            or processed == total
                        ):
                            self.progress.emit(
                                processed,
                                total,
                                attachment_count,
                                mail_name
                            )

                    if not self._stop_event.is_set():
                        submit_more()

            finally:
                # 取消尚未运行的任务；正在运行的任务允许自然结束。
                for future in futures:
                    future.cancel()
                executor.shutdown(wait=True)

            if self._stop_event.is_set():
                self.finished.emit({
                    "stopped": True,
                    "processed": processed,
                    "total": total,
                    "attachments": attachment_count,
                    "output": None,
                    "errors": error_count,
                })
                return

            rows.sort(key=lambda x: (x[2].lower(), x[0].lower()))
            output = save_attachment_csv(self.folder_path, rows)

            self.finished.emit({
                "stopped": False,
                "processed": processed,
                "total": total,
                "attachments": attachment_count,
                "output": output,
                "errors": error_count,
                "workers": self.max_workers,
            })

        except Exception:
            error = traceback.format_exc()
            write_error_log(error)
            self.failed.emit(error)


# ============================================================
# 主界面
# ============================================================

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.scan_thread = None
        self.scan_worker = None
        self.current_folder = get_default_mail_dir()
        self._analyze_after_scan = False
        self._matrix_position = {}
        self._closing_requested = False
        self._analysis_busy = False
        self._scan_locked_widgets = []
        self._analysis_locked_widgets = []

        self.setWindowTitle(APP_NAME)
        self.resize(1180, 820)
        self.setMinimumSize(980, 700)

        self._build_ui()
        self._auto_detect_initial_folder()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        title = QLabel("Email附件自动检测与分析")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        root.addWidget(title)

        # ----------------------------------------------------
        # ① 目录
        # ----------------------------------------------------
        folder_box = QGroupBox("① 邮件目录")
        folder_layout = QGridLayout(folder_box)

        self.folder_edit = QLineEdit()
        self.folder_edit.setText(self.current_folder)

        btn_browse = QPushButton("浏览目录")
        btn_browse.setObjectName("browseBtn")

        btn_auto = QPushButton("自动查找 EML")
        btn_open = QPushButton("打开目录")
        self.btn_help = QPushButton("帮助")
        self.btn_help.setObjectName("helpBtn")
        self.btn_help.setToolTip("打开使用手册、软件版本和系统信息")
        self._build_help_menu()

        btn_browse.clicked.connect(self.choose_folder)
        btn_auto.clicked.connect(self.auto_find_eml)
        btn_open.clicked.connect(self.open_current_folder)

        folder_layout.addWidget(QLabel("当前目录："), 0, 0)
        folder_layout.addWidget(self.folder_edit, 0, 1)
        folder_layout.addWidget(btn_browse, 0, 2)
        folder_layout.addWidget(btn_auto, 0, 3)
        folder_layout.addWidget(btn_open, 0, 4)
        folder_layout.addWidget(self.btn_help, 0, 5)
        self.btn_help.setMenu(self.help_menu)

        folder_layout.setColumnStretch(1, 1)
        folder_layout.setColumnStretch(0, 0)
        root.addWidget(folder_box)

        # ----------------------------------------------------
        # ② 参数
        # ----------------------------------------------------
        options_box = QGroupBox("② 编号分析规则")
        options = QGridLayout(options_box)

        self.start_spin = QSpinBox()
        self.end_spin = QSpinBox()
        self.step_spin = QSpinBox()

        for spin in (self.start_spin, self.end_spin, self.step_spin):
            spin.setRange(0, 2_000_000_000)
            spin.setMinimumHeight(30)

        self.start_spin.setValue(1)
        self.end_spin.setValue(100)
        self.step_spin.setRange(1, 2_000_000_000)
        self.step_spin.setValue(2)

        self.xlsx_check = QCheckBox("同时生成彩色 Excel（需要 openpyxl）")
        self.xlsx_check.setChecked(False)
        self.xlsx_check.setToolTip(
            "勾选后生成彩色 XLSX；需要安装 openpyxl。未勾选时不加载 openpyxl。"
        )

        # 紧凑单行参数区：避免原先 3 个输入框拉伸占满整行。
        self.thread_spin = QComboBox()
        thread_values = [1, 2, 4, 8, 16, 32, 64]
        for value in thread_values:
            self.thread_spin.addItem("{} 线程".format(value), value)
        self.system_info = get_runtime_capabilities()
        rec_threads, safe_cap, thread_reason = recommend_scan_threads(
            self.system_info["cpu"],
            self.system_info["memory_gb"],
            self.system_info["pointer_bits"]
        )
        self.thread_safe_cap = safe_cap
        self.system_info["recommended_threads"] = rec_threads
        self.system_info["safe_cap"] = safe_cap
        self.thread_recommend_label = QLabel()
        self.thread_recommend_label.setWordWrap(True)
        self.thread_recommend_label.setObjectName("threadHintLabel")
        self._set_thread_value(rec_threads)
        self.thread_spin.currentIndexChanged.connect(self.on_thread_choice_changed)
        self.thread_spin.setToolTip(
            "EML 扫描以磁盘 I/O 为主。线程越多不一定越快；老电脑优先 4~8，新电脑可尝试 16/32/64。"
        )

        options.setHorizontalSpacing(8)
        options.setVerticalSpacing(4)
        options.setContentsMargins(10, 8, 10, 8)
        labels = []
        for text, widget in (("开始", self.start_spin), ("结束", self.end_spin), ("步长", self.step_spin)):
            lab = QLabel(text)
            lab.setFixedWidth(34)
            labels.append(lab)
            widget.setFixedWidth(105)
            options.addWidget(lab, 0, len(labels) * 2 - 2)
            options.addWidget(widget, 0, len(labels) * 2 - 1)
        thread_label = QLabel("线程")
        thread_label.setFixedWidth(34)
        options.addWidget(thread_label, 0, 6)
        self.thread_spin.setFixedWidth(110)
        options.addWidget(self.thread_spin, 0, 7)
        self.xlsx_check.setMinimumWidth(190)
        options.addWidget(self.xlsx_check, 0, 8)
        options.addWidget(self.thread_recommend_label, 1, 0, 1, 9)

        root.addWidget(options_box)

        # ----------------------------------------------------
        # ③ 操作
        # ----------------------------------------------------
        action_box = QGroupBox("③ 操作")
        actions = QHBoxLayout(action_box)

        self.btn_auto_run = QPushButton("自动检测并分析")
        self.btn_auto_run.setObjectName("autoRunBtn")

        self.btn_scan = QPushButton("重新扫描邮件")
        self.btn_scan.setObjectName("scanBtn")

        self.btn_analyze = QPushButton("只分析已有附件列表")
        self.btn_analyze.setObjectName("analyzeBtn")

        self.btn_stop = QPushButton("停止扫描")
        self.btn_stop.setObjectName("stopBtn")
        self.btn_stop.setEnabled(False)

        self.btn_auto_run.clicked.connect(self.auto_run)
        self.btn_scan.clicked.connect(self.start_scan)
        self.btn_analyze.clicked.connect(self.analyze_existing)
        self.btn_stop.clicked.connect(self.stop_scan)

        actions.addWidget(self.btn_auto_run)
        actions.addWidget(self.btn_scan)
        actions.addWidget(self.btn_analyze)
        actions.addWidget(self.btn_stop)

        root.addWidget(action_box)

        # ----------------------------------------------------
        # ④ 进度
        # ----------------------------------------------------
        progress_box = QGroupBox("④ 扫描进度")
        progress_layout = QGridLayout(progress_box)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)

        self.status_label = QLabel("等待操作")
        self.count_label = QLabel("已处理：0 / 0    附件：0")

        progress_layout.addWidget(self.progress, 0, 0, 1, 2)
        progress_layout.addWidget(self.count_label, 1, 0)
        progress_layout.addWidget(self.status_label, 1, 1)

        root.addWidget(progress_box)

        # ----------------------------------------------------
        # ⑤ 结果预览 + 日志
        # ----------------------------------------------------
        splitter = QSplitter(Qt.Horizontal)

        result_box = QGroupBox(
            "⑤ 数据分析结果预览  【推荐在这里查看缺失 / 异常 / 重复】"
        )
        result_layout = QVBoxLayout(result_box)
        result_layout.setContentsMargins(8, 10, 8, 8)
        result_layout.setSpacing(7)

        # 统计卡片
        summary_frame = QFrame()
        summary_layout = QGridLayout(summary_frame)
        summary_layout.setContentsMargins(0, 0, 0, 0)
        summary_layout.setHorizontalSpacing(6)

        self.card_valid = self._create_summary_card("有效编号", "0")
        self.card_expected = self._create_summary_card("理论编号", "0")
        self.card_missing = self._create_summary_card("缺失", "0")
        self.card_abnormal = self._create_summary_card("异常", "0")
        self.card_duplicate = self._create_summary_card("重复", "0")

        for col, card in enumerate((
            self.card_valid,
            self.card_expected,
            self.card_missing,
            self.card_abnormal,
            self.card_duplicate
        )):
            summary_layout.addWidget(card, 0, col)
            summary_layout.setColumnStretch(col, 1)

        result_layout.addWidget(summary_frame)

        # 搜索
        search_layout = QHBoxLayout()

        self.preview_search = QLineEdit()
        self.preview_search.setPlaceholderText(
            "输入编号定位，例如：1669；支持定位正常、缺失、重复编号"
        )

        self.btn_preview_search = QPushButton("定位编号")
        self.btn_preview_search.setObjectName("findBtn")

        self.btn_preview_clear = QPushButton("清除定位")

        self.preview_search.returnPressed.connect(self.find_preview_number)
        self.btn_preview_search.clicked.connect(self.find_preview_number)
        self.btn_preview_clear.clicked.connect(self.clear_preview_search)

        search_layout.addWidget(self.preview_search, 1)
        search_layout.addWidget(self.btn_preview_search)
        search_layout.addWidget(self.btn_preview_clear)

        result_layout.addLayout(search_layout)

        # 小型图例
        legend_layout = QHBoxLayout()
        legend_layout.addWidget(QLabel("图例："))

        legend_layout.addWidget(self._create_legend_label("正常", QColor(146, 208, 80)))
        legend_layout.addWidget(self._create_legend_label("缺失", QColor(255, 0, 0)))
        legend_layout.addWidget(self._create_legend_label("异常", QColor(255, 217, 102)))
        legend_layout.addWidget(self._create_legend_label("重复", QColor(91, 155, 213)))
        legend_layout.addStretch(1)

        legend_layout.addWidget(QLabel("网格列数："))

        self.matrix_columns_spin = QSpinBox()
        self.matrix_columns_spin.setRange(UI_MIN_COLUMNS, UI_MAX_COLUMNS)
        self.matrix_columns_spin.setValue(UI_DEFAULT_COLUMNS)
        self.matrix_columns_spin.setToolTip(
            "列数直接控制编号面板布局，支持 6~36 列。修改后立即重排；窗口不足时允许横向滚动。"
        )
        self.matrix_columns_spin.valueChanged.connect(self.refresh_matrix_preview)

        legend_layout.addWidget(self.matrix_columns_spin)

        result_layout.addLayout(legend_layout)

        # Tab
        self.result_tabs = QTabWidget()

        # 总览
        matrix_page = QWidget()
        matrix_layout = QVBoxLayout(matrix_page)
        matrix_layout.setContentsMargins(2, 2, 2, 2)

        self.matrix_table = QTableWidget(0, 0)
        self.matrix_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.matrix_table.setAlternatingRowColors(False)
        self.matrix_table.verticalHeader().setVisible(False)
        self.matrix_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.matrix_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.matrix_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.matrix_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.matrix_table.horizontalHeader().setVisible(False)
        self.matrix_table.setShowGrid(True)
        self.matrix_table.setWordWrap(False)
        self.matrix_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        matrix_layout.addWidget(self.matrix_table)
        self.result_tabs.addTab(matrix_page, "编号总览")

        # 明细
        self.missing_table = self._create_number_table(["缺失编号"])
        self.abnormal_table = self._create_number_table(["异常编号"])
        self.duplicate_table = self._create_number_table(["重复编号", "出现次数"])

        self.result_tabs.addTab(self.missing_table, "缺失 0")
        self.result_tabs.addTab(self.abnormal_table, "异常 0")
        self.result_tabs.addTab(self.duplicate_table, "重复 0")

        self.missing_table.cellDoubleClicked.connect(
            lambda row, _col: self._locate_from_detail(
                self.missing_table, row, 1
            )
        )
        self.abnormal_table.cellDoubleClicked.connect(
            lambda row, _col: self._locate_from_detail(
                self.abnormal_table, row, 2
            )
        )
        self.duplicate_table.cellDoubleClicked.connect(
            lambda row, _col: self._locate_from_detail(
                self.duplicate_table, row, 3
            )
        )

        result_layout.addWidget(self.result_tabs, 1)

        # 扫描/分析期间需要保护的控件。初始化后不再依赖 getattr 空列表，确保锁定逻辑真的生效。
        self._scan_locked_widgets = [
            self.folder_edit, btn_browse, btn_auto, btn_open, self.btn_help,
            self.start_spin, self.end_spin, self.step_spin,
            self.thread_spin, self.xlsx_check,
            self.btn_auto_run, self.btn_scan, self.btn_analyze,
            self.preview_search, self.btn_preview_search, self.btn_preview_clear,
            self.matrix_columns_spin, self.result_tabs
        ]
        self._analysis_locked_widgets = [
            self.folder_edit, btn_browse, btn_auto, btn_open, self.btn_help,
            self.start_spin, self.end_spin, self.step_spin,
            self.thread_spin, self.xlsx_check,
            self.btn_auto_run, self.btn_scan, self.btn_analyze,
            self.preview_search, self.btn_preview_search, self.btn_preview_clear,
            self.matrix_columns_spin, self.result_tabs
        ]

        # 运行日志
        log_box = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_box)
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)

        splitter.addWidget(result_box)
        splitter.addWidget(log_box)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([900, 250])

        root.addWidget(splitter, 1)

        # ----------------------------------------------------
        # 样式
        # ----------------------------------------------------
        self.setStyleSheet("""
            QMainWindow {
                background: #eef2f7;
            }

            QGroupBox {
                font-weight: bold;
                color: #24364b;
                border: 1px solid #d7dee8;
                border-radius: 9px;
                margin-top: 10px;
                padding: 11px 9px 9px 9px;
                background: #ffffff;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 5px;
                color: #2f5f93;
            }

            QLabel#threadHintLabel {
                color: #6b7785;
                font-size: 12px;
                padding: 2px 0 0 38px;
            }

            QLabel#titleLabel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #245b9b, stop:1 #3c82c8);
                color: white;
                font-size: 22px;
                font-weight: bold;
                padding: 14px;
                border-radius: 10px;
            }

            QPushButton {
                min-height: 34px;
                padding: 0 14px;
                border: 1px solid #c6d0dc;
                border-radius: 7px;
                background: #ffffff;
                color: #24364b;
                font-weight: 600;
            }

            QPushButton:hover {
                background: #edf4fb;
                border-color: #7aa6d6;
            }

            QPushButton:pressed {
                background: #dbeaf8;
            }

            QPushButton:disabled {
                color: #999999;
                background: #eceff3;
                border-color: #d8dde3;
            }

            QPushButton#autoRunBtn {
                background: #2f80ed;
                color: white;
                border-color: #2f80ed;
            }

            QPushButton#scanBtn {
                background: #27ae60;
                color: white;
                border-color: #27ae60;
            }

            QPushButton#analyzeBtn {
                background: #8e5ad8;
                color: white;
                border-color: #8e5ad8;
            }

            QPushButton#stopBtn {
                background: #e74c3c;
                color: white;
                border-color: #e74c3c;
            }

            QPushButton#browseBtn {
                background: #607d9b;
                color: white;
                border-color: #607d9b;
            }

            QPushButton#findBtn {
                background: #f39c12;
                color: white;
                border-color: #f39c12;
            }

            QPushButton#helpBtn {
                background: #556b7f;
                color: white;
                border-color: #556b7f;
            }

            QLineEdit, QSpinBox {
                min-height: 29px;
                border: 1px solid #c5cfda;
                border-radius: 6px;
                padding: 2px 8px;
                background: #ffffff;
            }

            QProgressBar {
                min-height: 24px;
                border: 1px solid #b8c5d3;
                border-radius: 7px;
                text-align: center;
                background: #ffffff;
            }

            QProgressBar::chunk {
                background: #27ae60;
                border-radius: 6px;
            }

            QPlainTextEdit, QTableWidget {
                border: 1px solid #d7dee8;
                border-radius: 6px;
                background: #ffffff;
                gridline-color: #e4e9ef;
            }

            QTableWidget::item:selected {
                background: #cfe5ff;
                color: #17202a;
            }

            QTabBar::tab {
                background: #e9eef5;
                border: 1px solid #d3dce7;
                border-bottom: none;
                padding: 7px 13px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }

            QTabBar::tab:selected {
                background: #ffffff;
                color: #245b9b;
                font-weight: bold;
            }
        """)

        # 初始清空
        self._reset_results()

    # --------------------------------------------------------
    # UI 工具
    # --------------------------------------------------------

    def _create_summary_card(self, title: str, value: str):
        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(8, 5, 8, 5)
        layout.setSpacing(1)

        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignCenter)

        value_label = QLabel(value)
        value_label.setAlignment(Qt.AlignCenter)
        value_label.setObjectName("summaryValue")
        value_label.setStyleSheet(
            "font-size: 18px; font-weight: bold; color: #245b9b;"
        )

        layout.addWidget(title_label)
        layout.addWidget(value_label)

        frame._value_label = value_label
        return frame

    def _set_summary_card(self, card, value: int):
        card._value_label.setText(format_int(value))

    def _create_legend_label(self, text: str, color: QColor):
        label = QLabel("  {}  ".format(text))
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet(
            "background-color: rgb({},{},{}); "
            "border: 1px solid #aeb8c3; border-radius: 4px; padding: 2px 7px;"
            .format(color.red(), color.green(), color.blue())
        )
        if color.red() > 200 and color.green() < 120:
            label.setStyleSheet(
                label.styleSheet() + "color:white; font-weight:bold;"
            )
        return label

    def _create_number_table(self, headers):
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setAlternatingRowColors(True)
        table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        table.verticalHeader().setVisible(False)

        if len(headers) == 1:
            table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.Stretch
            )
        else:
            table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.Stretch
            )
            table.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.ResizeToContents
            )

        return table

    def _fill_single_number_table(
        self,
        table: QTableWidget,
        numbers,
        color: QColor
    ):
        table.setUpdatesEnabled(False)
        try:
            table.clearContents()
            table.setRowCount(len(numbers))

            for row, number in enumerate(numbers):
                item = QTableWidgetItem(str(number))
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(color)
                if color.red() > 200 and color.green() < 120:
                    item.setForeground(QColor(255, 255, 255))
                table.setItem(row, 0, item)
        finally:
            table.setUpdatesEnabled(True)

    def _fill_duplicate_table(
        self,
        duplicates: Dict[int, int],
        color: QColor
    ):
        items = sorted(duplicates.items())

        self.duplicate_table.setUpdatesEnabled(False)
        try:
            self.duplicate_table.clearContents()
            self.duplicate_table.setRowCount(len(items))

            for row, (number, count) in enumerate(items):
                item1 = QTableWidgetItem(str(number))
                item2 = QTableWidgetItem(str(count))

                item1.setTextAlignment(Qt.AlignCenter)
                item2.setTextAlignment(Qt.AlignCenter)

                item1.setBackground(color)
                item2.setBackground(color)

                self.duplicate_table.setItem(row, 0, item1)
                self.duplicate_table.setItem(row, 1, item2)
        finally:
            self.duplicate_table.setUpdatesEnabled(True)

    def _reset_results(self):
        self._set_summary_card(self.card_valid, 0)
        self._set_summary_card(self.card_expected, 0)
        self._set_summary_card(self.card_missing, 0)
        self._set_summary_card(self.card_abnormal, 0)
        self._set_summary_card(self.card_duplicate, 0)

        self.matrix_table.clearContents()
        self.matrix_table.setRowCount(0)
        self.matrix_table.setColumnCount(0)

        self.missing_table.setRowCount(0)
        self.abnormal_table.setRowCount(0)
        self.duplicate_table.setRowCount(0)

        self.result_tabs.setTabText(0, "编号总览")
        self.result_tabs.setTabText(1, "缺失 0")
        self.result_tabs.setTabText(2, "异常 0")
        self.result_tabs.setTabText(3, "重复 0")

        self._matrix_position.clear()

    # --------------------------------------------------------
    # 预览
    # --------------------------------------------------------

    def _calculate_matrix_columns(self) -> int:
        # 以前这里把用户设置再与窗口宽度取 min，导致“改了列数但没有效果”。
        # 现在严格使用用户设置；窗口不足时交给 QTableWidget 横向滚动。
        return max(UI_MIN_COLUMNS, min(UI_MAX_COLUMNS, int(self.matrix_columns_spin.value())))

    def refresh_matrix_preview(self):
        if not getattr(self, "_preview_data", None):
            return

        expected, numbers, missing, duplicates = self._preview_data
        self._render_matrix(expected, numbers, missing, duplicates)

    def _render_matrix(
        self,
        expected,
        actual_numbers,
        missing_numbers,
        duplicate_numbers
    ):
        columns = self._calculate_matrix_columns()
        rows = (len(expected) + columns - 1) // columns if expected else 1

        missing_set = set(missing_numbers)
        actual_counter = Counter(actual_numbers)

        green = QColor(146, 208, 80)
        red = QColor(255, 0, 0)
        blue = QColor(91, 155, 213)

        self._matrix_position = {}
        self.matrix_table.setUpdatesEnabled(False)

        try:
            self.matrix_table.clearContents()
            self.matrix_table.setColumnCount(columns)
            self.matrix_table.setRowCount(rows)

            for col in range(columns):
                self.matrix_table.setColumnWidth(col, UI_CELL_WIDTH)

            for row in range(rows):
                self.matrix_table.setRowHeight(row, 29)

            # 与截图/旧版 CSV 一致：先向下填满一列，再进入下一列。
            rows_per_column = rows
            for index, number in enumerate(expected):
                col = index // rows_per_column
                row = index % rows_per_column

                if number in missing_set:
                    text = "缺失:{}".format(number)
                    bg = red
                    fg = QColor(255, 255, 255)
                elif actual_counter[number] > 1:
                    text = "重复:{}({}次)".format(number, actual_counter[number])
                    bg = blue
                    fg = QColor(255, 255, 255)
                else:
                    text = str(number)
                    bg = green
                    fg = QColor(20, 70, 20)

                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(bg)
                item.setForeground(fg)
                self.matrix_table.setItem(row, col, item)

                self._matrix_position[number] = (row, col)

        finally:
            self.matrix_table.setUpdatesEnabled(True)

    def show_analysis_preview(
        self,
        numbers,
        expected,
        missing,
        abnormal,
        duplicates
    ):
        self._preview_data = (expected, numbers, missing, duplicates)

        self._set_summary_card(self.card_valid, len(numbers))
        self._set_summary_card(self.card_expected, len(expected))
        self._set_summary_card(self.card_missing, len(missing))
        self._set_summary_card(self.card_abnormal, len(abnormal))
        self._set_summary_card(self.card_duplicate, len(duplicates))

        self._render_matrix(expected, numbers, missing, duplicates)

        self._fill_single_number_table(
            self.missing_table,
            missing,
            QColor(255, 0, 0)
        )
        self._fill_single_number_table(
            self.abnormal_table,
            abnormal,
            QColor(255, 217, 102)
        )
        self._fill_duplicate_table(
            duplicates,
            QColor(91, 155, 213)
        )

        self.result_tabs.setTabText(
            1, "缺失 {}".format(len(missing))
        )
        self.result_tabs.setTabText(
            2, "异常 {}".format(len(abnormal))
        )
        self.result_tabs.setTabText(
            3, "重复 {}".format(len(duplicates))
        )
        self.result_tabs.setCurrentIndex(0)

    def _locate_from_detail(self, table, row: int, tab_index: int):
        item = table.item(row, 0)
        if not item:
            return

        value = item.text().strip()
        self.preview_search.setText(value)

        try:
            number = int(value)
        except Exception:
            return

        self.result_tabs.setCurrentIndex(0)

        position = self._matrix_position.get(number)
        if position is None:
            self.status_label.setText(
                "编号 {} 不在理论编号范围，只存在于异常明细。".format(number)
            )
            return

        r, c = position
        matrix_item = self.matrix_table.item(r, c)
        if matrix_item:
            self.matrix_table.setCurrentCell(r, c)
            self.matrix_table.scrollToItem(
                matrix_item,
                QAbstractItemView.PositionAtCenter
            )
            self.status_label.setText("已定位编号：{}".format(number))

    def find_preview_number(self):
        keyword = self.preview_search.text().strip()
        if not keyword:
            return

        # 优先按整数精确定位。
        number = None
        try:
            number = int(keyword)
        except Exception:
            pass

        if number is not None:
            position = self._matrix_position.get(number)

            if position is not None:
                row, col = position
                item = self.matrix_table.item(row, col)
                if item:
                    self.result_tabs.setCurrentIndex(0)
                    self.matrix_table.setCurrentCell(row, col)
                    self.matrix_table.scrollToItem(
                        item,
                        QAbstractItemView.PositionAtCenter
                    )
                    self.status_label.setText("已定位编号：{}".format(number))
                    return

            # 理论范围之外，尝试在异常表定位。
            for table_index, table in (
                (2, self.abnormal_table),
            ):
                for row in range(table.rowCount()):
                    item = table.item(row, 0)
                    if item and item.text().strip() == keyword:
                        self.result_tabs.setCurrentIndex(table_index)
                        table.selectRow(row)
                        table.scrollToItem(
                            item,
                            QAbstractItemView.PositionAtCenter
                        )
                        self.status_label.setText("已定位异常编号：{}".format(keyword))
                        return

        # 非纯数字时，按文本查找。
        for row in range(self.matrix_table.rowCount()):
            for col in range(self.matrix_table.columnCount()):
                item = self.matrix_table.item(row, col)
                if item and keyword in item.text().strip():
                    self.result_tabs.setCurrentIndex(0)
                    self.matrix_table.setCurrentCell(row, col)
                    self.matrix_table.scrollToItem(
                        item,
                        QAbstractItemView.PositionAtCenter
                    )
                    self.status_label.setText("已定位：{}".format(keyword))
                    return

        for tab_index, table in (
            (1, self.missing_table),
            (2, self.abnormal_table),
            (3, self.duplicate_table)
        ):
            for row in range(table.rowCount()):
                item = table.item(row, 0)
                if item and keyword in item.text().strip():
                    self.result_tabs.setCurrentIndex(tab_index)
                    table.selectRow(row)
                    table.scrollToItem(
                        item,
                        QAbstractItemView.PositionAtCenter
                    )
                    self.status_label.setText("已定位：{}".format(keyword))
                    return

        QMessageBox.information(
            self,
            "未找到",
            "预览结果中没有找到编号：{}".format(keyword)
        )

    def clear_preview_search(self):
        self.preview_search.clear()
        self.matrix_table.clearSelection()
        self.missing_table.clearSelection()
        self.abnormal_table.clearSelection()
        self.duplicate_table.clearSelection()

    def _set_thread_value(self, value: int):
        value = int(value)
        idx = self.thread_spin.findData(value)
        if idx >= 0:
            self.thread_spin.setCurrentIndex(idx)

    def on_thread_choice_changed(self, _index: int):
        value = int(self.thread_spin.currentData())
        cpu = self.system_info["cpu"]
        memory_gb = self.system_info["memory_gb"]
        pointer_bits = self.system_info["pointer_bits"]
        if value > self.thread_safe_cap:
            self.thread_recommend_label.setText(
                "[提示] 当前电脑建议不超过 {} 线程；选择 {} 线程可能降低速度或增加系统压力。".format(
                    self.thread_safe_cap, value
                )
            )
        else:
            self.thread_recommend_label.setText(
                "系统：{} 核 / {} 位 / {}GB；推荐 {} 线程。{}".format(
                    cpu, pointer_bits, memory_gb or "未知",
                    recommend_scan_threads(cpu, memory_gb, pointer_bits)[0],
                    recommend_scan_threads(cpu, memory_gb, pointer_bits)[2]
                )
            )

    def selected_scan_threads(self) -> int:
        requested = int(self.thread_spin.currentData())
        if requested > self.thread_safe_cap:
            self.append_log(
                "线程数 {} 超过当前机器建议上限 {}，已自动降为 {}，避免老旧电脑卡顿。".format(
                    requested, self.thread_safe_cap, self.thread_safe_cap
                )
            )
            return self.thread_safe_cap
        return requested

    # --------------------------------------------------------
    # UI / 目录
    # --------------------------------------------------------

    def append_log(self, text: str):
        self.log_edit.appendPlainText(
            "[{}] {}".format(time.strftime("%H:%M:%S"), text)
        )
        self.log_edit.moveCursor(QTextCursor.End)

    def current_folder_path(self) -> str:
        return os.path.abspath(self.folder_edit.text().strip())

    def _apply_scan_lock(self, locked: bool):
        """扫描期间锁定所有可能修改状态的控件，只保留停止按钮。"""
        for widget in getattr(self, "_scan_locked_widgets", []):
            widget.setEnabled(not locked)
        self.btn_stop.setEnabled(locked)
        self.btn_stop.setFocusPolicy(Qt.StrongFocus if locked else Qt.NoFocus)

    def set_busy(self, busy: bool):
        self._apply_scan_lock(busy)

    def set_analysis_busy(self, busy: bool):
        # 分析期间同样锁定核心操作，但不显示全局忙碌光标。
        self._analysis_busy = bool(busy)
        for widget in self._analysis_locked_widgets:
            widget.setEnabled(not busy)
        self.btn_stop.setEnabled(False)

    def _auto_detect_initial_folder(self):
        # 启动时只设置默认目录，不递归搜索 EML，避免老旧电脑启动时扫描磁盘。
        default_dir = get_default_mail_dir()
        self.current_folder = default_dir
        self.folder_edit.setText(default_dir)
        self.append_log("默认邮件目录：{}".format(default_dir))

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择包含 EML 邮件的目录",
            self.current_folder
        )

        if folder:
            self.current_folder = folder
            self.folder_edit.setText(folder)
            self.append_log("已选择目录：{}".format(folder))

    def auto_find_eml(self):
        base = get_default_mail_dir()
        folders = find_eml_folders(base)

        if not folders:
            QMessageBox.information(
                self,
                "未找到 EML",
                "程序目录及其子目录中没有找到 .eml 文件。"
            )
            return

        if len(folders) == 1:
            self.current_folder = folders[0]
            self.folder_edit.setText(folders[0])
            self.append_log(
                "找到 EML 目录：{}".format(folders[0])
            )
            return

        from PyQt5.QtWidgets import QInputDialog

        item, ok = QInputDialog.getItem(
            self,
            "选择邮件目录",
            "发现多个包含 EML 的目录：",
            folders,
            0,
            False
        )

        if ok and item:
            self.current_folder = item
            self.folder_edit.setText(item)
            self.append_log(
                "已选择 EML 目录：{}".format(item)
            )

    def _build_help_menu(self):
        self.help_menu = QMenu(self)
        self.help_menu.addAction(
            QAction("使用手册", self, triggered=self.open_help_manual)
        )
        self.help_menu.addAction(
            QAction("软件版本", self, triggered=self.show_version_info)
        )
        self.help_menu.addAction(
            QAction("系统信息", self, triggered=self.show_system_info)
        )

    def show_version_info(self):
        QMessageBox.information(
            self,
            "软件版本",
            "{}\n版本：{}\n\n"
            "功能：扫描 EML 邮件附件并分析附件编号。".format(
                APP_NAME, APP_VERSION
            )
        )

    def show_system_info(self):
        info = self.system_info
        memory = "{} GB".format(info["memory_gb"]) if info["memory_gb"] else "未知"
        message = (
            "操作系统：{system}\n"
            "Python：{python}\n"
            "运行环境：{bits}\n"
            "CPU 逻辑核心：{cpu}\n"
            "物理内存：{memory}\n"
            "推荐扫描线程：{rec}\n"
            "线程安全上限：{cap}\n"
            "程序目录：{app}\n"
            "默认邮件目录：{mail}"
        ).format(
            system=sys.platform,
            python=sys.version.split()[0],
            bits="{} 位".format(info["pointer_bits"]),
            cpu=info["cpu"],
            memory=memory,
            rec=info.get("recommended_threads", "-"),
            cap=info.get("safe_cap", "-"),
            app=get_app_dir(),
            mail=get_default_mail_dir(),
        )
        QMessageBox.information(self, "系统信息", message)

    def open_help_manual(self):
        """调用系统默认 Word/文档关联程序打开使用手册。程序不解析、不读取手册内容。"""
        manual_path = os.path.join(get_app_dir(), HELP_DOCX)
        # 仅做文件存在检查，不对 DOCX 内容进行任何处理。
        if not os.path.isfile(manual_path):
            manual_path = None

        if manual_path is None:
            QMessageBox.information(
                self,
                "找不到使用手册",
                "未找到使用手册：\n{}\n\n请将该 DOCX 文件放在程序 EXE 同目录。".format(HELP_DOCX)
            )
            return

        try:
            opened = QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(manual_path)))
            if not opened:
                raise RuntimeError("系统未能调用默认的 Word/文档程序打开文件。")
            self.append_log("已调用系统默认程序打开使用手册：{}".format(manual_path))
        except Exception as e:
            QMessageBox.warning(
                self,
                "打开使用手册失败",
                "请确认电脑已安装可打开 DOCX 的程序。\n\n{}".format(e)
            )

    def open_current_folder(self):
        folder = self.current_folder_path()

        if not os.path.isdir(folder):
            QMessageBox.warning(
                self,
                "目录错误",
                "当前目录不存在。"
            )
            return

        try:
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
        except Exception as e:
            QMessageBox.warning(
                self,
                "打开失败",
                str(e)
            )

    # --------------------------------------------------------
    # 扫描
    # --------------------------------------------------------

    def validate_folder(self) -> Optional[str]:
        folder = self.current_folder_path()

        if not os.path.isdir(folder):
            QMessageBox.warning(
                self,
                "目录错误",
                "请选择有效的邮件目录。"
            )
            return None

        self.current_folder = folder
        return folder

    def start_scan(self):
        if self._analysis_busy:
            self.status_label.setText("当前正在分析，请稍候。")
            return
        if self.scan_thread is not None:
            self.status_label.setText("扫描任务已经在运行，请先停止当前任务。")
            self.append_log("忽略重复扫描请求：当前任务仍在运行。")
            return

        folder = self.validate_folder()
        if not folder:
            return

        self.progress.setValue(0)
        self.count_label.setText("已处理：0 / 0    附件：0")
        self.status_label.setText("正在准备扫描")
        self.append_log("开始扫描：{}".format(folder))
        self._reset_results()

        actual_threads = self.selected_scan_threads()
        requested_threads = int(self.thread_spin.currentData())
        if actual_threads != requested_threads:
            self.status_label.setText("线程数已自动调整为 {}，开始扫描".format(actual_threads))
        else:
            self.status_label.setText("正在准备扫描（{} 线程）".format(actual_threads))

        self.set_busy(True)

        self.scan_thread = QThread(self)
        self.scan_worker = ScanWorker(folder, actual_threads)
        self.scan_worker.moveToThread(self.scan_thread)

        self.scan_thread.started.connect(self.scan_worker.run)
        self.scan_worker.progress.connect(self.on_scan_progress)
        self.scan_worker.log.connect(self.append_log)
        self.scan_worker.finished.connect(self.on_scan_finished)
        self.scan_worker.failed.connect(self.on_scan_failed)

        self.scan_worker.finished.connect(self.scan_thread.quit)
        self.scan_worker.failed.connect(self.scan_thread.quit)
        # 在工作线程事件循环结束前安排 worker 删除，避免无父 QObject 残留。
        self.scan_worker.finished.connect(self.scan_worker.deleteLater)
        self.scan_worker.failed.connect(self.scan_worker.deleteLater)
        self.scan_thread.finished.connect(self.cleanup_scan_thread)

        self.scan_thread.start()

    def stop_scan(self):
        if self.scan_worker:
            self.scan_worker.request_stop()
            self.status_label.setText(
                "已请求停止，正在完成当前邮件……"
            )
            self.btn_stop.setEnabled(False)
            self.append_log("用户请求停止扫描。")

    def on_scan_progress(
        self,
        processed,
        total,
        attachments,
        current_name
    ):
        percent = int(processed * 100 / total) if total else 0
        self.progress.setValue(percent)
        self.count_label.setText(
            "已处理：{} / {}    已提取附件：{}".format(
                format_int(processed),
                format_int(total),
                format_int(attachments)
            )
        )
        self.status_label.setText("当前：{}".format(current_name))

    def on_scan_finished(self, result: dict):
        # 自动扫描后紧接分析时，先保持控件锁定，避免扫描结束瞬间用户抢先点击其他操作。
        if result.get("stopped"):
            self.set_busy(False)
            self.status_label.setText("扫描已停止")
            self.append_log(
                "扫描已停止：已处理 {}/{}，附件 {}。".format(
                    result.get("processed", 0),
                    result.get("total", 0),
                    result.get("attachments", 0)
                )
            )

            self.append_log("已保留原有附件名称列表，没有覆盖输出文件。")
            return

        self.progress.setValue(100)
        self.status_label.setText("扫描完成")

        self.append_log(
            "扫描完成：邮件 {} 封，附件 {} 个，读取失败 {} 封。".format(
                result["processed"],
                result["attachments"],
                result.get("errors", 0)
            )
        )
        self.append_log(
            "附件列表已保存：{}".format(result["output"])
        )

        if getattr(self, "_analyze_after_scan", False):
            # 扫描完成后直接进入分析，整个自动流程期间持续锁定控件。
            self._analyze_after_scan = False
            self.analyze_existing()
        else:
            self.set_busy(False)
            self.append_log(
                "扫描完成提示：处理邮件 {} 封，提取附件 {} 个。结果文件已写入。".format(
                    format_int(result["processed"]),
                    format_int(result["attachments"])
                )
            )

    def on_scan_failed(self, error: str):
        self.set_busy(False)
        self.status_label.setText("扫描失败")
        self.append_log(
            "扫描发生异常，详细信息已写入 {}".format(ERROR_LOG)
        )
        write_error_log(error)

        QMessageBox.critical(
            self,
            "扫描异常",
            "扫描过程中发生异常，但程序不会闪退。\n"
            "详细错误已写入：{}".format(
                os.path.join(get_app_dir(), ERROR_LOG)
            )
        )

    def cleanup_scan_thread(self):
        worker = self.scan_worker
        thread = self.scan_thread
        self.scan_worker = None
        self.scan_thread = None

        # worker 已经通过 finished/failed 信号安排 deleteLater；这里仅释放 Python 引用。
        if thread:
            thread.deleteLater()

        if self._closing_requested:
            self._closing_requested = False
            self.close()
        elif not self._analysis_busy:
            # 扫描线程真正退出后再恢复控件。
            self.set_busy(False)

    # --------------------------------------------------------
    # 自动流程
    # --------------------------------------------------------

    def auto_run(self):
        folder = self.validate_folder()
        if not folder:
            return

        csv_path = os.path.join(folder, ATTACHMENT_CSV)

        if os.path.exists(csv_path):
            answer = QMessageBox.question(
                self,
                "已有附件列表",
                "检测到已有“{}”。\n\n"
                "是：直接使用已有列表分析\n"
                "否：重新扫描邮件后分析\n"
                "取消：不执行".format(ATTACHMENT_CSV),
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.Yes
            )

            if answer == QMessageBox.Yes:
                self.analyze_existing()
                return

            if answer == QMessageBox.No:
                self._analyze_after_scan = True
                self.start_scan()
                return

            return

        self._analyze_after_scan = True
        self.start_scan()

    # --------------------------------------------------------
    # 分析
    # --------------------------------------------------------

    def analyze_existing(self):
        folder = self.validate_folder()
        if not folder:
            return

        start = self.start_spin.value()
        end = self.end_spin.value()
        step = self.step_spin.value()

        if end < start:
            QMessageBox.warning(
                self,
                "参数错误",
                "结束编号不能小于开始编号。"
            )
            return

        if step <= 0:
            QMessageBox.warning(
                self,
                "参数错误",
                "步长必须大于 0。"
            )
            return

        expected_count = ((end - start) // step) + 1
        if expected_count > MAX_EXPECTED_NUMBERS:
            QMessageBox.warning(
                self,
                "编号范围过大",
                "理论编号数量为 {}，超过安全上限 {}。\n\n"
                "为避免老旧电脑内存不足或长时间无响应，请缩小编号范围或增大步长。".format(
                    format_int(expected_count), format_int(MAX_EXPECTED_NUMBERS)
                )
            )
            return

        input_csv = os.path.join(folder, ATTACHMENT_CSV)

        if not os.path.exists(input_csv):
            QMessageBox.warning(
                self,
                "找不到文件",
                "没有找到：{}\n请先扫描邮件生成附件列表。".format(
                    input_csv
                )
            )
            return

        # 不再设置 QApplication.OverrideCursor。老机器分析期间如果再弹 QMessageBox，
        # 很容易让用户看到“鼠标一直沙漏”的状态。
        self.set_analysis_busy(True)
        try:
            self.status_label.setText("正在分析编号……")
            QApplication.processEvents()

            numbers, invalid_rows, total_rows = read_numbers_from_csv(
                input_csv
            )

            if not numbers:
                self._reset_results()

                self.set_analysis_busy(False)
                QMessageBox.warning(
                    self,
                    "没有有效编号",
                    "附件列表中没有读取到任何可识别的编号。"
                )
                self.status_label.setText(
                    "分析结束：无有效编号"
                )
                return

            expected, missing, abnormal, duplicates = analyze_numbers(
                numbers,
                start,
                end,
                step
            )

            result_path = save_matrix_csv(
                folder,
                start,
                end,
                step,
                expected,
                numbers,
                missing,
                abnormal,
                duplicates
            )

            invalid_path = save_invalid_rows(
                folder,
                invalid_rows
            )

            xlsx_path = None
            if self.xlsx_check.isChecked():
                xlsx_path = try_save_styled_xlsx(
                    folder,
                    start,
                    end,
                    step,
                    expected,
                    numbers,
                    missing,
                    abnormal,
                    duplicates
                )

                if xlsx_path is None:
                    self.append_log(
                        "彩色 Excel 生成失败，已保留 CSV 结果。"
                    )

            self.show_analysis_preview(
                numbers,
                expected,
                missing,
                abnormal,
                duplicates
            )

            self.status_label.setText("编号分析完成")
            self.append_log("编号分析完成。")
            self.append_log("CSV 结果：{}".format(result_path))

            if invalid_path:
                self.append_log(
                    "无法识别编号：{}".format(invalid_path)
                )

            if xlsx_path:
                self.append_log(
                    "彩色 Excel：{}".format(xlsx_path)
                )

            msg = (
                "分析完成。\n\n"
                "有效附件编号：{}\n"
                "理论编号：{}\n"
                "缺失编号：{}\n"
                "异常编号：{}\n"
                "重复编号：{}\n"
                "无法识别：{}\n\n"
                "结果文件：{}"
            ).format(
                format_int(len(numbers)),
                format_int(len(expected)),
                format_int(len(missing)),
                format_int(len(abnormal)),
                format_int(len(duplicates)),
                format_int(len(invalid_rows)),
                result_path
            )

            if invalid_path:
                msg += "\n无法识别文件：{}".format(invalid_path)

            if xlsx_path:
                msg += "\n彩色 Excel：{}".format(xlsx_path)

            # 分析完成默认不再弹模态窗口，避免抢占鼠标焦点；结果已经显示在下方预览区。
            self.append_log("分析完成摘要：" + msg.replace("\n", " "))
            self.status_label.setText("编号分析完成")

        except Exception:
            error = traceback.format_exc()
            write_error_log(error)
            self.status_label.setText("分析发生异常")
            self.append_log(
                "分析异常，详细信息已写入错误日志。"
            )

            self.set_analysis_busy(False)
            QMessageBox.critical(
                self,
                "分析异常",
                "分析过程中发生异常，程序未退出。\n"
                "详细错误已写入：{}".format(
                    os.path.join(get_app_dir(), ERROR_LOG)
                )
            )

        finally:
            self.set_analysis_busy(False)
            # 仅在没有扫描任务时解锁；自动扫描后分析不会误解锁中途控件。
            if self.scan_thread is None:
                self.set_busy(False)

    def closeEvent(self, event):
        if self.scan_thread and self.scan_thread.isRunning():
            self._closing_requested = True
            if self.scan_worker:
                self.scan_worker.request_stop()
            self.status_label.setText("正在安全停止扫描，停止后自动退出……")
            self.append_log("收到退出请求：等待扫描线程安全结束。")
            event.ignore()
            return
        event.accept()


# ============================================================
# 全局异常保护
# ============================================================

def global_exception_hook(exc_type, exc_value, exc_tb):
    error = "".join(
        traceback.format_exception(
            exc_type,
            exc_value,
            exc_tb
        )
    )
    write_error_log(error)

    try:
        QMessageBox.critical(
            None,
            "程序发生异常",
            "程序发生未处理异常，但错误已记录。\n\n"
            "错误日志：{}".format(
                os.path.join(get_app_dir(), ERROR_LOG)
            )
        )
    except Exception:
        pass


def main():
    sys.excepthook = global_exception_hook

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("EmailNumberCheck")
    app.setFont(QFont("Microsoft YaHei UI", 9))

    window = MainWindow()
    window.show()

    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
