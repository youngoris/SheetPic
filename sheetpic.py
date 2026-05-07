"""
SheetPic - 批量图片提取 & 嵌入 工具
支持两种模式:
  - 提取图片: 从Excel下载/导出嵌入图片
  - 嵌入图片: 将URL图片下载后嵌入Excel单元格
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
import threading
import platform
import requests
import concurrent.futures
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as XlImage
from PIL import Image as PILImage
from io import BytesIO
import webbrowser
import datetime
import mimetypes
import re
import json
import sys
import time
import subprocess
import urllib.request

# ==========================================
# 版本号
# ==========================================
APP_VERSION = "1.0.29"

# ==========================================
# 语言与配置
# ==========================================
LANG_MAP = {
    'zh': {
        'title': "表图 - 表格图片提取 & 嵌入工具",
        'menu_lang': "语言",
        'footer_text': "Build by Andre  |  v{}",
        'tab_extract': "提取图片",
        'tab_embed': "嵌入图片",
        'sec_source': "数据来源",
        'sec_settings': "匹配与保存",
        'sec_embed_settings': "嵌入设置",
        'btn_browse': "📂 选择文件",
        'btn_clip': "📋 剪贴板",
        'lbl_dest': "保存位置:",
        'btn_dest': "修改",
        'lbl_sheet': "工作表:",
        'unnamed': "未命名",
        # Extract
        'lbl_img': "图片来源 (默认智能合并)",
        'lbl_code': "文件名列 (ID/SKU)",
        'opt_auto': "★ [智能合并] 优先下载数据最多的列 (推荐)",
        'type_url': "[链接] {} (含 {} 个URL)",
        'type_img': "[图片] {} (含 {} 张嵌入图)",
        'msg_skip': "❌ {}: [空] 未检测到有效图片",
        'done_msg': "耗时: {:.1f}s\n成功: {}\n失败: {}\n跳过: {}\n保存至: {}",
        # Embed
        'lbl_url_col': "图片URL列 (含链接的列)",
        'lbl_sku_col': "SKU/ID列 (用于排序)",
        'lbl_img_size': "最大边长 (px)",
        'chk_original': "插入原图 (不缩放)",
        'chk_del_url': "嵌入后删除原URL列",
        'chk_write_original': "写入原文件 (保留格式)",
        'msg_no_url': "❌ 未检测到包含URL的列",
        'msg_embed_done': "耗时: {:.1f}s\n嵌入成功: {}\n下载失败: {}\n输出文件: {}",
        'msg_dl_fail': "[下载失败]",
        'msg_dl_skip': "[无URL]",
        'msg_invalid_url': "⚠️ {}: [无效URL] {}",
        'msg_conn_err': "❌ {}: [连接失败] {}",
        'msg_ssl_err': "❌ {}: [SSL错误] {}",
        'msg_too_large': "❌ {}: [文件过大] {}MB",
        'msg_bad_image': "❌ {}: [图片格式错误] {}",
        'log_embed_start': "开始嵌入图片处理...",
        'log_embed_save': "正在保存Excel文件...",
        'embed_status_run': "嵌入: {}/{} (成功: {} | 失败: {})",
        # Shared
        'btn_start': "开始处理",
        'btn_stop': "停止",
        'status_idle': "准备就绪",
        'status_run': "进度: {}/{} (成功: {} | 失败: {} | 跳过: {})",
        'status_stop': "正在停止...",
        'log_ready': "已就绪。请加载含图片的表格文件。",
        'log_header': "✅ 锁定表头: 第 {} 行",
        'log_stats': "📊 列分析: 列 {} 含 {} 条有效数据 (类型: {})",
        'msg_404': "❌ {}: [404] 链接失效/不存在",
        'msg_timeout': "⚠️ {}: [超时] 网络连接卡顿",
        'msg_err': "❌ {}: [错误] {}",
        'menu_help': "帮助",
        'menu_check_update': "检查更新",
        'update_available': "⬆ 发现新版本 {}！点击下载",
        'update_none': "✅ 当前已是最新版本",
        'update_check_fail': "检查更新失败",
    },
    'en': {
        'title': "SheetPic - Spreadsheet Image Extract & Embed",
        'menu_lang': "Language",
        'footer_text': "Build by Andre  |  v{}",
        'tab_extract': "Extract",
        'tab_embed': "Embed",
        'sec_source': "Data Source",
        'sec_settings': "Settings",
        'sec_embed_settings': "Embed Settings",
        'btn_browse': "📂 File",
        'btn_clip': "📋 Clip",
        'lbl_dest': "Output:",
        'btn_dest': "Change",
        'lbl_sheet': "Sheet:",
        'unnamed': "Unnamed",
        # Extract
        'lbl_img': "Image Source (Auto Merge)",
        'lbl_code': "Filename Column",
        'opt_auto': "★ [Auto Merge] Priority by count",
        'type_url': "[Link] {} ({} URLs)",
        'type_img': "[Image] {} ({} Embedded)",
        'msg_skip': "❌ {}: [Skip] No valid image found",
        'done_msg': "Time: {:.1f}s\nSuccess: {}\nFailed: {}\nSkipped: {}\nPath: {}",
        # Embed
        'lbl_url_col': "Image URL Column",
        'lbl_sku_col': "SKU/ID Column (for ordering)",
        'lbl_img_size': "Max Dimension (px)",
        'chk_original': "Original Size (no resize)",
        'chk_del_url': "Delete URL column after embedding",
        'chk_write_original': "Write to original file (preserve format)",
        'msg_no_url': "❌ No URL column detected",
        'msg_embed_done': "Time: {:.1f}s\nEmbedded: {}\nFailed: {}\nOutput: {}",
        'msg_dl_fail': "[Download Failed]",
        'msg_dl_skip': "[No URL]",
        'msg_invalid_url': "⚠️ {}: [Invalid URL] {}",
        'msg_conn_err': "❌ {}: [Connection Error] {}",
        'msg_ssl_err': "❌ {}: [SSL Error] {}",
        'msg_too_large': "❌ {}: [File Too Large] {}MB",
        'msg_bad_image': "❌ {}: [Bad Image] {}",
        'log_embed_start': "Starting image embedding...",
        'log_embed_save': "Saving Excel file...",
        'embed_status_run': "Embed: {} / {} (OK: {} | Fail: {})",
        # Shared
        'btn_start': "Start",
        'btn_stop': "Stop",
        'status_idle': "Ready",
        'status_run': "{} / {} (OK: {} Fail: {} Skip: {})",
        'status_stop': "Stopping...",
        'log_ready': "Ready. Load a table with images.",
        'log_header': "✅ Header at Row {}",
        'log_stats': "📊 Col Stats: {} has {} valid items ({})",
        'msg_404': "❌ {}: [404] Not Found",
        'msg_timeout': "⚠️ {}: [Timeout] Connection failed",
        'msg_err': "❌ {}: [Error] {}",
        'menu_help': "Help",
        'menu_check_update': "Check for Updates",
        'update_available': "⬆ New version {} available! Click to download",
        'update_none': "✅ Already up to date",
        'update_check_fail': "Update check failed",
    }
}

COLORS = {
    'bg': '#F0F0F0', 'card': '#FFFFFF', 'primary': '#2563EB', 'primary_hov': '#1D4ED8',
    'danger': '#DC2626', 'text': '#1F2937', 'text_sub': '#666666', 'success': '#10B981',
    'border': '#DDDDDD', 'disabled_bg': '#CCCCCC', 'disabled_fg': '#555555'
}

GITHUB_URL = "https://github.com/youngoris/SheetPic"


# ==========================================
# Header-row detection (exposed for testing)
# ==========================================
HEADER_KEYWORDS = {
    # Chinese
    '图片', '图', '图像', '主图', '链接', '网址', '地址', '编号', '编码', '货号',
    '型号', '商品', '产品', '名称', '品名', '规格', '颜色', '尺寸', '尺码',
    '价格', '单价', '售价', '成本', '数量', '库存', '单位', '重量', '材质',
    '描述', '备注', '分类', '类目', '品牌', '日期', '时间', '订单', '客户',
    'sku', '条码', '条形码', '序号',
    # English
    'image', 'img', 'photo', 'picture', 'thumbnail', 'url', 'link', 'href',
    'id', 'code', 'sku', 'name', 'title', 'product', 'item', 'brand',
    'price', 'cost', 'qty', 'quantity', 'stock', 'size', 'color', 'colour',
    'weight', 'desc', 'description', 'note', 'remark', 'category', 'date',
    'time', 'order', 'customer', 'no', 'no.', 'number',
}


def _is_blank(v):
    if v is None:
        return True
    if isinstance(v, float):
        try:
            import math
            return math.isnan(v)
        except Exception:
            return False
    if isinstance(v, str):
        return v.strip() == ''
    return False


def _looks_like_url(s):
    if not isinstance(s, str):
        return False
    s = s.strip().lower()
    return s.startswith('http://') or s.startswith('https://') or s.startswith('//')


def _cell_type(v):
    """Classify a cell into one of: blank/str/num/date/url."""
    if _is_blank(v):
        return 'blank'
    if isinstance(v, bool):
        return 'num'
    if isinstance(v, (int, float)):
        return 'num'
    try:
        import datetime as _dt
        if isinstance(v, (_dt.datetime, _dt.date)):
            return 'date'
    except Exception:
        pass
    if isinstance(v, str):
        if _looks_like_url(v):
            return 'url'
        # numeric-looking strings still count as numbers (e.g., "123")
        s = v.strip()
        try:
            float(s.replace(',', ''))
            return 'num'
        except Exception:
            pass
        return 'str'
    return 'str'


def _row_signature(row):
    """Return (n_filled, type_counts dict, values list)."""
    vals = [v for v in row if not _is_blank(v)]
    types = {'str': 0, 'num': 0, 'date': 0, 'url': 0}
    for v in row:
        t = _cell_type(v)
        if t == 'blank':
            continue
        types[t] = types.get(t, 0) + 1
    return len(vals), types, vals


def _score_header_row(df_raw, scan_rows=15):
    """Score the first `scan_rows` rows of `df_raw` and return the best index.

    `df_raw` is a pandas DataFrame loaded with header=None.
    """
    import math as _math
    if df_raw is None or df_raw.empty:
        return 0
    n_total_rows = len(df_raw)
    n_cols = df_raw.shape[1]
    if n_cols == 0:
        return 0

    scan_rows = min(scan_rows, n_total_rows)
    # Widest non-trivial row width across the whole sample → expected col count
    row_widths = [_row_signature(df_raw.iloc[i].tolist())[0] for i in range(n_total_rows)]
    max_width = max(row_widths) if row_widths else 0
    if max_width == 0:
        return 0

    # Most common width across data-region rows (skip first 3 to avoid title bias)
    from collections import Counter
    tail_widths = [w for w in row_widths[3:] if w > 0]
    if tail_widths:
        mode_width = Counter(tail_widths).most_common(1)[0][0]
    else:
        mode_width = max_width

    expected_width = max(mode_width, int(max_width * 0.6))

    best_idx = 0
    best_score = -_math.inf

    for idx in range(scan_rows):
        row_vals = df_raw.iloc[idx].tolist()
        n_filled, types, vals = _row_signature(row_vals)
        if n_filled == 0:
            continue

        # Fill ratio relative to expected width
        fill_ratio = min(1.0, n_filled / expected_width) if expected_width else 0
        # String purity
        str_ratio = types['str'] / n_filled
        # No URLs in headers
        url_penalty = -0.5 if types['url'] > 0 else 0
        # Numbers in a "header" row are suspicious — but tolerated up to ~30%
        num_ratio = (types['num'] + types['date']) / n_filled

        # Uniqueness (case-insensitive, stripped)
        normalized = [str(v).strip().lower() for v in vals]
        uniq_ratio = len(set(normalized)) / len(normalized) if normalized else 0

        # Average label length — headers are short
        avg_len = sum(len(str(v)) for v in vals) / len(vals)
        # Penalize very long cells (likely descriptions/titles)
        len_score = 1.0 if avg_len <= 12 else max(0.0, 1.0 - (avg_len - 12) / 30.0)

        # Keyword match
        kw_hits = 0
        for v in vals:
            if not isinstance(v, str):
                continue
            low = v.strip().lower()
            if low in HEADER_KEYWORDS:
                kw_hits += 1
                continue
            # Substring match for common Chinese keywords
            for kw in HEADER_KEYWORDS:
                if len(kw) >= 2 and kw in low:
                    kw_hits += 1
                    break
        kw_score = min(1.0, kw_hits / max(1, n_filled))

        # "Followed by data" — look at next 3 rows: should be ≥ as wide and
        # contain MORE numbers/dates/URLs than this row (mixed types).
        followed_score = 0.0
        look = min(3, n_total_rows - idx - 1)
        if look > 0:
            wider_or_equal = 0
            more_mixed = 0
            for j in range(1, look + 1):
                nf, tt, _ = _row_signature(df_raw.iloc[idx + j].tolist())
                if nf >= max(1, n_filled - 1):
                    wider_or_equal += 1
                # data rows typically have more non-string content than the header
                if (tt['num'] + tt['date'] + tt['url']) > types['num'] + types['date'] + types['url']:
                    more_mixed += 1
            followed_score = (wider_or_equal / look) * 0.5 + (more_mixed / look) * 0.5
        else:
            # last row of the sheet can't be a header
            followed_score = -1.0

        # Sparse-row penalty (likely a merged title spanning few cells)
        sparse_penalty = -0.6 if n_filled < max(2, expected_width * 0.5) else 0.0

        score = (
            fill_ratio * 2.0 +
            str_ratio * 1.5 +
            uniq_ratio * 1.5 +
            len_score * 1.0 +
            kw_score * 1.5 +
            followed_score * 2.0 +
            url_penalty +
            sparse_penalty -
            num_ratio * 1.2
        )

        # Tie-breaker: prefer the LATER row (titles come first)
        if score > best_score + 1e-9 or (abs(score - best_score) <= 1e-9 and idx > best_idx):
            best_score = score
            best_idx = idx

    return best_idx


class SheetPicApp:
    def __init__(self, root):
        self.root = root
        self.setup_lang()
        self.root.title(f"{self.T['title']}  v{APP_VERSION}")
        self.root.configure(bg=COLORS['bg'])
        if platform.system() == "Darwin":
            self.root.geometry("520x750")
        else:
            self.root.geometry("500x730")

        self.default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        self.file_path = None
        self.df = None
        self.wb = None
        self.ws = None
        self.header_row = 0
        self.is_running = False
        self.sheet_names = []
        self._ui_thread = threading.current_thread()

        # Extract state
        self.sorted_img_cols = []

        # Embed state
        self.embed_url_col_idx = 0
        self.embed_sku_col_idx = 0
        self.embed_url_cols = []

        self.setup_style()
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.after(2000, lambda: self.check_update(auto=True))

    CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".sheetpic_config")

    def setup_lang(self):
        # 1. 读取用户手动设置
        saved = self._load_config().get('lang')
        if saved in LANG_MAP:
            self.lang = saved
            self.T = LANG_MAP[self.lang]
            return

        # 2. 自动检测系统语言
        self.lang = self._detect_system_lang()
        self.T = LANG_MAP[self.lang]

    def _detect_system_lang(self):
        try:
            if platform.system() == "Darwin":
                import subprocess
                result = subprocess.run(
                    ['defaults', 'read', '-g', 'AppleLanguages'],
                    capture_output=True, text=True, timeout=3
                )
                if result.returncode == 0 and 'zh' in result.stdout.lower():
                    return 'zh'
            elif platform.system() == "Windows":
                # Windows API: GetUserDefaultUILanguage → LANGID
                import ctypes
                lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
                # 0x0804=zh-CN, 0x0404=zh-TW, 0x0C04=zh-HK, 0x1004=zh-SG
                if lang_id in (0x0804, 0x0404, 0x0C04, 0x1004):
                    return 'zh'
                # 也检查 MUI language list
                buf = ctypes.create_unicode_buffer(256)
                ctypes.windll.kernel32.GetUserPreferredUILanguages(
                    0x08, None, buf, ctypes.byref(ctypes.c_uint(256)))
                if 'zh' in buf.value.lower():
                    return 'zh'
            else:
                for var in ('LANG', 'LC_ALL', 'LC_MESSAGES'):
                    val = os.environ.get(var, '')
                    if 'zh' in val.lower():
                        return 'zh'
        except Exception:
            pass
        return 'en'

    def _load_config(self):
        try:
            with open(self.CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}

    def _save_config(self, lang):
        cfg = self._load_config()
        cfg['lang'] = lang
        with open(self.CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False)

    def switch_lang(self, lang):
        if lang == self.lang:
            return
        self.lang = lang
        self.T = LANG_MAP[lang]
        self._save_config(lang)
        messagebox.showinfo(
            self.T['title'],
            "语言已切换，程序将重启以应用更改。\nLanguage changed, restarting..." if lang == 'zh'
            else "Language changed. The app will restart to apply.\n语言已切换，程序将重启。"
        )
        self.root.destroy()
        os.execl(sys.executable, sys.executable, *sys.argv)

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        if platform.system() == "Darwin":
            base_font = ("PingFang SC", 11)
            bold_font = ("PingFang SC", 11, "bold")
        else:
            base_font = ("Microsoft YaHei UI", 11)
            bold_font = ("Microsoft YaHei UI", 11, "bold")

        style.configure(".", background=COLORS['card'], foreground=COLORS['text'], font=base_font)
        style.configure("TFrame", background=COLORS['card'])
        style.configure("TEntry", fieldbackground="#F5F5F5", bordercolor=COLORS['border'], padding=5)
        style.configure("TButton", background="#E8E8E8", foreground=COLORS['text'], borderwidth=0, font=base_font)
        style.map("TButton", background=[('active', '#D8D8D8'), ('disabled', COLORS['disabled_bg'])],
                   foreground=[('disabled', COLORS['disabled_fg'])])
        style.configure("Primary.TButton", background=COLORS['primary'], foreground="white",
                         font=bold_font, borderwidth=0)
        style.map("Primary.TButton", background=[('active', COLORS['primary_hov']), ('disabled', COLORS['disabled_bg'])],
                   foreground=[('disabled', COLORS['disabled_fg'])])
        style.configure("Danger.TButton", background=COLORS['danger'], foreground="white",
                         font=bold_font, borderwidth=0)
        style.map("Danger.TButton", background=[('disabled', COLORS['disabled_bg'])],
                   foreground=[('disabled', COLORS['disabled_fg'])])
        style.configure("Green.Horizontal.TProgressbar", background=COLORS['success'],
                         troughcolor="#DDDDDD", bordercolor=COLORS['card'], thickness=6)
        # Notebook tab style
        style.configure("TNotebook", background="#FFFFFF", borderwidth=0)
        style.configure("TNotebook.Tab", font=base_font, padding=[16, 6],
                         background="#E0E0E0", foreground="#555555")
        style.map("TNotebook.Tab",
                   background=[("selected", "#FFFFFF"), ("!selected", "#E0E0E0")],
                   foreground=[("selected", "#1F2937"), ("!selected", "#888888")])

    def setup_ui(self):
        # === 菜单栏 ===
        menubar = tk.Menu(self.root)
        lang_menu = tk.Menu(menubar, tearoff=0)
        lang_menu.add_command(label="中文", command=lambda: self.switch_lang('zh'))
        lang_menu.add_command(label="English", command=lambda: self.switch_lang('en'))
        menubar.add_cascade(label=self.T['menu_lang'], menu=lang_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label=self.T['menu_check_update'], command=lambda: self.check_update(auto=False))
        menubar.add_cascade(label=self.T['menu_help'], menu=help_menu)
        self.root.config(menu=menubar)

        # === card1: 数据来源 (共享) ===
        card1 = tk.Frame(self.root, bg=COLORS['card'], padx=15, pady=15)
        card1.pack(fill='x', padx=15, pady=(20, 5))
        tk.Label(card1, text=self.T['sec_source'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10, "bold")).pack(anchor='w', pady=(0, 5))
        row1 = tk.Frame(card1, bg=COLORS['card'])
        row1.pack(fill='x')
        self.entry_path = ttk.Entry(row1)
        self.entry_path.pack(side='left', fill='x', expand=True, padx=(0, 5), ipady=3)
        self._setup_dnd(self.entry_path)
        ttk.Button(row1, text=self.T['btn_browse'], width=10, command=self.select_file).pack(side='left', padx=2)
        ttk.Button(row1, text=self.T['btn_clip'], width=8, command=self.load_clipboard).pack(side='left')

        # 工作表选择 (xlsx多Sheet)
        row_sheet = tk.Frame(card1, bg=COLORS['card'])
        row_sheet.pack(fill='x', pady=(10, 0))
        tk.Label(row_sheet, text=self.T['lbl_sheet'], bg=COLORS['card'], width=8, anchor='w').pack(side='left')
        self.combo_sheet = ttk.Combobox(row_sheet, state="disabled")
        self.combo_sheet.pack(side='left', fill='x', expand=True, padx=5, ipady=3)
        self.combo_sheet.bind('<<ComboboxSelected>>', self.on_sheet_changed)

        # 输出目录 (共享)
        row_dest = tk.Frame(card1, bg=COLORS['card'])
        row_dest.pack(fill='x', pady=(10, 0))
        tk.Label(row_dest, text=self.T['lbl_dest'], bg=COLORS['card'], width=8, anchor='w').pack(side='left')
        self.entry_dest = ttk.Entry(row_dest)
        self.entry_dest.insert(0, self.default_dir)
        self.entry_dest.pack(side='left', fill='x', expand=True, padx=5, ipady=3)
        ttk.Button(row_dest, text=self.T['btn_dest'], width=6, command=self.select_folder).pack(side='left')

        # === Notebook: 两个 Tab ===
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='x', padx=15, pady=10)

        tab_extract = tk.Frame(self.notebook, bg=COLORS['card'], padx=15, pady=15)
        tab_embed = tk.Frame(self.notebook, bg=COLORS['card'], padx=15, pady=15)
        self.notebook.add(tab_extract, text=f"  {self.T['tab_extract']}  ")
        self.notebook.add(tab_embed, text=f"  {self.T['tab_embed']}  ")

        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_changed)

        # --- Tab 1: 提取图片 ---
        self._build_extract_tab(tab_extract)

        # --- Tab 2: 嵌入图片 ---
        self._build_embed_tab(tab_embed)

        # === 动作区 (共享) ===
        action_frame = tk.Frame(self.root, bg=COLORS['bg'])
        action_frame.pack(fill='x', padx=15, pady=5)
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", mode="determinate",
                                         style="Green.Horizontal.TProgressbar")
        self.progress.pack(fill='x', pady=(0, 5))
        self.lbl_status = tk.Label(action_frame, text="...", bg=COLORS['bg'],
                                    fg=COLORS['text_sub'], font=("Arial", 10))
        self.lbl_status.pack(anchor='e')
        btn_box = tk.Frame(action_frame, bg=COLORS['bg'])
        btn_box.pack(fill='x', pady=5)
        self.btn_run = ttk.Button(btn_box, text=self.T['btn_start'], style="Primary.TButton",
                                   command=self.start_thread, state='disabled')
        self.btn_run.pack(side='left', fill='x', expand=True, padx=(0, 5), ipady=5)
        self.btn_stop = ttk.Button(btn_box, text=self.T['btn_stop'], style="Danger.TButton",
                                    command=self.stop_thread, state='disabled')
        self.btn_stop.pack(side='right', fill='x', expand=True, padx=(5, 0), ipady=5)

        # === 页脚 (先打包，确保不被日志区挤掉) ===
        footer = tk.Frame(self.root, bg=COLORS['bg'])
        footer.pack(side='bottom', fill='x', padx=15, pady=8)
        tk.Label(footer, text=self.T['footer_text'].format(APP_VERSION),
                 font=("Arial", 10), bg=COLORS['bg'], fg=COLORS['text_sub']).pack(side='left')
        self.lbl_update = tk.Label(footer, text="", font=("Arial", 10, "bold"),
                                   bg=COLORS['bg'], fg=COLORS['primary'], cursor="hand2")
        self.lbl_update.pack(side='left', padx=(10, 0))
        lbl_link = tk.Label(footer, text="GitHub", font=("Arial", 10),
                             bg=COLORS['bg'], fg=COLORS['primary'], cursor="hand2")
        lbl_link.pack(side='right')
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_URL))

        # === 日志区 ===
        log_frame = tk.Frame(self.root, bg=COLORS['card'], bd=1, relief="flat")
        log_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, font=("Consolas", 10),
                                                   bd=0, highlightthickness=0)
        self.log_text.pack(fill='both', expand=True)
        self.log_text.configure(bg="#F5F5F5", fg="#444", padx=10, pady=10, state='normal')

        self.mode = 'extract'
        self.log(self.T['log_ready'])

    def _build_extract_tab(self, parent):
        """构建「提取图片」Tab"""
        tk.Label(parent, text=self.T['sec_settings'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10, "bold")).pack(anchor='w', pady=(0, 5))

        # 上下两行布局
        tk.Label(parent, text=self.T['lbl_img'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10)).pack(anchor='w')
        self.combo_img = ttk.Combobox(parent, state="disabled")
        self.combo_img.pack(fill='x', pady=(2, 6))

        tk.Label(parent, text=self.T['lbl_code'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10)).pack(anchor='w')
        self.combo_code = ttk.Combobox(parent, state="disabled")
        self.combo_code.pack(fill='x', pady=(2, 0))

    def _build_embed_tab(self, parent):
        """构建「嵌入图片」Tab"""
        tk.Label(parent, text=self.T['sec_embed_settings'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10, "bold")).pack(anchor='w', pady=(0, 5))

        # URL列
        row_url = tk.Frame(parent, bg=COLORS['card'])
        row_url.pack(fill='x', pady=(0, 8))
        tk.Label(row_url, text=self.T['lbl_url_col'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10)).pack(anchor='w')
        self.combo_url = ttk.Combobox(row_url, state="disabled")
        self.combo_url.pack(fill='x', pady=(2, 0))

        # SKU列
        row_sku = tk.Frame(parent, bg=COLORS['card'])
        row_sku.pack(fill='x', pady=(0, 8))
        tk.Label(row_sku, text=self.T['lbl_sku_col'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10)).pack(anchor='w')
        self.combo_sku = ttk.Combobox(row_sku, state="disabled")
        self.combo_sku.pack(fill='x', pady=(2, 0))

        # 最大边长 + 3 个选项 同行排列（4 列：[输入框] [chk] [chk] [chk]）
        row_size = tk.Frame(parent, bg=COLORS['card'])
        row_size.pack(fill='x')
        tk.Label(row_size, text=self.T['lbl_img_size'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 10)).pack(anchor='w')

        size_frame = tk.Frame(parent, bg=COLORS['card'])
        size_frame.pack(fill='x', pady=(2, 0))
        self.entry_max_dim = ttk.Entry(size_frame, width=8)
        self.entry_max_dim.grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky='nw', pady=(2, 0))
        self.entry_max_dim.insert(0, "500")

        def _mk_chk(text, var, **kw):
            return tk.Checkbutton(size_frame, text=text, variable=var,
                                  bg=COLORS['card'], fg=COLORS['text_sub'],
                                  font=("Arial", 10),
                                  activebackground=COLORS['card'], **kw)

        # 第二列：插入原图 + 嵌入后删除原URL列（上下两行）
        self.var_original = tk.BooleanVar(value=False)
        chk_original = _mk_chk(self.T['chk_original'], self.var_original,
                               command=self._toggle_max_dim)
        chk_original.grid(row=0, column=1, sticky='w', padx=(0, 12))

        self.var_del_url = tk.BooleanVar(value=False)
        chk_del = _mk_chk(self.T['chk_del_url'], self.var_del_url)
        chk_del.grid(row=1, column=1, sticky='w', padx=(0, 12))

        # 第三列：写入原文件
        self.var_write_original = tk.BooleanVar(value=False)
        chk_wo = _mk_chk(self.T['chk_write_original'], self.var_write_original)
        chk_wo.grid(row=0, column=2, sticky='w')

    # ==========================================
    # 通用方法
    # ==========================================

    def _setup_dnd(self, widget):
        """Enable drag-and-drop on a widget if tkdnd is available."""
        try:
            widget.drop_target_register('DND_Files')
            widget.dnd_bind('<<Drop>>', self._on_drop)
        except (tk.TclError, AttributeError):
            pass

    def _on_drop(self, event):
        """Handle dropped files."""
        path = event.data.strip()
        # tkdnd may wrap paths in braces or quote them
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if path.startswith('"') and path.endswith('"'):
            path = path[1:-1]
        # Handle multiple files: take the first one
        if ' ' in path and not os.path.exists(path):
            # Try splitting by spaces, find first existing file
            for part in path.split():
                if os.path.exists(part):
                    path = part
                    break
        if os.path.exists(path):
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, path)
            self.file_path = path
            self.combo_sheet.set('')
            self.combo_sheet['values'] = []
            self.combo_sheet.config(state='disabled')
            threading.Thread(target=self.analyze_data, daemon=True).start()

    def on_tab_changed(self, event):
        tab = self.notebook.index(self.notebook.select())
        self.mode = 'extract' if tab == 0 else 'embed'

    def _update_sheet_combo(self, selected):
        self.combo_sheet['values'] = self.sheet_names
        if selected in self.sheet_names:
            self.combo_sheet.set(selected)
        elif self.sheet_names:
            self.combo_sheet.current(0)
        if len(self.sheet_names) > 1:
            self.combo_sheet.config(state='readonly')
        else:
            self.combo_sheet.config(state='disabled')

    def on_sheet_changed(self, event=None):
        if not self.wb:
            return
        name = self.combo_sheet.get()
        if name not in self.wb.sheetnames:
            return
        self.ws = self.wb[name]
        self.log(f">>> Sheet: {name}")
        threading.Thread(target=self._reload_sheet_data, daemon=True).start()

    def _reload_sheet_data(self):
        self.root.after(0, lambda: self.progress.config(mode='indeterminate'))
        self.root.after(0, lambda: self.progress.start(15))
        self.df = None
        self.header_row = 0
        try:
            selected_sheet = self.combo_sheet.get()
            self.header_row = self.find_robust_header(self.file_path, sheet_name=selected_sheet)
            if self.header_row > 0:
                self.root.after(0, lambda: self.log(self.T['log_header'].format(self.header_row + 1)))
            self.df = pd.read_excel(self.file_path, header=self.header_row, sheet_name=selected_sheet)
        except Exception as e:
            self.root.after(0, lambda: self.log(f"❌ Error: {e}"))
        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.progress.config(mode='determinate'))
        self.root.after(0, lambda: self.progress.__setitem__('value', 0))
        if self.df is not None and not self.df.empty:
            self.process_df()

    def log(self, msg):
        if getattr(self, '_ui_thread', None) is not threading.current_thread():
            try:
                self.root.after(0, self.log, msg)
            except (RuntimeError, tk.TclError):
                pass
            return

        now = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{now} {msg}\n")
        self.log_text.see(tk.END)

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("Data", "*.xlsx;*.xls;*.csv;*.html"), ("All", "*.*")])
        if p:
            self.file_path = p
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, os.path.basename(p))
            self.combo_sheet.set('')
            self.combo_sheet['values'] = []
            self.combo_sheet.config(state='disabled')
            threading.Thread(target=self.analyze_data, daemon=True).start()

    def select_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.entry_dest.delete(0, tk.END)
            self.entry_dest.insert(0, d)

    def load_clipboard(self):
        self.log(">>> Reading clipboard...")
        try:
            self.df = pd.read_clipboard()
            if not self.df.empty:
                self.file_path = "Clipboard"
                self.wb = None
                self.combo_sheet.set('')
                self.combo_sheet['values'] = []
                self.combo_sheet.config(state='disabled')
                self.entry_path.delete(0, tk.END)
                self.entry_path.insert(0, "Clipboard Data")
                self.process_df()
            else:
                self.log("❌ Clipboard empty")
        except Exception as e:
            self.log(f"❌ Error: {e}")

    def find_robust_header(self, file_path, sheet_name=0):
        """Locate the header row in an Excel sheet.

        Strategy: score the first ~15 rows on multiple signals and pick the
        best. A real header row is characterized by:
          - High fill ratio (close to the widest row in the sheet)
          - Mostly string cells with short labels
          - Unique non-null values (no duplicate column names)
          - Followed by data rows of comparable width but with mixed types
            (numbers / dates / mixed strings)
          - Bonus when cells contain common header keywords
          - Penalty when the row is sparse (likely a merged title)
        """
        try:
            if os.path.splitext(file_path)[1].lower() == '.csv':
                return 0
            df_raw = pd.read_excel(file_path, header=None, nrows=40, sheet_name=sheet_name)
            return _score_header_row(df_raw)
        except Exception:
            return 0

    def analyze_data(self):
        self.root.after(0, lambda: self.progress.config(mode='indeterminate'))
        self.root.after(0, lambda: self.progress.start(15))
        self.df = None
        if self.wb:
            try:
                self.wb.close()
            except Exception:
                pass
        self.wb = None
        self.ws = None
        self.header_row = 0

        try:
            ext = os.path.splitext(self.file_path)[1].lower() if self.file_path != "Clipboard" else ""

            selected_sheet = 0
            if ext == '.xlsx':
                try:
                    self.wb = openpyxl.load_workbook(self.file_path, data_only=True)
                    self.ws = self.wb.active
                    self.sheet_names = self.wb.sheetnames
                    selected_sheet = self.ws.title
                    self.root.after(0, lambda: self._update_sheet_combo(selected_sheet))
                except Exception:
                    pass

            if ext in ['.xlsx', '.xls']:
                self.header_row = self.find_robust_header(self.file_path, sheet_name=selected_sheet)
                if self.header_row > 0:
                    self.log(self.T['log_header'].format(self.header_row + 1))

            if ext == '.csv':
                try:
                    self.df = pd.read_csv(self.file_path, encoding='utf-8-sig', on_bad_lines='skip')
                except Exception:
                    self.df = pd.read_csv(self.file_path, encoding='gbk', on_bad_lines='skip')
            elif ext == '.html':
                self.df = pd.read_html(self.file_path)[0]
            else:
                self.df = pd.read_excel(self.file_path, header=self.header_row, sheet_name=selected_sheet)

        except Exception as e:
            self.log(f"❌ Error: {e}")

        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.progress.config(mode='determinate'))
        self.root.after(0, lambda: self.progress.__setitem__('value', 0))
        if self.df is not None and not self.df.empty:
            self.process_df()

    def process_df(self):
        unnamed = self.T['unnamed']
        # Rename "Unnamed: N" placeholders, then de-duplicate so every column
        # has a unique label. Without this, duplicate names (common in real
        # spreadsheets, e.g. two "条码" columns or several blank header cells)
        # cause `df[name]` to return a DataFrame instead of a Series, breaking
        # `.str.contains` and silently hiding image/URL columns.
        new_cols = []
        seen = {}
        for c in self.df.columns:
            base = unnamed if str(c).startswith("Unnamed") else str(c)
            n = seen.get(base, 0)
            seen[base] = n + 1
            new_cols.append(base if n == 0 else f"{base}.{n}")
        self.df.columns = new_cols
        cols = list(self.df.columns)

        # --- Extract: 扫描嵌入图 + URL ---
        embed_counts = {}
        if self.wb:
            for img in getattr(self.ws, '_images', []):
                try:
                    c = img.anchor._from.col
                    embed_counts[c] = embed_counts.get(c, 0) + 1
                except (AttributeError, IndexError):
                    pass

        url_counts = {}
        for i in range(len(cols)):
            # Access by position to avoid duplicate-name pitfalls.
            series = self.df.iloc[:, i]
            real_count = self._count_http_values(series)
            if real_count > 0:
                url_counts[i] = real_count

        all_img_indices = set(embed_counts.keys()) | set(url_counts.keys())
        self.sorted_img_cols = []
        for idx in all_img_indices:
            count = max(embed_counts.get(idx, 0), url_counts.get(idx, 0))
            type_str = "embed" if idx in embed_counts else "url"
            self.sorted_img_cols.append({'idx': idx, 'count': count, 'type': type_str})
        self.sorted_img_cols.sort(key=lambda x: x['count'], reverse=True)

        # Extract combo 选项
        img_opts = []
        if self.sorted_img_cols:
            img_opts.append(self.T['opt_auto'])
        for item in self.sorted_img_cols:
            i = item['idx']
            col_letter = get_column_letter(i + 1)
            display_name = f"{cols[i]} ({col_letter})"
            if item['type'] == 'embed':
                label = self.T['type_img'].format(display_name, item['count'])
                self.log(self.T['log_stats'].format(col_letter, item['count'], "Embedded"))
            else:
                label = self.T['type_url'].format(display_name, item['count'])
                self.log(self.T['log_stats'].format(col_letter, item['count'], "URL"))
            img_opts.append(label)

        code_opts = [f"{c} ({get_column_letter(i+1)})" for i, c in enumerate(cols)]

        # --- Embed: URL列 + SKU列 ---
        self.embed_url_cols = [{'idx': i, 'count': c} for i, c in url_counts.items()]
        self.embed_url_cols.sort(key=lambda x: x['count'], reverse=True)

        url_opts = []
        for item in self.embed_url_cols:
            i = item['idx']
            col_letter = get_column_letter(i + 1)
            display_name = f"{cols[i]} ({col_letter})"
            url_opts.append(f"{display_name} - {item['count']} URLs")

        sku_opts = code_opts[:]  # same list

        self.root.after(0, lambda: self.update_ui_lists(img_opts, code_opts, url_opts, sku_opts))

    @staticmethod
    def _count_http_values(series):
        sample = series.dropna().head(50)
        if sample.empty:
            return 0
        sample_text = sample.astype(str)
        if not sample_text.str.contains("http", case=False, na=False, regex=False).any():
            return 0
        full_text = series.dropna().astype(str)
        return int(full_text.str.contains("http", case=False, na=False, regex=False).sum())

    def update_ui_lists(self, img_opts, code_opts, url_opts, sku_opts):
        # Extract combos
        self.combo_img['values'] = img_opts
        if img_opts:
            self.combo_img.current(0)
        self.combo_code['values'] = code_opts
        best = next((x for x in code_opts if any(k in x.lower() for k in ["code", "sku", "条码", "货号"])), None)
        if best:
            self.combo_code.set(best)
        elif code_opts:
            self.combo_code.current(0)
        self.combo_img.config(state='readonly')
        self.combo_code.config(state='readonly')

        # Embed combos
        self.combo_url['values'] = url_opts
        if url_opts:
            self.combo_url.current(0)
            self.embed_url_col_idx = self.embed_url_cols[0]['idx']
        else:
            if self.mode == 'embed':
                self.log(self.T['msg_no_url'])

        self.combo_sku['values'] = sku_opts
        best_sku = next((x for x in sku_opts if any(k in x.lower()
                         for k in ["code", "sku", "条码", "货号", "id"])), None)
        if best_sku:
            self.combo_sku.set(best_sku)
        elif sku_opts:
            self.combo_sku.current(0)
        self.combo_url.config(state='readonly')
        self.combo_sku.config(state='readonly')
        self.combo_url.bind('<<ComboboxSelected>>', self._on_url_selected)
        self.combo_sku.bind('<<ComboboxSelected>>', self._on_sku_selected)

        # Enable start button
        if img_opts or url_opts:
            self.btn_run.config(state='normal')

    def _get_col_index(self, s):
        match = re.search(r'\(([A-Z]+)\)', s)
        if match:
            return column_index_from_string(match.group(1)) - 1
        return 0

    def _on_url_selected(self, event):
        self.embed_url_col_idx = self._get_col_index(self.combo_url.get())

    def _on_sku_selected(self, event):
        self.embed_sku_col_idx = self._get_col_index(self.combo_sku.get())

    def _toggle_max_dim(self):
        if self.var_original.get():
            self.entry_max_dim.config(state='disabled')
        else:
            self.entry_max_dim.config(state='normal')

    # ==========================================
    # 线程控制
    # ==========================================

    def start_thread(self):
        self.is_running = True
        self.btn_run.config(state='disabled')
        self.btn_stop.config(state='normal')
        if self.mode == 'extract':
            threading.Thread(target=self.run_extract_process, daemon=True).start()
        else:
            threading.Thread(target=self.run_embed_process, daemon=True).start()

    def stop_thread(self):
        self.is_running = False
        self.log(">>> Stopping...")
        self.btn_stop.config(state='disabled')
        self.lbl_status.config(text=self.T['status_stop'])
        self.progress.stop()

    # ==========================================
    # 提取图片处理
    # ==========================================

    def run_extract_process(self):
        t_start = time.time()
        self._process_start_time = t_start
        dest = self.entry_dest.get()
        fname = "Clipboard" if self.file_path == "Clipboard" else os.path.splitext(os.path.basename(self.file_path))[0]
        out_dir = os.path.join(dest, f"{fname}_Img")
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        idx_code = self._get_col_index(self.combo_code.get())
        selection = self.combo_img.get()
        target_cols = []

        if "★" in selection:
            target_cols = self.sorted_img_cols
        else:
            sel_idx = self._get_col_index(selection)
            for item in self.sorted_img_cols:
                if item['idx'] == sel_idx:
                    target_cols = [item]
                    break

        img_map_row_col = {}
        if self.wb:
            for img in getattr(self.ws, '_images', []):
                r = img.anchor._from.row
                c = img.anchor._from.col
                if r not in img_map_row_col:
                    img_map_row_col[r] = {}
                img_map_row_col[r][c] = img

        success = 0
        fail = 0
        skipped = 0
        self.progress['maximum'] = len(self.df)
        tasks = []

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            for i in range(len(self.df)):
                if not self.is_running:
                    break

                code = str(self.df.iloc[i, idx_code]).strip()
                base_name = "".join([c for c in code if c.isalnum() or c in '-_'])
                if not base_name:
                    base_name = f"Row_{i+1}"

                row_images = []
                for col_info in target_cols:
                    c_idx = col_info['idx']
                    if col_info['type'] == 'embed':
                        excel_row = self.header_row + 1 + i
                        if excel_row in img_map_row_col and c_idx in img_map_row_col[excel_row]:
                            img_obj = img_map_row_col[excel_row][c_idx]
                            row_images.append(('embed', img_obj))
                    elif col_info['type'] == 'url':
                        val = str(self.df.iloc[i, c_idx]).strip()
                        if not val or val.lower() == 'nan' or "http" not in val.lower():
                            if val and val.lower() != 'nan':
                                self.root.after(0, lambda v=val: self.log(
                                    self.T['msg_invalid_url'].format(base_name, v[:60])))
                            continue
                        if not val.startswith("http"):
                            m = re.search(r'(https?://[^\s;]+)', val)
                            if m:
                                val = m.group(1)
                            else:
                                self.root.after(0, lambda v=val: self.log(
                                    self.T['msg_invalid_url'].format(base_name, v[:60])))
                                continue
                        val = self.clean_url(val.split('?')[0].split('!')[0])
                        row_images.append(('url', val))

                if not row_images:
                    skipped += 1
                    self.root.after(0, self.update_progress_ext, i+1+len(tasks), len(self.df), success, fail, skipped, self.T['msg_skip'].format(base_name))
                    continue

                for img_idx, (src_type, src_data) in enumerate(row_images):
                    suffix = f"-{img_idx}" if img_idx > 0 else ""
                    final_name = f"{base_name}{suffix}"

                    if src_type == 'embed':
                        try:
                            ext = ".png" if src_data.format == "png" else ".jpg"
                            path = os.path.join(out_dir, final_name + ext)
                            with open(path, "wb") as f:
                                f.write(src_data._data())
                            success += 1
                        except Exception:
                            fail += 1
                    else:
                        tasks.append(executor.submit(self.download_url, src_data, final_name, out_dir))

                self.root.after(0, self.update_progress_ext, i+1, len(self.df), success, fail, skipped, "Process")

            done_count = i if 'i' in locals() else 0
            for future in concurrent.futures.as_completed(tasks):
                if not self.is_running:
                    break
                is_ok, msg = future.result()
                if is_ok:
                    success += 1
                else:
                    fail += 1
                self.root.after(0, self.update_progress_ext, len(self.df), len(self.df), success, fail, skipped, msg)

        duration = time.time() - t_start
        self.root.after(0, lambda: self.extract_finish(success, fail, skipped, out_dir, duration))

    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

    def download_url(self, url, filename_base, out_dir):
        if not self.is_running:
            return False, "Stopped"
        for attempt in range(2):
            try:
                headers = {'User-Agent': 'Mozilla/5.0'}
                r = requests.get(url, headers=headers, timeout=10, stream=True)
                if not self.is_running:
                    return False, "Stopped"
                if r.status_code == 200:
                    cl = int(r.headers.get('Content-Length', 0))
                    if cl > self.MAX_FILE_SIZE:
                        return False, self.T['msg_too_large'].format(filename_base, cl // 1024 // 1024)
                    ct = r.headers.get('Content-Type', '').lower()
                    ext = mimetypes.guess_extension(ct)
                    if not ext:
                        ext = ".jpg"
                    path = os.path.join(out_dir, filename_base + ext)
                    written = 0
                    with open(path, 'wb') as f:
                        for chunk in r.iter_content(8192):
                            if not self.is_running:
                                return False, "Stopped"
                            written += len(chunk)
                            if written > self.MAX_FILE_SIZE:
                                f.close()
                                os.remove(path)
                                return False, self.T['msg_too_large'].format(filename_base, written // 1024 // 1024)
                            f.write(chunk)
                    return True, "OK"
                elif r.status_code == 404:
                    return False, self.T['msg_404'].format(filename_base)
                else:
                    return False, self.T['msg_err'].format(filename_base, f"HTTP {r.status_code} ({url[:60]})")
            except requests.exceptions.Timeout:
                if attempt == 0:
                    continue
                return False, self.T['msg_timeout'].format(filename_base)
            except requests.exceptions.SSLError as e:
                return False, self.T['msg_ssl_err'].format(filename_base, str(e)[:80])
            except requests.exceptions.ConnectionError as e:
                return False, self.T['msg_conn_err'].format(filename_base, str(e)[:80])
            except Exception as e:
                if attempt == 0:
                    continue
                return False, self.T['msg_err'].format(filename_base, f"{type(e).__name__}: {str(e)[:60]}")
        return False, self.T['msg_err'].format(filename_base, "Max retries exceeded")

    def _format_eta(self, current, total):
        if current <= 0 or not hasattr(self, '_process_start_time'):
            return ""
        elapsed = time.time() - self._process_start_time
        if elapsed < 1 or current < 2:
            return ""
        remaining = elapsed / current * (total - current)
        if remaining < 60:
            return f"  ETA {int(remaining)}s"
        return f"  ETA {int(remaining//60)}m{int(remaining%60)}s"

    def update_progress_ext(self, current, total, success, fail, skipped, msg):
        if not self.is_running:
            return
        self.progress['value'] = current
        eta = self._format_eta(current, total)
        self.lbl_status.config(text=self.T['status_run'].format(current, total, success, fail, skipped) + eta)
        if "OK" not in msg and "Process" not in msg:
            self.log(msg)

    def extract_finish(self, success, fail, skipped, path, duration):
        if self.is_running:
            self.lbl_status.config(text="Done")
            self.progress['value'] = self.progress['maximum']
            msg = self.T['done_msg'].format(duration, success, fail, skipped, path)
            self.log("-" * 20)
            self.log(msg.replace("\n", " "))
            if success > 0:
                messagebox.showinfo("Done", msg)
                self._open_folder(path)
        else:
            self.lbl_status.config(text="Stopped")
        self.is_running = False
        self.btn_run.config(state='normal')
        self.btn_stop.config(state='disabled')

    # ==========================================
    # 嵌入图片处理
    # ==========================================

    def clean_url(self, url):
        original = url
        url = re.sub(r'!\d+x\d+', '', url)
        url = re.sub(r'\?imageView2/[^&]*', '', url)
        url = re.sub(r'\?x-oss-process=[^&]*', '', url)
        url = re.sub(r'[?&](width|height|w|h|size|resize|quality|format)=[^&]*', '', url)
        url = re.sub(r'\?\d+$', '', url)
        url = re.sub(r'\?&+', '?', url)
        url = re.sub(r'\?$', '', url)
        if url != original:
            self.root.after(0, lambda: self.log(f"URL clean: {original[:60]}... -> {url[:60]}..."))
        return url

    def download_to_bytesio(self, url, max_dim=None):
        if not self.is_running:
            return False, "Stopped"
        for attempt in range(2):
            try:
                headers = {'User-Agent': 'Mozilla/5.0'}
                r = requests.get(url, headers=headers, timeout=10)
                if not self.is_running:
                    return False, "Stopped"
                if r.status_code == 200:
                    cl = int(r.headers.get('Content-Length', 0))
                    if cl > self.MAX_FILE_SIZE:
                        return False, self.T['msg_too_large'].format(url[:50], cl // 1024 // 1024)
                    try:
                        pil_img = PILImage.open(BytesIO(r.content))
                    except Exception as e:
                        return False, self.T['msg_bad_image'].format(url[:50], str(e)[:60])
                    if max_dim:
                        pil_img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)
                    buf = BytesIO()
                    if pil_img.mode in ('RGBA', 'LA', 'P'):
                        pil_img = pil_img.convert('RGBA')
                        pil_img.save(buf, format='PNG')
                    else:
                        pil_img = pil_img.convert('RGB')
                        pil_img.save(buf, format='JPEG', quality=85)
                    buf.seek(0)
                    return True, buf
                elif r.status_code == 404:
                    return False, self.T['msg_404'].format(url[:50])
                else:
                    return False, self.T['msg_err'].format(url[:50], f"HTTP {r.status_code} ({url[:60]})")
            except requests.exceptions.Timeout:
                if attempt == 0:
                    continue
                return False, self.T['msg_timeout'].format(url[:50])
            except requests.exceptions.SSLError as e:
                return False, self.T['msg_ssl_err'].format(url[:50], str(e)[:80])
            except requests.exceptions.ConnectionError as e:
                return False, self.T['msg_conn_err'].format(url[:50], str(e)[:80])
            except Exception as e:
                if attempt == 0:
                    continue
                return False, self.T['msg_err'].format(url[:50], f"{type(e).__name__}: {str(e)[:60]}")
        return False, self.T['msg_err'].format(url[:50], "Max retries exceeded")

    def run_embed_process(self):
        t_start = time.time()
        self._process_start_time = t_start
        dest = self.entry_dest.get()
        fname = "Clipboard" if self.file_path == "Clipboard" else os.path.splitext(os.path.basename(self.file_path))[0]

        self.root.after(0, lambda: self.log(self.T['log_embed_start']))

        if self.var_original.get():
            max_dim = None
        else:
            try:
                max_dim = int(self.entry_max_dim.get())
            except ValueError:
                max_dim = 500

        url_col_idx = self.embed_url_col_idx
        sku_col_idx = self.embed_sku_col_idx
        write_original = self.var_write_original.get()

        del_url = self.var_del_url.get()

        header_row_excel = self.header_row + 1
        if write_original:
            out_file, ws, wb_out, img_header_col, header_row_excel = \
                self._embed_setup_original(fname, url_col_idx, del_url)
        else:
            out_file, ws, wb_out, img_header_col, header_row_excel = \
                self._embed_setup_new(fname, url_col_idx, del_url, header_row_excel)

        orig_cols = list(self.df.columns)
        total = len(self.df)
        self.progress['maximum'] = total

        # Collect URLs
        rows_data = []
        for i in range(total):
            url_raw = str(self.df.iloc[i, url_col_idx]).strip()
            url = url_raw
            if url and url.lower() != 'nan' and 'http' in url.lower():
                if not url.startswith("http"):
                    m = re.search(r'(https?://[^\s;]+)', url)
                    if m:
                        url = m.group(1)
                    else:
                        self.root.after(0, lambda u=url_raw: self.log(
                            self.T['msg_invalid_url'].format(f"Row {i+1}", u[:60])))
                        url = None
                if url:
                    url = self.clean_url(url)
            else:
                if url_raw and url_raw.lower() != 'nan':
                    self.root.after(0, lambda u=url_raw: self.log(
                        self.T['msg_invalid_url'].format(f"Row {i+1}", u[:60])))
                url = None
            rows_data.append(url)

        success = 0
        fail = 0
        row_results = [None] * total

        # Download concurrently
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            futures = {}
            for i, url in enumerate(rows_data):
                if not self.is_running:
                    break
                if url:
                    futures[executor.submit(self.download_to_bytesio, url, max_dim)] = i
                else:
                    futures[executor.submit(lambda: (False, "No URL"))] = i

            completed = 0
            for future in concurrent.futures.as_completed(futures):
                if not self.is_running:
                    break
                row_idx = futures[future]
                try:
                    is_ok, result = future.result()
                except Exception as e:
                    is_ok, result = False, str(e)
                row_results[row_idx] = (is_ok, result)
                completed += 1
                self.root.after(0, self.update_progress_emb, completed, total, success, fail)

        # Embed images into sheet
        for i, result in enumerate(row_results):
            if not self.is_running:
                break
            if result is None:
                result = (False, "Stopped")

            is_ok, data = result

            excel_row = header_row_excel + 1 + i
            if not write_original:
                # Write cell values for new-workbook mode
                out_col = 1
                for j in range(len(orig_cols)):
                    if del_url and j == url_col_idx:
                        out_col += 1
                    else:
                        cell_val = str(self.df.iloc[i, j]) if self.df.iloc[i, j] is not None else ""
                        if cell_val.lower() == 'nan':
                            cell_val = ""
                        ws.cell(row=excel_row, column=out_col, value=cell_val)
                        out_col += 1
                        if not del_url and j == url_col_idx:
                            out_col += 1

            img_col_letter = get_column_letter(img_header_col)
            if is_ok:
                try:
                    xl_img = XlImage(data)
                    row_height_pt = 40
                    ws.row_dimensions[excel_row].height = row_height_pt
                    img_ratio = xl_img.width / xl_img.height if xl_img.height > 0 else 1
                    col_width = row_height_pt * 1.33 * img_ratio / 7 + 1
                    ws.column_dimensions[img_col_letter].width = max(col_width, 12)
                    scaled_h = int(row_height_pt * 1.33)
                    scaled_w = int(scaled_h * img_ratio)
                    xl_img.width = scaled_w
                    xl_img.height = scaled_h
                    ws.add_image(xl_img, f"{img_col_letter}{excel_row}")
                    success += 1
                except Exception:
                    ws.cell(row=excel_row, column=img_header_col, value=self.T['msg_dl_fail'])
                    fail += 1
            else:
                ws.cell(row=excel_row, column=img_header_col, value=self.T['msg_dl_fail'])
                fail += 1

        # Handle stopped rows
        if not self.is_running:
            for i in range(total):
                if row_results[i] is None:
                    excel_row = header_row_excel + 1 + i
                    if not write_original:
                        out_col = 1
                        for j in range(len(orig_cols)):
                            if del_url and j == url_col_idx:
                                out_col += 1
                            else:
                                cell_val = str(self.df.iloc[i, j]) if self.df.iloc[i, j] is not None else ""
                                if cell_val.lower() == 'nan':
                                    cell_val = ""
                                ws.cell(row=excel_row, column=out_col, value=cell_val)
                                out_col += 1
                                if not del_url and j == url_col_idx:
                                    out_col += 1
                    ws.cell(row=excel_row, column=img_header_col, value=self.T['msg_dl_skip'])

        # Clear URL column cell values when deleting URL column in original workbook
        if del_url and write_original:
            for r in range(header_row_excel + 1, ws.max_row + 1):
                ws.cell(row=r, column=img_header_col).value = None

        self.root.after(0, lambda: self.log(self.T['log_embed_save']))
        wb_out.save(out_file)
        wb_out.close()

        duration = time.time() - t_start
        self.root.after(0, lambda: self.embed_finish(success, fail, out_file, duration))

    def _embed_setup_new(self, fname, url_col_idx, del_url=False, header_row_excel=1):
        """Create a new workbook for embedding. Returns (out_file, ws, wb, img_header_col, header_row_excel)."""
        dest = self.entry_dest.get()
        out_file = os.path.join(dest, f"{fname}_Embedded.xlsx")
        wb_out = openpyxl.Workbook()
        ws = wb_out.active

        orig_cols = list(self.df.columns)
        out_col = 1
        img_header_col = 1
        for i, col_name in enumerate(orig_cols):
            if del_url and i == url_col_idx:
                ws.cell(row=header_row_excel, column=out_col, value="图片")
                img_header_col = out_col
                out_col += 1
            else:
                ws.cell(row=header_row_excel, column=out_col, value=col_name)
                out_col += 1
                if not del_url and i == url_col_idx:
                    ws.cell(row=header_row_excel, column=out_col, value="图片")
                    img_header_col = out_col
                    out_col += 1

        return out_file, ws, wb_out, img_header_col, header_row_excel

    def _embed_setup_original(self, fname, url_col_idx, del_url=False):
        """Load original workbook, insert image column. Returns (out_file, ws, wb, img_header_col, header_row_excel)."""
        dest = self.entry_dest.get()
        out_file = os.path.join(dest, f"{fname}_WithImages.xlsx")

        if self.file_path == "Clipboard" or not os.path.exists(self.file_path):
            # Clipboard mode: fallback to new workbook
            return self._embed_setup_new(fname, url_col_idx, del_url)

        wb_out = openpyxl.load_workbook(self.file_path)
        ws = wb_out.active

        # Find header row and URL column in the Excel sheet
        url_col_name = str(self.df.columns[url_col_idx])
        header_row_excel = self.header_row + 1  # 0-based DataFrame → 1-based Excel

        url_excel_col = None
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=header_row_excel, column=col_idx).value
            if cell_val is not None and str(cell_val).strip() == url_col_name:
                url_excel_col = col_idx
                break

        if url_excel_col is None:
            # Fallback: search all rows for the header
            for r in range(1, min(ws.max_row + 1, 20)):
                for c in range(1, ws.max_column + 1):
                    cell_val = ws.cell(row=r, column=c).value
                    if cell_val is not None and str(cell_val).strip() == url_col_name:
                        url_excel_col = c
                        header_row_excel = r
                        break
                if url_excel_col:
                    break

        if url_excel_col is None:
            # Last resort: use column index directly
            url_excel_col = url_col_idx + 1

        if del_url:
            # Replace URL column with image column
            img_header_col = url_excel_col
            ws.cell(row=header_row_excel, column=img_header_col, value="图片")
        else:
            # Insert a new column after the URL column for images
            img_header_col = url_excel_col + 1
            ws.insert_cols(img_header_col)
            ws.cell(row=header_row_excel, column=img_header_col, value="图片")

        return out_file, ws, wb_out, img_header_col, header_row_excel

    def update_progress_emb(self, current, total, success, fail):
        if not self.is_running:
            return
        self.progress['value'] = current
        eta = self._format_eta(current, total)
        self.lbl_status.config(text=self.T['embed_status_run'].format(current, total, success, fail) + eta)

    def embed_finish(self, success, fail, path, duration):
        if self.is_running:
            self.lbl_status.config(text="Done")
            self.progress['value'] = self.progress['maximum']
            msg = self.T['msg_embed_done'].format(duration, success, fail, path)
            self.log("-" * 20)
            self.log(msg.replace("\n", " "))
            if success > 0:
                messagebox.showinfo("Done", msg)
                self._open_folder(os.path.dirname(path))
        else:
            self.lbl_status.config(text="Stopped")
        self.is_running = False
        self.btn_run.config(state='normal')
        self.btn_stop.config(state='disabled')

    # ==========================================
    # 通用工具
    # ==========================================

    def _open_folder(self, path):
        try:
            if platform.system() == "Darwin":
                subprocess.run(["open", path], check=False)
            else:
                os.startfile(path)
        except Exception:
            pass

    def check_update(self, auto=False):
        """Check GitHub for latest release. auto=True suppresses log messages."""
        API_URL = "https://api.github.com/repos/youngoris/SheetPic/releases/latest"
        RAW_URL = "https://raw.githubusercontent.com/youngoris/SheetPic/main/sheetpic.py"
        PROXIES = [
            "https://ghfast.top/",
            "https://gh-proxy.com/",
        ]
        RELEASES_URL = GITHUB_URL + "/releases/latest"

        def _fetch(url, timeout=5):
            req = urllib.request.Request(url, headers={"User-Agent": "SheetPic"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read().decode("utf-8", errors="ignore")

        def _do():
            remote_ver = None
            dl_url = RELEASES_URL

            # Source 1: GitHub API (works globally, may be blocked in China)
            try:
                text = _fetch(API_URL)
                data = json.loads(text)
                tag = data.get("tag_name", "").lstrip("v")
                if tag:
                    remote_ver = tuple(int(x) for x in tag.split("."))
                    dl_url = data.get("html_url", RELEASES_URL)
            except Exception:
                pass

            # Source 2: Raw file via proxy (fallback for China)
            if remote_ver is None:
                for proxy in PROXIES:
                    try:
                        text = _fetch(proxy + RAW_URL, timeout=8)
                        m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', text)
                        if m:
                            remote_ver = tuple(int(x) for x in m.group(1).split("."))
                            break
                    except Exception:
                        continue

            if remote_ver is None:
                if not auto:
                    self.root.after(0, lambda: self.log(self.T['update_check_fail']))
                return

            local_ver = tuple(int(x) for x in APP_VERSION.split("."))
            if remote_ver > local_ver:
                remote_tag = ".".join(str(x) for x in remote_ver)
                self.root.after(0, lambda: self._show_update(remote_tag, dl_url))
            elif not auto:
                self.root.after(0, lambda: self.log(self.T['update_none']))

        threading.Thread(target=_do, daemon=True).start()

    def _show_update(self, version, url):
        self.lbl_update.config(text=self.T['update_available'].format(version))
        self.lbl_update.bind("<Button-1>", lambda e: webbrowser.open(url))

    def on_closing(self):
        self.is_running = False
        if self.wb:
            try:
                self.wb.close()
            except Exception:
                pass
        self.root.destroy()
        os._exit(0)


if __name__ == "__main__":
    root = tk.Tk()
    # Try loading tkdnd for drag-and-drop support
    try:
        root.tk.call('package', 'require', 'tkdnd')
    except tk.TclError:
        pass
    app = SheetPicApp(root)
    root.mainloop()
