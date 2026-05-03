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

# ==========================================
# 版本号
# ==========================================
APP_VERSION = "1.0.8"

# ==========================================
# 语言与配置
# ==========================================
LANG_MAP = {
    'zh': {
        'title': "SheetPic - 图片提取 & 嵌入助手",
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
    },
    'en': {
        'title': "SheetPic - Image Extract & Embed Tool",
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
    }
}

COLORS = {
    'bg': '#F0F0F0', 'card': '#FFFFFF', 'primary': '#2563EB', 'primary_hov': '#1D4ED8',
    'danger': '#DC2626', 'text': '#1F2937', 'text_sub': '#666666', 'success': '#10B981',
    'border': '#DDDDDD', 'disabled_bg': '#CCCCCC', 'disabled_fg': '#555555'
}

GITHUB_URL = "https://github.com/youngoris/SheetPic"


class SheetPicApp:
    def __init__(self, root):
        self.root = root
        self.setup_lang()
        self.root.title(f"{self.T['title']}  v{APP_VERSION}")
        self.root.configure(bg=COLORS['bg'])
        if platform.system() == "Darwin":
            self.root.geometry("520x700")
        else:
            self.root.geometry("500x680")

        self.default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        self.file_path = None
        self.df = None
        self.wb = None
        self.ws = None
        self.header_row = 0
        self.is_running = False

        # Extract state
        self.sorted_img_cols = []

        # Embed state
        self.embed_url_col_idx = 0
        self.embed_sku_col_idx = 0
        self.embed_url_cols = []

        self.setup_style()
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

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
            base_font = ("PingFang SC", 13)
            bold_font = ("PingFang SC", 13, "bold")
        else:
            base_font = ("Microsoft YaHei UI", 9)
            bold_font = ("Microsoft YaHei UI", 10, "bold")

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
        self.root.config(menu=menubar)

        # === card1: 数据来源 (共享) ===
        card1 = tk.Frame(self.root, bg=COLORS['card'], padx=15, pady=15)
        card1.pack(fill='x', padx=15, pady=(20, 5))
        tk.Label(card1, text=self.T['sec_source'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8, "bold")).pack(anchor='w', pady=(0, 5))
        row1 = tk.Frame(card1, bg=COLORS['card'])
        row1.pack(fill='x')
        self.entry_path = ttk.Entry(row1)
        self.entry_path.pack(side='left', fill='x', expand=True, padx=(0, 5), ipady=3)
        self._setup_dnd(self.entry_path)
        ttk.Button(row1, text=self.T['btn_browse'], width=10, command=self.select_file).pack(side='left', padx=2)
        ttk.Button(row1, text=self.T['btn_clip'], width=8, command=self.load_clipboard).pack(side='left')

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
                                    fg=COLORS['text_sub'], font=("Arial", 8))
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
                 font=("Arial", 8), bg=COLORS['bg'], fg=COLORS['text_sub']).pack(side='left')
        lbl_link = tk.Label(footer, text="GitHub", font=("Arial", 8),
                             bg=COLORS['bg'], fg=COLORS['primary'], cursor="hand2")
        lbl_link.pack(side='right')
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open(GITHUB_URL))

        # === 日志区 ===
        log_frame = tk.Frame(self.root, bg=COLORS['card'], bd=1, relief="flat")
        log_frame.pack(fill='both', expand=True, padx=15, pady=(5, 5))
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, font=("Consolas", 8),
                                                   bd=0, highlightthickness=0)
        self.log_text.pack(fill='both', expand=True)
        self.log_text.configure(bg="#F5F5F5", fg="#444", padx=10, pady=10, state='normal')

        self.mode = 'extract'
        self.log(self.T['log_ready'])

    def _build_extract_tab(self, parent):
        """构建「提取图片」Tab"""
        tk.Label(parent, text=self.T['sec_settings'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8, "bold")).pack(anchor='w', pady=(0, 5))

        row_cols = tk.Frame(parent, bg=COLORS['card'])
        row_cols.pack(fill='x')
        col_box1 = tk.Frame(row_cols, bg=COLORS['card'])
        col_box1.pack(side='left', fill='x', expand=True, padx=(0, 5))
        tk.Label(col_box1, text=self.T['lbl_img'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8)).pack(anchor='w')
        self.combo_img = ttk.Combobox(col_box1, state="disabled")
        self.combo_img.pack(fill='x', pady=(2, 0))
        col_box2 = tk.Frame(row_cols, bg=COLORS['card'])
        col_box2.pack(side='left', fill='x', expand=True, padx=(5, 0))
        tk.Label(col_box2, text=self.T['lbl_code'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8)).pack(anchor='w')
        self.combo_code = ttk.Combobox(col_box2, state="disabled")
        self.combo_code.pack(fill='x', pady=(2, 0))

    def _build_embed_tab(self, parent):
        """构建「嵌入图片」Tab"""
        tk.Label(parent, text=self.T['sec_embed_settings'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8, "bold")).pack(anchor='w', pady=(0, 5))

        # URL列
        row_url = tk.Frame(parent, bg=COLORS['card'])
        row_url.pack(fill='x', pady=(0, 8))
        tk.Label(row_url, text=self.T['lbl_url_col'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8)).pack(anchor='w')
        self.combo_url = ttk.Combobox(row_url, state="disabled")
        self.combo_url.pack(fill='x', pady=(2, 0))

        # SKU列
        row_sku = tk.Frame(parent, bg=COLORS['card'])
        row_sku.pack(fill='x', pady=(0, 8))
        tk.Label(row_sku, text=self.T['lbl_sku_col'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8)).pack(anchor='w')
        self.combo_sku = ttk.Combobox(row_sku, state="disabled")
        self.combo_sku.pack(fill='x', pady=(2, 0))

        # 最大边长
        row_size = tk.Frame(parent, bg=COLORS['card'])
        row_size.pack(fill='x')
        tk.Label(row_size, text=self.T['lbl_img_size'], bg=COLORS['card'],
                 fg=COLORS['text_sub'], font=("Arial", 8)).pack(anchor='w')
        size_frame = tk.Frame(parent, bg=COLORS['card'])
        size_frame.pack(fill='x', pady=(2, 0))
        self.entry_max_dim = ttk.Entry(size_frame, width=8)
        self.entry_max_dim.pack(side='left', padx=(0, 6))
        self.entry_max_dim.insert(0, "500")

        # 插入原图
        self.var_original = tk.BooleanVar(value=False)
        chk_original = tk.Checkbutton(parent, text=self.T['chk_original'],
                                       variable=self.var_original, bg=COLORS['card'],
                                       fg=COLORS['text_sub'], font=("Arial", 8),
                                       activebackground=COLORS['card'],
                                       command=self._toggle_max_dim)
        chk_original.pack(anchor='w', pady=(6, 0))

        # 删除URL列
        self.var_del_url = tk.BooleanVar(value=False)
        chk_del = tk.Checkbutton(parent, text=self.T['chk_del_url'],
                                  variable=self.var_del_url, bg=COLORS['card'],
                                  fg=COLORS['text_sub'], font=("Arial", 8),
                                  activebackground=COLORS['card'])
        chk_del.pack(anchor='w', pady=(8, 0))

        # 写入原文件
        self.var_write_original = tk.BooleanVar(value=False)
        chk_wo = tk.Checkbutton(parent, text=self.T['chk_write_original'],
                                 variable=self.var_write_original, bg=COLORS['card'],
                                 fg=COLORS['text_sub'], font=("Arial", 8),
                                 activebackground=COLORS['card'])
        chk_wo.pack(anchor='w', pady=(8, 0))

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
            self.analyze_data()

    def on_tab_changed(self, event):
        tab = self.notebook.index(self.notebook.select())
        self.mode = 'extract' if tab == 0 else 'embed'

    def log(self, msg):
        now = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{now} {msg}\n")
        self.log_text.see(tk.END)

    def select_file(self):
        p = filedialog.askopenfilename(filetypes=[("Data", "*.xlsx;*.xls;*.csv;*.html"), ("All", "*.*")])
        if p:
            self.file_path = p
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, os.path.basename(p))
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
                self.entry_path.delete(0, tk.END)
                self.entry_path.insert(0, "Clipboard Data")
                self.process_df()
            else:
                self.log("❌ Clipboard empty")
        except Exception as e:
            self.log(f"❌ Error: {e}")

    def find_robust_header(self, file_path):
        try:
            if os.path.splitext(file_path)[1].lower() == '.csv':
                return 0
            df_raw = pd.read_excel(file_path, header=None, nrows=30)
            row_counts = df_raw.count(axis=1)
            if row_counts.empty:
                return 0
            mode_col_count = row_counts.value_counts().idxmax()
            for idx, count in row_counts.items():
                if count >= mode_col_count:
                    return idx
            return 0
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

            if ext == '.xlsx':
                try:
                    self.wb = openpyxl.load_workbook(self.file_path, data_only=True)
                    self.ws = self.wb.active
                except Exception:
                    pass

            if ext in ['.xlsx', '.xls']:
                self.header_row = self.find_robust_header(self.file_path)
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
                self.df = pd.read_excel(self.file_path, header=self.header_row)

        except Exception as e:
            self.log(f"❌ Error: {e}")

        self.root.after(0, lambda: self.progress.stop())
        self.root.after(0, lambda: self.progress.config(mode='determinate'))
        self.root.after(0, lambda: self.progress.__setitem__('value', 0))
        if self.df is not None and not self.df.empty:
            self.process_df()

    def process_df(self):
        self.df = self.df.astype(str)
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
        for i, c in enumerate(cols):
            sample_has_url = self.df[c].head(50).str.contains("http", case=False).any()
            if sample_has_url:
                real_count = self.df[c].str.contains("http", case=False, na=False).sum()
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
                            continue
                        if not val.startswith("http"):
                            m = re.search(r'(https?://[^\s;]+)', val)
                            if m:
                                val = m.group(1)
                        val = val.split('?')[0].split('!')[0]
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
                        return False, self.T['msg_err'].format(filename_base, f"File too large ({cl//1024//1024}MB)")
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
                                return False, self.T['msg_err'].format(filename_base, "File too large")
                            f.write(chunk)
                    return True, "OK"
                elif r.status_code == 404:
                    return False, self.T['msg_404'].format(filename_base)
                else:
                    return False, self.T['msg_err'].format(filename_base, f"HTTP {r.status_code}")
            except requests.exceptions.Timeout:
                if attempt == 0:
                    continue
                return False, self.T['msg_timeout'].format(filename_base)
            except Exception as e:
                if attempt == 0:
                    continue
                return False, self.T['msg_err'].format(filename_base, str(e))
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
                        return False, self.T['msg_err'].format(url[:50], f"File too large ({cl//1024//1024}MB)")
                    pil_img = PILImage.open(BytesIO(r.content))
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
                    return False, self.T['msg_err'].format(url[:50], f"HTTP {r.status_code}")
            except requests.exceptions.Timeout:
                if attempt == 0:
                    continue
                return False, self.T['msg_timeout'].format(url[:50])
            except Exception as e:
                if attempt == 0:
                    continue
                return False, self.T['msg_err'].format(url[:50], str(e))
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

        if write_original:
            out_file, ws, wb_out, img_header_col, header_row_excel = \
                self._embed_setup_original(fname, url_col_idx)
        else:
            out_file, ws, wb_out, img_header_col = \
                self._embed_setup_new(fname, url_col_idx)

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
                url = self.clean_url(url)
            else:
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

            if write_original:
                excel_row = header_row_excel + 1 + i
            else:
                excel_row = i + 2
                # Write cell values for new-workbook mode
                out_col = 1
                for j in range(len(orig_cols)):
                    if j == url_col_idx:
                        out_col += 1
                    else:
                        cell_val = str(self.df.iloc[i, j]) if self.df.iloc[i, j] is not None else ""
                        if cell_val.lower() == 'nan':
                            cell_val = ""
                        ws.cell(row=excel_row, column=out_col, value=cell_val)
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
                    if write_original:
                        excel_row = header_row_excel + 1 + i
                    else:
                        excel_row = i + 2
                        out_col = 1
                        for j in range(len(orig_cols)):
                            if j == url_col_idx:
                                out_col += 1
                            else:
                                cell_val = str(self.df.iloc[i, j]) if self.df.iloc[i, j] is not None else ""
                                if cell_val.lower() == 'nan':
                                    cell_val = ""
                                ws.cell(row=excel_row, column=out_col, value=cell_val)
                                out_col += 1
                    ws.cell(row=excel_row, column=img_header_col, value=self.T['msg_dl_skip'])

        self.root.after(0, lambda: self.log(self.T['log_embed_save']))
        wb_out.save(out_file)
        wb_out.close()

        duration = time.time() - t_start
        self.root.after(0, lambda: self.embed_finish(success, fail, out_file, duration))

    def _embed_setup_new(self, fname, url_col_idx):
        """Create a new workbook for embedding. Returns (out_file, ws, wb, img_header_col)."""
        dest = self.entry_dest.get()
        out_file = os.path.join(dest, f"{fname}_Embedded.xlsx")
        wb_out = openpyxl.Workbook()
        ws = wb_out.active

        orig_cols = list(self.df.columns)
        out_col = 1
        img_header_col = 1
        for i, col_name in enumerate(orig_cols):
            ws.cell(row=1, column=out_col, value=col_name)
            out_col += 1
            if i == url_col_idx:
                ws.cell(row=1, column=out_col, value="图片")
                img_header_col = out_col
                out_col += 1

        return out_file, ws, wb_out, img_header_col

    def _embed_setup_original(self, fname, url_col_idx):
        """Load original workbook, insert image column. Returns (out_file, ws, wb, img_header_col, header_row_excel)."""
        dest = self.entry_dest.get()
        out_file = os.path.join(dest, f"{fname}_WithImages.xlsx")

        if self.file_path == "Clipboard" or not os.path.exists(self.file_path):
            # Clipboard mode: fallback to new workbook
            return self._embed_setup_new(fname, url_col_idx) + (2,)

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
