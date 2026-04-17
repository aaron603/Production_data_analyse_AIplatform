"""
main.py
-------
产线数据分析AI平台 – Zillnk Efficiency Improvement Group
Entry point: launches the tkinter GUI.
"""

import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

# Ensure the project root is on sys.path so 'core' package is importable
_ROOT = os.path.dirname(os.path.abspath(__file__))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.data_extractor import (read_barcodes, run_extraction,
                                  run_extraction_all_pass, generate_missing_report,
                                  generate_duplicate_report,
                                  run_extraction_traverse,
                                  generate_folder_direct_excel)
from core.cpk_calculator import analyze_xlsx_folder, analyze_json_folder
from core.html_report import generate_report
from core.html_fail_report import generate_fail_report
from core.html_comprehensive_report import generate_comprehensive_report


# ============================================================================
# Helpers
# ============================================================================

def _ts() -> str:
    return datetime.now().strftime('%H:%M:%S')


_HELP_TEXT = """\
产线数据分析AI平台  –  使用说明
=====================================

【功能一：本地数据分析】

─────────────────────────────────────
一、输入 / 输出配置
─────────────────────────────────────
1. DUT条码 Excel 文件（可选）
   "最后一次pass数据"模式必填；"所选文件夹分析"模式若填写则只分析Excel中的条码。
   格式：第一列=序列号（主条码），第二列=产品编码（可选）。
   其他四种模式无需填写，程序自动扫描工站目录发现全部条码。

2. 输出目录
   分析结果存放位置。每次点击"开始分析"，程序自动在此目录下创建以
   "产品类别_时间戳"命名的子文件夹，本次所有输出均存入其中，不会覆盖
   历史记录。产品类别来自工站文件夹下 TestResult/ 的第一级子目录名称。

   <输出目录>/
     <产品类别>_<YYYYMMDD_HHMMSS>/          ← 本次运行专属目录
       <工站类型>/xlsx/                      提取出的测试结果 .xlsx 文件
       <工站类型>/json/                      提取出的测试测量 .json 文件
       missing_barcodes.xlsx                缺失条码报表（最后一次pass/全量模式）
       duplicate_barcodes.xlsx              重复测试条码报表（全部成功数据模式）
       fault_barcodes.xlsx                  故障条码列表（所有数据/仅失败模式）
       folder_direct_fail_analysis.xlsx     失败条码明细（所选文件夹分析·遍历模式）
       rule_suggestions_*.yaml              规则建议模板（所有数据/仅失败模式）
       cpk_report.html                      CPK测试数据分析报告
       fail_analysis_report.html            失败分析报告（所选文件夹分析·遍历模式）
       fault_database.db                    故障分析关系库（SQLite，可持续积累）
       analysis_log_<时间戳>.txt            本次运行完整过程日志

─────────────────────────────────────
二、测试工站配置
─────────────────────────────────────
   · 工站类型标签：自定义名称，如 FT1、FT2、Aging（用于报告分组）
   · 测试数据文件夹：该工站的数据根目录（包含 TestResult 子目录）
   · 点击 [＋ 添加工站] 增加配置行；在输入框内按 ↑/↓ 可在行间快速切换

   工站合并配置（点击展开）：
     将指定工站类型的数据归并到目标类型中联合分析。
     示例：FT2 → FT1，则 FT2 数据并入 FT1 统一计算 CPK。
     配置自动保存，下次启动恢复。

─────────────────────────────────────
三、分析模式详解
─────────────────────────────────────
 最后一次pass数据  [需要DUT条码Excel]
   从DUT条码Excel读取条码，取每个条码最新一次测试通过的记录做CPK。
   最常用模式，对应某批次发货模块的过程能力分析。

 所选文件夹分析  [无需Excel，可选填DUT条码过滤]
   · 场景A：配置文件夹直接包含目标 xlsx/json 文件 → 直接CPK分析
   · 场景B：配置文件夹为 TestResult 结构根目录 → 先多层遍历提取所有
     记录（pass+fail），再CPK分析。若存在失败记录，额外生成：
       fail_analysis_report.html（帕累托图+汇总卡片+失败条码表）
       folder_direct_fail_analysis.xlsx（3个Sheet：失败条码/失败测试项/从未成功）
   若填写DUT条码Excel，场景B只分析Excel中列出的条码。

 全部成功数据  [无需Excel]
   自动遍历所有工站目录，提取全部测试通过的记录做CPK。
   同一模块多次通过均纳入，体现完整过程能力。
   额外输出：duplicate_barcodes.xlsx（重复测试条码统计）

 所有数据（含失败）  [无需Excel]
   遍历全部记录（pass + fail），做全量CPK并自动建立故障分析库。
   支持跨站比对：同一模块在不同机台的测试数据关联分析。
   额外输出：fault_barcodes.xlsx + rule_suggestions_*.yaml

 仅失败数据  [无需Excel]
   仅对失败/异常记录做CPK，深度分析失败规律。
   输出高频失败测试项、设备错误 Top 排行、未分类故障样本。
   额外输出：fault_barcodes.xlsx + rule_suggestions_*.yaml

─────────────────────────────────────
四、分析文件类型
─────────────────────────────────────
   · xlsx — 读取 Test_Result_*.xlsx 做CPK（默认，适用所有产品）
   · json — 读取 *_MEASUREMENT_*.json 做CPK
            （适用 B3/B40 等有 MEASUREMENT 文件的产品，数据更完整）
   注：同一次分析中 xlsx 和 json 数量可能不同，这是正常现象，
       取决于产品类型（部分产品无 json 文件）。

─────────────────────────────────────
五、故障分析方式（所有数据/仅失败模式）
─────────────────────────────────────
   · 基础版（规则库）
     内置 16 类种子规则，按关键词匹配故障类型。
     匹配优先级：设备/通信错误 > 失败测试项名称 > 日志关键词。

   · 增强版（规则库+Ollama）
     规则库基础上，调用本地 Ollama LLM（localhost:11434）辅助
     分析未分类故障，给出根因推断和处置建议。
     需预先安装并启动 Ollama（推荐模型 qwen2.5:7b）。

─────────────────────────────────────
六、故障关系知识库（持续积累）
─────────────────────────────────────
   分析结束后在输出目录生成 rule_suggestions_*.yaml，包含：
     · 未分类故障样本（无规则匹配，最需要填写）
     · 高频失败测试项 Top10（附出现次数）
     · 设备/通信错误参考列表

   工程师知识积累流程：
     1. 用文本编辑器打开 rule_suggestions_*.yaml
     2. 找到 fault_type/suggestion 为空的条目，根据经验填写
     3. 菜单 → 工具 → 加载故障关系描述文件… → 选择该文件
     4. 系统自动合并到 fault_database.db（相同关键词更新，不重复导入）
     5. 下次分析时新规则立即生效，未分类故障比例持续下降

─────────────────────────────────────
七、操作控制
─────────────────────────────────────
   · 点击"开始分析"启动后台分析，按键变为"停止分析"
   · 点击"停止分析"可随时中止（当前文件处理完后停止）
   · 点击"查看运行日志"弹出实时日志窗口（关闭后隐藏，内容保留）
   · 所有配置（工站/模式/文件类型/合并规则）自动保存，重启后恢复

─────────────────────────────────────
八、HTML 报告说明
─────────────────────────────────────
   · 分析完成后自动在浏览器打开 comprehensive_report.html（综合报告）
   · 综合报告共6个 Tab：
       总览        — KPI卡片 + 良率趋势 + 失败类型分布 + 测试大类汇总
       失败分析    — Top25高频失败项柱图（点击查看明细） + 失败记录表
       CPK分析     — Cpk横向柱图（色标分级） + 完整CPK表
       数据分布    — 按测试项切换的堆叠直方图（pass蓝/fail红）+ 统计面板
       失败模式    — 失败类型统计 + 时序热图 + 多失败项SN分析
       故障回放    — 左侧SN列表（可搜索/过滤） + 右侧逐项测试结果展开
   · 同时也生成 cpk_report.html（CPK专项报告，按工站+测试大项分组）
   · 综合报告使用 Chart.js（需浏览器可访问 cdn.jsdelivr.net，离线时图表不显示）

─────────────────────────────────────────────
【功能二 / 三】深科技 / 立讯 MES 数据分析
─────────────────────────────────────────────
   功能待实现，敬请期待。

如有问题，请联系 Zillnk Efficiency Improvement Group。
"""


# ============================================================================
# Tooltip helper
# ============================================================================

_MODE_HINTS = {
    '最后一次pass数据':   '⚠ 需要DUT条码Excel\n仅取每个条码最近一次成功测试记录做CPK',
    '所选文件夹分析':     '✓ 无需Excel（可选填条码过滤）\n直接对所配置文件夹做CPK分析\n若填写DUT条码Excel，仅分析Excel中的条码',
    '全部成功数据':       '✓ 无需Excel，自动扫描\n收集全部工站测试成功记录（含同一模块多次）做CPK',
    '所有数据（含失败）': '✓ 无需Excel，自动扫描\n成功+失败全量CPK，自动建立故障分析库\n重点：跨站比对（同一模块在不同设备的测试数据关联分析）',
    '仅失败数据':         '✓ 无需Excel，自动扫描\n仅失败记录CPK，自动更新故障分析库\n重点：失败规律总结，高频失败项/设备错误模式挖掘',
}


class _ToolTip:
    """Lightweight balloon tooltip for any tkinter widget."""

    def __init__(self, widget, text_fn):
        self._widget = widget
        self._text_fn = text_fn   # callable() → str
        self._tip = None
        widget.bind('<Enter>', self._show)
        widget.bind('<Leave>', self._hide)
        widget.bind('<ButtonPress>', self._hide)

    def _show(self, _e=None):
        text = self._text_fn()
        if not text:
            return
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f'+{x}+{y}')
        tk.Label(self._tip, text=text, justify='left',
                 bg='#fffde7', fg='#333', relief='solid', bd=1,
                 font=('Segoe UI', 8), padx=6, pady=4).pack()

    def _hide(self, _e=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ============================================================================
# Product name inference for report title
# ============================================================================

_TR_NAMES = {'testresult', 'test_result', 'testresults', 'testdata', 'test_data'}
_DEBUG_PREFIXES = ('debug',)


def _get_product_category(station_configs: list) -> str:
    """
    Derive the product category from station folder configs.

    Strategy (tried in order for each configured folder):
    1. {folder}/TestResult/ exists as a subdirectory → list non-debug
       subdirs; the first one is the product category.
    2. The configured folder path itself contains a TestResult segment →
       take the immediately following path segment.

    Returns '' if nothing determinable across all configs.
    """
    from pathlib import Path

    def _is_debug(name: str) -> bool:
        return name.lower().startswith(_DEBUG_PREFIXES)

    for cfg in station_configs:
        folder = cfg.get('folder', '').strip()
        if not folder or not os.path.isdir(folder):
            continue

        # Strategy 1: TestResult is a subdirectory of the configured folder
        tr_path = os.path.join(folder, 'TestResult')
        if os.path.isdir(tr_path):
            try:
                cats = [
                    d for d in os.listdir(tr_path)
                    if os.path.isdir(os.path.join(tr_path, d)) and not _is_debug(d)
                ]
                if cats:
                    return cats[0]
            except OSError:
                pass

        # Strategy 2: TestResult appears as a segment inside the path itself
        parts = Path(folder).parts
        for i, part in enumerate(parts):
            if part.lower() in _TR_NAMES:
                if i + 1 < len(parts):
                    return parts[i + 1]
                break

    return ''


def _infer_product_name(station_configs: list) -> str:
    """Legacy wrapper used by HTML report title — delegates to _get_product_category."""
    return _get_product_category(station_configs)


# ============================================================================
# Station row widget
# ============================================================================

class StationRow:
    """One row in the station config table: [type entry] [folder entry] [Browse] [Delete]"""

    def __init__(self, parent_frame, delete_callback):
        self._del_cb = delete_callback
        self.frame = tk.Frame(parent_frame, bg='#f5f6fa')
        self.var_type = tk.StringVar()
        self.var_folder = tk.StringVar()

        self._type_entry = tk.Entry(self.frame, textvariable=self.var_type, width=10,
                                    font=('Segoe UI', 9))
        self._type_entry.pack(side='left', padx=(0, 3))

        self._folder_entry = tk.Entry(self.frame, textvariable=self.var_folder, width=28,
                                      font=('Segoe UI', 9))
        self._folder_entry.pack(side='left', padx=(0, 3), fill='x', expand=True)

        tk.Button(self.frame, text='浏览', font=('Segoe UI', 8),
                  command=self._browse, relief='flat', bg='#e0e4f0',
                  padx=5).pack(side='left', padx=(0, 3))

        tk.Button(self.frame, text='✕', font=('Segoe UI', 8, 'bold'), fg='#c62828',
                  command=self._on_delete,
                  relief='flat', bg='#fce4e4', padx=5).pack(side='left')

    def _browse(self):
        d = filedialog.askdirectory(title='选择测试数据文件夹')
        if d:
            self.var_folder.set(d)

    def _on_delete(self):
        self._del_cb(self)

    def pack(self, **kw):
        self.frame.pack(**kw)

    def destroy(self):
        self.frame.destroy()

    def get(self) -> dict:
        return {'type': self.var_type.get().strip(),
                'folder': self.var_folder.get().strip()}


# ============================================================================
# Merge rule row widget
# ============================================================================

class MergeRuleRow:
    """One station merge rule: [source entry] → 合并到 → [target entry] [Delete]"""

    def __init__(self, parent_frame, delete_callback):
        self._del_cb = delete_callback
        self.frame = tk.Frame(parent_frame, bg='#f9f9f9')
        self.var_src = tk.StringVar()
        self.var_dst = tk.StringVar()

        self._src_entry = tk.Entry(self.frame, textvariable=self.var_src, width=10,
                                   font=('Segoe UI', 9))
        self._src_entry.pack(side='left', padx=(0, 4))

        tk.Label(self.frame, text='→ 合并到', bg='#f9f9f9',
                 font=('Segoe UI', 8), fg='#555').pack(side='left', padx=(0, 4))

        self._dst_entry = tk.Entry(self.frame, textvariable=self.var_dst, width=10,
                                   font=('Segoe UI', 9))
        self._dst_entry.pack(side='left', padx=(0, 4))

        tk.Button(self.frame, text='✕', font=('Segoe UI', 8, 'bold'), fg='#c62828',
                  command=lambda: self._del_cb(self),
                  relief='flat', bg='#fce4e4', padx=5).pack(side='left')

    def pack(self, **kw):
        self.frame.pack(**kw)

    def destroy(self):
        self.frame.destroy()

    def get(self) -> dict:
        return {'src': self.var_src.get().strip(),
                'dst': self.var_dst.get().strip()}


# ============================================================================
# Local Analysis Tab
# ============================================================================

class LocalAnalysisTab:
    """Tab 1: 测试站本地测试数据分析"""

    def __init__(self, notebook: ttk.Notebook):
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text='  本地数据分析  ')

        self._report_path = None
        self._station_rows = []
        self._merge_rows = []
        self._merge_expanded = False
        self._stop_event = threading.Event()
        self._config_path = os.path.join(_ROOT, 'app_config.json')

        self._build_ui()
        self._load_config()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        # Fill the whole tab with the same background color first,
        # so there is no visible empty strip below the content sections.
        bg_fill = tk.Frame(self.frame, bg='#f0f2f5')
        bg_fill.pack(fill='both', expand=True)

        outer = tk.Frame(bg_fill, bg='#f0f2f5')
        outer.pack(fill='x', expand=False, padx=10, pady=6)
        self._outer_frame = outer

        # ── Section 1: Input / Output config ──────────────────────────
        sec1 = self._make_section(outer, '输入 / 输出配置')

        inp_row = tk.Frame(sec1, bg='white')
        inp_row.pack(fill='x', pady=(2, 3))
        inp_row.columnconfigure(1, weight=1)   # Excel entry expands
        inp_row.columnconfigure(4, weight=2)   # OutDir entry expands more

        tk.Label(inp_row, text='DUT条码(可选):', anchor='w',
                 bg='white', font=('Segoe UI', 9)).grid(
                     row=0, column=0, sticky='w', padx=(0, 4))
        self._var_excel = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_excel,
                 font=('Segoe UI', 9)).grid(
                     row=0, column=1, sticky='ew', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_file(
                      self._var_excel, '选择DUT条码Excel',
                      [('Excel', '*.xlsx *.xls')]),
                  bg='#e0e4f0', relief='flat', padx=5).grid(
                      row=0, column=2, padx=(0, 12))

        tk.Label(inp_row, text='输出目录:', anchor='w',
                 bg='white', font=('Segoe UI', 9)).grid(
                     row=0, column=3, sticky='w', padx=(0, 4))
        self._var_outdir = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_outdir,
                 font=('Segoe UI', 9)).grid(
                     row=0, column=4, sticky='ew', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_dir(self._var_outdir, '选择输出目录'),
                  bg='#e0e4f0', relief='flat', padx=5).grid(row=0, column=5)

        # ── Sections 2 + 3: Side by side ──────────────────────────────
        h_frame = tk.Frame(outer, bg='#f0f2f5')
        h_frame.pack(fill='x', pady=(0, 6))
        h_frame.columnconfigure(0, weight=1)
        h_frame.columnconfigure(1, weight=0, minsize=270)
        h_frame.rowconfigure(0, weight=1)

        # Left: Station config
        sec2 = tk.LabelFrame(h_frame, text='  测试工站配置  ',
                             font=('Segoe UI', 9, 'bold'),
                             bg='white', fg='#1a237e',
                             relief='groove', bd=1, padx=8, pady=6)
        sec2.grid(row=0, column=0, sticky='nsew', padx=(0, 4))
        self._build_station_list(sec2)
        self._build_merge_config(sec2)

        # Right: Analysis mode
        sec3 = tk.LabelFrame(h_frame, text='  分析模式  ',
                             font=('Segoe UI', 9, 'bold'),
                             bg='white', fg='#1a237e',
                             relief='groove', bd=1, padx=8, pady=6)
        sec3.grid(row=0, column=1, sticky='nsew')
        self._build_section3_content(sec3)

        # ── Section 4: Actions + progress ─────────────────────────────
        sec4 = self._make_section(outer, '操作')

        btn_row = tk.Frame(sec4, bg='white')
        btn_row.pack(fill='x', pady=(0, 4))

        self._btn_run = tk.Button(btn_row, text='开始分析',
                                  font=('Segoe UI', 9, 'bold'),
                                  bg='#3949ab', fg='white',
                                  relief='flat', padx=12, pady=4,
                                  command=self._on_run)
        self._btn_run.pack(side='left')

        tk.Button(btn_row, text='查看运行日志',
                  font=('Segoe UI', 9), bg='#455a64', fg='white',
                  relief='flat', padx=10, pady=4,
                  command=self._show_log).pack(side='left', padx=(8, 0))

        self._progress_label = tk.Label(sec4, text='就绪', anchor='w',
                                        font=('Segoe UI', 8), bg='white', fg='#555')
        self._progress_label.pack(fill='x')

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(sec4, variable=self._progress_var,
                                             maximum=100, mode='determinate')
        self._progress_bar.pack(fill='x', pady=(2, 0))

        # Build the detached log window (hidden until user opens it)
        self._build_log_window()

    # ── Station list (Section 2 upper part) ──────────────────────────────

    def _build_station_list(self, parent: tk.Frame):
        # Column header
        hdr = tk.Frame(parent, bg='#dde3f0')
        hdr.pack(fill='x', pady=(0, 2))
        tk.Label(hdr, text='工站类型', width=10, anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4, pady=1)
        tk.Label(hdr, text='测试数据文件夹路径', anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4)

        scroll_wrap = tk.Frame(parent, bg='#f5f6fa')
        scroll_wrap.pack(fill='x')

        self._station_canvas = tk.Canvas(
            scroll_wrap, bg='#f5f6fa', height=88, highlightthickness=0
        )
        vsb = ttk.Scrollbar(scroll_wrap, orient='vertical',
                             command=self._station_canvas.yview)
        self._station_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self._station_canvas.pack(side='left', fill='both', expand=True)

        self._rows_frame = tk.Frame(self._station_canvas, bg='#f5f6fa')
        self._canvas_win = self._station_canvas.create_window(
            (0, 0), window=self._rows_frame, anchor='nw'
        )

        def _on_rows_configure(_e):
            self._station_canvas.configure(
                scrollregion=self._station_canvas.bbox('all')
            )
            needed = self._rows_frame.winfo_reqheight() + 4
            new_h = max(28, min(needed, 130))
            self._station_canvas.configure(height=new_h)

        def _on_canvas_resize(e):
            self._station_canvas.itemconfig(self._canvas_win, width=e.width)

        self._rows_frame.bind('<Configure>', _on_rows_configure)
        self._station_canvas.bind('<Configure>', _on_canvas_resize)

        def _on_wheel(e):
            self._station_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

        self._station_canvas.bind('<MouseWheel>', _on_wheel)
        self._rows_frame.bind('<MouseWheel>', _on_wheel)

        for stype in ('FT1', 'FT2', 'Aging'):
            self._add_station_row(preset_type=stype)

        tk.Button(parent, text='＋ 添加工站', command=self._add_station_row,
                  font=('Segoe UI', 8), bg='#e8f5e9', relief='flat',
                  padx=6, pady=2).pack(anchor='w', pady=(4, 0))

    # ── Station merge config (Section 2 lower part, collapsible) ─────────

    def _build_merge_config(self, parent: tk.Frame):
        # Separator
        tk.Frame(parent, bg='#dde3f0', height=1).pack(fill='x', pady=(8, 0))

        # Toggle header
        merge_hdr = tk.Frame(parent, bg='white')
        merge_hdr.pack(fill='x', pady=(0, 0))
        self._merge_toggle_btn = tk.Button(
            merge_hdr, text='▶  工站合并配置',
            font=('Segoe UI', 8, 'bold'), bg='white', fg='#3949ab',
            relief='flat', anchor='w', cursor='hand2', padx=4, pady=3,
            command=self._toggle_merge,
        )
        self._merge_toggle_btn.pack(fill='x')

        # Merge body (hidden initially)
        self._merge_body = tk.Frame(parent, bg='#f9f9f9', relief='groove', bd=1)

        tk.Label(self._merge_body,
                 text='将指定工站类型的数据合并到目标工站类型中分析',
                 bg='#f9f9f9', font=('Segoe UI', 8), fg='#666').pack(
            anchor='w', padx=6, pady=(4, 2))

        self._merge_rows_frame = tk.Frame(self._merge_body, bg='#f9f9f9')
        self._merge_rows_frame.pack(fill='x', padx=6)

        tk.Button(self._merge_body, text='＋ 添加合并规则',
                  command=self._add_merge_row,
                  font=('Segoe UI', 8), bg='#e3f2fd', relief='flat',
                  padx=6, pady=2).pack(anchor='w', padx=6, pady=(4, 6))

    def _toggle_merge(self):
        self._merge_expanded = not self._merge_expanded
        if self._merge_expanded:
            self._merge_body.pack(fill='x', pady=(0, 4))
            self._merge_toggle_btn.configure(text='▼  工站合并配置')
        else:
            self._merge_body.pack_forget()
            self._merge_toggle_btn.configure(text='▶  工站合并配置')

    def _add_merge_row(self, src: str = '', dst: str = ''):
        row = MergeRuleRow(self._merge_rows_frame, self._delete_merge_row)
        if src:
            row.var_src.set(src)
        if dst:
            row.var_dst.set(dst)
        row.pack(fill='x', pady=1)
        self._merge_rows.append(row)

    def _delete_merge_row(self, row):
        self._merge_rows.remove(row)
        row.destroy()

    # ── Log window (popup) ────────────────────────────────────────────────

    def _build_log_window(self):
        """Create the detached log Toplevel (hidden until user opens it)."""
        self._log_win = tk.Toplevel(self.frame)
        self._log_win.title('运行日志 — 产线数据分析AI平台')
        self._log_win.geometry('760x420')
        self._log_win.configure(bg='#1e1e2e')
        self._log_win.protocol('WM_DELETE_WINDOW', self._log_win.withdraw)
        self._log_win.withdraw()   # hidden until user clicks 查看运行日志

        hdr = tk.Frame(self._log_win, bg='#1a237e')
        hdr.pack(fill='x')
        tk.Label(hdr, text='运行日志', font=('Segoe UI', 9, 'bold'),
                 bg='#1a237e', fg='white', padx=8, pady=4).pack(side='left')
        tk.Button(hdr, text='清空', font=('Segoe UI', 8),
                  bg='#3949ab', fg='white', relief='flat', padx=8,
                  command=self._clear_log).pack(side='right', padx=4, pady=2)

        self._log = scrolledtext.ScrolledText(
            self._log_win, font=('Consolas', 8),
            bg='#1e1e2e', fg='#a8d8a8', insertbackground='white',
            state='disabled', wrap='word',
        )
        self._log.pack(fill='both', expand=True, padx=2, pady=(0, 2))

    def _show_log(self):
        """Show (or bring to front) the log popup window."""
        self._log_win.deiconify()
        self._log_win.lift()
        self._log.see('end')

    # ── Section 3 content (analysis mode) ────────────────────────────────

    _CPK_MODE_LABELS = [
        '最后一次pass数据',
        '所选文件夹分析',
        '全部成功数据',
        '所有数据（含失败）',
        '仅失败数据',
    ]
    _CPK_MODE_VALUES = [
        'latest_pass',
        'folder_direct',
        'all_pass',
        'all_with_fail',
        'fail_only',
    ]

    def _build_section3_content(self, container: tk.Frame):
        """Fill the analysis mode panel (container = the LabelFrame sec3)."""
        # ── File type selection ────────────────────────────────────────
        type_row = tk.Frame(container, bg='white')
        type_row.pack(fill='x', pady=(0, 6))

        tk.Label(type_row, text='分析文件类型：', bg='white',
                 font=('Segoe UI', 9)).pack(side='left')

        self._var_file_type = tk.StringVar(value='xlsx')
        for val, lbl in (('xlsx', 'xlsx'), ('json', 'json')):
            tk.Radiobutton(type_row, text=lbl, variable=self._var_file_type, value=val,
                           bg='white', font=('Segoe UI', 9),
                           activebackground='white').pack(side='left', padx=(0, 10))

        # ── Mode combobox ──────────────────────────────────────────────
        tk.Label(container, text='分析模式：', bg='white',
                 font=('Segoe UI', 9), anchor='w').pack(fill='x', pady=(0, 2))

        self._mode_display_var = tk.StringVar(value=self._CPK_MODE_LABELS[0])
        mode_combo = ttk.Combobox(container, textvariable=self._mode_display_var,
                                  values=self._CPK_MODE_LABELS, state='readonly',
                                  width=22, font=('Segoe UI', 9))
        mode_combo.pack(fill='x', pady=(0, 4))
        _ToolTip(mode_combo, lambda: _MODE_HINTS.get(self._mode_display_var.get(), ''))

        # ── Fault analysis method (shown only for all_with_fail / fail_only) ──
        self._fault_frame = tk.Frame(container, bg='white')

        tk.Frame(self._fault_frame, bg='#cccccc', height=1).pack(fill='x', pady=(2, 6))

        fault_lbl = tk.Label(self._fault_frame, text='故障分析方式：', bg='white',
                             font=('Segoe UI', 9), anchor='w')
        fault_lbl.pack(fill='x')
        _ToolTip(fault_lbl,
                 lambda: '基础版：仅使用规则库匹配\n增强版：规则库 + Ollama LLM 辅助分类（需本地Ollama）')

        self._fault_level_var = tk.StringVar(value='基础版（规则库）')
        ttk.Combobox(self._fault_frame, textvariable=self._fault_level_var,
                     values=['基础版（规则库）', '增强版（规则库+Ollama）'],
                     state='readonly', width=22,
                     font=('Segoe UI', 9)).pack(fill='x', pady=(2, 0))

        # Show/hide fault panel when mode changes
        self._mode_display_var.trace_add('write', lambda *_: self._update_fault_panel())
        self._update_fault_panel()

    def _update_fault_panel(self):
        """Show fault analysis options only for modes that update the fault DB."""
        mode = self._mode_display_var.get()
        if mode in ('所有数据（含失败）', '仅失败数据'):
            self._fault_frame.pack(fill='x')
        else:
            self._fault_frame.pack_forget()

    # ── Section helper ────────────────────────────────────────────────────

    def _make_section(self, parent, title: str, expand: bool = False) -> tk.Frame:
        lf = tk.LabelFrame(parent, text=f'  {title}  ',
                           font=('Segoe UI', 9, 'bold'),
                           bg='white', fg='#1a237e',
                           relief='groove', bd=1, padx=8, pady=6)
        if expand:
            lf.pack(fill='both', expand=True, pady=(0, 6))
        else:
            lf.pack(fill='x', pady=(0, 6))
        return lf

    # ── Station rows ──────────────────────────────────────────────────────

    def _add_station_row(self, preset_type: str = ''):
        row = StationRow(self._rows_frame, self._delete_station_row)
        if preset_type:
            row.var_type.set(preset_type)
        row.pack(fill='x', pady=1, padx=2)
        self._station_rows.append(row)

        def _nav(event, r=row):
            idx = self._station_rows.index(r)
            focused = event.widget
            delta = -1 if event.keysym == 'Up' else 1
            target_idx = idx + delta
            if 0 <= target_idx < len(self._station_rows):
                target = self._station_rows[target_idx]
                if focused is r._folder_entry:
                    target._folder_entry.focus_set()
                else:
                    target._type_entry.focus_set()
            return 'break'

        for widget in (row._type_entry, row._folder_entry):
            widget.bind('<Up>',   _nav)
            widget.bind('<Down>', _nav)

    def _delete_station_row(self, row):
        self._station_rows.remove(row)
        row.destroy()

    # ── Browse helpers ────────────────────────────────────────────────────

    def _browse_file(self, var, title, filetypes):
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _browse_dir(self, var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    # ── Log helpers ───────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        def _do():
            self._log.configure(state='normal')
            self._log.insert('end', f'[{_ts()}] {msg}\n')
            self._log.see('end')
            self._log.configure(state='disabled')
        self.frame.after(0, _do)

    def _clear_log(self):
        self._log.configure(state='normal')
        self._log.delete('1.0', 'end')
        self._log.configure(state='disabled')

    # ── Config persistence ────────────────────────────────────────────────

    def _load_config(self):
        import json
        try:
            with open(self._config_path, encoding='utf-8') as f:
                cfg = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return

        self._var_excel.set(cfg.get('excel_path', ''))
        self._var_outdir.set(cfg.get('out_dir', ''))

        stations = cfg.get('stations')
        if stations:
            for row in list(self._station_rows):
                row.destroy()
            self._station_rows.clear()
            for s in stations:
                self._add_station_row(preset_type=s.get('type', ''))
                self._station_rows[-1].var_folder.set(s.get('folder', ''))

        # Restore analysis mode
        saved_mode = cfg.get('cpk_mode', self._CPK_MODE_LABELS[0])
        if saved_mode in self._CPK_MODE_LABELS:
            self._mode_display_var.set(saved_mode)
        saved_level = cfg.get('fault_level', '基础版（规则库）')
        self._fault_level_var.set(saved_level)

        # Restore file type
        saved_type = cfg.get('file_type', 'xlsx')
        if saved_type in ('xlsx', 'json'):
            self._var_file_type.set(saved_type)

        # Restore merge rules
        merge_rules = cfg.get('merge_rules', [])
        for rule in merge_rules:
            self._add_merge_row(src=rule.get('src', ''), dst=rule.get('dst', ''))
        if merge_rules:
            # Auto-expand merge config if there are rules
            self._toggle_merge()

    def save_config(self):
        import json
        cfg = {
            'excel_path':   self._var_excel.get().strip(),
            'out_dir':      self._var_outdir.get().strip(),
            'stations':     [r.get() for r in self._station_rows],
            'cpk_mode':     self._mode_display_var.get(),
            'fault_level':  self._fault_level_var.get(),
            'file_type':    self._var_file_type.get(),
            'merge_rules':  [r.get() for r in self._merge_rows
                             if r.get()['src'] and r.get()['dst']],
        }
        try:
            with open(self._config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except OSError:
            pass

    # ── Progress helpers ──────────────────────────────────────────────────

    def _set_progress(self, pct: float, label: str = ''):
        def _do():
            self._progress_var.set(pct)
            display = f'[{pct:.0f}%]  {label}' if label else f'[{pct:.0f}%]'
            self._progress_label.configure(text=display)
        self.frame.after(0, _do)

    def _on_stop(self):
        self._stop_event.set()
        self._btn_run.configure(state='disabled')
        self._log_msg('[INFO] 正在中止分析，请稍候...')

    def _set_buttons(self, running: bool):
        def _do():
            if running:
                self._btn_run.configure(
                    text='停止分析', bg='#c62828', command=self._on_stop,
                    state='normal'
                )
            else:
                self._btn_run.configure(
                    text='开始分析', bg='#3949ab', command=self._on_run,
                    state='normal'
                )
        self.frame.after(0, _do)

    # ── Main run logic ────────────────────────────────────────────────────

    def _on_run(self):
        excel_path = self._var_excel.get().strip()
        out_dir = self._var_outdir.get().strip()

        # Determine CPK mode
        label = self._mode_display_var.get()
        try:
            cpk_mode = self._CPK_MODE_VALUES[self._CPK_MODE_LABELS.index(label)]
        except (ValueError, IndexError):
            cpk_mode = 'latest_pass'

        fault_enabled = cpk_mode in ('all_with_fail', 'fail_only')
        fault_level = self._fault_level_var.get()
        file_type = self._var_file_type.get()

        # Collect merge rules
        merge_rules = [r.get() for r in self._merge_rows
                       if r.get()['src'] and r.get()['dst']]

        # Validation
        if cpk_mode == 'latest_pass':
            if not excel_path or not os.path.isfile(excel_path):
                messagebox.showerror(
                    '错误',
                    '【最后一次pass数据】模式需要DUT条码Excel文件\n'
                    '请选择文件，或切换到其他不需要Excel的分析模式'
                )
                return
        elif excel_path and not os.path.isfile(excel_path):
            messagebox.showerror('错误', 'DUT条码Excel文件路径无效，请重新选择')
            return

        if not out_dir:
            messagebox.showerror('错误', '请选择输出目录')
            return

        station_configs = [r.get() for r in self._station_rows]
        station_configs = [c for c in station_configs if c['type'] and c['folder']]
        if not station_configs:
            messagebox.showerror('错误', '请至少配置一个测试工站（类型 + 文件夹）')
            return

        self.save_config()
        self._report_path = None
        self._stop_event.clear()
        self._set_buttons(running=True)
        self._set_progress(0, '正在准备...')
        self._clear_log()

        threading.Thread(
            target=self._run_analysis,
            args=(excel_path, out_dir, station_configs, cpk_mode,
                  fault_enabled, fault_level, file_type, merge_rules),
            daemon=True,
        ).start()

    def _run_analysis(self, excel_path: str, out_dir: str, station_configs: list,
                      cpk_mode: str = 'latest_pass',
                      fault_enabled: bool = False,
                      fault_level: str = '基础版（规则库）',
                      file_type: str = 'xlsx',
                      merge_rules: list = None):
        import time
        t_start = time.time()
        merge_rules = merge_rules or []

        def elapsed():
            return f'{time.time() - t_start:.1f}s'

        # Create a timestamped subdirectory for this run so each analysis
        # result set is isolated and never overwrites a previous run.
        # Prefix with product category (the folder directly under TestResult/).
        run_ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        product_prefix = _get_product_category(station_configs)
        if not product_prefix:
            self.frame.after(0, lambda: messagebox.showwarning(
                '未识别目录结构',
                '未在测试工站文件夹下找到 TestResult 目录，\n'
                '无法自动获取产品类别名称。\n\n'
                '输出文件夹将仅使用时间标签命名。\n'
                '（预期结构：工站文件夹 → TestResult → 产品类别 → 工站类型 → …）'
            ))
        run_folder = f'{product_prefix}_{run_ts}' if product_prefix else run_ts
        out_dir = os.path.join(out_dir, run_folder)
        os.makedirs(out_dir, exist_ok=True)
        log_filename = f'analysis_log_{run_ts}.txt'
        log_path = os.path.join(out_dir, log_filename)
        try:
            _log_file = open(log_path, 'w', encoding='utf-8')
        except OSError:
            _log_file = None

        def _log(msg: str):
            self._log_msg(msg)
            if _log_file:
                try:
                    _log_file.write(f'[{_ts()}] {msg}\n')
                    _log_file.flush()
                except OSError:
                    pass

        _MODE_DESC = {
            'latest_pass':   '最后一次pass数据',
            'folder_direct': '所选文件夹分析',
            'all_pass':      '全部成功数据',
            'all_with_fail': '所有数据（含失败）',
            'fail_only':     '仅失败数据',
        }

        try:
            _log(f'本次输出目录: {out_dir}')
            _log(f'日志文件: {log_path}')
            _log(f'分析模式: {_MODE_DESC.get(cpk_mode, cpk_mode)}'
                 + (f'  |  故障分析: {fault_level}' if fault_enabled else '')
                 + f'  |  文件类型: {file_type}')
            if merge_rules:
                merge_desc = ', '.join(f'{r["src"]}→{r["dst"]}' for r in merge_rules)
                _log(f'工站合并规则: {merge_desc}')

            # ──────────────────────────────────────────────────────────
            # MODE: folder_direct — traverse ALL configured directories
            # ──────────────────────────────────────────────────────────
            if cpk_mode == 'folder_direct':
                _log('=' * 56)
                _log('【所选文件夹分析】遍历所有配置工站目录（pass + fail 全量提取）')

                # Log all configured folders
                for cfg in station_configs:
                    folder = cfg.get('folder', '')
                    stype  = cfg.get('type', '')
                    if folder:
                        exists = '✓' if os.path.isdir(folder) else '✗ 不存在'
                        _log(f'  [{stype}] {folder}  [{exists}]')

                # ── Read barcode filter from DUT Excel (optional) ─────
                fd_barcodes = None
                if excel_path and os.path.isfile(excel_path):
                    try:
                        fd_barcodes = read_barcodes(excel_path)
                        _log(f'  [INFO] 从DUT条码Excel读取到 {len(fd_barcodes)} 个条码，'
                             f'仅分析这些条码')
                    except Exception as exc:
                        _log(f'  [WARN] 读取DUT条码Excel失败: {exc}，将分析所有条码')

                def _trav_progress(done, total, bc):
                    pct = 5 + 60 * done / max(total, 1)
                    self._set_progress(pct, f'遍历中 ({done}/{total}): {bc}')

                self._set_progress(5, '开始遍历工站目录...')
                extraction_summary, fail_data = run_extraction_traverse(
                    station_configs=station_configs,
                    output_base_dir=out_dir,
                    file_type=file_type,
                    log_cb=_log,
                    progress_cb=_trav_progress,
                    stop_event=self._stop_event,
                    barcodes=fd_barcodes,
                )

                if self._stop_event.is_set():
                    _log('[INFO] 分析已中止')
                    self._set_progress(0, '已中止')
                    return

                # ── CPK analysis on extracted files ───────────────────
                _log('\n' + '=' * 56)
                _log('【CPK 分析】')
                all_analysis = {}
                station_list = list(extraction_summary.keys())
                for idx, stype in enumerate(station_list):
                    if self._stop_event.is_set():
                        break
                    if file_type == 'json':
                        analysis_dir = extraction_summary[stype]['json_dir']
                        ext = '.json'
                    else:
                        analysis_dir = extraction_summary[stype]['xlsx_dir']
                        ext = '.xlsx'
                    try:
                        file_count = len([
                            f for f in os.listdir(analysis_dir)
                            if f.lower().endswith(ext)
                        ])
                    except OSError:
                        file_count = 0
                    _log(f'\n  工站 [{stype}]  —  共 {file_count} 个{ext}文件')
                    self._set_progress(
                        67 + 15 * idx / max(len(station_list), 1),
                        f'CPK 分析: {stype} ({idx+1}/{len(station_list)})'
                    )
                    if file_type == 'json':
                        station_result = analyze_json_folder(analysis_dir, log_cb=_log)
                    else:
                        station_result = analyze_xlsx_folder(analysis_dir, log_cb=_log)
                    if station_result:
                        all_analysis[stype] = station_result
                    else:
                        _log(f'  [WARN] 工站 [{stype}] 无可分析数据')

                # ── CPK HTML report ────────────────────────────────────
                self._set_progress(83, '生成 CPK HTML 报告...')
                _log('\n' + '=' * 56)
                _log('【生成CPK HTML报告】')
                report_path = os.path.join(out_dir, 'cpk_report.html')
                from collections import Counter as _Counter
                station_info = dict(_Counter(
                    c['type'] for c in station_configs if c['type'] and c['folder']
                ))
                product_name = _infer_product_name(station_configs)
                report_title = (f'{product_name}测试数据分析报告 - Zillnk'
                                if product_name else '测试数据分析报告 - Zillnk')
                generate_report(
                    analysis_data=all_analysis,
                    output_path=report_path,
                    title=report_title,
                    station_info=station_info,
                )
                self._report_path = report_path
                _log(f'  CPK 报告: {report_path}  '
                     f'({os.path.getsize(report_path) // 1024} KB)')

                # ── Check if there are any failures ────────────────────
                has_failures = any(
                    sdata.get('fail_barcodes') or sdata.get('never_pass_barcodes')
                    for sdata in fail_data.values()
                )

                if has_failures and not self._stop_event.is_set():
                    # ── 3-sheet fail Excel ─────────────────────────────
                    self._set_progress(88, '生成失败分析Excel...')
                    _log('\n' + '=' * 56)
                    _log('【生成失败分析Excel（3 Sheet）】')
                    fail_excel_path = os.path.join(
                        out_dir, 'folder_direct_fail_analysis.xlsx'
                    )
                    try:
                        generate_folder_direct_excel(
                            fail_data=fail_data,
                            output_path=fail_excel_path,
                            log_cb=_log,
                        )
                    except Exception as exc:
                        _log(f'  [ERROR] 失败分析Excel生成失败: {exc}')

                    # ── Failure HTML report ────────────────────────────
                    self._set_progress(93, '生成失败分析HTML报告...')
                    _log('\n' + '=' * 56)
                    _log('【生成失败分析HTML报告】')
                    fail_html_path = os.path.join(out_dir, 'fail_analysis_report.html')
                    try:
                        gen_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        generate_fail_report(
                            fail_data=fail_data,
                            output_path=fail_html_path,
                            title=product_name,
                            generated_at=gen_at,
                        )
                        _log(f'  失败分析报告: {fail_html_path}  '
                             f'({os.path.getsize(fail_html_path) // 1024} KB)')
                    except Exception as exc:
                        _log(f'  [ERROR] 失败分析HTML生成失败: {exc}')
                else:
                    _log('\n  [INFO] 无失败记录，跳过失败分析报告生成')

                # ── Comprehensive HTML report ──────────────────────────
                self._set_progress(96, '生成综合分析报告...')
                _log('\n' + '=' * 56)
                _log('【生成综合分析报告】')
                comp_path = os.path.join(out_dir, 'comprehensive_report.html')
                try:
                    generate_comprehensive_report(
                        analysis_data=all_analysis,
                        output_path=comp_path,
                        title=product_name,
                        generated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        fail_data=fail_data,
                    )
                    _log(f'  综合报告: {comp_path}  '
                         f'({os.path.getsize(comp_path) // 1024} KB)')
                except Exception as exc:
                    _log(f'  [ERROR] 综合报告生成失败: {exc}')
                    comp_path = report_path

                _log(f'\n【完成】总耗时: {elapsed()}')
                self._set_progress(100, f'完成！耗时 {elapsed()}')
                self.frame.after(800, lambda p=comp_path: webbrowser.open(
                    'file:///' + p.replace(os.sep, '/')
                ))
                return

            # ──────────────────────────────────────────────────────────
            # MODE: all_pass — direct walk, no barcode list needed
            # ──────────────────────────────────────────────────────────
            if cpk_mode == 'all_pass':
                _log('=' * 56)
                _log('【全部成功数据】直接遍历工站文件夹，提取全部 pass 记录')

                def _ap_progress(done, total, bc):
                    pct = 5 + 55 * done / max(total, 1)
                    self._set_progress(pct, f'提取中 ({done}/{total}): {bc}')

                self._set_progress(5, '开始遍历工站目录...')
                extraction_summary = run_extraction_all_pass(
                    station_configs=station_configs,
                    output_base_dir=out_dir,
                    log_cb=_log,
                    progress_cb=_ap_progress,
                    stop_event=self._stop_event,
                )
                if self._stop_event.is_set():
                    _log('[INFO] 分析已中止')
                    self._set_progress(0, '已中止')
                    return
                total_extracted = sum(len(v['results']) for v in extraction_summary.values())
                _log(f'\n  共提取 {total_extracted} 条 pass 记录')

                # ── Duplicate barcode report ───────────────────────────
                _log('\n' + '=' * 56)
                _log('【重复条码统计】生成重复测试条码报表')
                dup_path = os.path.join(out_dir, 'duplicate_barcodes.xlsx')
                try:
                    generate_duplicate_report(
                        summary=extraction_summary,
                        output_path=dup_path,
                        log_cb=_log,
                    )
                except Exception as exc:
                    _log(f'  [ERROR] 重复条码报表生成失败: {exc}')

                self._set_progress(62, f'提取完成，共 {total_extracted} 条记录')

            # ──────────────────────────────────────────────────────────
            # MODES: latest_pass / all_with_fail / fail_only
            # ──────────────────────────────────────────────────────────
            else:
                # ── Step 1: read / discover barcodes ──────────────────
                _log('=' * 56)
                if excel_path and cpk_mode == 'latest_pass':
                    _log('【第1步】读取DUT条码Excel')
                    _log(f'  文件: {excel_path}')
                    try:
                        barcodes = read_barcodes(excel_path)
                    except Exception as exc:
                        _log(f'  [ERROR] 读取条码失败: {exc}')
                        self._set_buttons(running=False)
                        self._set_progress(0, '失败 - 请检查 Excel 文件')
                        self.frame.after(0, lambda e=exc: messagebox.showerror(
                            '读取Excel失败',
                            f'无法读取DUT条码Excel文件：\n\n{e}\n\n请确认文件未被其他程序占用，且格式正确。'
                        ))
                        return
                else:
                    _log('【第1步】自动扫描工站目录发现条码')
                    from core.data_extractor import discover_barcodes
                    valid_folders = [c['folder'] for c in station_configs
                                     if os.path.isdir(c.get('folder', ''))]
                    barcodes = discover_barcodes(valid_folders)
                    _log(f'  自动发现条码: {len(barcodes)} 个')

                unique_bc = list(dict.fromkeys(barcodes))
                if len(unique_bc) < len(barcodes):
                    _log(
                        f'  [WARN] 发现重复条码: 原始 {len(barcodes)} 条 → '
                        f'去重后 {len(unique_bc)} 条'
                    )
                    barcodes = unique_bc
                _log(f'  条码总数: {len(barcodes)} 个')
                if barcodes:
                    _log(f'  样例: {barcodes[:3]} ...')

                from collections import Counter
                type_counts = Counter(c['type'] for c in station_configs if c['type'])
                _log(f'\n  工站配置: {len(station_configs)} 条记录，'
                     f'涉及类型: {dict(type_counts)}')
                for cfg in station_configs:
                    exists = '✓' if os.path.isdir(cfg['folder']) else '✗ 不存在'
                    _log(f'    [{cfg["type"]}] {cfg["folder"]}  [{exists}]')

                _log(f'  输出目录: {out_dir}')
                self._set_progress(5, f'读取到 {len(barcodes)} 个条码')

                # ── Step 2: extract files ──────────────────────────────
                _log('\n' + '=' * 56)
                _log(f'【第2步】遍历工站目录，提取测试记录（模式: {_MODE_DESC.get(cpk_mode)}）')

                def progress_cb(done, total, bc):
                    pct = 5 + 55 * done / max(total, 1)
                    self._set_progress(pct, f'提取中 ({done}/{total}): {bc}')

                # Map all_with_fail → all for data_extractor (extractor uses 'all')
                extractor_mode = 'all' if cpk_mode == 'all_with_fail' else cpk_mode

                extraction_summary = run_extraction(
                    barcodes=barcodes,
                    station_configs=station_configs,
                    output_base_dir=out_dir,
                    log_cb=_log,
                    progress_cb=progress_cb,
                    stop_event=self._stop_event,
                    mode=extractor_mode,
                )

                if self._stop_event.is_set():
                    _log('[INFO] 分析已中止')
                    self._set_progress(0, '已中止')
                    return

                # ── Step 3: missing barcodes report ───────────────────
                _log('\n' + '=' * 56)
                _log('【第3步】生成缺失条码汇总报表')
                missing_path = os.path.join(out_dir, 'missing_barcodes.xlsx')
                try:
                    generate_missing_report(
                        summary=extraction_summary,
                        output_path=missing_path,
                        log_cb=_log,
                    )
                    total_missing = sum(
                        sum(1 for r in info['results'] if r['status'] != 'success')
                        for info in extraction_summary.values()
                    )
                    if total_missing == 0:
                        _log('  所有条码均已成功提取，缺失报表为空')
                    else:
                        _log(
                            f'  [注意] 共 {total_missing} 个条码缺失/异常，'
                            f'详见: {missing_path}'
                        )
                except Exception as exc:
                    _log(f'  [ERROR] 缺失报表生成失败: {exc}')

                self._set_progress(62, '缺失报表已生成')

            # ── Step 4: CPK analysis ───────────────────────────────────
            _log('\n' + '=' * 56)
            _log('【第4步】CPK 分析')
            all_analysis = {}
            station_list = list(extraction_summary.keys())

            for idx, stype in enumerate(station_list):
                if self._stop_event.is_set():
                    _log('[INFO] CPK 分析已中止')
                    self._set_progress(0, '已中止')
                    return
                if file_type == 'json':
                    analysis_dir = extraction_summary[stype]['json_dir']
                    ext = '.json'
                else:
                    analysis_dir = extraction_summary[stype]['xlsx_dir']
                    ext = '.xlsx'
                try:
                    file_count = len([
                        f for f in os.listdir(analysis_dir)
                        if f.lower().endswith(ext)
                    ])
                except OSError:
                    file_count = 0

                _log(f'\n  工站 [{stype}]  —  共 {file_count} 个{ext}文件')
                self._set_progress(
                    62 + 28 * idx / max(len(station_list), 1),
                    f'CPK 分析: {stype} ({idx+1}/{len(station_list)})'
                )
                if file_type == 'json':
                    station_result = analyze_json_folder(analysis_dir, log_cb=_log)
                else:
                    station_result = analyze_xlsx_folder(analysis_dir, log_cb=_log)
                if station_result:
                    all_analysis[stype] = station_result
                else:
                    _log(f'  [WARN] 工站 [{stype}] 无可分析数据')

            if not all_analysis:
                _log('[WARN] 所有工站均无可分析的 xlsx 数据，HTML 报告将为空')

            # ── Step 5: generate HTML report ───────────────────────────
            _log('\n' + '=' * 56)
            _log('【第5步】生成 HTML 报告')
            self._set_progress(92, '生成 HTML 报告...')

            report_path = os.path.join(out_dir, 'cpk_report.html')
            from collections import Counter as _Counter
            station_info = dict(_Counter(
                c['type'] for c in station_configs if c['type'] and c['folder']
            ))
            product_name = _infer_product_name(station_configs)
            report_title = (f'{product_name}测试数据分析报告 - Zillnk'
                            if product_name else '测试数据分析报告 - Zillnk')
            generate_report(
                analysis_data=all_analysis,
                output_path=report_path,
                title=report_title,
                station_info=station_info,
            )

            report_kb = os.path.getsize(report_path) // 1024
            self._report_path = report_path
            _log(f'  HTML 报告: {report_path}  ({report_kb} KB)')

            # ── Comprehensive report ───────────────────────────────────
            _log('\n' + '=' * 56)
            _log('【生成综合分析报告】')
            comp_path = os.path.join(out_dir, 'comprehensive_report.html')
            try:
                generate_comprehensive_report(
                    analysis_data=all_analysis,
                    output_path=comp_path,
                    title=product_name,
                    generated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                )
                _log(f'  综合报告: {comp_path}  '
                     f'({os.path.getsize(comp_path) // 1024} KB)')
            except Exception as exc:
                _log(f'  [ERROR] 综合报告生成失败: {exc}')
                comp_path = report_path

            # ── Step 6: fault analysis (optional) ─────────────────────
            if fault_enabled and not self._stop_event.is_set():
                _log('\n' + '=' * 56)
                _log(f'【第6步】故障分析及定位关系库建立（{fault_level}）')
                self._set_progress(93, '故障分析中...')
                try:
                    from core.fault_analyzer import (run_fault_analysis,
                                                    generate_fault_barcode_list,
                                                    generate_rule_suggestions_yaml)
                    # Map all_with_fail → all for fault_analyzer (uses 'all' internally)
                    fa_mode = 'all' if cpk_mode == 'all_with_fail' else cpk_mode
                    fault_summary = run_fault_analysis(
                        station_configs=station_configs,
                        out_dir=out_dir,
                        level=fault_level,
                        mode=fa_mode,
                        log_cb=_log,
                        stop_event=self._stop_event,
                    )
                    for stype, stats in fault_summary.items():
                        if stype.startswith('__'):
                            continue
                        _log(
                            f'  [{stype}] 总 {stats["total"]} 条记录，'
                            f'已分类 {stats["classified"]}，'
                            f'未分类 {stats["unclassified"]}'
                        )
                    # Generate fault barcode list Excel + rule suggestions YAML
                    if not self._stop_event.is_set():
                        db_path = os.path.join(out_dir, 'fault_database.db')
                        fault_list_path = os.path.join(out_dir, 'fault_barcodes.xlsx')
                        generate_fault_barcode_list(
                            db_path=db_path,
                            output_path=fault_list_path,
                            log_cb=_log,
                        )
                        yaml_path = os.path.join(
                            out_dir,
                            f'rule_suggestions_{datetime.now().strftime("%Y%m%d_%H%M%S")}.yaml'
                        )
                        generate_rule_suggestions_yaml(
                            db_path=db_path,
                            output_path=yaml_path,
                            log_cb=_log,
                        )
                except Exception as exc:
                    import traceback as _tb
                    _log(f'  [ERROR] 故障分析失败: {exc}')
                    _log(_tb.format_exc())

            # ── Final summary ──────────────────────────────────────────
            _log('\n' + '=' * 56)
            _log(f'【完成】总耗时: {elapsed()}')
            for stype, info in extraction_summary.items():
                res = info['results']
                ok  = sum(1 for r in res if r['status'] == 'success')
                bad = len(res) - ok
                cpk_pts = sum(
                    len(pts)
                    for pts in all_analysis.get(stype, {}).values()
                )
                _log(
                    f'  [{stype}] 提取: {ok}/{len(res)} 成功，{bad} 缺失  |  '
                    f'CPK 子项: {cpk_pts}'
                )
            _log('=' * 56)
            _log(f'日志已保存: {log_path}')

            self._set_progress(100, f'完成！耗时 {elapsed()}')

            self.frame.after(800, lambda p=comp_path: webbrowser.open(
                'file:///' + p.replace(os.sep, '/')
            ))

        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            _log(f'[ERROR] 未预期的错误: {exc}')
            _log(tb)
            self._set_progress(0, '发生错误，请查看日志')
            self.frame.after(0, lambda e=exc, t=tb: messagebox.showerror(
                '分析出错',
                f'分析过程中发生错误：\n\n{e}\n\n详细信息已写入运行日志，点击"查看运行日志"可查看完整追踪。'
            ))
        finally:
            if _log_file:
                _log_file.close()
            self._set_buttons(running=False)


# ============================================================================
# Placeholder tabs for future modules
# ============================================================================

class PlaceholderTab:
    def __init__(self, notebook: ttk.Notebook, title: str, description: str):
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text=f'  {title}  ')

        outer = tk.Frame(self.frame, bg='#f0f2f5')
        outer.pack(fill='both', expand=True)

        tk.Label(outer, text=title,
                 font=('Segoe UI', 18, 'bold'),
                 bg='#f0f2f5', fg='#1a237e').pack(pady=(80, 12))

        tk.Label(outer, text=description,
                 font=('Segoe UI', 11), bg='#f0f2f5', fg='#555',
                 justify='center').pack()

        tk.Label(outer, text='（功能待实现）',
                 font=('Segoe UI', 10),
                 bg='#f0f2f5', fg='#aaa').pack(pady=(8, 0))


# ============================================================================
# Help window
# ============================================================================

def _show_help(root: tk.Tk):
    win = tk.Toplevel(root)
    win.title('使用帮助 — 产线数据分析AI平台')
    win.geometry('620x480')
    win.resizable(True, True)
    win.configure(bg='#f0f2f5')
    win.transient(root)
    win.grab_set()

    tk.Label(win, text='使用帮助', font=('Segoe UI', 12, 'bold'),
             bg='#1a237e', fg='white').pack(fill='x', ipady=6)

    txt = scrolledtext.ScrolledText(
        win, font=('Segoe UI', 9), wrap='word',
        bg='white', fg='#212121', relief='flat',
        padx=12, pady=8
    )
    txt.pack(fill='both', expand=True, padx=8, pady=8)
    txt.insert('1.0', _HELP_TEXT)
    txt.configure(state='disabled')

    tk.Button(win, text='关闭', command=win.destroy,
              font=('Segoe UI', 9), bg='#3949ab', fg='white',
              relief='flat', padx=16, pady=4).pack(pady=(0, 10))


# ============================================================================
# Fault rule import dialog
# ============================================================================

def _import_fault_rules(root: tk.Tk, config_path: str):
    """Load a YAML fault rules file and import into the fault database."""
    yaml_path = filedialog.askopenfilename(
        title='选择故障关系描述文件',
        filetypes=[('YAML文件', '*.yml *.yaml'), ('所有文件', '*.*')]
    )
    if not yaml_path:
        return

    # Find the output dir from app_config.json to locate the DB
    import json
    out_dir = ''
    try:
        with open(config_path, encoding='utf-8') as f:
            cfg = json.load(f)
        out_dir = cfg.get('out_dir', '')
    except Exception:
        pass

    if not out_dir:
        out_dir = filedialog.askdirectory(
            title='选择包含 fault_database.db 的输出目录'
        )
        if not out_dir:
            return

    db_path = os.path.join(out_dir, 'fault_database.db')
    if not os.path.isfile(db_path):
        messagebox.showerror(
            '错误',
            f'未找到故障数据库文件:\n{db_path}\n\n请先运行一次"所有数据"或"仅失败数据"分析模式。'
        )
        return

    try:
        # Simple YAML parser for our format (no external dependency)
        rules = _parse_fault_rules_yaml(yaml_path)
        if not rules:
            messagebox.showwarning('警告', '文件中未找到有效规则，请检查YAML格式')
            return

        from core.fault_db import init_db, get_rules, add_rule, update_rule
        init_db(db_path)
        existing = {r['keywords']: r for r in get_rules(db_path)}

        added, updated = 0, 0
        for rule in rules:
            kw = rule.get('keywords', '').strip()
            ft = rule.get('fault_type', '').strip()
            sg = rule.get('suggestion', '').strip()
            if not kw or not ft:
                continue
            if kw in existing:
                update_rule(db_path, existing[kw]['id'],
                            keywords=kw, fault_type=ft, suggestion=sg)
                updated += 1
            else:
                add_rule(db_path, kw, ft, sg)
                added += 1

        messagebox.showinfo(
            '导入完成',
            f'故障规则导入成功！\n\n新增: {added} 条\n更新: {updated} 条\n\n数据库: {db_path}'
        )
    except Exception as exc:
        messagebox.showerror('导入失败', f'导入故障规则时出错:\n{exc}')


def _parse_fault_rules_yaml(path: str) -> list:
    """
    Minimal YAML parser for fault rules files.
    Supports the format:
        rules:
          - keywords: "..."
            fault_type: "..."
            suggestion: "..."
    """
    rules = []
    current = {}
    in_rules = False

    with open(path, encoding='utf-8') as f:
        for line in f:
            stripped = line.rstrip()
            if not stripped or stripped.lstrip().startswith('#'):
                continue

            if stripped.strip() == 'rules:':
                in_rules = True
                continue

            if not in_rules:
                continue

            indent = len(line) - len(line.lstrip())

            if stripped.strip().startswith('- '):
                if current:
                    rules.append(current)
                current = {}
                # Handle inline first key: "- keywords: ..."
                rest = stripped.strip()[2:]
                if ':' in rest:
                    k, v = rest.split(':', 1)
                    current[k.strip()] = v.strip().strip('"').strip("'")
            elif ':' in stripped and indent > 0:
                k, v = stripped.strip().split(':', 1)
                current[k.strip()] = v.strip().strip('"').strip("'")

    if current:
        rules.append(current)

    return [r for r in rules if r.get('keywords') and r.get('fault_type')]


# ============================================================================
# Main application window
# ============================================================================

class CPKAnalysisPlatform:

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('产线数据分析AI平台 — Zillnk Efficiency Improvement Group')
        root.configure(bg='#1a237e')

        self._apply_style()
        self._build_menu()
        self._build_ui()

        # Auto-fit window height to content; keep width fixed at 1000
        root.update_idletasks()
        root.geometry(f'1000x{root.winfo_reqheight()}')
        root.minsize(820, root.winfo_reqheight())

        root.bind('<F11>', self._toggle_fullscreen)
        root.bind('<Escape>', lambda _e: root.attributes('-fullscreen', False))
        root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TNotebook', background='#1a237e', borderwidth=0)
        style.configure('TNotebook.Tab',
                        background='#3949ab', foreground='white',
                        padding=[12, 5], font=('Segoe UI', 9))
        style.map('TNotebook.Tab',
                  background=[('selected', '#f0f2f5')],
                  foreground=[('selected', '#1a237e')])
        style.configure('TFrame', background='#f0f2f5')
        style.configure('TProgressbar',
                        troughcolor='#e0e0e0', background='#3949ab',
                        thickness=10)

    def _build_menu(self):
        menubar = tk.Menu(self.root, bg='#1a237e', fg='white',
                          activebackground='#3949ab', activeforeground='white',
                          relief='flat')

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0,
                             bg='white', fg='#212121',
                             activebackground='#3949ab', activeforeground='white')
        tools_menu.add_command(
            label='加载故障关系描述文件…',
            command=lambda: _import_fault_rules(
                self.root,
                os.path.join(_ROOT, 'app_config.json')
            )
        )
        menubar.add_cascade(label=' 工具 ', menu=tools_menu)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0,
                            bg='white', fg='#212121',
                            activebackground='#3949ab', activeforeground='white')
        help_menu.add_command(
            label='使用帮助',
            command=lambda: _show_help(self.root)
        )
        help_menu.add_separator()
        help_menu.add_command(
            label='关于',
            command=lambda: messagebox.showinfo(
                '关于',
                '产线数据分析AI平台  v2.0\n\nZillnk Efficiency Improvement Group\n\n'
                '本地测试站数据分析、CPK计算、故障定位关系库建立。'
            )
        )
        menubar.add_cascade(label=' 帮助 ', menu=help_menu)
        self.root.configure(menu=menubar)

    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill='both', expand=True)

        self._local_tab = LocalAnalysisTab(nb)
        PlaceholderTab(nb, '深科技 MES 数据分析',
                       '从深科技 MES 导出的测试数据 CPK 分析\n支持批次、工站、产品型号多维度分析')
        PlaceholderTab(nb, '立讯 MES 数据分析',
                       '从立讯 MES 导出的测试数据 CPK 分析\n支持批次、工站、产品型号多维度分析')

    def _on_close(self):
        self._local_tab.save_config()
        self.root.destroy()

    def _toggle_fullscreen(self, _event=None):
        current = self.root.attributes('-fullscreen')
        self.root.attributes('-fullscreen', not current)


# ============================================================================
# Entry point
# ============================================================================

def main():
    root = tk.Tk()
    CPKAnalysisPlatform(root)
    root.mainloop()


if __name__ == '__main__':
    main()
