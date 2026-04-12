"""
main.py
-------
CPK 分析平台 v1.0 – Zillnk
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

from core.data_extractor import read_barcodes, run_extraction, generate_missing_report
from core.cpk_calculator import analyze_xlsx_folder
from core.html_report import generate_report


# ============================================================================
# Helpers
# ============================================================================

def _ts() -> str:
    return datetime.now().strftime('%H:%M:%S')


_HELP_TEXT = """\
CPK 分析平台 v1.0  –  使用说明
=====================================

【功能一：本地数据分析】

1. 发货 Excel 文件
   选择包含 "PrdSN" 列的发货清单 Excel 文件。
   程序将逐行读取其中的模块条码进行分析。

2. 输出目录
   选择分析结果的存放目录。程序将自动创建以下内容：
     <输出目录>/
       <工站类型>/xlsx/          ← 提取出的测试 xlsx 文件（保留原文件名）
       <工站类型>/json/          ← 对应的测试 json 文件（保留原文件名）
       missing_barcodes.xlsx    ← 缺失条码汇总报表
       cpk_report.html          ← CPK 分析 HTML 报告
       analysis_log_<时间戳>.txt ← 本次运行完整过程日志

3. 测试工站配置
   为每类测试工站填写：
     · 工站类型标签（如 FT1、FT2、Aging …）
     · 测试数据文件夹路径（该工站的数据根目录）
   同一工站类型可添加多行（对应多台同类设备，程序自动合并处理）。
   点击 [+ 添加工站] 可增加配置行。
   在工站类型或路径输入框内，按 ↑ / ↓ 可快速在行间切换焦点。

4. 开始分析 / 停止分析
   点击"开始分析"后按键变为"停止分析"，可随时中止当前运行。
   分析流程：
     ① 读取条码列表
     ② 遍历所有工站目录，提取每个条码最新一次成功测试的 xlsx/json
        （自动跳过 debug/、file_bk/ 等辅助目录）
     ③ 生成缺失条码报表 missing_barcodes.xlsx
     ④ 对每个工站的 xlsx 数据进行 CPK 计算
        （非数值型子项、固定值子项、样本量不足子项自动跳过）
     ⑤ 生成 HTML 报告（自动在浏览器中打开）

5. 配置记忆
   程序关闭或开始分析时自动保存当前配置（发货文件路径、输出目录、
   工站列表）到同目录下的 app_config.json，下次启动自动恢复。

6. HTML 报告说明
   · 顶部显示工站种类数及各类工站台数
   · 默认仅显示统计指标（均值、标准差、上下限等）
   · 在搜索框输入 "cpk" 可展示 Cp/Cpl/Cpu/Cpk 列（内部使用）
   · 正态分布图始终显示 LSL/USL 限值线（即使数据落在限值外）
   · 点击表格行可切换下方正态分布图
   · 数值范围搜索：选择测试子项 + 输入范围，查询命中条码
     最小值 = 最大值时为精确匹配，结果按值排序并标注"精确匹配"

【功能二 / 三】深科技 / 立讯 MES 数据分析
   功能待实现，敬请期待。

如有问题，请联系 Zillnk 质量工程部。
"""


# ============================================================================
# Station row widget (used in the station config list)
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

        self._folder_entry = tk.Entry(self.frame, textvariable=self.var_folder, width=44,
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
# Local Analysis Tab
# ============================================================================

class LocalAnalysisTab:
    """Tab 1: 测试站本地测试数据分析"""

    def __init__(self, notebook: ttk.Notebook):
        self.frame = ttk.Frame(notebook)
        notebook.add(self.frame, text='  本地数据分析  ')

        self._report_path = None
        self._station_rows = []
        self._stop_event = threading.Event()
        self._config_path = os.path.join(_ROOT, 'app_config.json')

        self._build_ui()
        self._load_config()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self):
        # Root container that fills the entire tab and expands on resize
        outer = tk.Frame(self.frame, bg='#f0f2f5')
        outer.pack(fill='both', expand=True, padx=10, pady=6)

        # ── Section 1: Inputs ─────────────────────────────────────────
        sec1 = self._make_section(outer, '输入 / 输出配置')

        # Row 1: Excel file  +  Output dir
        inp_row = tk.Frame(sec1, bg='white')
        inp_row.pack(fill='x', pady=(2, 3))

        tk.Label(inp_row, text='发货Excel:', width=8, anchor='w',
                 bg='white', font=('Segoe UI', 9)).pack(side='left')
        self._var_excel = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_excel,
                 font=('Segoe UI', 9), width=30).pack(side='left', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_file(
                      self._var_excel, '选择发货 Excel',
                      [('Excel', '*.xlsx *.xls')]),
                  bg='#e0e4f0', relief='flat', padx=5).pack(side='left', padx=(0, 14))

        tk.Label(inp_row, text='输出目录:', width=8, anchor='w',
                 bg='white', font=('Segoe UI', 9)).pack(side='left')
        self._var_outdir = tk.StringVar()
        tk.Entry(inp_row, textvariable=self._var_outdir,
                 font=('Segoe UI', 9), width=30).pack(side='left', padx=(0, 3))
        tk.Button(inp_row, text='浏览…', font=('Segoe UI', 8),
                  command=lambda: self._browse_dir(self._var_outdir, '选择输出目录'),
                  bg='#e0e4f0', relief='flat', padx=5).pack(side='left')


        # ── Section 2: Station config (scrollable canvas) ─────────────
        sec2 = self._make_section(outer, '测试工站配置')

        # Column header
        hdr = tk.Frame(sec2, bg='#dde3f0')
        hdr.pack(fill='x', pady=(0, 2))
        tk.Label(hdr, text='工站类型', width=10, anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4, pady=1)
        tk.Label(hdr, text='测试数据文件夹路径', anchor='w',
                 bg='#dde3f0', font=('Segoe UI', 8, 'bold')).pack(side='left', padx=4)

        # Scrollable area
        scroll_wrap = tk.Frame(sec2, bg='#f5f6fa')
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

        # Keep inner frame width = canvas width; update scroll region on resize
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

        # Mouse-wheel scroll
        def _on_wheel(e):
            self._station_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

        self._station_canvas.bind('<MouseWheel>', _on_wheel)
        self._rows_frame.bind('<MouseWheel>', _on_wheel)

        # Default station rows
        for stype in ('FT1', 'FT2', 'Aging'):
            self._add_station_row(preset_type=stype)

        tk.Button(sec2, text='＋ 添加工站', command=self._add_station_row,
                  font=('Segoe UI', 8), bg='#e8f5e9', relief='flat',
                  padx=6, pady=2).pack(anchor='w', pady=(4, 0))

        # ── Section 3: Actions + progress ────────────────────────────
        sec3 = self._make_section(outer, '操作')

        btn_row = tk.Frame(sec3, bg='white')
        btn_row.pack(fill='x', pady=(0, 4))

        self._btn_run = tk.Button(btn_row, text='开始分析',
                                  font=('Segoe UI', 9, 'bold'),
                                  bg='#3949ab', fg='white',
                                  relief='flat', padx=12, pady=4,
                                  command=self._on_run)
        self._btn_run.pack(side='left')

        self._progress_label = tk.Label(sec3, text='就绪', anchor='w',
                                        font=('Segoe UI', 8), bg='white', fg='#555')
        self._progress_label.pack(fill='x')

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(sec3, variable=self._progress_var,
                                             maximum=100, mode='determinate')
        self._progress_bar.pack(fill='x', pady=(2, 0))

        # ── Section 4: Log (expands to fill remaining space) ──────────
        sec4 = self._make_section(outer, '运行日志', expand=True)

        self._log = scrolledtext.ScrolledText(
            sec4, font=('Consolas', 8),
            bg='#1e1e2e', fg='#a8d8a8', insertbackground='white',
            state='disabled', wrap='word'
        )
        self._log.pack(fill='both', expand=True)

        tk.Button(sec4, text='清空日志',
                  font=('Segoe UI', 8), bg='#f5f5f5',
                  relief='flat', padx=6,
                  command=self._clear_log).pack(anchor='e', pady=(3, 0))

    # ── Section helper ───────────────────────────────────────────────────

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

    # ── Station rows ─────────────────────────────────────────────────────

    def _add_station_row(self, preset_type: str = ''):
        row = StationRow(self._rows_frame, self._delete_station_row)
        if preset_type:
            row.var_type.set(preset_type)
        row.pack(fill='x', pady=1, padx=2)
        self._station_rows.append(row)

        # ↑ / ↓ navigate between rows; focus the same column (type or folder)
        def _nav(event, r=row):
            idx = self._station_rows.index(r)
            focused = event.widget
            delta = -1 if event.keysym == 'Up' else 1
            target_idx = idx + delta
            if 0 <= target_idx < len(self._station_rows):
                target = self._station_rows[target_idx]
                # Stay in the same column
                if focused is r._folder_entry:
                    target._folder_entry.focus_set()
                else:
                    target._type_entry.focus_set()
            return 'break'   # prevent default cursor movement

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
        """Thread-safe log append."""
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
        """Restore last session's input paths and station list from JSON."""
        import json
        try:
            with open(self._config_path, encoding='utf-8') as f:
                cfg = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return   # first run or corrupt file — keep defaults

        self._var_excel.set(cfg.get('excel_path', ''))
        self._var_outdir.set(cfg.get('out_dir', ''))

        stations = cfg.get('stations')
        if stations:
            # Replace default rows with saved ones
            for row in list(self._station_rows):
                row.destroy()
            self._station_rows.clear()
            for s in stations:
                self._add_station_row(preset_type=s.get('type', ''))
                self._station_rows[-1].var_folder.set(s.get('folder', ''))

    def save_config(self):
        """Persist current inputs and station list to JSON."""
        import json
        cfg = {
            'excel_path': self._var_excel.get().strip(),
            'out_dir':    self._var_outdir.get().strip(),
            'stations':   [r.get() for r in self._station_rows],
        }
        try:
            with open(self._config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except OSError:
            pass   # non-critical — silently ignore write errors

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

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror('错误', '请选择有效的发货 Excel 文件')
            return
        if not out_dir:
            messagebox.showerror('错误', '请选择输出目录')
            return

        station_configs = [r.get() for r in self._station_rows]
        station_configs = [c for c in station_configs
                           if c['type'] and c['folder']]
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
            args=(excel_path, out_dir, station_configs),
            daemon=True,
        ).start()

    def _run_analysis(self, excel_path: str, out_dir: str, station_configs: list):
        import time
        t_start = time.time()

        def elapsed():
            return f'{time.time() - t_start:.1f}s'

        # ── Open per-run log file ────────────────────────────────────────
        os.makedirs(out_dir, exist_ok=True)
        log_filename = f'analysis_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
        log_path = os.path.join(out_dir, log_filename)
        try:
            _log_file = open(log_path, 'w', encoding='utf-8')
        except OSError:
            _log_file = None

        def _log(msg: str):
            """Write to GUI log and to the run log file simultaneously."""
            self._log_msg(msg)
            if _log_file:
                try:
                    _log_file.write(f'[{_ts()}] {msg}\n')
                    _log_file.flush()
                except OSError:
                    pass

        try:
            _log(f'日志文件: {log_path}')

            # ── Step 1: read barcodes ─────────────────────────────────
            _log('=' * 56)
            _log('【第1步】读取发货 Excel 条码列表')
            _log(f'  文件: {excel_path}')
            try:
                barcodes = read_barcodes(excel_path)
            except Exception as exc:
                _log(f'  [ERROR] 读取条码失败: {exc}')
                self._set_buttons(running=False)
                self._set_progress(0, '失败 - 请检查 Excel 文件')
                return

            unique_bc = list(dict.fromkeys(barcodes))   # preserve order, dedupe
            if len(unique_bc) < len(barcodes):
                _log(
                    f'  [WARN] 发现重复条码: 原始 {len(barcodes)} 条 → '
                    f'去重后 {len(unique_bc)} 条'
                )
                barcodes = unique_bc
            _log(f'  条码总数: {len(barcodes)} 个')
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

            # ── Step 2: extract files ─────────────────────────────────
            _log('\n' + '=' * 56)
            _log('【第2步】遍历工站目录，提取最新成功测试记录')

            def progress_cb(done, total, bc):
                pct = 5 + 55 * done / max(total, 1)
                self._set_progress(pct, f'提取中 ({done}/{total}): {bc}')

            extraction_summary = run_extraction(
                barcodes=barcodes,
                station_configs=station_configs,
                output_base_dir=out_dir,
                log_cb=_log,
                progress_cb=progress_cb,
                stop_event=self._stop_event,
            )

            if self._stop_event.is_set():
                _log('[INFO] 分析已中止')
                self._set_progress(0, '已中止')
                return

            # ── Step 3: generate missing barcodes Excel ───────────────
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

            # ── Step 4: CPK analysis ──────────────────────────────────
            _log('\n' + '=' * 56)
            _log('【第4步】CPK 分析')
            all_analysis = {}
            station_list = list(extraction_summary.keys())

            for idx, stype in enumerate(station_list):
                if self._stop_event.is_set():
                    _log('[INFO] CPK 分析已中止')
                    self._set_progress(0, '已中止')
                    return
                xlsx_dir = extraction_summary[stype]['xlsx_dir']
                try:
                    xlsx_count = len([
                        f for f in os.listdir(xlsx_dir)
                        if f.lower().endswith('.xlsx')
                    ])
                except OSError:
                    xlsx_count = 0

                _log(f'\n  工站 [{stype}]  —  共 {xlsx_count} 个xlsx文件')
                self._set_progress(
                    62 + 28 * idx / max(len(station_list), 1),
                    f'CPK 分析: {stype} ({idx+1}/{len(station_list)})'
                )
                station_result = analyze_xlsx_folder(xlsx_dir, log_cb=_log)
                if station_result:
                    all_analysis[stype] = station_result
                else:
                    _log(f'  [WARN] 工站 [{stype}] 无可分析数据')

            if not all_analysis:
                _log('[WARN] 所有工站均无可分析的 xlsx 数据，HTML 报告将为空')

            # ── Step 5: generate HTML report ──────────────────────────
            _log('\n' + '=' * 56)
            _log('【第5步】生成 HTML 报告')
            self._set_progress(92, '生成 HTML 报告...')

            report_path = os.path.join(out_dir, 'cpk_report.html')
            # Count configured folders per station type for the HTML header
            from collections import Counter as _Counter
            station_info = dict(_Counter(
                c['type'] for c in station_configs if c['type'] and c['folder']
            ))
            generate_report(
                analysis_data=all_analysis,
                output_path=report_path,
                station_info=station_info,
            )

            report_kb = os.path.getsize(report_path) // 1024
            self._report_path = report_path
            _log(f'  HTML 报告: {report_path}  ({report_kb} KB)')

            # ── Final summary ─────────────────────────────────────────
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

            self.frame.after(800, lambda: webbrowser.open(
                'file:///' + report_path.replace(os.sep, '/')
            ))

        except Exception as exc:
            import traceback
            _log(f'[ERROR] 未预期的错误: {exc}')
            _log(traceback.format_exc())
            self._set_progress(0, '发生错误，请查看日志')
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
    win.title('使用帮助 — CPK 分析平台')
    win.geometry('620x480')
    win.resizable(True, True)
    win.configure(bg='#f0f2f5')

    # Make it float above the main window
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
# Main application window
# ============================================================================

class CPKAnalysisPlatform:

    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('CPK 分析平台 v1.0 — Zillnk Quality Engineering')
        root.geometry('1000x720')
        root.minsize(820, 560)
        root.configure(bg='#1a237e')

        self._apply_style()
        self._build_menu()
        self._build_ui()

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
                'CPK 分析平台 v1.0\n\nZillnk Quality Engineering\n\n'
                '用于本地测试站数据 CPK 分析与质量管控。'
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