"""
data_extractor.py
-----------------
Traverses local test station folders, finds test records matching barcodes,
selects the latest successful test for each barcode, copies xlsx/json files
to the output directories, and generates a missing-barcodes Excel report.

Real-world folder layout observed:
  <station_root>/
    [X11|X11_X11|...]/          ← optional grouping sub-folder
      <barcode> or <bc1_bc2>/   ← barcode folder (name *contains* the barcode)
        <YYYYMMDDHHMMSS>/        ← one folder per test attempt
          Test_Result_<ts>_<bc>.xlsx
          <bc>_..._MEASUREMENT_...json
          ...
"""

import json as _json_mod
import os
import re
import shutil
import pandas as pd
from collections import defaultdict
from datetime import datetime


# ---------------------------------------------------------------------------
# Barcode input helpers
# ---------------------------------------------------------------------------

def read_barcodes(excel_path: str) -> list:
    """Read barcode list from 'PrdSN' column of the input Excel file."""
    df = pd.read_excel(excel_path)
    col_map = {c.strip().lower(): c for c in df.columns}
    key = col_map.get('prdsn')
    if key is None:
        raise ValueError(
            f"Input Excel has no 'PrdSN' column. Found: {list(df.columns)}"
        )
    barcodes = df[key].dropna().astype(str).str.strip().tolist()
    return [b for b in barcodes if b]


# ---------------------------------------------------------------------------
# Timestamp parsing
# ---------------------------------------------------------------------------

_TS_PATTERNS = [
    r'(?<!\d)(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?!\d)',
    r'(?<!\d)(\d{4})(\d{2})(\d{2})[_\-](\d{2})(\d{2})(\d{2})(?!\d)',
    r'(\d{4})[_\-](\d{2})[_\-](\d{2})[_ ](\d{2})[_\-](\d{2})[_\-](\d{2})',
    r'(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})',
]


def _parse_timestamp(name: str) -> datetime:
    for pattern in _TS_PATTERNS:
        m = re.search(pattern, name)
        if m:
            try:
                g = m.groups()
                return datetime(int(g[0]), int(g[1]), int(g[2]),
                                int(g[3]), int(g[4]), int(g[5]))
            except ValueError:
                continue
    return datetime.min


def _is_timestamp_folder(name: str) -> bool:
    return _parse_timestamp(name) != datetime.min


# ---------------------------------------------------------------------------
# xlsx helpers
# ---------------------------------------------------------------------------

def _open_excel_safe(path: str):
    try:
        return pd.ExcelFile(path)
    except Exception:
        return None


def is_test_successful(xlsx_path: str) -> bool:
    """Return True iff no 'fail' result in any sheet (case-insensitive)."""
    xl = _open_excel_safe(xlsx_path)
    if xl is None:
        return False
    try:
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if 'result' not in df.columns:
                continue
            results = df['result'].dropna().astype(str).str.strip().str.lower()
            if (results == 'fail').any():
                return False
        return True
    except Exception:
        return False


def get_earliest_start_time(xlsx_path: str) -> datetime:
    xl = _open_excel_safe(xlsx_path)
    if xl is None:
        return datetime.min
    try:
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if 'start_time' not in df.columns:
                continue
            times = pd.to_datetime(df['start_time'], errors='coerce').dropna()
            if not times.empty:
                return times.min().to_pydatetime()
    except Exception:
        pass
    return datetime.min


def _find_xlsx_for_barcode(folder: str, barcode: str):
    candidates = []
    try:
        for name in os.listdir(folder):
            if not name.lower().endswith('.xlsx'):
                continue
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                if barcode in name:
                    return full
                candidates.append(full)
    except OSError:
        pass
    return candidates[0] if candidates else None


def _find_measurement_json(folder: str, barcode: str):
    candidates = []
    try:
        for name in os.listdir(folder):
            if not name.lower().endswith('.json'):
                continue
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                if 'DutInfo' in name:
                    continue
                if barcode in name:
                    return full
                if 'MEASUREMENT' in name.upper():
                    candidates.append(full)
    except OSError:
        pass
    return candidates[0] if candidates else None


# ---------------------------------------------------------------------------
# Core discovery
# ---------------------------------------------------------------------------

def find_test_records(barcode: str, station_root: str) -> list:
    """
    Walk station_root and return all test records for barcode.

    Supports any layout where, at some depth:
      - A folder whose name contains the barcode (the "barcode folder")
      - Has timestamp-named (YYYYMMDDHHMMSS) subdirectories directly beneath it
      - Each timestamp subfolder contains the xlsx/json test files

    Standard structure (confirmed):
      <root>/TestResult/<product>/<station>/[variable middle levels]/<barcode>/<timestamp>/files

    The number of levels between <station> and <barcode> may vary per engineer.
    Legacy structure (also supported):
      <root>/[<grouping>/]<barcode_or_combined>/<timestamp>/files

    Rules:
      - Folders named "debug" (case-insensitive) are always skipped.
      - Non-data helper dirs (file_bk, RU1_Log_*, TM1_Log, ...) are skipped
        once we are inside a timestamp folder.
      - Once a barcode-match folder is found, we handle its timestamp
        subdirs directly and stop recursing deeper into it.

    Each record: (ts_folder_path, xlsx_path, json_path, is_success, test_time)
    xlsx_path may be None for json-only records.
    """
    records = []
    norm_root = os.path.normpath(station_root)

    # Folder names to prune (partial match prefix or exact name, lower-case)
    _SKIP_EXACT = {'debug', 'file_bk', 'env_comp', 'tm1_log'}
    _SKIP_PREFIX = ('ru1_log_', 'ru2_log_', 'tm1_log', 'tm2_log',
                    'gain flatness', 'tx aclr', 'peak power',
                    'rx sensitivity', 'efficency')

    def _should_skip_dir(name: str) -> bool:
        nl = name.lower()
        return nl in _SKIP_EXACT or any(nl.startswith(p) for p in _SKIP_PREFIX)

    for root, dirs, files in os.walk(norm_root):
        norm_root_r = os.path.normpath(root)
        rel = os.path.relpath(norm_root_r, norm_root)
        depth = 0 if rel == '.' else rel.count(os.sep) + 1

        # Hard depth limit — allow up to 10 to accommodate variable intermediate levels
        if depth > 10:
            dirs.clear()
            continue

        folder_name = os.path.basename(norm_root_r)

        # Always prune skip-list dirs before recursing
        dirs[:] = [d for d in dirs if not _should_skip_dir(d)]

        # If we are inside a timestamp folder, stop recursing (nothing useful deeper)
        if _is_timestamp_folder(folder_name):
            dirs.clear()
            continue

        # If the current folder name contains the barcode, process its
        # timestamp subdirs and stop descending into this folder.
        if barcode in folder_name:
            for ts_dir in list(dirs):
                if not _is_timestamp_folder(ts_dir):
                    continue
                ts_path = os.path.join(norm_root_r, ts_dir)
                xlsx = _find_xlsx_for_barcode(ts_path, barcode)
                json_f = _find_measurement_json(ts_path, barcode)

                if xlsx is not None:
                    success = is_test_successful(xlsx)
                    test_time = get_earliest_start_time(xlsx)
                elif json_f is not None:
                    success = _json_result_pass(json_f)
                    test_time = _json_start_time(json_f)
                else:
                    continue

                if test_time == datetime.min:
                    test_time = _parse_timestamp(ts_dir)

                records.append((ts_path, xlsx, json_f, success, test_time))

            # Do not descend further into this barcode folder
            dirs.clear()

    return records


# ---------------------------------------------------------------------------
# JSON helpers
# ---------------------------------------------------------------------------

def _json_result_pass(json_path: str) -> bool:
    for enc in ('utf-8', 'utf-8-sig', 'gbk', 'latin-1'):
        try:
            with open(json_path, encoding=enc) as f:
                data = _json_mod.load(f)
            result = data.get('DutInfo', {}).get('Result', 'Fail')
            return str(result).strip().lower() == 'pass'
        except UnicodeDecodeError:
            continue
        except Exception:
            return False
    return False


def _json_start_time(json_path: str) -> datetime:
    for enc in ('utf-8', 'utf-8-sig', 'gbk', 'latin-1'):
        try:
            with open(json_path, encoding=enc) as f:
                data = _json_mod.load(f)
            ts_str = data.get('DutInfo', {}).get('StartTime', '')
            if ts_str:
                return datetime.strptime(ts_str[:19], '%Y-%m-%d %H:%M:%S')
        except UnicodeDecodeError:
            continue
        except Exception:
            break
    return datetime.min


# ---------------------------------------------------------------------------
# Public extraction entry point
# ---------------------------------------------------------------------------

def run_extraction(
    barcodes: list,
    station_configs: list,
    output_base_dir: str,
    log_cb=None,
    progress_cb=None,
    stop_event=None,
) -> dict:
    """
    Run full extraction for all barcodes across all station types.

    Multiple configs with the same 'type' are merged: records from all
    folders are pooled, then the latest successful one wins.

    Result dict per barcode includes extra fields for the missing report:
      total_records, pass_records, latest_any_time, found_in, note
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    # Group folders by station type
    type_to_folders: dict = defaultdict(list)
    for cfg in station_configs:
        stype = cfg.get('type', '').strip()
        sfolder = cfg.get('folder', '').strip()
        if stype and sfolder:
            type_to_folders[stype].append(sfolder)

    summary = {}
    n_types = len(type_to_folders)
    total = len(barcodes) * n_types
    done = 0

    for stype, folders in type_to_folders.items():
        if stop_event and stop_event.is_set():
            _log(f"[INFO] 用户中止，跳过剩余工站")
            break
        valid_folders = [f for f in folders if os.path.isdir(f)]
        invalid_folders = [f for f in folders if not os.path.isdir(f)]

        if invalid_folders:
            for f in invalid_folders:
                _log(f"  [WARN] 工站 [{stype}] 文件夹不存在，已跳过: {f}")

        if not valid_folders:
            _log(f"[SKIP] 工站 [{stype}]: 无有效文件夹，跳过该工站")
            continue

        xlsx_out = os.path.join(output_base_dir, stype, 'xlsx')
        json_out = os.path.join(output_base_dir, stype, 'json')
        # Clear existing output dirs to avoid accumulating files from prior runs
        for _d in (xlsx_out, json_out):
            if os.path.exists(_d):
                shutil.rmtree(_d)
            os.makedirs(_d)

        _log(f"\n{'='*60}")
        _log(f"[INFO] 开始处理工站类型: [{stype}]")
        _log(f"[INFO] 共配置 {len(valid_folders)} 个文件夹:")
        for f in valid_folders:
            _log(f"       → {f}")
        _log(f"[INFO] 待查询条码: {len(barcodes)} 个")
        _log(f"[INFO] 输出目录: xlsx → {xlsx_out}")
        _log(f"{'='*60}")

        results = []
        not_found_n, no_pass_n, no_xlsx_n, ok_n, err_n = 0, 0, 0, 0, 0

        for bc in barcodes:
            all_records = []
            read_errors = []

            for folder in valid_folders:
                try:
                    recs = find_test_records(bc, folder)
                    all_records.extend(recs)
                except Exception as exc:
                    read_errors.append(f"{folder}: {exc}")

            if read_errors:
                for err in read_errors:
                    _log(f"  [ERROR] {bc} 读取文件夹时出错 — {err}")
                err_n += 1

            # Build summary fields
            total_recs = len(all_records)
            pass_recs = sum(1 for _, _, _, s, _ in all_records if s)
            latest_any = max(
                (t for _, _, _, _, t in all_records), default=datetime.min
            )
            latest_any_str = (
                latest_any.strftime('%Y-%m-%d %H:%M:%S')
                if latest_any != datetime.min else ''
            )

            if not all_records:
                not_found_n += 1
                _log(f"  [WARN] {bc} — 未找到任何测试记录")
                r = {
                    'status': 'not_found', 'barcode': bc,
                    'message': '未找到任何测试记录',
                    'xlsx': None, 'json': None,
                    'total_records': 0, 'pass_records': 0,
                    'latest_any_time': '', 'found_in': '', 'note': '',
                }

            else:
                # Records with xlsx that passed (needed for CPK)
                successful_xlsx = [
                    (p, x, j, t) for p, x, j, s, t in all_records
                    if s and x is not None
                ]
                # All pass records (including json-only)
                all_pass = [
                    (p, x, j, t) for p, x, j, s, t in all_records if s
                ]

                if not all_pass:
                    no_pass_n += 1
                    fail_n = total_recs - pass_recs
                    _log(
                        f"  [WARN] {bc} — 找到 {total_recs} 条记录，"
                        f"通过 {pass_recs} 条，失败 {fail_n} 条，"
                        f"无通过记录，最近测试: {latest_any_str}"
                    )
                    r = {
                        'status': 'no_pass', 'barcode': bc,
                        'message': f'找到{total_recs}条记录，均未全部通过',
                        'xlsx': None, 'json': None,
                        'total_records': total_recs, 'pass_records': pass_recs,
                        'latest_any_time': latest_any_str, 'found_in': '', 'note': '',
                    }

                elif not successful_xlsx:
                    no_xlsx_n += 1
                    _log(
                        f"  [WARN] {bc} — 找到 {total_recs} 条记录，"
                        f"{pass_recs} 条通过，但全为 json-only（无 xlsx），"
                        f"无法进行 CPK 分析"
                    )
                    r = {
                        'status': 'no_xlsx', 'barcode': bc,
                        'message': '通过记录存在但无xlsx，无法CPK分析',
                        'xlsx': None, 'json': None,
                        'total_records': total_recs, 'pass_records': pass_recs,
                        'latest_any_time': latest_any_str, 'found_in': '',
                        'note': '仅有json文件，缺少xlsx',
                    }

                else:
                    # Pick latest successful xlsx record
                    successful_xlsx.sort(key=lambda x: x[3], reverse=True)
                    ts_folder, src_xlsx, src_json, test_time = successful_xlsx[0]

                    # Determine which station folder this record came from
                    found_in = ''
                    for folder in valid_folders:
                        if os.path.normpath(folder) in os.path.normpath(ts_folder):
                            found_in = folder
                            break

                    dest_xlsx = os.path.join(xlsx_out, os.path.basename(src_xlsx))
                    copy_note = ''
                    try:
                        shutil.copy2(src_xlsx, dest_xlsx)
                    except Exception as exc:
                        dest_xlsx = None
                        copy_note = f'xlsx复制失败: {exc}'
                        _log(f"  [ERROR] {bc} xlsx 复制失败: {exc}")

                    dest_json = None
                    if src_json:
                        dest_json = os.path.join(
                            json_out, os.path.basename(src_json)
                        )
                        try:
                            shutil.copy2(src_json, dest_json)
                        except Exception as exc:
                            dest_json = None
                            _log(f"  [WARN] {bc} json 复制失败: {exc}")

                    has_json_note = '' if dest_json else '（无对应json）'
                    ok_n += 1
                    time_str = test_time.strftime('%Y-%m-%d %H:%M:%S')
                    multi_note = (
                        f"（共{total_recs}条记录，{pass_recs}条通过，取最新）"
                        if total_recs > 1 else ''
                    )
                    _log(
                        f"  [OK]   {bc}  测试时间:{time_str}"
                        f"{has_json_note}{multi_note}"
                    )
                    r = {
                        'status': 'success', 'barcode': bc, 'message': 'OK',
                        'xlsx': dest_xlsx, 'json': dest_json,
                        'total_records': total_recs, 'pass_records': pass_recs,
                        'latest_any_time': time_str, 'found_in': found_in,
                        'note': copy_note,
                    }

            results.append(r)
            done += 1
            if progress_cb:
                progress_cb(done, total, bc)

            if stop_event and stop_event.is_set():
                _log(f"  [INFO] 用户中止，已处理 {len(results)}/{len(barcodes)} 个条码")
                break

        missing_n = not_found_n + no_pass_n + no_xlsx_n
        _log(f"\n[汇总] 工站 [{stype}]  总条码: {len(barcodes)}")
        _log(f"       成功提取: {ok_n}  |  无通过记录: {no_pass_n}"
             f"  |  无xlsx: {no_xlsx_n}  |  未找到: {not_found_n}"
             f"  |  读取错误: {err_n}")
        _log(f"       缺失（需关注）: {missing_n} 个条码")

        summary[stype] = {
            'xlsx_dir': xlsx_out,
            'json_dir': json_out,
            'results': results,
        }

    return summary


# ---------------------------------------------------------------------------
# Missing barcodes Excel report
# ---------------------------------------------------------------------------

_STATUS_LABEL = {
    'not_found': '未找到测试记录',
    'no_pass':   '无通过测试记录',
    'no_xlsx':   '有通过记录但无xlsx',
    'error':     '读取异常',
}


def generate_missing_report(summary: dict, output_path: str, log_cb=None) -> str:
    """
    Generate an Excel report listing all non-success barcodes per station type.

    Each station type gets one sheet.  A summary row is appended at the bottom.
    Returns the absolute path of the written file.
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # remove default blank sheet

    # Colours
    HDR_FILL   = PatternFill('solid', fgColor='1A237E')
    HDR_FONT   = Font(color='FFFFFF', bold=True, size=10)
    WARN_FILL  = PatternFill('solid', fgColor='FFF3E0')
    ERR_FILL   = PatternFill('solid', fgColor='FFEBEE')
    SUM_FILL   = PatternFill('solid', fgColor='E8EAF6')
    SUM_FONT   = Font(bold=True, size=10)
    THIN       = Side(style='thin', color='BBBBBB')
    BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CENTER     = Alignment(horizontal='center', vertical='center')
    LEFT       = Alignment(horizontal='left',   vertical='center', wrap_text=True)

    COLUMNS = [
        ('PrdSN',       18),
        ('状态',         18),
        ('总测试次数',    10),
        ('通过次数',      10),
        ('最新测试时间',  20),
        ('备注',         40),
    ]

    total_missing_all = 0

    for stype, info in summary.items():
        results   = info['results']
        total_bc  = len(results)
        ok_n      = sum(1 for r in results if r['status'] == 'success')
        missing   = [r for r in results if r['status'] != 'success']
        total_missing_all += len(missing)

        ws = wb.create_sheet(title=stype[:31])   # sheet name max 31 chars

        # ── Header row ──────────────────────────────────────────────────
        for col_idx, (col_name, col_w) in enumerate(COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill      = HDR_FILL
            cell.font      = HDR_FONT
            cell.alignment = CENTER
            cell.border    = BORDER
            ws.column_dimensions[get_column_letter(col_idx)].width = col_w

        ws.row_dimensions[1].height = 20
        ws.freeze_panes = 'A2'

        # ── Data rows ────────────────────────────────────────────────────
        for row_idx, r in enumerate(missing, 2):
            status_label = _STATUS_LABEL.get(r['status'], r['status'])
            row_fill = ERR_FILL if r['status'] == 'not_found' else WARN_FILL

            values = [
                r['barcode'],
                status_label,
                r.get('total_records', ''),
                r.get('pass_records', ''),
                r.get('latest_any_time', ''),
                r.get('note', '') or r.get('message', ''),
            ]
            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.fill      = row_fill
                cell.border    = BORDER
                cell.alignment = CENTER if col_idx != len(values) else LEFT
                cell.font      = Font(size=10)

        # ── Summary row ──────────────────────────────────────────────────
        sum_row = len(missing) + 2
        ws.cell(row=sum_row, column=1, value='汇总').font = SUM_FONT
        ws.cell(row=sum_row, column=1).fill = SUM_FILL
        ws.cell(row=sum_row, column=1).alignment = CENTER
        summary_text = (
            f"总条码: {total_bc}  |  成功提取: {ok_n}"
            f"  |  缺失/异常: {len(missing)}"
        )
        cell = ws.cell(row=sum_row, column=2, value=summary_text)
        cell.font = SUM_FONT
        cell.fill = SUM_FILL
        cell.alignment = LEFT
        ws.merge_cells(
            start_row=sum_row, start_column=2,
            end_row=sum_row, end_column=len(COLUMNS)
        )
        for col_idx in range(1, len(COLUMNS) + 1):
            ws.cell(row=sum_row, column=col_idx).border = BORDER

        _log(f"  [缺失报表] 工站 [{stype}]: 总{total_bc}条  成功{ok_n}"
             f"  缺失/异常{len(missing)}条")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    _log(f"\n[INFO] 缺失条码报表已保存: {output_path}"
         f"  (共 {total_missing_all} 个缺失/异常条码)")
    return output_path