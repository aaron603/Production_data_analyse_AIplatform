"""
cpk_calculator.py
-----------------
Reads all xlsx/json files from a folder and calculates CPK statistics per
sheet / point_name, using the most recently dated file's limits as the
unified LSL/USL.
"""

import json as _json_mod
import os
import re
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path


# ---------------------------------------------------------------------------
# Core CPK math
# ---------------------------------------------------------------------------

def calculate_cpk(values: list, lsl=None, usl=None) -> dict:
    """
    Calculate basic statistics and CPK indices.

    Returns a dict with keys:
      n, mean, std, min, max, lsl, usl, cp, cpl, cpu, cpk
    All float/int values; CPK fields are None when not computable.
    """
    arr = np.array(values, dtype=float)
    n = int(len(arr))

    if n < 2:
        return _empty_stats(n, arr, lsl, usl)

    mean = float(np.mean(arr))
    std = float(np.std(arr, ddof=1))

    result = {
        'n': n,
        'mean': round(mean, 6),
        'std': round(std, 6),
        'min': round(float(np.min(arr)), 6),
        'max': round(float(np.max(arr)), 6),
        'lsl': lsl,
        'usl': usl,
        'cp': None,
        'cpl': None,
        'cpu': None,
        'cpk': None,
    }

    if std == 0:
        return result   # Fixed-value item – skip CPK

    if lsl is not None and usl is not None:
        result['cp'] = round((usl - lsl) / (6.0 * std), 4)

    if lsl is not None:
        result['cpl'] = round((mean - lsl) / (3.0 * std), 4)

    if usl is not None:
        result['cpu'] = round((usl - mean) / (3.0 * std), 4)

    candidates = [v for v in [result['cpl'], result['cpu']] if v is not None]
    if candidates:
        result['cpk'] = round(min(candidates), 4)

    return result


def _empty_stats(n, arr, lsl, usl) -> dict:
    return {
        'n': n,
        'mean': round(float(np.mean(arr)), 6) if n > 0 else None,
        'std': None,
        'min': round(float(np.min(arr)), 6) if n > 0 else None,
        'max': round(float(np.max(arr)), 6) if n > 0 else None,
        'lsl': lsl,
        'usl': usl,
        'cp': None, 'cpl': None, 'cpu': None, 'cpk': None,
    }


# ---------------------------------------------------------------------------
# Folder-level analysis
# ---------------------------------------------------------------------------

def _get_file_time(xl: pd.ExcelFile) -> datetime:
    """Return the minimum start_time found in this ExcelFile, or datetime.min."""
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet)
            if 'start_time' not in df.columns:
                continue
            times = pd.to_datetime(df['start_time'], errors='coerce').dropna()
            if not times.empty:
                return times.min().to_pydatetime()
        except Exception:
            pass
    return datetime.min


def _file_time_from_name(stem: str) -> datetime:
    """
    Parse the 14-digit timestamp embedded in filenames like
    Test_Result_<YYYYMMDDHHMMSS>_<barcode>.
    Returns datetime.min if no such pattern is found.
    """
    m = re.search(r'_(\d{14})(?:_|$)', stem)
    if m:
        s = m.group(1)
        try:
            return datetime(int(s[0:4]), int(s[4:6]), int(s[6:8]),
                            int(s[8:10]), int(s[10:12]), int(s[12:14]))
        except ValueError:
            pass
    return datetime.min


def analyze_xlsx_folder(xlsx_folder: str, log_cb=None) -> dict:
    """
    Analyze all xlsx files in `xlsx_folder`.

    For each sheet × point_name combination:
      - Collect (barcode, value) pairs across all files
      - Use limits from the most recently timed file as LSL/USL
      - Skip items where all values are identical (std == 0)

    Returns:
    {
      sheet_name: {
        point_name: {
          'n': int,
          'mean': float,
          'std': float,
          'min': float,
          'max': float,
          'lsl': float | None,
          'usl': float | None,
          'cp': float | None,
          'cpl': float | None,
          'cpu': float | None,
          'cpk': float | None,
          'values': [(barcode, value), ...],   # raw data for distribution chart
        }
      }
    }
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    folder = Path(xlsx_folder)
    xlsx_files = sorted(folder.glob('*.xlsx'))

    if not xlsx_files:
        _log(f"  [WARN] CPK分析：{xlsx_folder} 目录中未找到 xlsx 文件")
        return {}

    _log(f"  [INFO] CPK分析目录: {xlsx_folder}")
    _log(f"  [INFO] 找到 {len(xlsx_files)} 个 xlsx 文件，开始读取数据...")

    # ── Pass 1: collect data ──────────────────────────────────────────────
    # collected[sheet][point_name] = {
    #   'values': [(barcode, float), ...],
    #   'latest_time': datetime,
    #   'lsl': float | None,
    #   'usl': float | None,
    # }
    collected: dict = {}

    for xlsx_path in xlsx_files:
        # Extract barcode from file name.
        # Standard pattern: Test_Result_<YYYYMMDDHHMMSS>_<barcode>.xlsx
        # The barcode follows the 14-digit timestamp segment.
        stem = xlsx_path.stem
        _m = re.search(r'_(\d{14})_(.+)$', stem)
        if _m:
            barcode = _m.group(2)
        else:
            # Fallback: last underscore-separated segment
            barcode = stem.rsplit('_', 1)[-1] if '_' in stem else stem

        xl = None
        try:
            xl = pd.ExcelFile(xlsx_path)
        except Exception as exc:
            _log(f"  [ERROR] 无法打开文件 {xlsx_path.name}: {exc}")
            continue

        # Prefer filename-embedded timestamp (reliable under parallel test stations).
        # Fall back to content start_time only if filename carries no timestamp.
        file_time = _file_time_from_name(xlsx_path.stem)
        if file_time == datetime.min:
            file_time = _get_file_time(xl)

        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet)
            except Exception as exc:
                _log(f"  [ERROR] {xlsx_path.name} 读取 Sheet[{sheet}] 失败: {exc}")
                continue

            if df.empty:
                continue

            required = {'point_name', 'data'}
            if not required.issubset(df.columns):
                missing_cols = required - set(df.columns)
                _log(f"  [WARN] {xlsx_path.name} Sheet[{sheet}] 缺少必要列: {missing_cols}，已跳过")
                continue

            if sheet not in collected:
                collected[sheet] = {}

            for _, row in df.iterrows():
                pname = str(row.get('point_name', '')).strip()
                if not pname or pname.lower() == 'nan':
                    continue

                raw_val = row.get('data')
                try:
                    val = float(raw_val)
                except (TypeError, ValueError):
                    continue

                # Row-level pass/fail status (from 'result' column if present)
                row_result = str(row.get('result', 'pass')).strip().lower()
                row_pass = (row_result != 'fail')

                if pname not in collected[sheet]:
                    collected[sheet][pname] = {
                        'values': [],
                        'latest_time': datetime.min,
                        'lsl': None,
                        'usl': None,
                    }

                # values stores (barcode, measurement_value, is_pass)
                collected[sheet][pname]['values'].append((barcode, val, row_pass))

                # Update limits only if this file is newer
                if file_time >= collected[sheet][pname]['latest_time']:
                    collected[sheet][pname]['latest_time'] = file_time

                    raw_lsl = row.get('limit_low')
                    raw_usl = row.get('limit_high')

                    try:
                        collected[sheet][pname]['lsl'] = (
                            float(raw_lsl) if pd.notna(raw_lsl) else None
                        )
                    except (TypeError, ValueError):
                        collected[sheet][pname]['lsl'] = None

                    try:
                        collected[sheet][pname]['usl'] = (
                            float(raw_usl) if pd.notna(raw_usl) else None
                        )
                    except (TypeError, ValueError):
                        collected[sheet][pname]['usl'] = None

    # ── Pass 2: calculate CPK ─────────────────────────────────────────────
    results: dict = {}
    skipped_fixed = 0
    skipped_small = 0
    analysed_total = 0

    for sheet, points in collected.items():
        results[sheet] = {}

        for pname, data in points.items():
            raw_values = [v for _, v, _ in data['values']]
            lsl = data['lsl']
            usl = data['usl']

            # Skip fixed-value items (std == 0, e.g. version strings forced to float)
            if len(set(raw_values)) <= 1:
                _log(f"  [SKIP] 固定值，跳过CPK: [{sheet}] {pname}"
                     f"  (值={raw_values[0] if raw_values else '-'}，N={len(raw_values)})")
                skipped_fixed += 1
                continue

            # Skip if too few samples
            if len(raw_values) < 2:
                _log(f"  [SKIP] 样本量不足，跳过CPK: [{sheet}] {pname}"
                     f"  (N={len(raw_values)})")
                skipped_small += 1
                continue

            stats = calculate_cpk(raw_values, lsl, usl)
            stats['values'] = data['values']   # [(barcode, value, is_pass), ...]
            n_pass = sum(1 for _, _, p in data['values'] if p)
            stats['n_pass'] = n_pass
            stats['n_fail'] = len(data['values']) - n_pass
            results[sheet][pname] = stats
            analysed_total += 1

    # Drop sheets that have zero analyzable items (all non-numeric or all-constant)
    empty_sheets = [s for s, pts in results.items() if not pts]
    for s in empty_sheets:
        del results[s]
    if empty_sheets:
        _log(f"  [INFO] 已隐藏无可分析数据的Sheet: {empty_sheets}")

    # Summary log
    total_sheets = len(results)
    _log(f"  [INFO] CPK分析完成: {total_sheets} 个Sheet，"
         f"{analysed_total} 个测试子项已分析，"
         f"跳过固定值/非数值 {skipped_fixed} 个，样本不足 {skipped_small} 个")

    return results


# ---------------------------------------------------------------------------
# JSON folder analysis  (mirrors analyze_xlsx_folder for *_MEASUREMENT_*.json)
# ---------------------------------------------------------------------------

def analyze_json_folder(json_folder: str, log_cb=None) -> dict:
    """
    Analyze all JSON measurement files in `json_folder`.

    JSON structure expected:
      {
        "DutInfo": {"SerialNumber": "...", "StartTime": "YYYY-MM-DD HH:MM:SS", ...},
        "TestResult": [
          {
            "CaseName": "...",          ← sheet name
            "TestPoints": [
              {
                "TestPointNumber": "...",  ← point_name
                "TestData": "...",         ← data value
                "LimitLow": "...",
                "LimitHigh": "...",
                "Result": "Pass/Fail",
                "StartTime": "..."
              }, ...
            ]
          }, ...
        ]
      }

    Returns the same dict shape as analyze_xlsx_folder.
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    folder = Path(json_folder)
    json_files = sorted(folder.glob('*.json'))

    if not json_files:
        _log(f"  [WARN] CPK分析：{json_folder} 目录中未找到 json 文件")
        return {}

    _log(f"  [INFO] CPK分析目录: {json_folder}")
    _log(f"  [INFO] 找到 {len(json_files)} 个 json 文件，开始读取数据...")

    # collected[case_name][point_number] = {values, latest_time, lsl, usl}
    collected: dict = {}

    for json_path in json_files:
        data = None
        for enc in ('utf-8', 'utf-8-sig', 'gbk', 'latin-1'):
            try:
                with open(json_path, encoding=enc) as f:
                    data = _json_mod.load(f)
                break
            except UnicodeDecodeError:
                continue
            except Exception as exc:
                _log(f"  [ERROR] 无法解析文件 {json_path.name}: {exc}")
                break

        if data is None:
            _log(f"  [ERROR] 无法读取文件 {json_path.name}")
            continue

        # Barcode
        dut_info = data.get('DutInfo', {})
        barcode = str(dut_info.get('SerialNumber', '')).strip() or json_path.stem

        # File time: prefer filename timestamp, fall back to DutInfo.StartTime
        file_time = _file_time_from_name(json_path.stem)
        if file_time == datetime.min:
            ts_str = dut_info.get('StartTime', '')
            if ts_str:
                try:
                    file_time = datetime.strptime(str(ts_str)[:19], '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    pass

        test_result = data.get('TestResult', [])
        if not isinstance(test_result, list):
            _log(f"  [WARN] {json_path.name}: TestResult 格式异常，已跳过")
            continue

        for case in test_result:
            case_name = str(case.get('CaseName', 'Unknown')).strip()
            test_points = case.get('TestPoints', [])
            if not isinstance(test_points, list):
                continue

            if case_name not in collected:
                collected[case_name] = {}

            for pt in test_points:
                pname = str(pt.get('TestPointNumber', '')).strip()
                if not pname or pname.lower() == 'nan':
                    continue

                raw_val = pt.get('TestData')
                try:
                    val = float(raw_val)
                except (TypeError, ValueError):
                    continue

                row_pass = str(pt.get('Result', 'Pass')).strip().lower() != 'fail'

                if pname not in collected[case_name]:
                    collected[case_name][pname] = {
                        'values': [],
                        'latest_time': datetime.min,
                        'lsl': None,
                        'usl': None,
                    }

                collected[case_name][pname]['values'].append((barcode, val, row_pass))

                if file_time >= collected[case_name][pname]['latest_time']:
                    collected[case_name][pname]['latest_time'] = file_time

                    raw_lsl = pt.get('LimitLow')
                    raw_usl = pt.get('LimitHigh')
                    try:
                        collected[case_name][pname]['lsl'] = (
                            float(raw_lsl) if raw_lsl not in (None, '') else None
                        )
                    except (TypeError, ValueError):
                        collected[case_name][pname]['lsl'] = None
                    try:
                        collected[case_name][pname]['usl'] = (
                            float(raw_usl) if raw_usl not in (None, '') else None
                        )
                    except (TypeError, ValueError):
                        collected[case_name][pname]['usl'] = None

    # ── Pass 2: calculate CPK (identical logic to analyze_xlsx_folder) ────
    results: dict = {}
    skipped_fixed = 0
    skipped_small = 0
    analysed_total = 0

    for sheet, points in collected.items():
        results[sheet] = {}

        for pname, data in points.items():
            raw_values = [v for _, v, _ in data['values']]
            lsl = data['lsl']
            usl = data['usl']

            if len(set(raw_values)) <= 1:
                _log(f"  [SKIP] 固定值，跳过CPK: [{sheet}] {pname}"
                     f"  (值={raw_values[0] if raw_values else '-'}，N={len(raw_values)})")
                skipped_fixed += 1
                continue

            if len(raw_values) < 2:
                _log(f"  [SKIP] 样本量不足，跳过CPK: [{sheet}] {pname}"
                     f"  (N={len(raw_values)})")
                skipped_small += 1
                continue

            stats = calculate_cpk(raw_values, lsl, usl)
            stats['values'] = data['values']
            n_pass = sum(1 for _, _, p in data['values'] if p)
            stats['n_pass'] = n_pass
            stats['n_fail'] = len(data['values']) - n_pass
            results[sheet][pname] = stats
            analysed_total += 1

    empty_sheets = [s for s, pts in results.items() if not pts]
    for s in empty_sheets:
        del results[s]
    if empty_sheets:
        _log(f"  [INFO] 已隐藏无可分析数据的Sheet: {empty_sheets}")

    total_sheets = len(results)
    _log(f"  [INFO] CPK分析完成: {total_sheets} 个Sheet，"
         f"{analysed_total} 个测试子项已分析，"
         f"跳过固定值/非数值 {skipped_fixed} 个，样本不足 {skipped_small} 个")

    return results