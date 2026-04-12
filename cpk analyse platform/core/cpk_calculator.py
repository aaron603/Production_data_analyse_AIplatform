"""
cpk_calculator.py
-----------------
Reads all xlsx files from a folder and calculates CPK statistics per
sheet / point_name, using the most recently dated file's limits as the
unified LSL/USL.
"""

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

                if pname not in collected[sheet]:
                    collected[sheet][pname] = {
                        'values': [],
                        'latest_time': datetime.min,
                        'lsl': None,
                        'usl': None,
                    }

                collected[sheet][pname]['values'].append((barcode, val))

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
            raw_values = [v for _, v in data['values']]
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
            stats['values'] = data['values']
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