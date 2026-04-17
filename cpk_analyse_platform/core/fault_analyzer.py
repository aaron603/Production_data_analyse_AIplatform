"""
fault_analyzer.py
-----------------
Traverse test record directories, extract structured fault data, and persist
to the fault database.

Confirmed directory structure (production):
    {station_root}/
      TestResult/
        {product_category}/       e.g. ORBI_B3, ORBI_B40
          {station_type}/         e.g. FT1, FT2, Aging
            debug*/               ← skipped
            {product_code}/       e.g. X11_X11, R1B
              {barcode}/          single (BC) or dual (BC1_BC2)
                {YYYYMMDDHHMMSS[_suffix]}/   ← test record
                  Test_Result_*.xlsx
                  *_MEASUREMENT_Zillnk.json
                  ate_test_log.log / ate_test_log.html
                  file_bk/env_config.yml
                  file_bk/env_comp/*.csv
                  RU1_Log_{BC}/
                  TM1_Log/
                  Failed_points_*.txt
                  screen_*.png / {item_name}/ (screenshot dirs)

Entry point: run_fault_analysis(station_configs, out_dir, level, log_cb, stop_event)
"""

import json
import re
from datetime import datetime
from pathlib import Path

from core import fault_db


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_OLLAMA_URL   = 'http://localhost:11434'
_OLLAMA_MODEL = 'qwen2.5:7b'
_LOG_EXCERPT_MAX = 2000

_SKIP_DIRS = frozenset({
    '__pycache__', '.git', 'node_modules',
    '$RECYCLE.BIN', 'RECYCLER', 'System Volume Information',
})

# Timestamp directory: starts with 14 digits (YYYYMMDDHHMMSS) optionally followed by _suffix
_TS_RE = re.compile(r'^\d{14}')

# Log line with structured test-point result
_CRITICAL_RE = re.compile(
    r'CRITICAL\s+-\s+<string>\s+-\s+(.+?),\s*data=(.+?)\((.+?)\),\s*'
    r'limit=(\[.+?\]),\s*result=(Pass|Fail)',
    re.IGNORECASE,
)

# Equipment / environment error patterns: (compiled_re, label, group_for_detail)
_EQUIP_PATTERNS = [
    (re.compile(r"could not open port '?(COM\d+)'?", re.I),        '串口设备断连',      1),
    (re.compile(r'SerialTimeoutException|Write timeout',  re.I),   '串口超时',          0),
    (re.compile(r'INSTRUMENT ERROR.*?ERR_CODE=(-\d+)',    re.I),   '仪器VISA指令错误',  1),
    (re.compile(r'SSH.*?retry.*?hostname=([\d.]+)',       re.I),   'DUT SSH失联',       1),
    (re.compile(r'(?:socket.*?timeout|connect.*?fail|TCPIP.*?fail)', re.I), '网络仪器连接失败', 0),
    (re.compile(r'SwitchBox.*?connect fail',              re.I),   '射频开关箱连接失败', 0),
    (re.compile(r'FileNotFoundError.*?COM\d+',            re.I),   '串口端口不存在',    0),
]

# Instrument type mapping from env_config.yml key prefixes
_INSTR_TYPE = {
    'SA':  '频谱仪 (Signal Analyzer)',
    'SG':  '信号源 (Signal Generator)',
    'PM':  '功率计 (Power Meter)',
    'VNA': '矢量网络分析仪 (VNA)',
    'PS':  '程控电源 (Power Supply)',
    'SW':  '开关控制器 (Switch)',
    'TM':  '测试主机 (Test Master)',
    'RU':  'DUT (RU/PA)',
}


# ---------------------------------------------------------------------------
# Ollama helpers (unchanged from original)
# ---------------------------------------------------------------------------

def _check_ollama() -> bool:
    try:
        import urllib.request
        with urllib.request.urlopen(f'{_OLLAMA_URL}/api/tags', timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def _ollama_analyze(context: dict) -> dict:
    """
    Send structured context to Ollama. Returns parsed dict or {'error': ...}.
    context keys: failed_items, instruments, equip_errors, first_fail_desc, log_excerpt
    """
    try:
        import urllib.request

        failed_summary = ''
        for fi in (context.get('failed_items') or [])[:10]:
            failed_summary += (
                f"  {fi['item']}: {fi['value']}{fi['unit']} "
                f"(限值{fi['lsl']}~{fi['usl']}, 偏差{fi.get('deviation', 'N/A')})\n"
            )

        instr_summary = ''
        ins = context.get('instruments') or {}
        for k, v in ins.items():
            if k in ('EQP_ID', 'LOCATION'):
                continue
            label = _INSTR_TYPE.get(k[:2], k)
            instr_summary += f"  {k}({label}): {v}\n"

        equip_summary = ''
        for ee in (context.get('equip_errors') or [])[:5]:
            equip_summary += f"  [{ee['label']}] {ee.get('detail', '')} — {ee['raw_line'][:120]}\n"

        prompt = (
            '你是射频测试设备故障分析助手（RRU/PA产线）。\n'
            f'【已知仪器配置】\n{instr_summary or "  (未获取)"}\n'
            f'【失败测试项（最多10条）】\n{failed_summary or "  (无结构化失败项)"}\n'
            f'【设备/通信错误】\n{equip_summary or "  (无)"}\n'
            f'【ATE首次失败描述】{context.get("first_fail_desc") or "(无)"}\n'
            f'【日志摘要】\n{(context.get("log_excerpt") or "")[:1500]}\n\n'
            '请判断：\n'
            '1. 故障来源：DUT硬件 / DUT软件固件 / 仪器仪表 / 射频链路 / 串口开关 / ATE程序 / 校准偏差 / 未知\n'
            '2. 最可能具体原因（指明仪器型号/测试项/错误类型）\n'
            '3. 建议处置措施\n'
            '以JSON严格返回，字段：fault_category(str), root_cause(str), suggestion(str), confidence(0.0-1.0)'
        )

        payload = json.dumps({
            'model': _OLLAMA_MODEL, 'prompt': prompt,
            'stream': False, 'format': 'json',
        }).encode()
        req = urllib.request.Request(
            f'{_OLLAMA_URL}/api/generate', data=payload,
            headers={'Content-Type': 'application/json'}, method='POST',
        )
        with urllib.request.urlopen(req, timeout=90) as resp:
            raw  = json.loads(resp.read().decode())
            text = raw.get('response', '')
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                return {'raw_response': text}
    except Exception as exc:
        return {'error': str(exc)}


# ---------------------------------------------------------------------------
# Directory traversal
# ---------------------------------------------------------------------------

def _find_testresult(root: Path):
    """Return TestResult Path if found at root or one level below, else None."""
    if root.name.lower() == 'testresult':
        return root
    candidate = root / 'TestResult'
    if candidate.is_dir():
        return candidate
    return None


def _iter_records(station_configs: list):
    """
    Yield one dict per test record directory found under each configured station folder.

    Yielded dict keys:
      barcode, barcode_full, station_label, station_machine,
      product_category, product_code, test_time, record_dir (Path)
    """
    for cfg in station_configs:
        if isinstance(cfg, dict):
            station_label = cfg.get('type', '')
            folder_path   = cfg.get('folder', '')
        else:
            station_label, folder_path = cfg[0], cfg[1]

        root = Path(folder_path)
        if not root.is_dir():
            continue

        # station_machine: physical machine name from folder or EQP_ID (set later)
        station_machine = root.name

        testresult = _find_testresult(root)
        if testresult:
            yield from _scan_testresult(
                testresult, station_label, station_machine
            )
        else:
            # Fallback: treat root as product_code level (e.g. copied apricot dirs)
            yield from _scan_barcode_level(
                root, station_label, station_machine, '', ''
            )


def _scan_testresult(testresult: Path, station_label: str, station_machine: str):
    """
    Navigate TestResult/{product_category}/{station_type}/{product_code}/...
    Skip any folder whose name starts with 'debug' (case-insensitive).
    """
    for prod_cat in testresult.iterdir():
        if not prod_cat.is_dir() or prod_cat.name in _SKIP_DIRS:
            continue
        for stype_dir in prod_cat.iterdir():
            if not stype_dir.is_dir() or stype_dir.name in _SKIP_DIRS:
                continue
            for prod_code in stype_dir.iterdir():
                if not prod_code.is_dir():
                    continue
                if prod_code.name.lower().startswith('debug'):
                    continue
                if prod_code.name in _SKIP_DIRS:
                    continue
                yield from _scan_barcode_level(
                    prod_code, station_label, station_machine,
                    prod_cat.name, prod_code.name,
                )


def _scan_barcode_level(prod_code_dir: Path, station_label: str, station_machine: str,
                        product_category: str, product_code: str):
    """Traverse {barcode}/{timestamp}/ under prod_code_dir."""
    for bc_dir in prod_code_dir.iterdir():
        if not bc_dir.is_dir() or bc_dir.name in _SKIP_DIRS:
            continue
        barcode_full = bc_dir.name
        # Primary barcode: first segment (handles both single and dual BC1_BC2)
        parts = barcode_full.split('_')
        primary = parts[0] if len(parts) >= 2 and all(
            p.isalnum() and len(p) >= 5 for p in parts[:2]
        ) else barcode_full

        for ts_dir in bc_dir.iterdir():
            if not ts_dir.is_dir() or not _TS_RE.match(ts_dir.name):
                continue
            raw_ts = re.sub(r'\D', '', ts_dir.name)[:14]
            try:
                test_time = datetime.strptime(raw_ts, '%Y%m%d%H%M%S').strftime(
                    '%Y-%m-%d %H:%M:%S'
                )
            except Exception:
                test_time = ts_dir.name

            yield {
                'barcode':          primary,
                'barcode_full':     barcode_full,
                'station_label':    station_label,
                'station_machine':  station_machine,
                'product_category': product_category,
                'product_code':     product_code,
                'test_time':        test_time,
                'record_dir':       ts_dir,
            }


# ---------------------------------------------------------------------------
# Per-record data extractors
# ---------------------------------------------------------------------------

def _read_log(record_dir: Path):
    """Return (text, path_str) for the best log file in record_dir."""
    priority = ('ate_test_log.log', 'ate_test_log.html', 'test_log.log',
                'log.txt', 'result.log')
    for name in priority:
        p = record_dir / name
        if p.exists():
            try:
                return p.read_text(encoding='utf-8', errors='replace'), str(p)
            except Exception:
                pass
    for p in record_dir.glob('*.log'):
        try:
            return p.read_text(encoding='utf-8', errors='replace'), str(p)
        except Exception:
            pass
    return '', ''


def _parse_critical_lines(log_text: str) -> tuple:
    """
    Extract structured test-point results from CRITICAL - <string> lines.

    Returns:
      failed_items : list of dicts {item, value, unit, lsl, usl, direction, deviation}
      status       : 'pass' | 'fail' | 'unknown'
    """
    failed, has_any = [], False
    for m in _CRITICAL_RE.finditer(log_text):
        item, raw_val, unit, raw_limit, result = m.groups()
        has_any = True
        if result.lower() != 'fail':
            continue
        # Parse numeric value
        try:
            val = float(raw_val.strip())
        except ValueError:
            val = raw_val.strip()

        # Parse limit list  e.g. [90, 120] or ['1.1.5', '1.1.5']
        lsl, usl = None, None
        try:
            limits = json.loads(raw_limit.replace("'", '"'))
            if isinstance(limits, list) and len(limits) == 2:
                lsl, usl = limits[0], limits[1]
        except Exception:
            pass

        # Direction and deviation (numeric only)
        direction, deviation = '', ''
        if isinstance(val, float) and isinstance(lsl, (int, float)) \
                and isinstance(usl, (int, float)):
            if val < lsl:
                direction = 'low'
                deviation = f'{val - lsl:+.3g}'
            elif val > usl:
                direction = 'high'
                deviation = f'{val - usl:+.3g}'

        failed.append({
            'item':      item.strip(),
            'value':     raw_val.strip(),
            'unit':      unit.strip(),
            'lsl':       lsl,
            'usl':       usl,
            'direction': direction,
            'deviation': deviation,
        })

    if not has_any:
        status = 'unknown'
    elif failed:
        status = 'fail'
    else:
        status = 'pass'

    return failed, status


def _detect_equip_errors(log_text: str) -> list:
    """
    Scan log for equipment / communication error patterns.
    Returns list of {label, detail, raw_line}.
    """
    errors = []
    seen_labels = set()
    for line in log_text.splitlines():
        for pattern, label, group in _EQUIP_PATTERNS:
            m = pattern.search(line)
            if m:
                detail = m.group(group) if group and group <= len(m.groups()) else ''
                key = f'{label}:{detail}'
                if key not in seen_labels:
                    seen_labels.add(key)
                    errors.append({
                        'label':    label,
                        'detail':   detail,
                        'raw_line': line.strip()[:200],
                    })
    return errors


def _parse_env_config(record_dir: Path) -> dict:
    """
    Parse file_bk/env_config.yml.
    Returns dict with 'instruments' (addr map) and metadata keys (EQP_ID, LOCATION).
    No external YAML library required — uses simple line parsing.
    """
    yml_path = record_dir / 'file_bk' / 'env_config.yml'
    if not yml_path.exists():
        return {}

    result = {}
    try:
        text = yml_path.read_text(encoding='utf-8', errors='replace')
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if ':' not in line:
                continue
            key, _, val = line.partition(':')
            key = key.strip()
            val = val.strip().strip("'\"")
            if not key or val in ('', '0', "''", '""'):
                continue
            result[key] = val
    except Exception:
        return {}

    # Build instrument summary: keep VISA addresses and metadata
    instruments = {}
    for prefix in ('SA', 'SG', 'PM', 'VNA', 'PS', 'SW', 'TM', 'RU'):
        num = int(result.get(f'{prefix}_NUM', 0) or 0)
        for i in range(1, num + 1):
            addr = result.get(f'{prefix}{i}', '')
            if addr and addr != '0':
                instruments[f'{prefix}{i}'] = addr
    # Metadata
    for meta in ('EQP_ID', 'LOCATION'):
        if meta in result:
            instruments[meta] = result[meta]

    return instruments  # {SA1: "TCPIP0::...", EQP_ID: "FT_1", ...}


def _read_measurement_json(record_dir: Path) -> dict:
    """
    Read *_MEASUREMENT_Zillnk.json and return DutInfo fields.
    Returns {} if not found.
    """
    for f in record_dir.glob('*_MEASUREMENT_Zillnk.json'):
        try:
            data = json.loads(f.read_bytes().decode('utf-8', 'replace'))
            di   = data.get('DutInfo', {})
            return {
                'result':          di.get('Result', ''),
                'first_fail_desc': di.get('FirstFailCaseDescription', ''),
                'product_name':    di.get('ProductName', ''),
                'rstate':          di.get('Rstate', ''),
                'eqp_id':          di.get('SiteName', ''),
            }
        except Exception:
            pass
    return {}


def _read_failed_points_txt(record_dir: Path) -> list:
    """
    Read Failed_points_*.txt (present in Apricot/RRU records).
    Returns list of failed item name strings.
    """
    for f in record_dir.glob('Failed_points_*.txt'):
        try:
            lines = f.read_text(encoding='utf-8', errors='replace').splitlines()
            items = []
            for ln in lines:
                ln = ln.strip()
                if not ln:
                    continue
                # Format: "{barcode} - {item_name}"
                if ' - ' in ln:
                    items.append(ln.split(' - ', 1)[1].strip())
                else:
                    items.append(ln)
            return items
        except Exception:
            pass
    return []


def _infer_status_from_dir(record_dir: Path) -> str:
    """Fallback status inference when log parsing yields 'unknown'."""
    if any(record_dir.glob('Failed_points_*.txt')):
        return 'fail'
    if any(record_dir.glob('Test_Result_*.xlsx')):
        return 'pass'
    return 'unknown'


def _quick_is_fail(record_dir: Path) -> bool:
    """
    Fast pre-check (no log read) whether a test record is a fail.
    Used in fail_only mode to skip pass records before heavy parsing.
    """
    # Failed_points_*.txt only exists for fail records (RRU/Apricot style)
    if any(record_dir.glob('Failed_points_*.txt')):
        return True
    # MEASUREMENT JSON result field (B3B40 style) — just scan first 1 KB
    for f in record_dir.glob('*_MEASUREMENT_Zillnk.json'):
        try:
            head = f.read_bytes()[:1024].decode('utf-8', 'replace')
            if '"Result": "Fail"' in head or '"Result":"Fail"' in head:
                return True
            if '"Result": "Pass"' in head or '"Result":"Pass"' in head:
                return False
        except Exception:
            pass
    # Scan just the tail of ate_test_log.log for final result line
    log_path = record_dir / 'ate_test_log.log'
    if log_path.exists():
        try:
            # Read last 2 KB for result line
            with open(log_path, 'rb') as f:
                f.seek(max(0, log_path.stat().st_size - 2048))
                tail = f.read().decode('utf-8', 'replace')
            if 'result=Fail' in tail:
                return True
            if 'result=Pass' in tail:
                return False
        except Exception:
            pass
    return None   # unknown — let caller decide


def generate_fault_barcode_list(db_path, output_path: str, log_cb=None) -> int:
    """
    Query fail/unknown records from the fault DB and write to an Excel file.

    Columns: 条码, 完整条码, 工站, 机台, 测试时间, 状态, 故障类型,
             失败测试项数量, 首次失败描述

    Returns the number of rows written.
    """
    import json as _json
    import pandas as pd

    def _log(msg):
        if log_cb:
            log_cb(msg)

    records = fault_db.get_records(db_path, limit=10000)
    fail_records = [r for r in records if r.get('status') in ('fail', 'unknown')]

    if not fail_records:
        _log('  [INFO] 故障条码列表：无失败/未知记录，跳过生成')
        return 0

    rows = []
    for r in fail_records:
        # Count failed test items
        failed_count = 0
        if r.get('failed_items'):
            try:
                failed_count = len(_json.loads(r['failed_items']))
            except Exception:
                pass

        rows.append({
            '条码':           r.get('barcode', ''),
            '完整条码':       r.get('barcode_full', ''),
            '工站':           r.get('station', ''),
            '机台':           r.get('station_machine', ''),
            '测试时间':       r.get('test_time', ''),
            '状态':           r.get('status', ''),
            '故障类型':       r.get('fault_type', ''),
            '失败测试项数量': failed_count,
            '首次失败描述':   r.get('first_fail_desc', ''),
        })

    df = pd.DataFrame(rows)
    # Sort by test_time desc, then barcode
    df.sort_values(['测试时间', '条码'], ascending=[False, True], inplace=True)
    df.reset_index(drop=True, inplace=True)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='故障条码列表', index=False)
            ws = writer.sheets['故障条码列表']
            # Auto-fit column widths
            for col in ws.columns:
                max_len = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in col
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    except Exception as exc:
        _log(f'  [ERROR] 故障条码列表写入失败: {exc}')
        return 0

    _log(f'  [INFO] 故障条码列表: {len(rows)} 条记录 → {output_path}')
    return len(rows)


def generate_rule_suggestions_yaml(db_path, output_path: str, log_cb=None) -> int:
    """
    Generate a YAML template for engineers to fill in fault relationship descriptions.

    Based on _generate_fail_patterns(): lists top unclassified failed items and
    high-frequency failed test items that lack a matching rule.
    Engineers fill in fault_type / suggestion, then load the file via the menu.

    Returns the number of suggestion entries written.
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    patterns = _generate_fail_patterns(db_path)
    total_fail = patterns.get('total_fail', 0)
    if total_fail == 0:
        _log('  [INFO] 规则建议YAML：无失败记录，跳过生成')
        return 0

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_str = datetime.now().strftime('%Y%m%d')

    lines = [
        f'# 故障规则建议模板 — 由系统自动生成',
        f'# 生成时间: {now_str}',
        f'# 基于本次分析: {total_fail} 条失败记录',
        f'#',
        f'# 使用说明:',
        f'#   1. 将下方每条建议的 fault_type 和 suggestion 填写完整',
        f'#   2. 保存文件',
        f'#   3. 通过菜单「工具 → 加载故障关系描述文件…」导入',
        f'#',
        f'# 注意: keywords 字段为系统建议关键词，可按需修改（逗号分隔）',
        f'',
        f'version: "1.0"',
        f'date: "{datetime.now().strftime("%Y-%m-%d")}"',
        f'',
        f'rules:',
    ]

    count = 0

    # Section 1: unclassified samples (highest priority — these have no rule match at all)
    unclassified = patterns.get('unclassified_samples', [])
    if unclassified:
        lines.append(f'  # ── 未分类故障样本（共 {len(unclassified)} 条，前 {min(len(unclassified), 10)} 条）')
        for sample in unclassified[:10]:
            safe = sample.replace('"', "'").replace('\n', ' ')[:120]
            lines += [
                f'  - keywords: "{safe}"',
                f'    fault_type: ""    # TODO: 请填写故障类型',
                f'    suggestion: ""    # TODO: 请填写处置建议',
                f'    # example_log: "{safe}"',
                f'',
            ]
            count += 1

    # Section 2: top failed test items without matching rules
    top_items = patterns.get('top_failed_items', [])
    if top_items:
        lines.append(f'  # ── 高频失败测试项（出现次数最多的前10项）')
        existing_kw = set()
        for item, cnt in top_items[:10]:
            if not item or item in existing_kw:
                continue
            existing_kw.add(item)
            safe_item = item.replace('"', "'")[:80]
            lines += [
                f'  - keywords: "{safe_item}"    # 出现 {cnt} 次',
                f'    fault_type: ""    # TODO: 请填写故障类型',
                f'    suggestion: ""    # TODO: 请填写处置建议',
                f'',
            ]
            count += 1

    # Section 3: equipment error labels as hints (read-only reference)
    top_equip = patterns.get('top_equip_errors', [])
    if top_equip:
        lines.append(f'  # ── 设备/通信错误参考（已有内置规则，如需细化可添加）')
        for label, cnt in top_equip[:5]:
            if not label:
                continue
            safe_label = label.replace('"', "'")
            lines += [
                f'  # - keywords: "{safe_label}"    # 出现 {cnt} 次  ← 参考，非必填',
            ]
        lines.append('')

    if count == 0:
        _log('  [INFO] 规则建议YAML：无未分类/高频项，跳过生成')
        return 0

    yaml_text = '\n'.join(lines)
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(yaml_text)
    except OSError as exc:
        _log(f'  [ERROR] 规则建议YAML写入失败: {exc}')
        return 0

    _log(f'  [INFO] 规则建议YAML模板: {count} 条建议 → {output_path}')
    _log(f'  [INFO] 请填写 fault_type/suggestion 后，通过「工具→加载故障关系描述文件…」导入')
    return count


def _generate_fail_patterns(db_path) -> dict:
    """
    Summarise patterns from fail records already in the DB.
    Called after fail_only traversal to produce rule-improvement candidates.

    Returns:
      {
        'top_failed_items': [(item_name, count), ...],   # most frequent failed test items
        'top_equip_errors': [(label, count), ...],       # most frequent equipment errors
        'top_fault_types':  [(fault_type, count), ...],  # fault classification breakdown
        'unclassified_samples': [first_fail_desc, ...],  # unclassified for manual review
      }
    """
    import json as _json
    from collections import Counter

    records = fault_db.get_records(db_path, limit=5000)
    fail_records = [r for r in records if r.get('status') == 'fail']

    item_counter  = Counter()
    equip_counter = Counter()
    type_counter  = Counter()
    unclassified_samples = []

    for r in fail_records:
        # Count failed test items
        if r.get('failed_items'):
            try:
                items = _json.loads(r['failed_items'])
                for fi in items:
                    item_counter[fi.get('item', '')] += 1
            except Exception:
                pass

        # Count equipment error labels
        if r.get('equip_errors'):
            try:
                errs = _json.loads(r['equip_errors'])
                for ee in errs:
                    equip_counter[ee.get('label', '')] += 1
            except Exception:
                pass

        # Fault type distribution
        type_counter[r.get('fault_type', '未分类故障')] += 1

        # Collect unclassified samples for manual review
        if r.get('fault_type') == '未分类故障' and r.get('first_fail_desc'):
            if len(unclassified_samples) < 20:
                unclassified_samples.append(r['first_fail_desc'])

    return {
        'total_fail':         len(fail_records),
        'top_failed_items':   item_counter.most_common(20),
        'top_equip_errors':   equip_counter.most_common(10),
        'top_fault_types':    type_counter.most_common(20),
        'unclassified_samples': unclassified_samples,
    }


def _extract_excerpt(log_text: str) -> str:
    """Return a short excerpt prioritising error/fail lines."""
    priority = ('error', 'fail', 'exception', 'timeout', 'fatal',
                'critical - <string>', '失败', '超时', '异常', '错误', 'traceback')
    lines = log_text.splitlines()
    hits  = [l for l in lines if any(kw in l.lower() for kw in priority)]
    chosen = hits[:20] if hits else lines[:20]
    return '\n'.join(chosen)[:_LOG_EXCERPT_MAX]


# ---------------------------------------------------------------------------
# Rule-based keyword matching (fallback when structured parsing is insufficient)
# ---------------------------------------------------------------------------

def _match_rules(log_text: str, failed_items: list, equip_errors: list,
                 rules: list) -> tuple:
    """
    Returns (fault_type: str, matched_rule_id: int | None).
    Priority: equip_errors → failed_items names → log keywords.
    """
    lower = log_text.lower()

    # 1. Equipment errors take highest priority (environment/instrument issue)
    if equip_errors:
        label = equip_errors[0]['label']
        for rule in rules:
            kws = [k.strip().lower() for k in rule['keywords'].split(',') if k.strip()]
            if any(kw in label.lower() for kw in kws):
                return rule['fault_type'], rule['id']
        return label, None   # use the error label directly as fault_type

    # 2. Failed item names
    item_text = ' '.join(fi['item'].lower() for fi in failed_items)
    if item_text:
        for rule in rules:
            kws = [k.strip().lower() for k in rule['keywords'].split(',') if k.strip()]
            if any(kw in item_text for kw in kws):
                return rule['fault_type'], rule['id']

    # 3. Full log keyword scan
    for rule in rules:
        kws = [k.strip() for k in rule['keywords'].split(',') if k.strip()]
        if any(kw.lower() in lower for kw in kws):
            return rule['fault_type'], rule['id']

    return '未分类故障', None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_fault_analysis(station_configs: list, out_dir: str,
                       level: str = '基础版（规则库）',
                       mode: str = 'all',
                       log_cb=None, stop_event=None) -> dict:
    """
    Traverse all station test directories, extract structured fault data,
    and persist to fault_database.db.

    Args:
        station_configs : list of dicts {type, folder} or (type, folder) tuples
        out_dir         : output directory; fault_database.db stored here
        level           : '基础版（规则库）' | '增强版（规则库+Ollama）'
        mode            : 'all'       — process all records; pass records stored with minimal
                                        analysis; emphasis on cross-station comparison
                          'fail_only' — skip pass records; deep analysis of fail records only;
                                        generates pattern summary for rule improvement
        log_cb          : optional callable(str) for progress messages
        stop_event      : optional threading.Event for cancellation

    Returns:
        {station_label: {total, classified, unclassified},
         '__stats__': [...],
         '__cross_station__': [...],
         '__fail_patterns__': {...}  (fail_only mode only)}
    """
    def _log(msg):
        if log_cb:
            log_cb(msg)

    def _stopped():
        return stop_event is not None and stop_event.is_set()

    db_path = Path(out_dir) / 'fault_database.db'
    fault_db.init_db(db_path)
    rules = fault_db.get_rules(db_path)
    _log(f'  [INFO] 故障分析启动，规则库 {len(rules)} 条，DB: {db_path}')

    use_ollama = '增强版' in level
    ollama_ok  = False
    if use_ollama:
        ollama_ok = _check_ollama()
        if ollama_ok:
            _log(f'  [INFO] Ollama 已连接，启用增强版分析（{_OLLAMA_MODEL}）')
        else:
            _log('  [WARN] 未检测到 Ollama (localhost:11434)，降级为基础版规则库')

    fail_only = (mode == 'fail_only')
    _log(f'  [INFO] 分析模式: {"只分析失败数据（深度模式）" if fail_only else "分析所有数据（含跨站比对）"}')

    fault_db.clear_records(db_path)

    per_station: dict = {}

    for rec in _iter_records(station_configs):
        if _stopped():
            _log('  [INFO] 故障分析已中止')
            break

        slabel = rec['station_label']
        if slabel not in per_station:
            per_station[slabel] = {'total': 0, 'classified': 0, 'unclassified': 0}

        record_dir = rec['record_dir']

        # ── fail_only mode: quick pre-filter (skip pass records) ─────────
        if fail_only:
            quick = _quick_is_fail(record_dir)
            if quick is False:
                continue          # confirmed pass — skip entirely
            # quick is True or None (unknown) → proceed with full analysis

        per_station[slabel]['total'] += 1

        # ── 1. Read primary log ──────────────────────────────────────────
        log_text, log_path = _read_log(record_dir)

        # ── 2. Structured failure extraction ────────────────────────────
        failed_items, log_status = _parse_critical_lines(log_text)

        # ── 3. Equipment error detection ────────────────────────────────
        equip_errors = _detect_equip_errors(log_text) if log_text else []

        # ── 4. MEASUREMENT JSON (B3B40 / similar) ───────────────────────
        mj = _read_measurement_json(record_dir)
        first_fail_desc = mj.get('first_fail_desc', '')
        if log_status == 'unknown' and mj.get('result'):
            log_status = mj['result'].lower()
        station_machine = rec['station_machine']
        if mj.get('eqp_id'):
            station_machine = mj['eqp_id']

        # ── 5. Failed_points.txt (RRU/Apricot style) ────────────────────
        _read_failed_points_txt(record_dir)   # result used for status inference below

        # ── 6. Instrument config ─────────────────────────────────────────
        instruments = _parse_env_config(record_dir) or None

        # ── 7. Final status ──────────────────────────────────────────────
        if log_status == 'unknown':
            log_status = _infer_status_from_dir(record_dir)

        # In fail_only mode, skip confirmed pass records that only surfaced here
        if fail_only and log_status == 'pass':
            per_station[slabel]['total'] -= 1
            continue

        # In all mode: pass records get minimal fault_type, skip heavy analysis
        if not fail_only and log_status == 'pass':
            fault_db.add_record(
                db_path,
                barcode          = rec['barcode'],
                barcode_full     = rec['barcode_full'],
                station          = slabel,
                station_machine  = station_machine,
                product_category = rec['product_category'],
                product_code     = rec['product_code'],
                test_time        = rec['test_time'],
                status           = 'pass',
                fault_type       = '测试通过',
                instruments      = instruments,
                log_path         = log_path,
            )
            # Don't count pass records toward classified/unclassified
            continue

        # ── 8. Fault classification (fail / unknown records) ─────────────
        fault_type, rule_id = _match_rules(
            log_text, failed_items, equip_errors, rules
        )
        if fault_type == '未分类故障' and first_fail_desc:
            fault_type, rule_id = _match_rules(
                first_fail_desc, failed_items, equip_errors, rules
            )

        # ── 9. LLM enhancement (optional) ───────────────────────────────
        llm_analysis = None
        if use_ollama and ollama_ok and not _stopped():
            context = {
                'failed_items':    failed_items,
                'instruments':     instruments,
                'equip_errors':    equip_errors,
                'first_fail_desc': first_fail_desc,
                'log_excerpt':     _extract_excerpt(log_text),
            }
            llm_result = _ollama_analyze(context)
            if llm_result and 'error' not in llm_result:
                llm_analysis = json.dumps(llm_result, ensure_ascii=False)
                if fault_type == '未分类故障' and 'fault_category' in llm_result:
                    fault_type = llm_result['fault_category']

        # ── 10. Persist ──────────────────────────────────────────────────
        fault_db.add_record(
            db_path,
            barcode          = rec['barcode'],
            barcode_full     = rec['barcode_full'],
            station          = slabel,
            station_machine  = station_machine,
            product_category = rec['product_category'],
            product_code     = rec['product_code'],
            test_time        = rec['test_time'],
            status           = log_status,
            fault_type       = fault_type,
            matched_rule_id  = rule_id,
            first_fail_desc  = first_fail_desc,
            failed_items     = failed_items  or None,
            equip_errors     = equip_errors  or None,
            instruments      = instruments,
            log_excerpt      = _extract_excerpt(log_text),
            log_path         = log_path,
            llm_analysis     = llm_analysis,
        )

        if fault_type == '未分类故障':
            per_station[slabel]['unclassified'] += 1
        else:
            per_station[slabel]['classified'] += 1

    # ── Summary ──────────────────────────────────────────────────────────
    global_stats   = fault_db.get_stats(db_path)
    total_all      = sum(v['total']      for v in per_station.values())
    classified_all = sum(v['classified'] for v in per_station.values())
    _log(
        f'  [INFO] 故障分析完成: 共处理 {total_all} 条失败/未知记录，'
        f'已分类 {classified_all}，未分类 {total_all - classified_all}'
    )
    for s in global_stats[:12]:
        if s['count'] > 0 and s['fault_type'] != '测试通过':
            _log(f'           {s["fault_type"]}: {s["count"]} 次')

    # Cross-station report (meaningful in 'all' mode where pass records exist)
    cross = fault_db.get_cross_station_barcodes(db_path)
    if cross:
        _log(f'  [INFO] 跨站条码（同一模块在多台设备均有测试记录）: {len(cross)} 个')
        for c in cross[:8]:
            flag = ' ← 含失败' if c['has_fail'] else ''
            _log(f'           {c["barcode"]}  →  [{c["machines"]}]{flag}  共{c["record_count"]}条')

    # fail_only mode: generate pattern summary for rule improvement
    fail_patterns = {}
    if fail_only:
        _log('  [INFO] 生成失败模式摘要（可用于规则库优化）...')
        fail_patterns = _generate_fail_patterns(db_path)
        _log(f'  [INFO] 共 {fail_patterns["total_fail"]} 条失败记录')
        if fail_patterns['top_failed_items']:
            _log('  [INFO] 高频失败测试项 TOP10:')
            for item, cnt in fail_patterns['top_failed_items'][:10]:
                if item:
                    _log(f'           {item}: {cnt} 次')
        if fail_patterns['top_equip_errors']:
            _log('  [INFO] 设备/通信错误 TOP5:')
            for label, cnt in fail_patterns['top_equip_errors'][:5]:
                _log(f'           {label}: {cnt} 次')
        if fail_patterns['unclassified_samples']:
            _log(f'  [INFO] 未分类故障样本（前5条，建议添加规则）:')
            for s in fail_patterns['unclassified_samples'][:5]:
                _log(f'           {s}')

    result = dict(per_station)
    result['__stats__']        = global_stats
    result['__cross_station__'] = cross
    result['__fail_patterns__'] = fail_patterns
    return result
