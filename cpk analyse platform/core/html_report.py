"""
html_report.py
--------------
Generates a self-contained HTML CPK analysis report.
No external CSS/JS dependencies – works fully offline.
"""

import json
import os
from datetime import datetime


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def generate_report(
    analysis_data: dict,
    output_path: str,
    title: str = "CPK 分析报告 - Zillnk",
    station_info: dict = None,
) -> str:
    """
    Build and write the HTML report.

    Parameters
    ----------
    analysis_data : dict
        {station_type: {sheet_name: {point_name: {stats + values}}}}
    output_path : str
        Absolute path to write the .html file.
    title : str
        Page title shown in browser tab.
    station_info : dict, optional
        {station_type: folder_count} — number of physical station folders
        configured per type, used for the header summary line.

    Returns
    -------
    str
        Absolute path of the written file.
    """
    html = _build_html(analysis_data, title, station_info)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return output_path


# ---------------------------------------------------------------------------
# Internal builders
# ---------------------------------------------------------------------------

def _station_summary(station_types: list, station_info: dict) -> str:
    """Build the header station-type summary string."""
    n = len(station_types)
    if not station_info:
        return f'共 {n} 种工站类型'
    parts = []
    for stype in station_types:
        cnt = station_info.get(stype)
        parts.append(f'{stype}&nbsp;×&nbsp;{cnt}台' if cnt else stype)
    return f'共 {n} 种工站类型：' + '，'.join(parts)


def _fmt(val, decimals=4):
    if val is None:
        return '-'
    try:
        return f"{float(val):.{decimals}f}"
    except Exception:
        return str(val)


def _build_html(data: dict, title: str, station_info: dict = None) -> str:
    station_types = list(data.keys())

    # Serialize all data as JSON so JS can access it
    # Structure: {station: {sheet: {point: {stats + values}}}}
    js_data = {}
    for stype, sheets in data.items():
        js_data[stype] = {}
        for sheet, points in sheets.items():
            js_data[stype][sheet] = {}
            for pname, stats in points.items():
                entry = {k: v for k, v in stats.items() if k != 'values'}
                entry['barcodes'] = [b for b, _ in stats.get('values', [])]
                entry['raw'] = [v for _, v in stats.get('values', [])]
                js_data[stype][sheet][pname] = entry

    js_data_str = json.dumps(js_data, ensure_ascii=False)

    # Build station tabs HTML
    station_tabs_nav = ''
    station_tabs_content = ''
    for i, stype in enumerate(station_types):
        active = 'active' if i == 0 else ''
        station_tabs_nav += (
            f'<button class="stab-btn {active}" '
            f'onclick="switchStation(\'{_esc_js(stype)}\')" '
            f'id="stab-{_esc_id(stype)}">{stype}</button>\n'
        )
        station_tabs_content += _build_station_panel(stype, data[stype], i == 0)

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>{title}</title>
<style>
{_CSS}
</style>
</head>
<body>
<div class="header">
  <div class="header-title">{title}</div>
  <div class="header-sub">生成时间: {now} &nbsp;|&nbsp; {_station_summary(station_types, station_info)}</div>
</div>

<!-- Station type tabs -->
<div class="stab-bar">
{station_tabs_nav}
</div>

<!-- Station panels -->
{station_tabs_content}

<div id="range-modal" class="modal-overlay" style="display:none">
  <div class="modal-box">
    <div class="modal-title">数值范围搜索结果</div>
    <div id="range-result"></div>
    <button class="btn btn-sm" onclick="document.getElementById('range-modal').style.display='none'">关闭</button>
  </div>
</div>

<script>
const ALL_DATA = {js_data_str};

// CPK visibility state per sheet
const cpkModeActive = {{}};

// Tracks currently displayed point per sheet: stationType + '|' + sheetName → pointName
const currentPoint = {{}};

// ── ID escaping: mirrors Python _esc_id() ────────────────────────────────
function escId(s) {{
  return String(s).replace(/[ /\\\\]/g, '_');
}}

// ── Station switching ────────────────────────────────────────────────────
function switchStation(stype) {{
  document.querySelectorAll('.station-panel').forEach(p => p.style.display = 'none');
  document.querySelectorAll('.stab-btn').forEach(b => b.classList.remove('active'));
  const panel = document.getElementById('panel-' + escId(stype));
  if (panel) panel.style.display = 'block';
  const btn = document.getElementById('stab-' + escId(stype));
  if (btn) btn.classList.add('active');
}}

// ── Sheet switching ──────────────────────────────────────────────────────
function switchSheet(stationType, sheetName) {{
  document.querySelectorAll('.sheet-panel[data-station="' + stationType + '"]')
    .forEach(p => p.style.display = 'none');
  document.querySelectorAll('.shtab-btn[data-station="' + stationType + '"]')
    .forEach(b => b.classList.remove('active'));
  const panelId = 'shpanel_' + escId(stationType) + '_' + escId(sheetName);
  const el = document.getElementById(panelId);
  if (el) el.style.display = 'block';
  const btnId = 'shbtn_' + escId(stationType) + '_' + escId(sheetName);
  const btn = document.getElementById(btnId);
  if (btn) btn.classList.add('active');
  drawFirstInSheet(stationType, sheetName);
}}

// ── Search / CPK toggle ──────────────────────────────────────────────────
function onSearch(stationType, sheetName) {{
  const inputId = 'search_' + escId(stationType) + '_' + escId(sheetName);
  const val = document.getElementById(inputId).value.trim().toLowerCase();
  const tableId = 'table_' + escId(stationType) + '_' + escId(sheetName);
  const table = document.getElementById(tableId);
  if (!table) return;

  const showCpk = val.includes('cpk');
  cpkModeActive[stationType + '|' + sheetName] = showCpk;

  table.querySelectorAll('.cpk-col').forEach(el => {{
    el.style.display = showCpk ? '' : 'none';
  }});

  const panelId = 'shpanel_' + escId(stationType) + '_' + escId(sheetName);
  const panel = document.getElementById(panelId);
  if (panel) {{
    panel.querySelectorAll('.cpk-chartinfo').forEach(e => {{
      e.style.display = showCpk ? 'inline' : 'none';
    }});
  }}

  const keyword = val.replace(/cpk/g, '').trim();
  table.querySelectorAll('tbody tr').forEach(row => {{
    const pointName = row.getAttribute('data-point') || '';
    row.style.display = (!keyword || pointName.toLowerCase().includes(keyword)) ? '' : 'none';
  }});
}}

// ── Row click → update chart ─────────────────────────────────────────────
function onRowClick(stationType, sheetName, pointName) {{
  drawChart(stationType, sheetName, pointName);
  const tableId = 'table_' + escId(stationType) + '_' + escId(sheetName);
  const tbl = document.getElementById(tableId);
  if (tbl) {{
    tbl.querySelectorAll('tbody tr').forEach(r => r.classList.remove('selected'));
    const row = tbl.querySelector('tr[data-point="' + pointName.replace(/"/g, '\\"') + '"]');
    if (row) row.classList.add('selected');
  }}
}}

// ── Normal distribution chart (Canvas) ──────────────────────────────────
function normalPDF(x, mu, sigma) {{
  return Math.exp(-0.5 * Math.pow((x - mu) / sigma, 2)) / (sigma * Math.sqrt(2 * Math.PI));
}}

function drawFirstInSheet(stationType, sheetName) {{
  const shData = (ALL_DATA[stationType] || {{}})[sheetName];
  if (!shData) return;
  const pts = Object.keys(shData);
  if (pts.length > 0) drawChart(stationType, sheetName, pts[0]);
}}

function drawChart(stationType, sheetName, pointName) {{
  const canvasId = 'chart_' + escId(stationType) + '_' + escId(sheetName);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  // Sync drawing resolution to CSS display width so chart fills the container
  const cssW = canvas.offsetWidth;
  if (cssW > 0) canvas.width = cssW;

  // Remember what's currently shown so we can redraw on window resize
  currentPoint[stationType + '|' + sheetName] = pointName;

  const stData = ALL_DATA[stationType];
  if (!stData) return;
  const shData = stData[sheetName];
  if (!shData) return;
  const pt = shData[pointName];
  if (!pt) return;

  const mu = pt.mean;
  const sigma = pt.std;
  const lsl = pt.lsl;
  const usl = pt.usl;
  const rawVals = pt.raw || [];

  const ctx = canvas.getContext('2d');
  const W = canvas.width;
  const H = canvas.height;
  const padL = 60, padR = 30, padT = 30, padB = 50;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;

  ctx.clearRect(0, 0, W, H);

  if (!sigma || sigma === 0 || rawVals.length < 2) {{
    ctx.fillStyle = '#888';
    ctx.font = '14px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('数据不足，无法绘制分布图', W / 2, H / 2);
    // Update info label
    _updateChartInfo(stationType, sheetName, pt, pointName);
    return;
  }}

  const span = Math.max(4 * sigma, (usl != null && lsl != null) ? (usl - lsl) * 0.6 : 4 * sigma);
  let xMin = mu - Math.max(4 * sigma, span * 0.55);
  let xMax = mu + Math.max(4 * sigma, span * 0.55);
  // Always keep LSL / USL inside the plot area with a small margin
  const margin = sigma * 0.4;
  if (lsl != null) xMin = Math.min(xMin, lsl - margin);
  if (usl != null) xMax = Math.max(xMax, usl + margin);

  // ── Histogram bins ───────────────────────────────────────────────────
  const nBins = Math.min(30, Math.max(8, Math.ceil(Math.sqrt(rawVals.length))));
  const binW = (xMax - xMin) / nBins;
  const bins = new Array(nBins).fill(0);
  rawVals.forEach(v => {{
    const idx = Math.min(nBins - 1, Math.max(0, Math.floor((v - xMin) / binW)));
    bins[idx]++;
  }});
  const maxCount = Math.max(...bins);

  // ── Normal curve ─────────────────────────────────────────────────────
  const nPts = 200;
  const step = (xMax - xMin) / nPts;
  const curve = [];
  let maxPDF = 0;
  for (let i = 0; i <= nPts; i++) {{
    const x = xMin + i * step;
    const y = normalPDF(x, mu, sigma);
    curve.push({{x, y}});
    if (y > maxPDF) maxPDF = y;
  }}

  // Scale: histogram bar tops align with the PDF peak
  const scaleY = plotH / Math.max(maxPDF * 1.1, 0.001);
  const histScale = (maxPDF * 0.9) / (maxCount || 1);

  const toCanvasX = x => padL + ((x - xMin) / (xMax - xMin)) * plotW;
  const toCanvasY = y => padT + plotH - y * scaleY;

  // ── Background ───────────────────────────────────────────────────────
  ctx.fillStyle = '#fafbfc';
  ctx.fillRect(0, 0, W, H);
  ctx.fillStyle = '#fff';
  ctx.fillRect(padL, padT, plotW, plotH);

  // ── Grid lines ───────────────────────────────────────────────────────
  ctx.strokeStyle = '#e0e0e0';
  ctx.lineWidth = 1;
  const nGridX = 6;
  for (let i = 0; i <= nGridX; i++) {{
    const x = padL + (i / nGridX) * plotW;
    ctx.beginPath(); ctx.moveTo(x, padT); ctx.lineTo(x, padT + plotH); ctx.stroke();
  }}

  // ── Histogram bars ───────────────────────────────────────────────────
  for (let i = 0; i < nBins; i++) {{
    const bx = toCanvasX(xMin + i * binW);
    const bx2 = toCanvasX(xMin + (i + 1) * binW);
    const bh = bins[i] * histScale * scaleY;
    ctx.fillStyle = 'rgba(70, 130, 180, 0.35)';
    ctx.fillRect(bx, padT + plotH - bh, bx2 - bx - 1, bh);
    ctx.strokeStyle = 'rgba(70, 130, 180, 0.6)';
    ctx.lineWidth = 0.5;
    ctx.strokeRect(bx, padT + plotH - bh, bx2 - bx - 1, bh);
  }}

  // ── Normal curve ─────────────────────────────────────────────────────
  ctx.beginPath();
  ctx.strokeStyle = '#1a73e8';
  ctx.lineWidth = 2.5;
  curve.forEach((pt, i) => {{
    const cx = toCanvasX(pt.x);
    const cy = toCanvasY(pt.y);
    if (i === 0) ctx.moveTo(cx, cy); else ctx.lineTo(cx, cy);
  }});
  ctx.stroke();

  // ── Limit lines ──────────────────────────────────────────────────────
  function drawVLine(xVal, color, label) {{
    if (xVal == null) return;
    const cx = toCanvasX(xVal);
    ctx.beginPath();
    ctx.setLineDash([6, 4]);
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.moveTo(cx, padT);
    ctx.lineTo(cx, padT + plotH);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = color;
    ctx.font = 'bold 11px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(label, cx, padT - 8);
    ctx.font = '10px sans-serif';
    ctx.fillText(_fmt4(xVal), cx, padT + plotH + 14);
  }}

  drawVLine(lsl, '#e53935', 'LSL');
  drawVLine(usl, '#e53935', 'USL');

  // Mean line
  const muX = toCanvasX(mu);
  ctx.beginPath();
  ctx.setLineDash([5, 3]);
  ctx.strokeStyle = '#43a047';
  ctx.lineWidth = 1.5;
  ctx.moveTo(muX, padT);
  ctx.lineTo(muX, padT + plotH);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.fillStyle = '#43a047';
  ctx.font = 'bold 11px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('μ', muX, padT - 8);

  // ── X axis ────────────────────────────────────────────────────────────
  ctx.strokeStyle = '#333';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  ctx.fillStyle = '#444';
  ctx.font = '10px sans-serif';
  ctx.textAlign = 'center';
  for (let i = 0; i <= nGridX; i++) {{
    const xVal = xMin + (i / nGridX) * (xMax - xMin);
    const cx = padL + (i / nGridX) * plotW;
    ctx.fillText(_fmt4(xVal), cx, padT + plotH + 28);
  }}

  // ── Point name label ─────────────────────────────────────────────────
  ctx.fillStyle = '#333';
  ctx.font = 'bold 13px sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText(pointName, padL, padT - 14);

  _updateChartInfo(stationType, sheetName, pt, pointName);
}}

function _fmt4(v) {{
  if (v == null) return '-';
  const n = parseFloat(v);
  if (isNaN(n)) return String(v);
  // Use up to 4 significant figures, no trailing zeros
  return parseFloat(n.toPrecision(4)).toString();
}}

function _updateChartInfo(stationType, sheetName, pt, pointName) {{
  const infoId = 'chartinfo_' + escId(stationType) + '_' + escId(sheetName);
  const el = document.getElementById(infoId);
  if (!el) return;
  const cpk = pt.cpk != null ? pt.cpk.toFixed(4) : '-';
  const cpkColor = pt.cpk == null ? '#666'
    : pt.cpk >= 1.33 ? '#2e7d32'
    : pt.cpk >= 1.0 ? '#e65100'
    : '#c62828';
  // Cpk label only visible when user has typed "cpk" in the search box
  const showCpk = cpkModeActive[stationType + '|' + sheetName] || false;
  el.innerHTML = `
    <span style="margin-right:16px"><b>测试项目:</b> ${{pointName}}</span>
    <span style="margin-right:16px"><b>N=</b>${{pt.n}}</span>
    <span style="margin-right:16px"><b>μ=</b>${{pt.mean != null ? pt.mean.toFixed(4) : '-'}}</span>
    <span style="margin-right:16px"><b>σ=</b>${{pt.std != null ? pt.std.toFixed(4) : '-'}}</span>
    <span class="cpk-chartinfo"
          style="display:${{showCpk ? 'inline' : 'none'}};color:${{cpkColor}};font-weight:bold">
      <b>Cpk=</b>${{cpk}}
    </span>
  `;
}}

// ── Range search ─────────────────────────────────────────────────────────
function doRangeSearch(stationType, sheetName) {{
  const sid = escId(stationType), shid = escId(sheetName);
  const ptSel = document.getElementById('range_point_' + sid + '_' + shid);
  const loEl = document.getElementById('range_lo_' + sid + '_' + shid);
  const hiEl = document.getElementById('range_hi_' + sid + '_' + shid);

  const pointName = ptSel ? ptSel.value : '';
  const lo = parseFloat(loEl ? loEl.value : '');
  const hi = parseFloat(hiEl ? hiEl.value : '');

  if (!pointName) {{ alert('请选择测试子项目'); return; }}
  if (isNaN(lo) || isNaN(hi)) {{ alert('请输入有效的数值范围'); return; }}

  const shData = ALL_DATA[stationType] && ALL_DATA[stationType][sheetName];
  if (!shData || !shData[pointName]) {{ alert('未找到数据'); return; }}

  const pt = shData[pointName];
  const isExact = lo === hi;
  const matched = [];
  pt.barcodes.forEach((bc, i) => {{
    const v = pt.raw[i];
    const hit = isExact
      ? Math.abs(v - lo) <= Math.abs(lo) * 1e-9 + 1e-12   // float-safe exact
      : v >= lo && v <= hi;
    if (hit) matched.push({{bc, val: v}});
  }});
  // For exact match, sort by value for easier reading
  if (isExact) matched.sort((a, b) => a.val - b.val);

  let html = isExact
    ? `<p>精确匹配: <b>${{pointName}}</b> = <b>${{lo}}</b> &nbsp;|&nbsp; 命中: <b>${{matched.length}}</b> 个</p>`
    : `<p>范围搜索: <b>${{pointName}}</b> &nbsp; [<b>${{lo}}</b> ~ <b>${{hi}}</b>] &nbsp;|&nbsp; 命中: <b>${{matched.length}}</b> 个</p>`;
  if (matched.length === 0) {{
    html += '<p style="color:#888">无匹配条码</p>';
  }} else {{
    html += '<table class="modal-table"><thead><tr><th>条码</th><th>测试值</th></tr></thead><tbody>';
    matched.forEach(item => {{
      html += `<tr><td>${{item.bc}}</td><td>${{item.val}}</td></tr>`;
    }});
    html += '</tbody></table>';
  }}

  document.getElementById('range-result').innerHTML = html;
  document.getElementById('range-modal').style.display = 'flex';
}}

// ── Window resize: redraw all tracked charts ─────────────────────────────
let _resizeTimer = null;
window.addEventListener('resize', function() {{
  clearTimeout(_resizeTimer);
  _resizeTimer = setTimeout(function() {{
    for (const [key, pname] of Object.entries(currentPoint)) {{
      const sep = key.indexOf('|');
      const stype = key.slice(0, sep);
      const sheet = key.slice(sep + 1);
      drawChart(stype, sheet, pname);
    }}
  }}, 120);
}});

// ── Init: script is at end of <body>, DOM is already parsed ──────────────
(function() {{
  for (const [stype, sheets] of Object.entries(ALL_DATA)) {{
    for (const [sheet, points] of Object.entries(sheets)) {{
      const pts = Object.keys(points);
      if (pts.length > 0) drawChart(stype, sheet, pts[0]);
    }}
  }}
}})();
</script>
</body>
</html>"""


def _esc_js(s: str) -> str:
    return s.replace("'", "\\'").replace('"', '\\"')


def _esc_id(s: str) -> str:
    # Replace special chars so string can be used as HTML id
    return s.replace(' ', '_').replace('/', '_').replace('\\', '_')


def _build_station_panel(stype: str, sheets: dict, visible: bool) -> str:
    display = 'block' if visible else 'none'
    panel_id = f'panel-{_esc_id(stype)}'

    sheet_names = list(sheets.keys())
    sheet_tabs_nav = ''
    sheet_tabs_content = ''

    for j, sheet in enumerate(sheet_names):
        sh_active = 'active' if j == 0 else ''
        btn_id = f'shbtn_{_esc_id(stype)}_{_esc_id(sheet)}'
        panel_id_sh = f'shpanel_{_esc_id(stype)}_{_esc_id(sheet)}'

        sheet_tabs_nav += (
            f'<button class="shtab-btn {sh_active}" '
            f'id="{btn_id}" '
            f'data-station="{stype}" '
            f'onclick="switchSheet(\'{_esc_js(stype)}\', \'{_esc_js(sheet)}\')">'
            f'{sheet}</button>\n'
        )
        sheet_tabs_content += _build_sheet_panel(
            stype, sheet, sheets[sheet], visible=(j == 0)
        )

    return f'''
<div class="station-panel" id="{panel_id}" style="display:{display}">
  <div class="shtab-bar">
{sheet_tabs_nav}
  </div>
{sheet_tabs_content}
</div>
'''


def _build_sheet_panel(stype: str, sheet: str, points: dict, visible: bool) -> str:
    display = 'block' if visible else 'none'
    sid = _esc_id(stype)
    shid = _esc_id(sheet)
    panel_id = f'shpanel_{sid}_{shid}'
    table_id = f'table_{sid}_{shid}'
    search_id = f'search_{sid}_{shid}'
    chart_id = f'chart_{sid}_{shid}'
    chart_info_id = f'chartinfo_{sid}_{shid}'
    range_point_id = f'range_point_{sid}_{shid}'
    range_lo_id = f'range_lo_{sid}_{shid}'
    range_hi_id = f'range_hi_{sid}_{shid}'

    # Table rows
    rows_html = ''
    for pname, stats in points.items():
        cpk_val = stats.get('cpk')
        cpk_color = ''
        if cpk_val is not None:
            if cpk_val >= 1.33:
                cpk_color = 'color:#2e7d32;font-weight:bold'
            elif cpk_val >= 1.0:
                cpk_color = 'color:#e65100;font-weight:bold'
            else:
                cpk_color = 'color:#c62828;font-weight:bold'

        rows_html += f'''<tr data-point="{pname}" onclick="onRowClick('{_esc_js(stype)}','{_esc_js(sheet)}','{_esc_js(pname)}')" style="cursor:pointer">
  <td>{pname}</td>
  <td>{stats.get("n", "-")}</td>
  <td>{_fmt(stats.get("mean"))}</td>
  <td>{_fmt(stats.get("std"))}</td>
  <td>{_fmt(stats.get("min"))}</td>
  <td>{_fmt(stats.get("max"))}</td>
  <td>{_fmt(stats.get("lsl"))}</td>
  <td>{_fmt(stats.get("usl"))}</td>
  <td class="cpk-col" style="display:none">{_fmt(stats.get("cp"))}</td>
  <td class="cpk-col" style="display:none">{_fmt(stats.get("cpl"))}</td>
  <td class="cpk-col" style="display:none">{_fmt(stats.get("cpu"))}</td>
  <td class="cpk-col" style="display:none;{cpk_color}">{_fmt(stats.get("cpk"))}</td>
</tr>'''

    # Point name options for range search
    option_items = ''.join(
        f'<option value="{pname}">{pname}</option>'
        for pname in points.keys()
    )

    return f'''
<div class="sheet-panel" id="{panel_id}" data-station="{stype}" style="display:{display}">
  <!-- Search bar -->
  <div class="search-bar">
    <div class="search-group">
      <label>搜索测试项目：</label>
      <input type="text" id="{search_id}"
             placeholder="搜索测试项目..."
             oninput="onSearch('{_esc_js(stype)}', '{_esc_js(sheet)}')" />
    </div>
    <div class="search-group range-group">
      <label>数值范围搜索：</label>
      <select id="{range_point_id}">{option_items}</select>
      <input type="number" id="{range_lo_id}" placeholder="最小值（相等则精确匹配）" step="any" style="width:160px"/>
      <span>~</span>
      <input type="number" id="{range_hi_id}" placeholder="最大值" step="any" style="width:90px"/>
      <button class="btn btn-sm" onclick="doRangeSearch('{_esc_js(stype)}', '{_esc_js(sheet)}')">查询条码</button>
    </div>
  </div>

  <!-- Data table -->
  <div class="table-wrapper">
    <table id="{table_id}" class="data-table">
      <thead>
        <tr>
          <th>测试子项目</th>
          <th>样本数</th>
          <th>均值</th>
          <th>标准差</th>
          <th>最小值</th>
          <th>最大值</th>
          <th>下限 LSL</th>
          <th>上限 USL</th>
          <th class="cpk-col" style="display:none">Cp</th>
          <th class="cpk-col" style="display:none">Cpl</th>
          <th class="cpk-col" style="display:none">Cpu</th>
          <th class="cpk-col" style="display:none">Cpk</th>
        </tr>
      </thead>
      <tbody>
{rows_html}
      </tbody>
    </table>
  </div>

  <!-- Normal distribution chart -->
  <div class="chart-section">
    <div class="chart-tip">点击上方表格行切换查看对应测试项目的正态分布图</div>
    <div id="{chart_info_id}" class="chart-info"></div>
    <canvas id="{chart_id}" height="320" style="width:100%;display:block"></canvas>
  </div>
</div>
'''



# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #222; font-size: 13px; }

.header {
  background: linear-gradient(135deg, #1a237e 0%, #283593 100%);
  color: white;
  padding: 14px 28px;
}
.header-title { font-size: 20px; font-weight: bold; letter-spacing: 1px; }
.header-sub { font-size: 11px; opacity: 0.75; margin-top: 4px; }

/* Station tabs */
.stab-bar {
  display: flex; flex-wrap: wrap; gap: 4px;
  background: #283593; padding: 6px 16px 0;
}
.stab-btn {
  padding: 7px 20px; border: none; cursor: pointer;
  background: rgba(255,255,255,0.15); color: white;
  border-radius: 4px 4px 0 0; font-size: 13px; font-weight: 500;
  transition: background 0.2s;
}
.stab-btn:hover { background: rgba(255,255,255,0.25); }
.stab-btn.active { background: #f0f2f5; color: #1a237e; font-weight: bold; }

/* Sheet tabs */
.shtab-bar {
  display: flex; flex-wrap: wrap; gap: 3px;
  background: #e8eaf6; padding: 6px 16px 0; border-bottom: 2px solid #c5cae9;
}
.shtab-btn {
  padding: 5px 16px; border: none; cursor: pointer;
  background: #c5cae9; color: #3949ab;
  border-radius: 4px 4px 0 0; font-size: 12px;
  transition: background 0.15s;
}
.shtab-btn:hover { background: #9fa8da; }
.shtab-btn.active { background: #fff; color: #1a237e; font-weight: bold; border-bottom: 2px solid #fff; }

.station-panel { background: #f0f2f5; }
.sheet-panel { padding: 14px 16px; }

/* Search bar */
.search-bar {
  display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
  background: #fff; border-radius: 6px; padding: 10px 14px;
  margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.search-group { display: flex; align-items: center; gap: 6px; }
.search-bar label { font-weight: 600; color: #444; white-space: nowrap; }
.search-bar input[type="text"] {
  padding: 5px 10px; border: 1px solid #bbb; border-radius: 4px; font-size: 13px; width: 240px;
}
.search-bar input[type="number"] {
  padding: 5px 6px; border: 1px solid #bbb; border-radius: 4px; font-size: 13px;
}
.search-bar select {
  padding: 5px 6px; border: 1px solid #bbb; border-radius: 4px; font-size: 13px; max-width: 220px;
}

/* Table */
.table-wrapper { overflow-x: auto; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); margin-bottom: 16px; }
.data-table { width: 100%; border-collapse: collapse; background: #fff; }
.data-table thead tr { background: #3949ab; color: white; }
.data-table th { padding: 9px 10px; text-align: center; font-size: 12px; white-space: nowrap; }
.data-table td { padding: 7px 10px; text-align: right; border-bottom: 1px solid #e8eaf6; font-size: 12px; white-space: nowrap; }
.data-table td:first-child { text-align: left; font-weight: 500; color: #1a237e; }
.data-table tbody tr:hover { background: #e8eaf6; }
.data-table tbody tr.selected { background: #c5cae9 !important; }

/* Chart section */
.chart-section {
  background: #fff; border-radius: 6px; padding: 14px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.1);
}
.chart-tip { font-size: 11px; color: #888; margin-bottom: 6px; }
.chart-info { font-size: 12px; color: #333; margin-bottom: 8px; min-height: 20px; }
canvas { display: block; max-width: 100%; }

/* Buttons */
.btn {
  padding: 6px 16px; border: none; border-radius: 4px;
  cursor: pointer; font-size: 13px; font-weight: 500;
  background: #3949ab; color: white; transition: background 0.2s;
}
.btn:hover { background: #283593; }
.btn-sm { padding: 4px 12px; font-size: 12px; }

/* Modal */
.modal-overlay {
  position: fixed; inset: 0; background: rgba(0,0,0,0.45);
  display: flex; align-items: center; justify-content: center; z-index: 1000;
}
.modal-box {
  background: #fff; border-radius: 8px; padding: 20px 24px;
  max-width: 600px; width: 90%; max-height: 80vh; overflow-y: auto;
  box-shadow: 0 8px 32px rgba(0,0,0,0.2);
}
.modal-title { font-size: 16px; font-weight: bold; color: #1a237e; margin-bottom: 12px; }
.modal-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 13px; }
.modal-table th { background: #e8eaf6; padding: 6px 10px; text-align: left; }
.modal-table td { padding: 5px 10px; border-bottom: 1px solid #eee; }
"""