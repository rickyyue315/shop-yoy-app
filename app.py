# -*- coding: utf-8 -*-
"""
Shop YoY 分析系統
=================

用法：
    python app.py                     # 本機執行（預設 http://127.0.0.1:5000）
    PORT=8080 gunicorn app:app        # 生產環境（Docker / Zeabur）

功能：
    1. 上載 Hugo_by Shop_YoY 匯出的 Excel
       （欄位：Month / Shop / Sold Qty / Net Sales / COGS / GP$ / GP%）
    2. 自動對照內置 stores-template.csv（店鋪主檔），補上 Regional / Class / Size / OM / Type 等屬性
    3. 依 Site、年份、月份（單月或多月累計）計算 YoY%，支援 Regional / OM 等進階篩選及分組
    4. 匯出 Excel 報告
"""

import csv
import io
import os
import re
import threading
import time
import uuid
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from flask import Flask, Response, jsonify, render_template, request, send_file

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_CSV = os.path.join(BASE_DIR, 'stores-template.csv')

MONTH_RE = re.compile(r'^(\d{4})-(0[1-9]|1[0-2])$')

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 上載上限 32 MB

# 私隱設定（可用環境變數覆寫）：
#   APP_PASSWORD    設定後，開啟網站需輸入密碼（用戶名稱隨意，密碼 = 此值）
#   DATA_TTL_HOURS  上載資料於記憶體保留幾多小時後自動刪除（預設 8；0 = 不自動過期）
APP_PASSWORD = (os.environ.get('APP_PASSWORD') or '').strip()
try:
    DATA_TTL_HOURS = float(os.environ.get('DATA_TTL_HOURS', '8'))
except ValueError:
    DATA_TTL_HOURS = 8.0

# 已上載資料（只存記憶體，不寫入硬碟／資料庫），保留最近數份
_DATASETS = {}
_DATASET_ORDER = []
_TIMES = {}
_DATASETS_MAX = 5
_LOCK = threading.Lock()

GROUP_FIELDS = {
    'shop': None,
    'regional': 'regional',
    'om': 'om',
    'class1': 'class1',
    'class2': 'class2',
    'size': 'size',
    'type': 'type',
}
GROUP_LABELS = {
    'shop': '店鋪',
    'regional': 'Regional',
    'om': 'OM',
    'class1': 'Class 1',
    'class2': 'Class 2',
    'size': 'Size',
    'type': 'Type',
}

# 精簡 SKU 店鋪（進階篩選快速選用；如日後名單有變，改呢個清單即可）
LEAN_SKU_SITES = {
    'HA19', 'HA21', 'HA33', 'HA37', 'HB30', 'HC44', 'HB86', 'HB77', 'HB91',
    'HB94', 'HC26', 'HC60', 'HC62', 'HC68', 'HC13', 'HC69', 'HB87', 'HB80',
}

# 按月矩陣的可用指標：metric -> (顯示種類, 標籤)
MATRIX_METRICS = {
    'salesYoy': ('yoy', '淨銷售 YoY%'),
    'qtyYoy': ('yoy', '銷量 YoY%'),
    'gpYoy': ('yoy', '毛利 GP$ YoY%'),
    'gpPp': ('pp', '毛利率 ±pp'),
    'salesVal': ('money', '淨銷售'),
    'qtyVal': ('qty', '銷量'),
    'gpVal': ('money', '毛利 GP$'),
    'gpPctVal': ('pct', '毛利率（值）'),
}


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------

def _norm(v):
    """正規化欄位名：去空白／底線、轉小寫。"""
    return re.sub(r'[\s_]+', '', str(v or '')).lower()


def _num(v):
    """把儲存格值轉成 float；空白或非數字一律當 0。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _month_str(v):
    """把 Month 欄位值轉成 'YYYY-MM'；不是有效月份就回傳空字串。"""
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m')
    s = str(v).strip()
    return s if MONTH_RE.match(s) else ''


def _period_label(months):
    """把月份清單化成易讀字串，如 '2026-01 至 2026-09'。"""
    if not months:
        return '（無）'
    nums = sorted({int(m[5:7]) for m in months})
    year = sorted(months)[0][:4]
    ranges = []
    s = p = nums[0]
    for n in nums[1:]:
        if n == p + 1:
            p = n
        else:
            ranges.append((s, p))
            s = p = n
    ranges.append((s, p))
    parts = []
    for a, b in ranges:
        if a == b:
            parts.append('%s-%02d' % (year, a))
        elif b == a + 1:
            parts.append('%s-%02d、%s-%02d' % (year, a, year, b))
        else:
            parts.append('%s-%02d 至 %s-%02d' % (year, a, year, b))
    return '、'.join(parts)


def load_template():
    """讀取內置店鋪主檔 stores-template.csv（每次上載時重新讀取）。"""
    tpl = {}
    if not os.path.exists(TEMPLATE_CSV):
        return tpl
    with open(TEMPLATE_CSV, encoding='utf-8-sig', newline='') as f:
        for r in csv.DictReader(f):
            code = (r.get('Site') or '').strip()
            if not code:
                continue
            tpl[code] = {
                'name': (r.get('Shop') or '').strip(),
                'regional': (r.get('Regional') or '').strip(),
                'class1': (r.get('Class 1') or '').strip(),
                'class2': (r.get('Class 2') or '').strip(),
                'size': (r.get('Size') or '').strip(),
                'om': (r.get('OM') or '').strip(),
                'type': (r.get('Type') or '').strip(),
                'focus': bool((r.get('需特別關注') or '').strip()),
            }
    return tpl


# ---------------------------------------------------------------------------
# Excel 解析
# ---------------------------------------------------------------------------

def parse_excel(fileobj):
    """
    解析 Hugo_by Shop_YoY 匯出 Excel，回傳 dataset：
        {
          'months': ['2024-10', ...],
          'years':  {'2024': [...], ...},
          'sites':  [{code, name, regional, class1, class2, size, om, type, focus, inTemplate}, ...],
          'rows':   [[monthIdx, siteIdx, qty, sales, gp], ...],
        }
    """
    try:
        wb = openpyxl.load_workbook(fileobj, read_only=True, data_only=True)
    except Exception:
        raise ValueError('無法讀取檔案，請確認為有效的 Excel (.xlsx) 檔案。')

    try:
        sheet = header_i = colmap = None
        for sh in wb.worksheets:
            if getattr(sh, 'sheet_state', 'visible') != 'visible':
                continue
            for ri, row in enumerate(sh.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                cells = {}
                for ci, cv in enumerate(row or ()):
                    if cv is None:
                        continue
                    cells[_norm(cv)] = ci
                if 'month' in cells and 'shop' in cells and 'netsales' in cells:
                    sheet, header_i, colmap = sh, ri, cells
                    break
            if sheet is not None:
                break

        if sheet is None:
            raise ValueError(
                '找不到資料表：檔案需包含 Month、Shop、Net Sales 欄位（Hugo_by Shop_YoY 匯出格式）。')

        i_month = colmap['month']
        i_shop = colmap['shop']
        i_sales = colmap['netsales']
        i_qty = colmap.get('soldqty')
        i_cogs = colmap.get('cogs')
        i_gp = colmap.get('gp$')

        def cell(row, i):
            return row[i] if (i is not None and i < len(row)) else None

        agg = {}          # (month, code) -> [qty, sales, gp]
        excel_names = {}  # code -> 店名（取自 Excel 的 'CODE - 店名' 格式）
        months_set = set()

        for row in sheet.iter_rows(min_row=header_i + 1, values_only=True):
            if row is None or len(row) <= max(i_month, i_shop):
                continue
            m = _month_str(row[i_month])
            if not m:
                continue  # 排除 Total 行、備註行等
            shop_raw = row[i_shop]
            if shop_raw is None or not str(shop_raw).strip():
                continue
            shop = str(shop_raw).strip()
            if ' - ' in shop:
                code, name = shop.split(' - ', 1)
                code, name = code.strip(), name.strip()
            else:
                code, name = shop, ''
            if code not in excel_names and name:
                excel_names[code] = name

            qty = _num(cell(row, i_qty))
            sales = _num(cell(row, i_sales))
            if i_gp is not None:
                gp = _num(cell(row, i_gp))
            else:
                gp = sales - _num(cell(row, i_cogs))

            key = (m, code)
            t = agg.get(key)
            if t is None:
                agg[key] = [qty, sales, gp]
            else:
                t[0] += qty
                t[1] += sales
                t[2] += gp
            months_set.add(m)

        if not agg:
            raise ValueError('找不到有效的資料列（Month 欄位需為 YYYY-MM 格式）。')

        months = sorted(months_set)
        m_idx = {m: i for i, m in enumerate(months)}

        # 對照店鋪主檔
        tpl = load_template()
        sites = []
        s_idx = {}
        for code in sorted({c for _, c in agg}):
            t = tpl.get(code)
            sites.append({
                'code': code,
                'name': (t or {}).get('name') or excel_names.get(code, '') or code,
                'regional': (t or {}).get('regional', ''),
                'class1': (t or {}).get('class1', ''),
                'class2': (t or {}).get('class2', ''),
                'size': (t or {}).get('size', ''),
                'om': (t or {}).get('om', ''),
                'type': (t or {}).get('type', ''),
                'focus': bool((t or {}).get('focus', False)),
                'inTemplate': t is not None,
                'leanSku': code in LEAN_SKU_SITES,
            })
            s_idx[code] = len(sites) - 1

        rows = []
        for (m, code), (qty, sales, gp) in agg.items():
            rows.append([
                m_idx[m], s_idx[code],
                int(round(qty)),
                round(sales, 2),
                round(gp, 2),
            ])
        rows.sort(key=lambda r: (r[0], r[1]))

        years = {}
        for m in months:
            years.setdefault(m[:4], []).append(m)

        return {
            'months': months,
            'years': years,
            'sites': sites,
            'rows': rows,
            'latestMonth': months[-1],
            'rowCount': len(rows),
            'siteCount': len(sites),
        }
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# YoY 計算（匯出用；與前端 index.html 的邏輯一致）
# ---------------------------------------------------------------------------

def _select_groups(ds, opts):
    """依篩選條件建立分組（店鋪或屬性分組）。

    回傳 (groups, site_group, pass_count)：
      groups      key -> 組別資料（含 label / n / cur / ly）
      site_group  site index -> group key（未通過篩選的不在此）
      pass_count  通過篩選的店鋪數
    """
    sites = ds['sites']
    groupby = opts.get('groupBy') or 'shop'
    gfield = GROUP_FIELDS.get(groupby)

    sel_codes = opts.get('siteCodes')
    sel_codes = {str(c) for c in sel_codes} if sel_codes else None
    fsets = {k: {str(v) for v in (opts.get(k) or [])}
             for k in ('regional', 'om', 'class1', 'class2', 'size', 'type')}

    def site_pass(s):
        if sel_codes is not None and s['code'] not in sel_codes:
            return False
        if opts.get('templateOnly') and not s['inTemplate']:
            return False
        if opts.get('focusOnly') and not s['focus']:
            return False
        if opts.get('leanSkuOnly') and not s.get('leanSku'):
            return False
        for k, vals in fsets.items():
            if vals and (s.get(k) or '') not in vals:
                return False
        return True

    groups = {}
    site_group = {}
    pass_count = 0
    for i, s in enumerate(sites):
        if not site_pass(s):
            continue
        pass_count += 1
        if gfield is None:
            key = 'site:' + s['code']
            groups[key] = {
                'label': (s['code'] + (' ' + s['name'] if s['name'] else '')).strip(),
                'code': s['code'], 'name': s['name'],
                'regional': s['regional'], 'om': s['om'],
                'class1': s['class1'], 'class2': s['class2'],
                'size': s['size'], 'type': s['type'], 'focus': s['focus'],
                'n': 1, 'cur': [0.0, 0.0, 0.0], 'ly': [0.0, 0.0, 0.0],
            }
        else:
            val = (s.get(gfield) or '').strip() or '（未分類）'
            key = gfield + ':' + val
            g = groups.get(key)
            if g is None:
                g = {'label': val, 'n': 0, 'cur': [0.0, 0.0, 0.0], 'ly': [0.0, 0.0, 0.0]}
                groups[key] = g
            g['n'] += 1
        site_group[i] = key
    return groups, site_group, pass_count


def compute_report(ds, opts):
    """累計模式：所選月份合併，與去年同期比較。"""
    year = int(opts.get('year'))
    lyear = year - 1
    months = {str(m) for m in (opts.get('months') or [])}
    month_nums = {m[5:7] for m in months}
    months_list = ds['months']

    groups, site_group, pass_count = _select_groups(ds, opts)

    ystr, lystr = str(year), str(lyear)
    for mi, si, qty, sales, gp in ds['rows']:
        key = site_group.get(si)
        if key is None:
            continue
        m = months_list[mi]
        if m[:4] == ystr and m in months:
            target = groups[key]['cur']
        elif m[:4] == lystr and m[5:7] in month_nums:
            target = groups[key]['ly']
        else:
            continue
        target[0] += qty
        target[1] += sales
        target[2] += gp

    hide_empty = opts.get('hideEmpty', True)
    out = []
    for g in groups.values():
        if hide_empty and not any(g['cur']) and not any(g['ly']):
            continue
        out.append(g)
    out.sort(key=lambda g: (-g['cur'][1], g['label']))

    totals = {'cur': [0.0, 0.0, 0.0], 'ly': [0.0, 0.0, 0.0]}
    for g in out:
        for j in range(3):
            totals['cur'][j] += g['cur'][j]
            totals['ly'][j] += g['ly'][j]

    return {'groups': out, 'totals': totals, 'siteCount': pass_count}


def compute_matrix(ds, opts):
    """按月矩陣：每個組別 × 每個所選月份（本期與去年同期分開存放）。"""
    year = int(opts.get('year'))
    lyear = year - 1
    months = sorted({str(m) for m in (opts.get('months') or [])})
    ly_of = {('%d-%s' % (lyear, m[5:7])): m for m in months}

    groups, site_group, pass_count = _select_groups(ds, opts)
    for g in groups.values():
        g['cells'] = {m: [0.0, 0.0, 0.0] for m in months}
        g['cellsLy'] = {m: [0.0, 0.0, 0.0] for m in months}
        g['aggCur'] = [0.0, 0.0, 0.0]
        g['aggLy'] = [0.0, 0.0, 0.0]

    months_list = ds['months']
    for mi, si, qty, sales, gp in ds['rows']:
        key = site_group.get(si)
        if key is None:
            continue
        m = months_list[mi]
        if m in months:
            cell = groups[key]['cells'][m]
        elif m in ly_of:
            cell = groups[key]['cellsLy'][ly_of[m]]  # 去年同月掛到同一欄
        else:
            continue
        cell[0] += qty
        cell[1] += sales
        cell[2] += gp

    for g in groups.values():
        for m in months:
            for j in range(3):
                g['aggCur'][j] += g['cells'][m][j]
                g['aggLy'][j] += g['cellsLy'][m][j]

    hide_empty = opts.get('hideEmpty', True)
    out = []
    for g in groups.values():
        if hide_empty and not any(g['aggCur']) and not any(g['aggLy']):
            continue
        out.append(g)
    out.sort(key=lambda g: (-g['aggCur'][1], g['label']))

    totals = {
        'cells': {m: [0.0, 0.0, 0.0] for m in months},
        'cellsLy': {m: [0.0, 0.0, 0.0] for m in months},
        'aggCur': [0.0, 0.0, 0.0],
        'aggLy': [0.0, 0.0, 0.0],
    }
    for g in out:
        for m in months:
            for j in range(3):
                totals['cells'][m][j] += g['cells'][m][j]
                totals['cellsLy'][m][j] += g['cellsLy'][m][j]
        for j in range(3):
            totals['aggCur'][j] += g['aggCur'][j]
            totals['aggLy'][j] += g['aggLy'][j]

    return {'groups': out, 'months': months, 'totals': totals, 'siteCount': pass_count}


def _matrix_value(cur, ly, metric):
    """按月矩陣的單格數值。cur / ly = [qty, sales, gp]；無法計算回傳 None。"""
    if metric == 'qtyYoy':
        return ((cur[0] - ly[0]) / ly[0]) if ly[0] > 0 else None
    if metric == 'gpYoy':
        return ((cur[2] - ly[2]) / ly[2]) if ly[2] > 0 else None
    if metric == 'gpPp':
        pc = (cur[2] / cur[1]) if cur[1] else None
        pl = (ly[2] / ly[1]) if ly[1] else None
        return (pc - pl) if (pc is not None and pl is not None) else None
    if metric == 'salesVal':
        return cur[1]
    if metric == 'qtyVal':
        return cur[0]
    if metric == 'gpVal':
        return cur[2]
    if metric == 'gpPctVal':
        return (cur[2] / cur[1]) if cur[1] else None
    # 預設：淨銷售 YoY%
    return ((cur[1] - ly[1]) / ly[1]) if ly[1] > 0 else None


def _filter_summary(opts):
    parts = ['Site：%s' % ('全部' if not opts.get('siteCodes') else '%d 間' % len(opts['siteCodes']))]
    for label, key in (('Regional', 'regional'), ('OM', 'om'), ('Class 1', 'class1'),
                       ('Class 2', 'class2'), ('Size', 'size'), ('Type', 'type')):
        vals = opts.get(key) or []
        if vals:
            parts.append('%s：%s' % (label, '、'.join(str(v) for v in vals)))
    if opts.get('focusOnly'):
        parts.append('只顯示需特別關注店鋪')
    if opts.get('leanSkuOnly'):
        parts.append('只顯示精簡SKU店鋪')
    if opts.get('templateOnly'):
        parts.append('只顯示 stores-template 內店鋪')
    return '；'.join(parts)


def _write_title_block(ws, ds, opts, months, ly_months, extra_line=''):
    """匯出檔的標題區（A1-A6）。"""
    year = int(opts['year'])
    ws['A1'] = 'Shop YoY 分析報告'
    ws['A1'].font = Font(size=16, bold=True, color='0F172A')
    ws['A2'] = '資料檔：%s（上載時間：%s）' % (ds.get('filename', ''), ds.get('uploadedAt', ''))
    ws['A2'].font = Font(size=10, color='475569')
    ws['A3'] = '比較期間：%s  vs  %s' % (_period_label(months), _period_label(ly_months))
    ws['A3'].font = Font(size=11, bold=True, color='0F172A')

    groupby = opts.get('groupBy') or 'shop'
    line = '分組：%s' % GROUP_LABELS.get(groupby, groupby)
    if extra_line:
        line += '　|　' + extra_line
    line += '　|　' + _filter_summary(opts)
    ws['A4'] = line
    ws['A4'].font = Font(size=10, color='475569')
    ws['A5'] = '產生時間：%s' % datetime.now().strftime('%Y-%m-%d %H:%M')
    ws['A5'].font = Font(size=10, color='475569')

    latest = ds.get('latestMonth', '')
    if latest and latest[:4] == str(year) and latest in months:
        ws['A6'] = '備註：%s 可能為不完整月份（若為進行中月份），YoY 數字請斟酌參考。' % latest
        ws['A6'].font = Font(size=10, bold=True, color='B45309')


def _row_metrics(g):
    """由組別資料計算 12 個顯示值：本期/去年/YoY% × (淨銷售、銷量、GP$)，再加 GP% 與 ±pp。"""
    sc, qc, gc = g['cur'][1], g['cur'][0], g['cur'][2]
    sl, ql, gl = g['ly'][1], g['ly'][0], g['ly'][2]
    pc = (gc / sc) if sc else None
    pl = (gl / sl) if sl else None
    y = lambda c, l: ((c - l) / l) if l and l > 0 else None
    return [
        sc, sl, y(sc, sl),
        qc, ql, y(qc, ql),
        gc, gl, y(gc, gl),
        pc, pl, (pc - pl) if (pc is not None and pl is not None) else None,
    ]


# ---------------------------------------------------------------------------
# Excel 匯出
# ---------------------------------------------------------------------------

def build_workbook(ds, opts, report):
    year = int(opts['year'])
    lyear = year - 1
    groupby = opts.get('groupBy') or 'shop'
    months = sorted(str(m) for m in (opts.get('months') or []))
    ly_months = ['%d-%s' % (lyear, m[5:7]) for m in months]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'YoY'

    thin = Side(style='thin', color='D9E1EC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(bold=True, color='FFFFFF', size=11)
    head_fill = PatternFill('solid', fgColor='0F172A')
    sub_fill = PatternFill('solid', fgColor='F1F5F9')
    total_fill = PatternFill('solid', fgColor='EEF2FF')
    na_font = Font(color='94A3B8')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')
    right = Alignment(horizontal='right', vertical='center')

    FMT_QTY = '#,##0'
    FMT_MONEY = '#,##0'
    FMT_PCT = '0.00%'
    FMT_YOY = '+0.0%;-0.0%;0.0%'
    FMT_PP = '+0.00" pp";-0.00" pp";0.00" pp"'

    # ---- 標題區 ----
    _write_title_block(ws, ds, opts, months, ly_months)

    header_row = 8

    # ---- 欄位定義 ----
    metric_groups = [
        ('淨銷售', ['本期', '去年同期', 'YoY%']),
        ('銷量', ['本期', '去年同期', 'YoY%']),
        ('毛利 GP$', ['本期', '去年同期', 'YoY%']),
        ('毛利率', ['本期', '去年', '± pp']),
    ]
    if groupby == 'shop':
        lead = [
            ('店鋪', 'label', 26), ('Site', 'text', 9), ('店名', 'text', 18),
            ('Regional', 'text', 9), ('OM', 'text', 9), ('Class 1', 'text', 8),
            ('Class 2', 'text', 8), ('Size', 'text', 7), ('Type', 'text', 6), ('需關注', 'text', 7),
        ]
    else:
        lead = [('分組', 'label', 24), ('店鋪數', 'qty', 9)]
    kinds = ['money', 'money', 'yoy', 'qty', 'qty', 'yoy', 'money', 'money', 'yoy', 'pct', 'pct', 'pp']
    kind_width = {'money': 15, 'qty': 11, 'yoy': 10, 'pct': 11, 'pp': 10, 'text': 9, 'label': 26}

    ncols = len(lead) + len(kinds)
    h1, h2 = header_row, header_row + 1

    # ---- 表頭 ----
    for rr in (h1, h2):
        for ci in range(1, ncols + 1):
            c = ws.cell(row=rr, column=ci)
            c.border = border
            c.alignment = center
            c.fill = head_fill if rr == h1 else sub_fill
            c.font = head_font if rr == h1 else Font(bold=True, color='334155')

    # 第一欄（及分組模式的店鋪數欄）合併兩行
    vmerged = [1] + ([2] if groupby != 'shop' else [])
    for ci in vmerged:
        ws.merge_cells(start_row=h1, start_column=ci, end_row=h2, end_column=ci)
        c = ws.cell(row=h1, column=ci)
        c.fill = head_fill
        c.font = head_font
        ws.cell(row=h2, column=ci).fill = head_fill

    ws.cell(row=h1, column=1, value=lead[0][0])
    col = 2
    if groupby == 'shop':
        ws.cell(row=h1, column=2, value='店鋪資料')
        ws.merge_cells(start_row=h1, start_column=2, end_row=h1, end_column=len(lead))
    else:
        ws.cell(row=h1, column=2, value='店鋪數')
    col = len(lead) + 1
    for gname, subs in metric_groups:
        ws.cell(row=h1, column=col, value=gname)
        if len(subs) > 1:
            ws.merge_cells(start_row=h1, start_column=col, end_row=h1, end_column=col + len(subs) - 1)
        col += len(subs)

    for ci, (name, kind, _w) in enumerate(lead, start=1):
        if ci in vmerged:
            continue
        ws.cell(row=h2, column=ci, value=name)
    col = len(lead) + 1
    for _gname, subs in metric_groups:
        for sname in subs:
            ws.cell(row=h2, column=col, value=sname)
            col += 1

    # 欄寬
    for ci, (name, kind, w) in enumerate(lead, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    col = len(lead) + 1
    for k in kinds:
        ws.column_dimensions[get_column_letter(col)].width = kind_width[k]
        col += 1

    # ---- 資料列 ----
    def write_cell(rr, ci, kind, v, bold=False, fill=None):
        c = ws.cell(row=rr, column=ci)
        c.border = border
        if fill:
            c.fill = fill
        if kind in ('label', 'text'):
            c.value = v if v is not None else ''
            c.font = Font(bold=bold)
            c.alignment = left if kind == 'label' else center
            return
        c.alignment = right
        if v is None:
            c.value = '—'
            c.font = na_font
            c.alignment = center
            return
        c.value = float(v)
        if kind == 'money':
            c.number_format = FMT_MONEY
        elif kind == 'qty':
            c.number_format = FMT_QTY
        elif kind == 'pct':
            c.number_format = FMT_PCT
        elif kind == 'yoy':
            c.number_format = FMT_YOY
        elif kind == 'pp':
            c.number_format = FMT_PP
        if kind in ('yoy', 'pp'):
            if v > 1e-9:
                c.font = Font(bold=bold, color='047857')
            elif v < -1e-9:
                c.font = Font(bold=bold, color='B91C1C')
            else:
                c.font = Font(bold=bold)
        else:
            c.font = Font(bold=bold)

    r = h2 + 1

    def emit(g, bold=False, fill=None):
        nonlocal r
        mvals = _row_metrics(g)
        if groupby == 'shop':
            lead_vals = [
                g['label'], g['code'], g['name'], g['regional'], g['om'],
                g['class1'], g['class2'], g['size'], g['type'],
                ('Y' if g.get('focus') else ''),
            ]
        else:
            lead_vals = [g['label'], g['n']]
        for ci, ((name, kind, _w), val) in enumerate(zip(lead, lead_vals), start=1):
            write_cell(r, ci, kind, val, bold=bold, fill=fill)
        for j, (kind, v) in enumerate(zip(kinds, mvals)):
            write_cell(r, len(lead) + 1 + j, kind, v, bold=bold, fill=fill)
        r += 1

    # 總計列（分組模式時「店鋪數」欄顯示通過篩選的店鋪總數）
    tot = report['totals']
    tot_g = {'cur': tot['cur'], 'ly': tot['ly'],
             'label': '總計', 'code': '', 'name': '', 'regional': '', 'om': '',
             'class1': '', 'class2': '', 'size': '', 'type': '', 'focus': False,
             'n': report.get('siteCount', 0)}
    emit(tot_g, bold=True, fill=total_fill)

    for g in report['groups']:
        emit(g)

    last_row = r - 1
    ws.auto_filter.ref = 'A%d:%s%d' % (h2, get_column_letter(ncols), last_row)
    ws.freeze_panes = 'B%d' % (h2 + 1)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def build_matrix_workbook(ds, opts, matrix):
    """按月矩陣匯出：每組別一行、每個月一欄（顯示所選指標），尾加合計欄。"""
    year = int(opts['year'])
    lyear = year - 1
    groupby = opts.get('groupBy') or 'shop'
    metric = opts.get('matrixMetric') or 'salesYoy'
    if metric == 'all':
        sections = [('salesYoy', '淨銷售 YoY%'), ('qtyYoy', '銷量 YoY%'),
                    ('gpYoy', '毛利 GP$ YoY%'), ('gpPp', '毛利率 ±pp')]
        metric_label = '全部指標（同步顯示）'
    elif metric == 'allv':
        sections = [('salesVal', '淨銷售（值）'), ('salesYoy', '淨銷售 YoY%'),
                    ('qtyVal', '銷量（值）'), ('qtyYoy', '銷量 YoY%'),
                    ('gpVal', '毛利 GP$（值）'), ('gpYoy', '毛利 GP$ YoY%'),
                    ('gpPctVal', '毛利率（值）'), ('gpPp', '毛利率 ±pp')]
        metric_label = '全部指標＋值（同步顯示）'
    else:
        _k, _label = MATRIX_METRICS.get(metric, ('yoy', '淨銷售 YoY%'))
        sections = [(metric, _label)]
        metric_label = _label
    multi = len(sections) > 1
    months = matrix['months']
    ly_months = ['%d-%s' % (lyear, m[5:7]) for m in months]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'YoY 按月'

    thin = Side(style='thin', color='D9E1EC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_font = Font(bold=True, color='FFFFFF', size=11)
    head_fill = PatternFill('solid', fgColor='0F172A')
    total_fill = PatternFill('solid', fgColor='EEF2FF')
    sub_fill = PatternFill('solid', fgColor='F1F5F9')
    na_font = Font(color='94A3B8')
    border_sec = Border(left=Side(style='medium', color='94A3B8'), right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')
    right = Alignment(horizontal='right', vertical='center')

    FMT = {'yoy': '+0.0%;-0.0%;0.0%', 'pp': '+0.00" pp";-0.00" pp";0.00" pp"',
           'money': '#,##0', 'qty': '#,##0', 'pct': '0.00%'}

    # ---- 標題區 ----
    _write_title_block(ws, ds, opts, months, ly_months,
                       extra_line='檢視：按月矩陣（指標：%s）' % metric_label)

    h1 = 8  # 表頭起始列
    h2 = h1 + 1 if multi else h1
    first_data = h2 + 1

    if groupby == 'shop':
        lead = [
            ('店鋪', 'label', 26), ('Site', 'text', 9), ('店名', 'text', 18),
            ('Regional', 'text', 9), ('OM', 'text', 9), ('Class 1', 'text', 8),
            ('Class 2', 'text', 8), ('Size', 'text', 7), ('Type', 'text', 6), ('需關注', 'text', 7),
        ]
    else:
        lead = [('分組', 'label', 24), ('店鋪數', 'qty', 9)]

    ncols = len(lead) + len(sections) * (len(months) + 1)

    # ---- 表頭 ----
    if multi:
        # 第一列：店鋪資料 + 各指標區塊（合併）
        for cc in range(1, len(lead) + 1):
            c = ws.cell(row=h1, column=cc)
            c.fill = head_fill
            c.border = border
        c = ws.cell(row=h1, column=1, value='店鋪資料' if groupby == 'shop' else '分組資料')
        c.font = head_font
        c.alignment = center
        ws.merge_cells(start_row=h1, start_column=1, end_row=h1, end_column=len(lead))
        col = len(lead) + 1
        for _mk, mlabel in sections:
            n_sec = len(months) + 1
            for cc in range(col, col + n_sec):
                c = ws.cell(row=h1, column=cc)
                c.fill = head_fill
                c.border = border_sec if cc == col else border
            c = ws.cell(row=h1, column=col, value=mlabel)
            c.font = head_font
            c.alignment = center
            if n_sec > 1:
                ws.merge_cells(start_row=h1, start_column=col, end_row=h1, end_column=col + n_sec - 1)
            col += n_sec
        # 第二列：欄名 + 月份 + 合計
        for ci, (name, _k, _w) in enumerate(lead, start=1):
            c = ws.cell(row=h2, column=ci, value=name)
            c.border = border
            c.fill = sub_fill
            c.font = Font(bold=True, color='334155')
            c.alignment = center
        col = len(lead) + 1
        for _mk, _ml in sections:
            first = True
            for m in months:
                c = ws.cell(row=h2, column=col, value=m)
                c.border = border_sec if first else border
                c.fill = sub_fill
                c.font = Font(bold=True, color='334155')
                c.alignment = center
                first = False
                col += 1
            c = ws.cell(row=h2, column=col, value='合計')
            c.border = border
            c.fill = sub_fill
            c.font = Font(bold=True, color='334155')
            c.alignment = center
            col += 1
    else:
        # 單一指標：單列深色表頭
        for ci, (name, _k, _w) in enumerate(lead, start=1):
            c = ws.cell(row=h1, column=ci, value=name)
            c.border = border
            c.fill = head_fill
            c.font = head_font
            c.alignment = center
        col = len(lead) + 1
        for m in months:
            c = ws.cell(row=h1, column=col, value=m)
            c.border = border
            c.fill = head_fill
            c.font = head_font
            c.alignment = center
            col += 1
        c = ws.cell(row=h1, column=ncols, value='合計')
        c.border = border
        c.fill = head_fill
        c.font = head_font
        c.alignment = center

    # 欄寬
    for ci, (_n, _k, w) in enumerate(lead, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci in range(len(lead) + 1, ncols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 11

    # ---- 儲存格寫入 ----
    def write_lead(rr, g, bold=False, fill=None):
        if groupby == 'shop':
            vals = [
                (g['label'], 'label'), (g['code'], 'text'), (g['name'], 'text'),
                (g['regional'], 'text'), (g['om'], 'text'), (g['class1'], 'text'),
                (g['class2'], 'text'), (g['size'], 'text'), (g['type'], 'text'),
                ('Y' if g.get('focus') else '', 'text'),
            ]
        else:
            vals = [(g['label'], 'label'), (g['n'], 'qty')]
        for ci, (v, k) in enumerate(vals, start=1):
            c = ws.cell(row=rr, column=ci)
            c.border = border
            if fill:
                c.fill = fill
            if k == 'qty':
                c.value = v if v is not None else ''
                c.number_format = FMT['qty']
                c.alignment = right
                c.font = Font(bold=bold)
            else:
                c.value = v if v is not None else ''
                c.alignment = left if k == 'label' else center
                c.font = Font(bold=bold)

    def write_cell(rr, ci, cur, ly, mk, bold=False, fill=None, sec_left=False):
        c = ws.cell(row=rr, column=ci)
        c.border = border_sec if sec_left else border
        c.alignment = right
        if fill:
            c.fill = fill
        v = _matrix_value(cur, ly, mk)
        if v is None:
            c.value = '—'
            c.font = na_font
            c.alignment = center
            return
        kind = MATRIX_METRICS.get(mk, ('yoy', ''))[0]
        c.value = float(v)
        c.number_format = FMT[kind]
        if kind in ('yoy', 'pp'):
            if v > 1e-9:
                c.font = Font(bold=bold, color='047857')
            elif v < -1e-9:
                c.font = Font(bold=bold, color='B91C1C')
            else:
                c.font = Font(bold=bold)
        else:
            c.font = Font(bold=bold)

    r = first_data

    # 總計列
    tot = matrix['totals']
    tot_g = {'label': '總計', 'n': matrix.get('siteCount', ''),
             'code': '', 'name': '', 'regional': '', 'om': '', 'class1': '',
             'class2': '', 'size': '', 'type': '', 'focus': False}
    write_lead(r, tot_g, bold=True, fill=total_fill)
    col = len(lead) + 1
    for mk, _ml in sections:
        first = True
        for m in months:
            write_cell(r, col, tot['cells'][m], tot['cellsLy'][m], mk, bold=True, fill=total_fill, sec_left=first)
            first = False
            col += 1
        write_cell(r, col, tot['aggCur'], tot['aggLy'], mk, bold=True, fill=total_fill)
        col += 1
    r += 1

    for g in matrix['groups']:
        write_lead(r, g)
        col = len(lead) + 1
        for mk, _ml in sections:
            first = True
            for m in months:
                write_cell(r, col, g['cells'][m], g['cellsLy'][m], mk, sec_left=first)
                first = False
                col += 1
            write_cell(r, col, g['aggCur'], g['aggLy'], mk)
            col += 1
        r += 1

    last_row = r - 1
    ws.auto_filter.ref = 'A%d:%s%d' % (h2, get_column_letter(ncols), last_row)
    ws.freeze_panes = 'B%d' % first_data

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
# 路由
# ---------------------------------------------------------------------------

def _get_dataset(token):
    """取出未過期的 dataset；已過期即由記憶體刪除。"""
    ds = _DATASETS.get(token)
    if ds is None:
        return None
    if DATA_TTL_HOURS > 0 and time.time() - _TIMES.get(token, 0) > DATA_TTL_HOURS * 3600:
        with _LOCK:
            _DATASETS.pop(token, None)
            _TIMES.pop(token, None)
            if token in _DATASET_ORDER:
                _DATASET_ORDER.remove(token)
        return None
    return ds


@app.before_request
def _require_auth():
    """若設定了 APP_PASSWORD，所有請求（除 /healthz）需先通過 Basic Auth。"""
    if not APP_PASSWORD or request.path == '/healthz':
        return None
    auth = request.authorization
    if not auth or auth.password != APP_PASSWORD:
        return Response('需要登入：請輸入存取密碼（用戶名稱隨意）', 401,
                        {'WWW-Authenticate': 'Basic realm="Shop YoY"'})
    return None


@app.get('/')
def index():
    return render_template('index.html')


@app.get('/healthz')
def healthz():
    return jsonify({'ok': True})


@app.post('/api/upload')
def api_upload():
    f = request.files.get('file')
    if f is None or not f.filename:
        return jsonify({'ok': False, 'error': '未收到檔案。'}), 400
    fn = f.filename
    if not fn.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'ok': False, 'error': '只支援 .xlsx 檔案（請由 Hugo_by Shop_YoY 匯出）。'}), 400
    try:
        ds = parse_excel(f.stream)
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    except Exception as e:  # noqa: BLE001
        app.logger.exception('parse error')
        return jsonify({'ok': False, 'error': '解析檔案時發生錯誤：%s' % e}), 500

    ds['filename'] = fn
    ds['uploadedAt'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    token = uuid.uuid4().hex
    now = time.time()
    with _LOCK:
        # 先清走過期資料，再存入新資料（最多保留 _DATASETS_MAX 份）
        if DATA_TTL_HOURS > 0:
            cutoff = now - DATA_TTL_HOURS * 3600
            for t in [x for x in _DATASET_ORDER if _TIMES.get(x, 0) < cutoff]:
                _DATASET_ORDER.remove(t)
                _DATASETS.pop(t, None)
                _TIMES.pop(t, None)
        _DATASETS[token] = ds
        _TIMES[token] = now
        _DATASET_ORDER.append(token)
        while len(_DATASET_ORDER) > _DATASETS_MAX:
            old = _DATASET_ORDER.pop(0)
            _DATASETS.pop(old, None)
            _TIMES.pop(old, None)
    ds.pop('_ts', None)
    return jsonify({'ok': True, 'token': token, 'dataset': ds, 'ttlHours': DATA_TTL_HOURS})


@app.post('/api/clear-all')
def api_clear_all():
    """即時清除所有已上載資料（由記憶體刪除）。"""
    with _LOCK:
        _DATASETS.clear()
        _DATASET_ORDER.clear()
        _TIMES.clear()
    return jsonify({'ok': True})


@app.post('/api/clear/<token>')
def api_clear(token):
    """即時清除指定的一份已上載資料。"""
    with _LOCK:
        _DATASETS.pop(token, None)
        _TIMES.pop(token, None)
        if token in _DATASET_ORDER:
            _DATASET_ORDER.remove(token)
    return jsonify({'ok': True})


@app.post('/api/report/<token>')
def api_report(token):
    ds = _get_dataset(token)
    if ds is None:
        return jsonify({'ok': False, 'error': '資料已過期，請重新上載 Excel。'}), 404
    opts = request.get_json(silent=True) or {}
    try:
        if (opts.get('view') or 'agg') == 'matrix':
            mx = compute_matrix(ds, opts)
            r2 = lambda arr: [round(x, 2) for x in arr]
            groups = [{
                'label': g['label'], 'n': g['n'],
                'cells': {m: r2(g['cells'][m]) for m in mx['months']},
                'ly': {m: r2(g['cellsLy'][m]) for m in mx['months']},
                'agg': {'cur': r2(g['aggCur']), 'ly': r2(g['aggLy'])},
            } for g in mx['groups']]
            tot = mx['totals']
            return jsonify({
                'ok': True, 'view': 'matrix', 'months': mx['months'], 'groups': groups,
                'totals': {
                    'cells': {m: r2(tot['cells'][m]) for m in mx['months']},
                    'ly': {m: r2(tot['cellsLy'][m]) for m in mx['months']},
                    'agg': {'cur': r2(tot['aggCur']), 'ly': r2(tot['aggLy'])},
                },
                'siteCount': mx['siteCount'],
            })

        rep = compute_report(ds, opts)
    except Exception as e:  # noqa: BLE001
        return jsonify({'ok': False, 'error': '計算失敗：%s' % e}), 400
    groups = []
    for g in rep['groups']:
        groups.append({
            'label': g['label'], 'n': g['n'],
            'cur': {'qty': g['cur'][0], 'sales': g['cur'][1], 'gp': g['cur'][2]},
            'ly': {'qty': g['ly'][0], 'sales': g['ly'][1], 'gp': g['ly'][2]},
        })
    tot = rep['totals']
    return jsonify({
        'ok': True,
        'groups': groups,
        'totals': {
            'cur': {'qty': tot['cur'][0], 'sales': tot['cur'][1], 'gp': tot['cur'][2]},
            'ly': {'qty': tot['ly'][0], 'sales': tot['ly'][1], 'gp': tot['ly'][2]},
        },
        'siteCount': rep['siteCount'],
    })


@app.post('/api/export/<token>')
def api_export(token):
    ds = _get_dataset(token)
    if ds is None:
        return jsonify({'ok': False, 'error': '資料已過期，請重新上載 Excel。'}), 404
    opts = request.get_json(silent=True) or {}
    if not opts.get('year') or not opts.get('months'):
        return jsonify({'ok': False, 'error': '請先選擇年份及至少一個月份。'}), 400
    try:
        if (opts.get('view') or 'agg') == 'matrix':
            mx = compute_matrix(ds, opts)
            bio = build_matrix_workbook(ds, opts, mx)
        else:
            rep = compute_report(ds, opts)
            bio = build_workbook(ds, opts, rep)
    except Exception as e:  # noqa: BLE001
        app.logger.exception('export error')
        return jsonify({'ok': False, 'error': '產生 Excel 時發生錯誤：%s' % e}), 500

    ys = sorted({int(str(m)[5:7]) for m in opts['months']})
    contiguous = all(ys[i] + 1 == ys[i + 1] for i in range(len(ys) - 1))
    if len(ys) == 1:
        mm = '%02d' % ys[0]
    elif contiguous:
        mm = '%02d-%02d' % (ys[0], ys[-1])
    else:
        mm = '-'.join('%02d' % m for m in ys)
    name = 'Shop_YoY_%d_%s.xlsx' % (int(opts['year']), mm)
    return send_file(
        bio, as_attachment=True, download_name=name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.errorhandler(413)
def too_large(_e):
    return jsonify({'ok': False, 'error': '檔案過大（上限 32 MB）。'}), 413


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
