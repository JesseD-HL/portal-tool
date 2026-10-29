"""
Buysheet Portal - Flask backend
================================
One file, read top to bottom:

  1. config + constants        where HANA is, what a size looks like, season codes
  2. helpers                   small pure functions (numbers, dates, seasons, ratios)
  3. HANA                      connection + the six read-only queries
  4. parsers                   buysheet (DATA sheet), legacy prebuys, the new prebuy template
  5. validation                every row checked against what SAP holds
  6. outputs                   buysheet DATA + HEADER OUT, creation templates, PR CSV
  7. diagnostics               what SAP really holds for anything the checks flagged
  8. routes                    the HTTP endpoints the portal calls

Endpoints:
  GET  /                    serve portal.html
  GET  /health              status + build stamp
  GET  /diagnostics         plain-text JSON for pasting into a chat
  POST /validate            upload a prebuy or buysheet, validate, cache the result
  GET  /generate            the buysheet: the .xlsm template filled (templates\\Buysheet_Template.xlsm), else a plain DATA workbook
  GET  /creation_templates  GENERIC / PPK creation workbook from the cache
  GET  /validation_report   SUMMARY / GENERIC VALIDATION / PPK VALIDATION / PPK COMPONENTS workbook from the cache
  POST /stock               the SAP Cut and Sold export (Excel / CSV): what is already in stock, joined by article number
  GET  /stock, POST /stock/clear, GET /validation
  GET  /req_csv             purchase-requisition CSV from the cache
  POST /delivery_split      split a delivery file proportionally

Run:  python server.py        (portal.html and config.json sit in the same folder as server.py)

Setup on a new machine: copy config.example.json to config.json and fill in your HANA host, user and password.
Nothing secret is kept in this file; the settings can also come from environment variables (HANA_HOST, HANA_PORT,
HANA_USER, HANA_PASSWORD, TOOLS_DIR, STOCK_VIEW) or from a file named by PORTAL_CONFIG. STOCK_VIEW (default
YOUR_COMPANY_SCHEMA.YOUR_STOCK_VIEW) is where live stock is read; OFF switches live stock off.

SAP placeholders: HANA schema/package names, calculation-view names, table names, and Z-transaction codes
below are all generic stand-ins (YOUR_COMPANY_SCHEMA, YOUR_ARTICLE_MASTER_VIEW, YOUR_TXN_1, etc.) rather
than this installation's real names. Find-and-replace each placeholder with your own before running this
against a real SAP HANA system.
"""

import os, re, io, csv, math, json, zipfile, traceback
from datetime import datetime
from collections import defaultdict

from flask import Flask, request, jsonify, send_file, send_from_directory, Response
from flask_cors import CORS
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from hdbcli import dbapi

app = Flask(__name__)
CORS(app)

# ═══════════════════════════════════════════════════════════════════════════════
# 1. CONFIG + CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════
HERE = os.path.dirname(os.path.abspath(__file__))

def load_settings():
    """Where HANA is and who connects: defaults, then config.json next to this file (or the file PORTAL_CONFIG names),
    then environment variables, which win. No credential is kept in this source file."""
    cfg = {'HANA_HOST': '', 'HANA_PORT': '33015', 'HANA_USER': '', 'HANA_PASSWORD': '', 'TOOLS_DIR': HERE,
           'STOCK_VIEW': 'YOUR_COMPANY_SCHEMA.YOUR_STOCK_VIEW'}       # live stock: the Cut and Sold report's own view (OFF = no live stock)
    path = os.environ.get('PORTAL_CONFIG') or os.path.join(HERE, 'config.json')
    source = 'built-in defaults'
    if os.path.exists(path):
        try:
            with open(path, encoding='utf-8-sig') as f:                    # utf-8-sig: Windows Notepad adds a BOM
                raw = json.load(f)
        except ValueError as exc:
            raise SystemExit(f"{path} is not valid JSON ({exc}). Compare it with config.example.json: every name and value is in double quotes.")
        cfg.update({str(k).upper(): v for k, v in raw.items() if str(k).upper() in cfg})
        source = os.path.basename(path)
    from_env = [k for k in cfg if os.environ.get(k)]
    for k in from_env:
        cfg[k] = os.environ[k]
    if any(k.startswith('HANA_') for k in from_env):                 # name the environment only when it supplied a HANA setting
        source = 'environment variables' if source == 'built-in defaults' else f"{source} + environment variables"
    try:
        cfg['HANA_PORT'] = int(str(cfg['HANA_PORT']).strip() or 33015)
    except ValueError:
        raise SystemExit(f"HANA_PORT must be a number (got '{cfg['HANA_PORT']}').")
    return cfg, source

_SETTINGS, SETTINGS_SOURCE = load_settings()
HANA_HOST     = str(_SETTINGS['HANA_HOST']).strip()
HANA_PORT     = _SETTINGS['HANA_PORT']
HANA_USER     = str(_SETTINGS['HANA_USER']).strip()
HANA_PASSWORD = str(_SETTINGS['HANA_PASSWORD'])
STOCK_VIEW    = '' if str(_SETTINGS['STOCK_VIEW']).strip().upper() in ('', 'OFF', 'NONE') else str(_SETTINGS['STOCK_VIEW']).strip()
TOOLS_DIR     = str(_SETTINGS['TOOLS_DIR'])            # where portal.html and templates/ live: next to server.py unless told otherwise

# Bump whenever this file changes. The portal shows it next to its own PAGE_BUILD so a stale file is obvious.
BUILD = "2026-10-14-diag35"

_cache = {}          # the last validation: rows, customers, validation (one session, one file at a time)

KNOWN_SIZES = {'XS','S','M','L','XL','XXL','2XL','3XL','4XL','5XL','6XL',
               '1X','2X','3X','4X','5X','OS','ONE SIZE','LT','XLT','2XLT','3XLT','4XLT','5XLT'}
SIZE_ORDER  = ['XS','S','M','L','XL','XXL','2XL','3XL','4XL','5XL','6XL',
               '1X','2X','3X','4X','5X','LT','XLT','2XLT','3XLT','4XLT','5XLT','OS','ONE SIZE']

# Only SPELLED-OUT sizes are translated ('SMALL' -> 'S'). Labels such as 2XL / XXL are never rewritten:
# every size is used exactly as the prebuy types it, because SAP's code for a size depends on the brand.
SIZE_ALIASES = {'SMALL':'S','MEDIUM':'M','LARGE':'L',
                'X SMALL':'XS','X-SMALL':'XS','EXTRA SMALL':'XS',
                'X LARGE':'XL','X-LARGE':'XL','EXTRA LARGE':'XL',
                'XX LARGE':'XXL','XX-LARGE':'XXL'}

SEASON_LABEL = {'1':'Q1','2':'Q2','3':'Q3','4':'Q4'}                       # style[5] -> quarter
SEASON_NAME  = {'1':'SPR','2':'SUM','3':'FALL','4':'HOL'}                  # Q1 SPR, Q2 SUM, Q3 FALL, Q4 HOL
SEASON_WORD  = {'SPR':1,'SPRING':1,'SUM':2,'SUMMER':2,'FALL':3,'AUTUMN':3,'HOL':4,'HOLIDAY':4}

GENDER_PREFIX = {
    'RVM':'MEN','RDM':'MEN','REM':'MEN','CPM':'MEN','CLM':'MEN','ANM':'MEN','HEB':'MEN',
    'RVW':'LADIES','RDW':'LADIES','REW':'LADIES','CPW':'LADIES','CLW':'LADIES','HEW':'LADIES',
    'RVB':'BOYS','RDB':'BOYS','REB':'BOYS','CPB':'BOYS','REG':'BOYS','RDG':'BOYS',
    'CAM':'MEN','CFW':'LADIES','CAB':'BOYS',
}
# weights / volume per (gender, 2-digit product code) for the creation templates
WEIGHTS_LOOKUP = {
    ('MEN',    'OW'): {'HOODIE':(1.7,0.16),'HOOD':(1.7,0.16),'DEFAULT':(1.1,0.067)},
    ('MEN',    'HD'): {'DEFAULT':(1.7,0.16)},
    ('MEN',    'FH'): {'DEFAULT':(1.7,0.16)},
    ('MEN',    'TL'): {'DEFAULT':(0.8,0.16)},
    ('MEN',    'TS'): {'POLO':(0.8,0.08),'DEFAULT':(0.8,0.067)},
    ('MEN',    'SS'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'TE'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'TX'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'JK'): {'DEFAULT':(1.1,0.16)},
    ('MEN',    'PA'): {'FLEECE':(1.4,0.16),'DEFAULT':(1.03,0.067)},
    ('MEN',    'SH'): {'DEFAULT':(0.45,0.16)},
    ('LADIES', 'OW'): {'HOODIE':(1.4,0.13),'HOOD':(1.4,0.13),'DEFAULT':(0.9,0.067)},
    ('LADIES', 'HD'): {'DEFAULT':(1.4,0.13)},
    ('LADIES', 'FH'): {'DEFAULT':(1.4,0.13)},
    ('LADIES', 'TL'): {'DEFAULT':(0.8,0.08)},
    ('LADIES', 'TS'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'SS'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'TE'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'TX'): {'DEFAULT':(0.3527,0.067)},
    ('LADIES', 'BR'): {'TANK':(0.3527,0.067),'DEFAULT':(0.23,0.067)},
    ('LADIES', 'JK'): {'DEFAULT':(1.1,0.067)},
    ('LADIES', 'SD'): {'SKIRT':(0.31,0.067),'DEFAULT':(0.6,0.067)},
    ('LADIES', 'LE'): {'DEFAULT':(0.5732,0.067)},
    ('LADIES', 'PA'): {'FLEECE':(0.77,0.1),'DEFAULT':(0.77,0.1)},
    ('LADIES', 'SH'): {'DEFAULT':(0.35,0.067)},
    ('BOYS',   'OW'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'HD'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'FH'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'TL'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TS'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'SS'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TE'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TX'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'JK'): {'DEFAULT':(0.88,0.128)},
    ('BOYS',   'PA'): {'FLEECE':(1.12,0.112),'DEFAULT':(0.824,0.0536)},
    ('BOYS',   'SH'): {'DEFAULT':(0.36,0.0536)},
}

def get_weights(style, description=''):
    """(weight, volume) for a style from the lookup above, keyed by gender prefix and product code."""
    style = str(style).strip().upper()
    desc  = str(description).strip().upper()
    gender = GENDER_PREFIX.get(style[:3])
    lookup = WEIGHTS_LOOKUP.get((gender, style[6:8])) if gender else None
    if not lookup:
        return None, None
    for kw, vals in lookup.items():
        if kw != 'DEFAULT' and kw in desc:
            return vals
    return lookup.get('DEFAULT', (None, None))

# ═══════════════════════════════════════════════════════════════════════════════
# 2. HELPERS - small, pure, no I/O
# ═══════════════════════════════════════════════════════════════════════════════
def num(v):
    """Tolerant number reader: '16.96', '$16.96', '1,200' -> float; blanks and text ('TBD') -> None."""
    if v is None or v == '': return None
    try: return float(str(v).replace('$', '').replace(',', '').strip())
    except ValueError: return None

def up(v):
    """Everything written to the buysheet is upper case."""
    return v.upper() if isinstance(v, str) else v

def iso_date(d):
    return d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else (str(d)[:10] if d else '')

def mmddyyyy(d):
    """Delivery date as MMDDYYYY for the PR file; empty (never garbage) when it cannot be read."""
    if not d: return ''
    if hasattr(d, 'strftime'): return d.strftime('%m%d%Y')
    t = str(d).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%Y%m%d', '%m%d%Y'):
        try: return datetime.strptime(t, fmt).strftime('%m%d%Y')
        except ValueError: pass
    return ''

def parse_season(text):
    """'Q2 2027', 'Q2-2027', 'q2/2027', '2027 SUM' or 'SUM 2027' -> 'Q2 2027' ('' when it cannot be read)."""
    t = str(text or '').strip().upper()
    m = re.search(r'\bQ\s*([1-4])\s*[-/ ]?\s*(\d{4})\b', t)
    if m:
        return f"Q{m.group(1)} {m.group(2)}"
    m = re.search(r'\b(20\d{2})\s*[-/ ]?\s*([A-Z]{3,7})\b', t) or re.search(r'\b([A-Z]{3,7})\s*[-/ ]?\s*(20\d{2})\b', t)
    if m:
        year, word = (m.group(1), m.group(2)) if m.group(1).isdigit() else (m.group(2), m.group(1))
        if word in SEASON_WORD:
            return f"Q{SEASON_WORD[word]} {year}"
    return ''

def season_text(label, style):
    """'2027 SUM' for the PR file: from the header season label when there is one, else from the style number."""
    m = re.fullmatch(r'Q([1-4]) (\d{4})', label or '')
    if m:
        return f"{m.group(2)} {SEASON_NAME[m.group(1)]}"
    s_num = style[5] if len(style) > 5 else '1'
    year = '20' + style[3:5] if len(style) > 4 else '2027'
    return f"{year} {SEASON_NAME.get(s_num, 'SPR')}"

def season_from_style(style):
    """'Q2-2027' label for the buysheet's top row, from the style number."""
    return f"{SEASON_LABEL.get(style[5], 'Q1')}-20{style[3:5]}" if len(style) > 5 else ''

def canonical_size(raw):
    r = str(raw).strip().upper()
    return SIZE_ALIASES.get(r, r)

def base_style(raw):
    """'CFW271SH01-S470' -> 'CFW271SH01'."""
    return str(raw).strip().split('-')[0]

generic_from_ppk = base_style

def color_from_ppk(raw):
    parts = str(raw).strip().split('-')
    return parts[1] if len(parts) > 1 else ''

def is_ppk(raw):
    return '-' in str(raw).strip() if raw else False

def strip_color(c):
    return re.sub(r'\s*\(.*?\)', '', str(c)).replace(' ', '').upper()

def normalize_color(c):
    c = re.sub(r'\s*\(.*?\)', '', str(c)).strip().upper()
    c = re.sub(r'\s+', ' ', c)
    return re.sub(r'\bHTR\b|\bHTHR\b', 'HEATHER', c).strip()

def build_ratio_sizes(detail):
    """SAP's components for a prepack -> ('XS-S-M', '1-3-4') in SAP's size order."""
    if not detail or not detail.get('variants'):
        return '', ''
    sv = sorted(detail['variants'], key=lambda x: x[0])
    return ('-'.join(str(v[1]) for v in sv),
            '-'.join(str(int(v[2])) if v[2] is not None else '?' for v in sv))

def dash_list(v):
    """'1-3-4-3-2-2', '1, 3, 4' or '1 3 4' -> ['1','3','4'] (upper case, no blanks).
    A real date means Excel turned something like 1-2 into a date: that cannot be read back, so it gives nothing."""
    if v is None or hasattr(v, 'strftime'):
        return []
    return [x for x in re.sub(r'[\s,;/]+', '-', str(v).strip().upper()).split('-') if x]

def split_by_ratio(total, weights):
    """Split a total across components in proportion to their ratio, so the parts always add back to
    exactly the total (largest-remainder rounding: no piece is lost or invented)."""
    whole = sum(weights)
    if total is None or not whole:
        return [None] * len(weights)
    total = int(round(total))
    raw = [total * w / whole for w in weights]
    parts = [int(x) for x in raw]
    for i in sorted(range(len(weights)), key=lambda i: raw[i] - parts[i], reverse=True)[:max(total - sum(parts), 0)]:
        parts[i] += 1
    return parts

def price_check_text(buy, sap):
    """Price check cell. Says WHICH side is missing instead of a generic 'no data'."""
    if buy and sap:
        return '✓' if round(buy, 2) == round(sap, 2) else f'✗ SAP={sap:.2f} BUY={buy:.2f}'
    if buy: return '⚠ NO SAP PRICE'
    if sap: return '⚠ NO BUY PRICE'
    return '—'


def compare_ratio(exists, buy_sizes, buy_ratio, detail):
    """The file's sizes + ratio against the components SAP holds for the prepack, size by size (the order typed does not matter).
    Returns (sizes check, ratio check, a plain-English note saying exactly which sizes differ)."""
    if not exists:
        return '—', '—', ''
    hana = {}
    for _, sz, q in (detail or {}).get('variants', []):
        try:
            hana[str(sz).strip().upper()] = int(round(float(q)))
        except (TypeError, ValueError):
            pass
    if not (buy_sizes and buy_ratio and hana):
        return '⚠ NO DATA', '⚠ NO DATA', ''
    szs = [x for x in buy_sizes.split('-') if x]
    rts = [x for x in buy_ratio.split('-') if x]
    if len(szs) != len(rts):
        return '✗ MISMATCH', '✗ MISMATCH', f"The file has {len(rts)} ratio number{'' if len(rts) == 1 else 's'} but {len(szs)} size{'' if len(szs) == 1 else 's'}."
    if len(set(szs)) != len(szs):
        return '✗ MISMATCH', '✗ MISMATCH', 'A size is listed more than once.'
    try:
        buy = {sz: int(rt) for sz, rt in zip(szs, rts)}
    except ValueError:
        return '✗ MISMATCH', '✗ MISMATCH', 'The ratio has something that is not a whole number.'
    order = sorted(set(buy) | set(hana), key=lambda x: (SIZE_ORDER.index(x) if x in SIZE_ORDER else 99, x))
    diffs = [f"{sz}: file {buy.get(sz, 'none')}, SAP {hana.get(sz, 'none')}" for sz in order if buy.get(sz) != hana.get(sz)]
    return ('✓' if set(buy) == set(hana) else '✗ MISMATCH'), ('✓' if not diffs else '✗ MISMATCH'), '; '.join(diffs)

def desc_diff_note(buy, sap):
    """When two descriptions differ only in punctuation or spacing (7/8" vs 7/8), say so."""
    squash = lambda t: re.sub(r'[^A-Z0-9]', '', str(t).upper())
    return 'differs only in punctuation or spacing' if buy and sap and squash(buy) == squash(sap) else ''

# ═══════════════════════════════════════════════════════════════════════════════
# 3. HANA - connection + read-only queries
# ═══════════════════════════════════════════════════════════════════════════════
def hana_configured():
    return bool(HANA_HOST and HANA_USER and HANA_PASSWORD)

def connect_hana():
    if not hana_configured():
        missing = [n for n, v in (('HANA_HOST', HANA_HOST), ('HANA_USER', HANA_USER), ('HANA_PASSWORD', HANA_PASSWORD)) if not v]
        raise RuntimeError(f"HANA connection is not set up ({', '.join(missing)} missing). Copy config.example.json to config.json "
                           f"next to server.py, fill in your HANA host, user and password, and restart the server.")
    return dbapi.connect(address=HANA_HOST, port=HANA_PORT, user=HANA_USER, password=HANA_PASSWORD)

def query_hana_generics(conn, generic_styles):
    if not generic_styles: return {}
    style_list = ', '.join(f"'{s}'" for s in generic_styles)
    sql = f"""
        SELECT "F_GENERIC" AS "Generic", "MARA_ARTICLE" AS "Article",
               "MARA_COLOR_CODE" AS "Color Code", "MARA_SIZE_CODE" AS "Size Code",
               "F_GENERIC_DESCRIPTION" AS "Description", "WRF_SORT_POSITION" AS "Size Sort",
               "F_PC_PER_CAR_FIX" AS "PC Per Ctn",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name"
        FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
        WHERE "F_GENERIC" IN ({style_list}) AND "F_ARTICLE_TYPE" <> 'C'
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    result = {}
    for row in rows:
        r = dict(zip(cols, row))
        generic    = str(r.get('Generic','') or '').strip()
        color_code = str(r.get('Color Code','') or '').strip()
        size_code  = str(r.get('Size Code','') or '').strip().upper()
        hana_desc  = str(r.get('Description','') or '').strip().upper()
        size_sort  = str(r.get('Size Sort','') or '99999').strip()
        pc_per_ctn = r.get('PC Per Ctn')
        color_name = str(r.get('Color Name','') or '').strip().upper()
        result[(generic, color_code, size_code)] = {
            'article':    str(r.get('Article','') or '').strip(),
            'desc':       hana_desc,
            'sort':       size_sort,
            'pc_per_ctn': pc_per_ctn,
            'color_name': color_name,
        }
    return result

def query_hana_ppks(conn, ppk_articles):
    if not ppk_articles: return {}, {}
    article_list = ', '.join(f"'{a}'" for a in ppk_articles)
    sql = f"""
        SELECT "MAST_ARTICLE" AS "PPK Article", "MARA_ARTICLE" AS "Component",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name",
               "MARA_SIZE_CODE" AS "Size Code", "STPO_PPK_QTY" AS "PPK Qty",
               "F_GENERIC_DESCRIPTION" AS "Description", "WRF_SORT_POSITION" AS "Size Sort",
               "MARM_PC_PER_PAC" AS "PC Per PAC"
        FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
        WHERE "MAST_ARTICLE" IN ({article_list}) AND "F_ARTICLE_TYPE" = 'C'
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    ppk_exists = {}; ppk_details = {}
    for row in rows:
        r          = dict(zip(cols, row))
        ppk_art    = str(r.get('PPK Article','') or '').strip()
        color_name = str(r.get('Color Name','') or '').strip().upper()
        size_code  = str(r.get('Size Code','') or '').strip().upper()
        ppk_qty    = r.get('PPK Qty')
        size_sort  = str(r.get('Size Sort','') or '99999').strip()
        hana_desc  = str(r.get('Description','') or '').strip().upper()
        pc_per_pac = r.get('PC Per PAC')
        ppk_exists[ppk_art] = True
        if ppk_art not in ppk_details:
            ppk_details[ppk_art] = {
                'color_name': color_name, 'hana_desc': hana_desc,
                'variants': [], 'seen': set(), 'pc_per_pac': pc_per_pac
            }
        if size_code not in ppk_details[ppk_art]['seen']:
            ppk_details[ppk_art]['seen'].add(size_code)
            ppk_details[ppk_art]['variants'].append((size_sort, size_code, ppk_qty))
            ppk_details[ppk_art].setdefault('components', []).append(
                {'article': str(r.get('Component') or '').strip(), 'code': color_from_ppk(ppk_art), 'color': color_name, 'size': size_code,
                 'qty': int(round(float(ppk_qty))) if ppk_qty not in (None, '') else None, 'sort': size_sort})
    return ppk_exists, ppk_details

def query_hana_ppks_for_generics(conn, generics):
    """Every prepack under these generics with its components -> {article: {'generic', 'comps': [(code, name, size, qty)], 'pc_per_pac', 'desc'}}.
    Used to find a prepack by CONTENT (color + sizes + ratio) when the typed article name is not found - the same
    search check_8_7.py does."""
    generics = sorted({str(g).strip() for g in generics if g})
    if not generics: return {}
    gen_list = ', '.join(f"'{g}'" for g in generics)
    sql = f"""
        SELECT "MAST_ARTICLE" AS "PPK", "MARA_ARTICLE" AS "Component", "F_GENERIC" AS "Generic", "MARA_COLOR_CODE" AS "Color Code",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name", "MARA_SIZE_CODE" AS "Size", "STPO_PPK_QTY" AS "Qty",
               "WRF_SORT_POSITION" AS "Sort", "MARM_PC_PER_PAC" AS "PC Per PAC", "F_GENERIC_DESCRIPTION" AS "Description"
        FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
        WHERE "F_GENERIC" IN ({gen_list}) AND "F_ARTICLE_TYPE" = 'C'
        ORDER BY "MAST_ARTICLE", "WRF_SORT_POSITION"
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    out = {}
    for row in rows:
        r = dict(zip(cols, row))
        art = str(r.get('PPK') or '').strip().upper()
        if not art: continue
        e = out.setdefault(art, {'generic': str(r.get('Generic') or '').strip().upper(), 'comps': [], 'pc_per_pac': r.get('PC Per PAC'),
                                 'desc': str(r.get('Description') or '').strip().upper(), 'seen': set()})
        size = str(r.get('Size') or '').strip().upper(); code = str(r.get('Color Code') or '').strip().upper()
        if (code, size) in e['seen']: continue
        e['seen'].add((code, size))
        try: qty = int(round(float(r.get('Qty') or 0)))
        except (TypeError, ValueError): qty = 0
        e['comps'].append((code, str(r.get('Color Name') or '').strip().upper(), size, qty, str(r.get('Sort') or '99999')))
        e.setdefault('components', []).append({'article': str(r.get('Component') or r.get('PPK') or '').strip(), 'code': code,
                                               'color': str(r.get('Color Name') or '').strip().upper(), 'size': size, 'qty': qty, 'sort': str(r.get('Sort') or '99999')})
    return out

def query_hana_ppk_descriptions(conn, ppk_articles):
    if not ppk_articles: return {}
    article_list = ', '.join(f"'{a}'" for a in ppk_articles)
    sql = f"""
        SELECT "MARA_ARTICLE" AS "Article", "MAKT_ARTICLE_DESCRIPTION" AS "Description"
        FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
        WHERE "MARA_ARTICLE" IN ({article_list})
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    result = {}
    for row in rows:
        r = dict(zip(cols, row))
        article = str(r.get('Article','') or '').strip()
        desc    = str(r.get('Description','') or '').strip().upper()
        if article and desc: result[article] = desc
    return result

def query_hana_prices(conn, articles):
    """Price per article from YOUR_PRICE_VIEW (already a per-piece amount).

    The view can hold more than one row per article (price history / future prices), so the
    choice is deterministic: a row valid today wins, otherwise the one with the latest start.
    Long IN lists are sent in chunks of 500.
    """
    articles = sorted({str(a).strip() for a in articles if a})
    if not articles: return {}
    today = datetime.now().strftime('%Y%m%d')
    norm  = lambda v, default: (str(v).replace('-', '')[:8] if v else default)
    best  = {}                                   # article -> (rank, price)
    cursor = conn.cursor()
    for i in range(0, len(articles), 500):
        article_list = ', '.join(f"'{a}'" for a in articles[i:i+500])
        sql = f"""
            SELECT "ARTICLE" AS "Article", "F_AMOUNT_PER_PC" AS "Price",
                   "VALIDITY_START_DATE" AS "Valid From", "VALIDITY_END_DATE" AS "Valid To"
            FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_PRICE_VIEW"
            WHERE "ARTICLE" IN ({article_list})
        """
        cursor.execute(sql)
        cols = [d[0] for d in cursor.description]
        for row in cursor.fetchall():
            r = dict(zip(cols, row))
            article = str(r.get('Article') or '').strip()
            price   = r.get('Price')
            if not article or price is None: continue
            start = norm(r.get('Valid From'), '00000000')
            end   = norm(r.get('Valid To'),   '99991231')
            rank  = (start <= today <= end, start)
            if article not in best or rank > best[article][0]:
                best[article] = (rank, float(price))
    return {art: price for art, (rank, price) in best.items()}

def query_hana_color_codes(conn, codes):
    """Color code -> set of descriptions SAP uses for it, read from the articles in HANA
    (the same source the standalone color_code_validator uses)."""
    codes = sorted({str(c).strip() for c in codes if c})
    out = {}
    if not codes: return out
    cursor = conn.cursor()
    for i in range(0, len(codes), 500):
        code_list = ', '.join(f"'{c}'" for c in codes[i:i+500])
        cursor.execute(f"""
            SELECT DISTINCT "MARA_COLOR_CODE" AS "Code", "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Description"
            FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
            WHERE "MARA_COLOR_CODE" IN ({code_list})
        """)
        for code, desc in cursor.fetchall():
            code = str(code or '').strip(); desc = str(desc or '').strip().upper()
            if code and desc: out.setdefault(code, set()).add(desc)
    return out

def query_variant_articles(conn, generics):
    """Get variant articles for PR CSV generation."""
    if not generics: return {}
    gen_list = ', '.join(f"'{g}'" for g in generics)
    sql = f"""
        SELECT "MARA_ARTICLE","F_GENERIC","MARA_COLOR_CODE","MARA_SIZE_CODE"
        FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW"
        WHERE "F_GENERIC" IN ({gen_list}) AND "F_ARTICLE_TYPE" != 'C'
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    result = {}
    for row in rows:
        article,gen,code,size = [str(x).strip().upper() if x else '' for x in row]
        result[(gen,code,size)] = article
    return result

# ═══════════════════════════════════════════════════════════════════════════════
# 4. PARSERS - every file becomes a list of row dicts (type OS or PPK) + the customers in it
# ═══════════════════════════════════════════════════════════════════════════════
def detect_file_type(wb):
    """A buysheet has a DATA sheet; anything else is a prebuy (legacy layout or the new template)."""
    return 'buysheet' if 'DATA' in wb.sheetnames else 'prebuy'

# ── 4a. buysheet (DATA sheet) ─────────────────────────────────────────────────
def read_storage_locations(wb):
    """customer -> SAP storage location, from the buysheet's HEADER sheet (HEADER, HEADER & FACTORY, ...).
    Columns are found by their heading, so a different column order does not matter."""
    out = {}
    for name in wb.sheetnames:
        if not name.strip().upper().startswith('HEADER'):
            continue
        ws, hdr_row = wb[name], None
        for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            cells = [str(v).strip().upper() if v is not None else '' for v in row]
            is_loc = lambda c: 'STORAGE' in c or c in ('WAREHOUSE', 'WHSE')          # the buysheet template's own heading is WHSE
            if 'CUSTOMER' in cells and any(is_loc(c) for c in cells):
                hdr_row, cust_i = ri, cells.index('CUSTOMER')
                loc_i = next(i for i, c in enumerate(cells) if is_loc(c))
                break
        if hdr_row is None:
            continue
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if len(row) > max(cust_i, loc_i) and row[cust_i] and row[loc_i] and str(row[loc_i]).strip():
                out.setdefault(str(row[cust_i]).strip().upper(), str(row[loc_i]).strip().upper())
    return out

def read_season_label(all_rows):
    """'Q2 2027' from a label such as Q2-2027 near the top of the DATA sheet ('' if there is none)."""
    for r in all_rows[:3]:
        for v in r:
            label = parse_season(v)
            if label:
                return label
    return ''

def parse_buysheet(wb):
    """Parse buysheet DATA sheet — same logic as validator."""
    ws = wb['DATA'] if 'DATA' in wb.sheetnames else wb.active
    all_rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    for i, r in enumerate(all_rows):
        if r[0] and str(r[0]).upper().strip() == 'CUSTOMER':
            header_idx = i; break
    if header_idx is None:
        header_idx = next((i for i, r in enumerate(all_rows) if any(v for v in r)), 0)

    header    = list(all_rows[header_idx])
    data_rows = [r for r in all_rows[header_idx+1:] if any(v for v in r)]
    header_upper = [str(h).upper().strip().replace('\n',' ').replace('\r',' ') if h else '' for h in header]

    def find_col(name):
        for i, h in enumerate(header_upper):
            if name in h: return i
        return None

    idx = {
        'style':        find_col('STYLE'),
        'color_code':   find_col('COLOR CODE'),
        'color':        next((i for i,h in enumerate(header_upper) if h == 'COLOR'), None),
        'desc':         find_col('DESCRIPTION'),
        'pc_ctn':       find_col('PC PER CTN'),
        'ratio':        find_col('SIZE RATIO'),
        'sizes_col':    find_col('SIZES'),
        'pc_per_pac':   find_col('PC PER PAC'),
        'pac_per_ctn':  find_col('PAC PER CTN'),
        'ppk_qty':      find_col('PPK QTY'),
        'pc_ctn_os':    find_col('PC PER CTN OS'),
        'price':        find_col('STORE COST (CAD)') or find_col('PRICE'),
        'msrp':         find_col('MSRP'),
        'delivery':     find_col('DELIVERY'),                              # the DELIVERY column only: a blank stays blank, CRD is never substituted
        'segment':      find_col('SEGMENT'),
        'customer':     find_col('CUSTOMER'),
        'warehouse':    next((i for i, h in enumerate(header_upper) if h in ('WAREHOUSE', 'WHSE', 'STORAGE LOCATION', 'STORAGE LOC')), None),   # exact heading only: 'IN WHSE DATE' is a date
    }
    size_indices = [(i, str(header[i]).strip().upper())
                    for i, h in enumerate(header_upper) if h in KNOWN_SIZES]

    storage_by_customer = read_storage_locations(wb)
    season_label = read_season_label(all_rows)

    def gv_cust(r):
        i = idx.get('customer')
        return r[i] if i is not None and i < len(r) else None

    rows = []
    for r in data_rows:
        si = idx['style']
        if si is None or not r[si]: continue
        raw_style = str(r[si]).strip()
        if not raw_style: continue

        cust_kind = re.search(r'\b(OS|PPK|COMPLEX)\b', str(gv_cust(r) or '').upper())
        cust_kind = cust_kind.group(1) if cust_kind else ''
        ppk = is_ppk(raw_style) or cust_kind in ('PPK', 'COMPLEX')      # a complex prepack's article is the bare style: the customer name says COMPLEX
        sizes = {}
        for si2, sz_label in size_indices:
            if si2 < len(r) and r[si2]:
                try:
                    qty = int(float(r[si2]))
                    if qty > 0: sizes[sz_label] = qty
                except: pass

        def gv(key):
            i = idx.get(key)
            return r[i] if i is not None and i < len(r) else None

        entry = {
            'type':        'PPK' if ppk else 'OS',
            'pack_kind':   'COMPLEX' if cust_kind == 'COMPLEX' else ('PPK' if ppk else 'OS'),
            'complex':     cust_kind == 'COMPLEX',
            'raw_style':   raw_style,
            'style':       base_style(raw_style) if ppk else raw_style,
            'color_code':  str(gv('color_code') or '').strip(),
            'color':       str(gv('color') or '').strip().upper(),
            'desc':        str(gv('desc') or '').strip(),
            'price':       gv('price'),
            'msrp':        gv('msrp'),
            'delivery':    gv('delivery'),
            'segment':     str(gv('segment') or '').strip(),
            'customer':    str(gv('customer') or '').strip(),
            'warehouse':   str(gv('warehouse') or '').strip() or storage_by_customer.get(str(gv('customer') or '').strip().upper(), ''),
            'season_label': season_label,
            'pc_ctn':      gv('pc_ctn'),
            'pc_ctn_os':   gv('pc_ctn_os'),
            'ratio':       str(gv('ratio') or '').strip(),
            'row_pc':      sum(int(x) for x in dash_list(gv('ratio')) if x.isdigit()),      # this color's pieces per pack (complex prepacks)
            'sizes_str':   str(gv('sizes_col') or '').strip(),
            'pc_per_pac':  gv('pc_per_pac'),
            'pac_per_ctn': gv('pac_per_ctn'),
            'ppk_qty':     gv('ppk_qty'),
            'sizes':       sizes,
            'size_order':  [lab for _, lab in size_indices],
            'pack_kind':   'COMPLEX' if re.search(r'\bCOMPLEX\b', str(gv('customer') or '').upper()) else ('PPK' if ppk else 'OS'),
        }
        rows.append(entry)

    customers = list(set(r['customer'] for r in rows if r['customer']))
    return rows, customers

# ── 4b. legacy prebuys: OPEN STOCK / PREPACK tabs with a header block above the grid ──
def get_header_info(ws):
    """Header block of a legacy prebuy: labels sit in one row and their values directly BELOW them
    (D1 'SELECT WAREHOUSE:' over D2 'SJCA'). A value to the right is accepted only when it is not itself a label."""
    info = {'customer': None, 'warehouse': None, 'segment': None, 'hang_fold': None, 'tickets': None, 'brand': None, 'terms': None, 'terms_desc': None}
    keys = {'CUSTOMER:': 'customer', 'CUSTOMER': 'customer', 'BRAND:': 'brand',
            'SELECT WAREHOUSE:': 'warehouse', 'WAREHOUSE:': 'warehouse', 'WAREHOUSE': 'warehouse',
            'SEGMENT:': 'segment', 'SEGMENT': 'segment', 'STOCK SEGMENT:': 'segment',
            'HANG/FOLD:': 'hang_fold', 'HANG/FOLD': 'hang_fold', 'PRICE TICKETS:': 'tickets', 'PRICE TICKETS': 'tickets',
            'TERMS/ALLOWANCES': 'terms', 'DESCRIPTION OF TERMS': 'terms_desc'}
    other_labels = {'SAMPLES NEEDED:', 'PACKS PER CTN'}
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=6, values_only=True)]

    def is_label(v):
        t = str(v).strip().upper()
        return t.endswith(':') or t in keys or t in other_labels

    def value_for(r, i):
        below = grid[r + 1][i] if r + 1 < len(grid) and i < len(grid[r + 1]) else None
        right = grid[r][i + 1] if i + 1 < len(grid[r]) else None
        for cand in (below, right):
            if cand not in (None, '') and not is_label(cand):
                return str(cand).strip()
        return None

    for r, row in enumerate(grid):
        for i, v in enumerate(row):
            key = keys.get(str(v).strip().upper()) if v else None
            if key and info[key] is None:
                info[key] = value_for(r, i)
    return info

def find_header_row(ws, required_cols):
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [str(v).strip().upper() if v else '' for v in row]
        if all(c in cells for c in required_cols):
            return ri, cells
    return None, None

def has_ppk_column(ws):
    for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
        cells = [str(v).strip().upper() if v else '' for v in row]
        if 'PPK' in cells or 'PPK QTY' in cells:
            return True
    return False

def parse_os_tab(ws):
    info      = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE'])
    if not header_row: return info['customer'], []

    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {
        'style':   ci('STYLE'), 'code': ci('COLOR CODE'),
        'color':   ci('COLOR DESCRIPTION'), 'desc': ci('DESCRIPTION'),
        'delivery':ci('DELIVERY'), 'price': ci('PRICE'), 'msrp': ci('MSRP'),
        'segment': ci('SEGMENT'), 'pc_ctn': ci('PC PER CTN OS'),
    }
    size_idx = {canonical_size(h): ci_n for ci_n, h in enumerate(header_row)
                if canonical_size(h) in KNOWN_SIZES}

    rows = []
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        si = idx['style']
        if si is None or not row[si]: continue
        style = str(row[si]).strip()
        if not re.match(r'^[A-Z]{3}\d', style.upper()): continue
        sizes = {}
        for sz, ci_n in size_idx.items():
            v = row[ci_n]
            if v and str(v).strip() not in ('','0','None'):
                try: sizes[sz] = int(float(v))
                except: pass
        if not sizes: continue
        row_seg = str(row[idx['segment']]).strip() if idx['segment'] is not None and row[idx['segment']] else info.get('segment','')
        rows.append({
            'type':'OS', 'raw_style':style, 'style':style,
            'customer': info['customer'], 'warehouse': info.get('warehouse',''),
            'color_code': str(row[idx['code']]).strip()          if idx['code']  is not None and row[idx['code']]  else '',
            'color':      str(row[idx['color']]).strip().upper()  if idx['color'] is not None and row[idx['color']] else '',
            'desc':       str(row[idx['desc']]).strip()           if idx['desc']  is not None and row[idx['desc']]  else '',
            'delivery':   row[idx['delivery']]                    if idx['delivery'] is not None                    else None,
            'price':      row[idx['price']]                       if idx['price'] is not None                       else None,
            'msrp':       row[idx['msrp']]                        if idx['msrp'] is not None                        else None,
            'segment':    row_seg,
            'pc_ctn_os':  row[idx['pc_ctn']]                      if idx['pc_ctn'] is not None and row[idx['pc_ctn']] else None,
            'sizes':      sizes,
            'size_order': list(size_idx), 'pack_kind': 'OS',
            'ratio':'', 'sizes_str':'', 'pc_per_pac':None, 'pac_per_ctn':None, 'ppk_qty':None,
        })
    return info['customer'], rows

def parse_ppk_tab(ws):
    info = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE','COLOR CODE'])
    if not header_row: return info['customer'], []

    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {'style':ci('STYLE'),'code':ci('COLOR CODE'),'color':ci('COLOR DESCRIPTION'),'price':ci('PRICE'),'msrp':ci('MSRP')}
    desc_ci  = ci('DESCRIPTION') or ci('VENDOR DESCRIPTION')
    size_idx = {canonical_size(h): ci_n for ci_n, h in enumerate(header_row) if canonical_size(h) in KNOWN_SIZES}

    rows = []
    current_style = None; current_price = None; current_msrp = None
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        if not any(v for v in row): continue
        style_val = row[idx['style']] if idx['style'] is not None else None
        if style_val:
            sv = str(style_val).strip()
            if re.match(r'^[A-Z]{3}\d', sv.upper()):
                current_style = sv
                if idx['price'] is not None and row[idx['price']]: current_price = row[idx['price']]
                if idx['msrp']  is not None and row[idx['msrp']]:  current_msrp  = row[idx['msrp']]
        code_val = row[idx['code']] if idx['code'] is not None else None
        if code_val and current_style:
            sizes = {}
            for sz, ci_n in size_idx.items():
                v = row[ci_n]
                if v and str(v).strip() not in ('','0','None'):
                    try: sizes[sz] = int(float(v))
                    except: pass
            if not sizes: continue
            vals = list(sizes.values())
            g = math.gcd(*vals) if len(vals) > 1 else vals[0]
            ratio_vals = [v//g for v in vals]
            raw_style  = f"{current_style}-{str(code_val).strip()}"
            rows.append({
                'type':'PPK', 'raw_style':raw_style, 'style':current_style,
                'customer': info['customer'], 'warehouse': info.get('warehouse',''),
                'color_code': str(code_val).strip(),
                'color':      str(row[idx['color']]).strip().upper() if idx['color'] is not None and row[idx['color']] else '',
                'desc':       str(row[desc_ci]).strip() if desc_ci is not None and row[desc_ci] else '',
                'price':      current_price,
                'msrp':       current_msrp,
                'segment':    info.get('segment',''),
                'delivery':   None,
                'sizes':      sizes,
                'sizes_str':  '-'.join(sizes.keys()),
                'ratio':      '-'.join(str(r) for r in ratio_vals),
                'pc_per_pac': sum(ratio_vals),
                'ppk_qty':    g,
                'pac_per_ctn':None,
                'pc_ctn_os':  None,
                'size_order': list(size_idx), 'pack_kind': 'COMPLEX' if 'COMPLEX' in ws.title.upper() else 'PPK',
            })
    return info['customer'], rows

def parse_simple_ppk_tab(ws):
    info = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE'])
    if not header_row: return info['customer'], []
    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {'style':ci('STYLE'),'code':ci('COLOR CODE'),'color':ci('COLOR DESCRIPTION'),
           'desc':ci('DESCRIPTION'),'ppk':ci('PPK'),'price':ci('PRICE'),'msrp':ci('MSRP')}
    rows = []
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        si = idx['style']
        if si is None or not row[si]: continue
        style = str(row[si]).strip()
        if not re.match(r'^[A-Z]{3}\d', style.upper()): continue
        ppk_qty = row[idx['ppk']] if idx['ppk'] is not None else None
        if not ppk_qty: continue
        rows.append({
            'type':'PPK', 'raw_style':style, 'style':style,
            'customer': info['customer'], 'warehouse': info.get('warehouse',''),
            'color_code': str(row[idx['code']]).strip()          if idx['code']  is not None and row[idx['code']]  else '',
            'color':      str(row[idx['color']]).strip().upper()  if idx['color'] is not None and row[idx['color']] else '',
            'desc':       str(row[idx['desc']]).strip()           if idx['desc']  is not None and row[idx['desc']]  else '',
            'price':      row[idx['price']]                       if idx['price'] is not None                       else None,
            'msrp':       row[idx['msrp']]                        if idx['msrp'] is not None                        else None,
            'ppk_qty':    int(float(ppk_qty)),
            'segment':    info.get('segment',''), 'delivery': None,
            'sizes':{},'sizes_str':'','ratio':'','pc_per_pac':None,'pac_per_ctn':None,'pc_ctn_os':None,
            'size_order': [], 'pack_kind': 'PPK',
        })
    return info['customer'], rows

# ── 4c. the new prebuy template: HEADER + OS grid (units by size) + PPK grid (sizes / ratio / qty as text) ──
TEMPLATE_MARK = 'HIGH LIFE PREBUY TEMPLATE'

def is_template(wb):
    """The new prebuy template announces itself in HEADER!A1."""
    return 'HEADER' in wb.sheetnames and str(wb['HEADER']['A1'].value or '').strip().upper().startswith(TEMPLATE_MARK)

def read_template_header(ws):
    """HEADER tab = label in column A, value in column B. Sales type segment and storage location ONCE here;
    they apply to every row and travel on to the buysheet and the PR file."""
    raw = {}
    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
        if row and row[0] and len(row) > 1 and row[1] not in (None, ''):
            raw[str(row[0]).strip().upper()] = row[1]
    def text(key, upper=False):
        v = str(raw.get(key, '') or '').strip()
        return v.upper() if upper else v
    return {'customer': text('CUSTOMER'), 'brand': text('BRAND'), 'season_label': parse_season(raw.get('SEASON')), 'segment': text('SEGMENT', True),
            'warehouse': text('WAREHOUSE / STORAGE LOCATION', True), 'currency': text('CURRENCY', True),
            'hang_fold': text('HANG / FOLD', True), 'tickets': text('PRICE TICKETS'),
            'terms': raw.get('TERMS / ALLOWANCES'), 'terms_desc': text('DESCRIPTION OF TERMS')}

def template_columns(ws):
    """(header row number, {HEADING: column index}) - columns are found by heading, never by position."""
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=12, values_only=True), 1):
        cells = [re.sub(r'\s+', ' ', str(v)).strip().upper() if v is not None else '' for v in row]
        if 'STYLE' in cells and 'COLOR CODE' in cells:
            return ri, {h: i for i, h in enumerate(cells) if h}
    return None, {}

def template_header_fields(header):
    """What every row takes from the HEADER tab."""
    return {'customer': header['customer'], 'segment': header['segment'], 'warehouse': header['warehouse'], 'hang_fold': header['hang_fold'],
            'tickets': header['tickets'], 'currency': header['currency'], 'brand': header['brand'], 'season_label': header['season_label'],
            'terms': header['terms'], 'terms_desc': header['terms_desc']}

def template_rows(ws, kind, header):
    """OS tab: a grid, one row per style + color, units in the size columns (only the sizes the prebuy has)."""
    hr, C = template_columns(ws)
    if hr is None:
        return []
    size_cols = [(h, i) for h, i in C.items() if h in KNOWN_SIZES]
    def cell(row, h):
        i = C.get(h)
        return row[i] if i is not None and i < len(row) else None
    out = []
    for ri, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), hr + 1):
        style = str(cell(row, 'STYLE') or '').strip().upper()
        if not style or str(cell(row, 'NOTES') or '').strip().upper().startswith('EXAMPLE'):
            continue
        units = {}
        for h, i in size_cols:
            q = num(row[i]) if i < len(row) else None
            if q and q > 0:
                units[h] = int(round(q))
        if not units:
            continue
        pc = cell(row, 'PC PER CTN OS')
        out.append({**template_header_fields(header), 'segment': str(cell(row, 'SEGMENT') or '').strip().upper() or header['segment'],
                    'style': style, 'color_code': str(cell(row, 'COLOR CODE') or '').strip().upper(),
                    'color': str(cell(row, 'COLOR DESCRIPTION') or '').strip().upper(), 'desc': str(cell(row, 'DESCRIPTION') or '').strip(),
                    'price': cell(row, 'SELL PRICE'), 'msrp': cell(row, 'MSRP'), 'delivery': cell(row, 'DELIVERY'),
                    'notes': str(cell(row, 'NOTES') or '').strip(), 'size_order': [h for h, _ in size_cols], 'source': f"{ws.title}!{ri}",
                    'type': 'OS', 'pack_kind': 'OS', 'raw_style': style, 'pc_ctn': pc, 'pc_ctn_os': pc, 'sizes': units, 'ratio': '',
                    'sizes_str': '', 'pc_per_pac': None, 'pac_per_ctn': None, 'ppk_qty': None})
    return out

def template_ppk_rows(ws, header, complex_mode=False):
    """PPK tab: a grid, one ROW per prepack. No size columns: SIZES, SIZE RATIO and PPK QTY are typed as text.
    The buysheet article is built from the two fields sales type: STYLE + '-' + COLOR CODE (CFW271SH01 + S470 -> CFW271SH01-S470).
    COMPLEX tab (complex_mode): one row per COLOR; rows with the same STYLE are one prepack. PPK QTY is that color's
    own pieces; those rows become one complex_group and the article STYLE-A (see apply_complex_split)."""
    hr, C = template_columns(ws)
    if hr is None:
        return []
    def cell(row, h):
        i = C.get(h)
        return row[i] if i is not None and i < len(row) else None
    out = []
    for ri, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), hr + 1):
        style = str(cell(row, 'STYLE') or '').strip().upper()
        code = str(cell(row, 'COLOR CODE') or '').strip().upper()
        if (not style and not code) or str(cell(row, 'NOTES') or '').strip().upper().startswith('EXAMPLE'):
            continue                  # an empty row or the EXAMPLE row; a half-filled one is kept so it shows up as incomplete
        qty = num(cell(row, 'QTY') if 'QTY' in C else cell(row, 'PPK QTY'))
        ratio = dash_list(cell(row, 'SIZE RATIO'))
        extra = ({'raw_style': style, 'complex_group': style,
                  'row_pc': sum(int(x) for x in ratio if x.isdigit())} if complex_mode else {'raw_style': f"{style}-{code}"})
        out.append({**template_header_fields(header), **extra, 'segment': str(cell(row, 'SEGMENT') or '').strip().upper() or header['segment'],
                    'style': style, 'color_code': code,
                    'color': str(cell(row, 'COLOR DESCRIPTION') or '').strip().upper(), 'desc': str(cell(row, 'DESCRIPTION') or '').strip(),
                    'price': cell(row, 'SELL PRICE'), 'msrp': cell(row, 'MSRP'), 'delivery': cell(row, 'DELIVERY'),
                    'ratio': '-'.join(ratio), 'sizes_str': '-'.join(dash_list(cell(row, 'SIZES'))),
                    'pc_per_pac': num(cell(row, 'PC PER PAC')) or (sum(int(x) for x in ratio if x.isdigit()) or None),   # no column on the template: pieces per pack = the ratio added up
                    'pac_per_ctn': num(cell(row, 'PAC PER CTN')),
                    'ppk_qty': int(round(qty)) if qty is not None else None,
                    'notes': str(cell(row, 'NOTES') or '').strip(), 'source': f"{ws.title}!{ri}",
                    'type': 'PPK', 'pack_kind': 'PPK', 'sizes': {}, 'pc_ctn': None, 'pc_ctn_os': None, 'size_order': []})
    return out

def apply_complex_split(ppk):
    """COMPLEX tab rows: the rows of one 'complex_group' (same style) are one prepack. PPK QTY on each row is that
    color's own pieces, as typed; the whole-prepack size is every color's ratio added together."""
    groups = {}
    for r in ppk:
        if r.get('complex_group'):
            groups.setdefault(r['complex_group'], []).append(r)
    for style, grp in groups.items():
        whole = sum(x['row_pc'] for x in grp)
        article = f"{style}-A"                     # a complex prepack is the style plus a letter suffix, never a color code
        for x in grp:
            x.update(complex=True, pack_kind='COMPLEX', pc_per_pac=whole, raw_style=article)

def parse_template(wb):
    header = read_template_header(wb['HEADER'])
    rows = template_rows(wb['OS'], 'OS', header) if 'OS' in wb.sheetnames else []
    ppk = template_ppk_rows(wb['PPK'], header) if 'PPK' in wb.sheetnames else []
    cx = template_ppk_rows(wb['COMPLEX'], header, complex_mode=True) if 'COMPLEX' in wb.sheetnames else []
    apply_complex_split(cx)
    ppk = ppk + cx
    return rows + ppk, ([header['customer']] if header['customer'] else [])

# ── 4d. entry point ──
def parse_prebuy(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    if detect_file_type(wb) == 'buysheet':
        rows, customers = parse_buysheet(wb)
        return rows, customers, 'buysheet'
    if is_template(wb):
        rows, customers = parse_template(wb)
        return rows, customers, 'template'
    all_rows=[]; customers=[]
    SKIP=['DO NOT','COLOR CODE','SETTINGS','REFERENCE','LOOKUP','TERMS']
    for sheet in wb.sheetnames:
        ws=wb[sheet]; snu=sheet.strip().upper()
        if any(x in snu for x in SKIP): continue
        if any(x in snu for x in ['PREPACK','COMPLEX']): cust,rows=parse_ppk_tab(ws)
        elif 'PPK' in snu: cust,rows=parse_ppk_tab(ws)
        elif has_ppk_column(ws): cust,rows=parse_simple_ppk_tab(ws)
        else: cust,rows=parse_os_tab(ws)
        if rows:
            hf = get_header_info(ws)
            for r in rows:
                r.setdefault('hang_fold', hf.get('hang_fold') or '')
                r.setdefault('tickets', hf.get('tickets') or '')
                r.setdefault('brand', hf.get('brand') or '')
                r.setdefault('terms', hf.get('terms') or '')
                r.setdefault('terms_desc', hf.get('terms_desc') or '')
            if cust and cust not in customers: customers.append(cust)
            all_rows.extend(rows)
    return all_rows, customers, 'prebuy'

# ═══════════════════════════════════════════════════════════════════════════════
# 5. VALIDATION - every row against what SAP holds
# ═══════════════════════════════════════════════════════════════════════════════
class SapView:
    """Everything the checks need from HANA for one file, fetched once and indexed."""

    def __init__(self, rows):
        os_rows  = [r for r in rows if r['type'] == 'OS']
        ppk_rows = [r for r in rows if r['type'] == 'PPK']
        self.generic_styles = sorted({r['style'] for r in os_rows})
        self.ppk_articles   = sorted({r['raw_style'] for r in ppk_rows})
        ppk_styles          = sorted({generic_from_ppk(a) for a in self.ppk_articles})

        conn = connect_hana()
        try:
            self.generic_exists = query_hana_generics(conn, self.generic_styles)            # (style, code, size) -> variant
            self.ppk_exists, self.ppk_details = query_hana_ppks(conn, self.ppk_articles)
            self.ppk_desc = query_hana_ppk_descriptions(conn, self.ppk_articles)           # prepack article -> its own name
            self.ppks_under_generic = query_hana_ppks_for_generics(conn, ppk_styles)         # every prepack SAP has for these styles
            ppk_gen_exists = query_hana_generics(conn, ppk_styles)
            self.generic_desc = query_hana_ppk_descriptions(conn, sorted(set(self.generic_styles) | set(ppk_styles)))   # the generic ARTICLE's own name
            variant_arts = {v['article'] for v in self.generic_exists.values() if v.get('article')}
            self.price_keys = set(self.ppk_articles) | variant_arts | set(self.generic_styles) | set(self.ppks_under_generic)   # found-by-content prepacks need a price too
            self.prices = query_hana_prices(conn, self.price_keys)
            self.code_names = query_hana_color_codes(conn, {r['color_code'] for r in os_rows})  # code -> names SAP uses for it anywhere
        finally:
            conn.close()

        # per-style facts, taken from the variants. A prepack row ('STYLE-COLOR') never supplies the style description.
        self.style_desc, self.style_pc_ctn, self.style_color = {}, {}, {}
        self.style_colors, self.color_sizes = defaultdict(set), defaultdict(set)
        self.style_pc_ctn_votes = defaultdict(lambda: defaultdict(int))
        for (g, c, s), v in self.generic_exists.items():
            self.style_colors[g].add(c); self.color_sizes[(g, c)].add(s)
            if v.get('color_name') and (g, c) not in self.style_color:
                self.style_color[(g, c)] = v['color_name']
            if v.get('desc') and g not in self.style_desc and '-' not in (v.get('article') or ''):
                self.style_desc[g] = v['desc']
            if v.get('pc_per_ctn') and s:                       # only real size variants vote, never the generic row
                self.style_pc_ctn_votes[g][str(v['pc_per_ctn']).strip()] += 1
        for g, votes in self.style_pc_ctn_votes.items():      # most common value across the style's variants
            self.style_pc_ctn[g] = max(votes, key=lambda k: votes[k])
        self.ppk_style_desc = {}
        for (g, c, s), v in ppk_gen_exists.items():
            if v.get('desc') and g not in self.ppk_style_desc and '-' not in (v.get('article') or ''):
                self.ppk_style_desc[g] = v['desc']

    def find_ppk_by_content(self, style, rows):
        """A prepack under this generic whose components are exactly these rows' (color code, size, units per pack).
        Works for one-color prepacks and for complex ones (several rows, one article). '' when none matches."""
        want = set()
        for r in rows:
            sizes, ratio = dash_list(r.get('sizes_str')), dash_list(r.get('ratio'))
            if not sizes or len(sizes) != len(ratio): return ''
            try: want |= {(r['color_code'], sz, int(q)) for sz, q in zip(sizes, ratio) if int(q) > 0}
            except ValueError: return ''
        if not want: return ''
        for art, e in self.ppks_under_generic.items():
            if e['generic'] == style and {(c, sz, q) for c, _, sz, q, _ in e['comps'] if q > 0} == want:
                return art
        return ''

    def detail_for(self, art):
        """ppk_details-shaped record for a prepack found by content, so the same checks apply."""
        e = self.ppks_under_generic.get(art)
        if not e: return {}
        return {'color_name': e['comps'][0][1] if e['comps'] else '', 'hana_desc': e['desc'], 'pc_per_pac': e['pc_per_pac'],
                'variants': [(sort, sz, q) for _, _, sz, q, sort in e['comps']], 'comps': e['comps'],
                'components': e.get('components', []), 'seen': set()}

    def pc_ctn_note(self, style):
        votes = self.style_pc_ctn_votes.get(style, {})
        if len(votes) <= 1: return ''
        return 'SAP variants disagree on PC/CTN: ' + ', '.join(f"{k} on {n} variant{'s' if n != 1 else ''}" for k, n in sorted(votes.items(), key=lambda kv: -kv[1]))

    def generic_description(self, style):
        """The generic's description: its own article record first, then its variants - never a prepack's."""
        return self.generic_desc.get(style) or self.style_desc.get(style) or self.ppk_style_desc.get(style) or ''

    def os_price(self, style, code, size):
        """SAP price for one OS size plus a plain-English trace. The variant's own price wins, then the generic's.
        Prepack prices are never used here: they are priced separately and can differ from the generic."""
        art = (self.generic_exists.get((style, code, size.upper())) or {}).get('article')
        steps = []
        if art:
            if art in self.prices:
                return self.prices[art], 'variant', f'variant {art}: {self.prices[art]:.2f}'
            steps.append(f'variant {art}: no price row')
        else:
            steps.append(f'variant {style}/{code}/{size}: not found in SAP')
        if style in self.prices:
            steps.append(f'generic {style}: {self.prices[style]:.2f}')
            return self.prices[style], 'generic', ' -> '.join(steps)
        steps.append(f'generic {style}: no price row')
        return None, '', ' -> '.join(steps)

    def os_color_check(self, style, code, buy_color):
        """Color code + description for one OS row, judged against HANA only -> (check, SAP name compared with, note).
        HANA only shows codes used on ARTICLES: a code created in the color master (YOUR_COLOR_MASTER_TABLE) but not on any article
        yet looks the same as one that does not exist, so that case is a warning, never 'needs creating'."""
        cmp = lambda c: normalize_color(c).replace(' ', '')
        buy = cmp(buy_color)
        on_style = self.style_color.get((style, code))
        if on_style:
            if not buy: return f'⚠ NO BUY COLOR (SAP={on_style})', on_style, 'The file has no color description for this code.'
            if buy == cmp(on_style): return '✓', on_style, ''
            return f'✗ SAP={on_style}', on_style, f'SAP calls {code} "{on_style}" on this style.'
        names = sorted(self.code_names.get(code, []))
        if names:
            if buy and not any(buy == cmp(n) for n in names):
                return f'✗ SAP={names[0]}', names[0], f'SAP calls {code} "{names[0]}" on other articles.'
            return '⚠ NEW ON STYLE', names[0], f'{code} ({names[0]}) is used on other SAP articles but not on this style yet.'
        if code:
            return ('⚠ NOT ON ANY ARTICLE', '',
                    f'No SAP article uses color code {code}. It may still exist in the color master (YOUR_COLOR_MASTER_TABLE): '
                    'this check only sees articles, so it cannot say whether the code needs creating.')
        return '⚠ NO COLOR CODE', '', 'The row has no color code.'

    def missing_info(self, style, code, size):
        """Why an OS size row is not in SAP: the whole style, the color on this style, or only this size.
        Always says what SAP does hold, so a wrong 'needs creating' can be spotted at a glance."""
        if style not in self.style_colors:
            return 'style', f'Style {style} is not in SAP - the generic needs creating.'
        if code not in self.style_colors[style]:
            norm = lambda c: re.sub(r'[^A-Z0-9]', '', str(c).upper())
            similar = [c for c in self.style_colors[style] if norm(c) == norm(code)]
            hint = f' SAP has {similar[0]} - the same code spelled differently.' if similar else ''
            known = sorted(self.code_names.get(code, []))
            known_note = (f' The code is already in SAP as "{known[0]}" - add it to this style.' if known
                          else ' The code is not on any article in SAP yet.')
            return 'color', f'Color {code} is not on style {style} in SAP.{hint}{known_note} SAP has: {sap_list(self.style_colors[style])}.'
        return 'size', f'Size {size} is missing for color {code} on {style}. SAP has: {sap_list(self.color_sizes[(style, code)])}.'


def sap_list(items, limit=10):
    items = sorted(items, key=lambda x: (SIZE_ORDER.index(x) if x in SIZE_ORDER else 99, x))
    return ', '.join(items[:limit]) + (f' (+{len(items) - limit} more)' if len(items) > limit else '')


def desc_check_text(buy, sap):
    """Style description against SAP's. A SAP name may carry a ' - COLOR' tail; the name without it also matches."""
    if not sap:
        return '⚠ NO DATA'
    return '✓' if buy.upper() in (sap, sap.rsplit(' - ', 1)[0].strip()) else '✗ MISMATCH'


def check_os_row(r, sap):
    """One open-stock prebuy row (style + color) -> one result per size."""
    style, code, desc = r['style'], r['color_code'], r['desc']
    price, msrp = num(r.get('price')), num(r.get('msrp'))

    sap_desc = sap.generic_description(style)
    desc_check = desc_check_text(desc, sap_desc)
    desc_note = desc_diff_note(desc, sap_desc) if desc_check.startswith('✗') else ''

    buy_pc = num(r.get('pc_ctn'))
    def pc_ctn_for(sz):
        """PC/CTN check for ONE size: that variant's own value; the style's majority only when the variant has none."""
        v = sap.generic_exists.get((style, code, sz.upper()), {})
        sap_pc, src = (v.get('pc_per_ctn'), f"variant {v.get('article')}") if v.get('pc_per_ctn') else (sap.style_pc_ctn.get(style), 'style majority')
        if buy_pc and sap_pc:
            ok = buy_pc == float(sap_pc)
            return ('✓' if ok else f'✗ SAP={int(float(sap_pc))} BUY={int(buy_pc)}'), (f'SAP {int(float(sap_pc))} per carton ({src})' if not ok else '')
        if buy_pc: return '⚠ NO SAP PC/CTN', 'SAP has no pieces-per-carton on this variant or style.'
        if sap_pc: return f'⚠ NO BUY PC/CTN (SAP={int(float(sap_pc))})', f'The prebuy has no PC PER CTN; SAP says {int(float(sap_pc))} ({src}).'
        return '—', ''

    total_qty, pc_os = sum(r['sizes'].values()), r.get('pc_ctn_os')
    if pc_os and total_qty:
        try: round_check = '✓' if total_qty % float(pc_os) == 0 else f'✗ {total_qty}/{pc_os}'
        except Exception: round_check = '⚠ NO DATA'
    else:
        round_check = '—'

    color_check, sap_color, color_note = sap.os_color_check(style, code, r.get('color', ''))
    common = {
        'style': style, 'color_code': code, 'buy_price': price, 'buy_msrp': msrp,
        'color_check': color_check, 'color_note': color_note, 'buy_color': r.get('color', ''), 'sap_color': sap_color,
        'code_in_sap': bool(sap.code_names.get(code)),
        'season': season_text(r.get('season_label'), style), 'season_src': 'header' if r.get('season_label') else 'style number',
        'customer': r.get('customer', ''), 'segment': r.get('segment', ''), 'delivery': iso_date(r.get('delivery')), 'warehouse': r.get('warehouse', ''),
        'buy_desc': desc, 'sap_desc': sap_desc, 'desc_check': desc_check, 'desc_note': desc_note,
        'round_check': round_check,
    }
    out = []
    for sz, qty in r['sizes'].items():
        exists = (style, code, sz.upper()) in sap.generic_exists
        kind, reason = ('', '') if exists else sap.missing_info(style, code, sz)
        sap_price, source, trace = sap.os_price(style, code, sz)
        pc_check, pc_note = pc_ctn_for(sz)
        out.append({**common, 'size': sz, 'qty': qty, 'sap_article': (sap.generic_exists.get((style, code, sz.upper())) or {}).get('article') or '',
                    'status': '✓ EXISTS' if exists else '✗ MISSING',
                    'pc_ctn_check': pc_check, 'pc_ctn_note': ' '.join(x for x in (pc_note, sap.pc_ctn_note(style)) if x),
                    'missing_kind': kind, 'missing_reason': reason, 'sap_price': sap_price, 'price_source': source,
                    'price_check': price_check_text(price, sap_price), 'price_trace': trace})
    return out


def check_ppk_row(r, sap, group=None):
    """One prepack row -> one result. Descriptions are not checked on prepacks (sales abbreviate them);
    the generic's / prepack's SAP name is carried for display only."""
    art, style, code = r['raw_style'], r['style'], r['color_code']
    buy_color, buy_ratio, buy_sizes, desc = r['color'], r.get('ratio', ''), r.get('sizes_str', ''), r['desc']
    price, msrp = num(r.get('price')), num(r.get('msrp'))
    pc_per_pac, pac_per_ctn, ppk_qty = r.get('pc_per_pac'), r.get('pac_per_ctn'), r.get('ppk_qty')

    exists = art in sap.ppk_exists
    detail = sap.ppk_details.get(art, {})
    match_note = ''
    if not exists:                                           # not under that name: look for it by content under the generic
        found = sap.find_ppk_by_content(style, group or [r])
        if found:
            exists, detail = True, sap.detail_for(found)
            match_note = f'Exists in SAP as {found}' + ('' if found == art else f' (the prebuy would call it {art})')
            art = found
    sap_color = detail.get('color_name', '') if detail else ''
    sap_sizes, sap_ratio = build_ratio_sizes(detail) if detail else ('', '')

    part = color_from_ppk(r['raw_style'])
    cc_match = ('✓' if code in part or part in code else '✗ MISMATCH') if (code and part) else '⚠ NO DATA'

    # prepack descriptions are NOT checked (they are often abbreviated); SAP's is shown for information only
    ppk_full = sap.ppk_desc.get(art, '')
    sap_desc = sap.generic_description(style) or (ppk_full.split(' - ')[0].strip() if ppk_full else '')
    desc_check, desc_note = '—', ''

    sap_price = sap.prices.get(art) if art in sap.prices else sap.prices.get(r['raw_style'])
    price_trace = f'PPK article {art}: {sap_price:.2f}' if sap_price is not None else f'PPK article {art}: no price row in YOUR_PRICE_VIEW'

    color_note = ''
    if not exists: color_check = '—'
    elif buy_color and sap_color:
        color_check = '✓' if strip_color(buy_color) == strip_color(sap_color) else f'✗ SAP={sap_color}'
        if color_check != '✓': color_note = f'SAP\'s components of {art} are "{sap_color}"; the prebuy says "{buy_color}".'
    else: color_check = '⚠ NO DATA'

    ratio_sum = sum(int(x) for x in buy_ratio.split('-') if x.strip().isdigit()) if buy_ratio else None
    buy_pc = float(pc_per_pac) if pc_per_pac else None
    sap_pc = float(detail.get('pc_per_pac') or 0) if detail and detail.get('pc_per_pac') else None
    if ratio_sum and buy_pc and sap_pc:
        pc_ppk_check = '✓' if ratio_sum == buy_pc == sap_pc else f'✗ RATIO={ratio_sum} BUY={int(buy_pc)} SAP={int(sap_pc)}'
    elif ratio_sum and buy_pc:
        pc_ppk_check = '✓' if ratio_sum == buy_pc else f'✗ RATIO={ratio_sum} BUY={int(buy_pc)}'
    else:
        pc_ppk_check = '⚠ NO DATA'

    sizes_check, ratio_check, ratio_note = compare_ratio(exists, buy_sizes, buy_ratio, detail)

    if ppk_qty and pac_per_ctn and pc_per_pac:
        try:
            full = float(pc_per_pac) * float(pac_per_ctn)
            round_check = '✓' if float(ppk_qty) % full == 0 else f'✗ {ppk_qty}/{full}'
        except Exception: round_check = '⚠ NO DATA'
    else:
        round_check = '—'

    if r.get('complex'):          # one article, several colors: check THIS color's components inside the prepack
        cc_match = '—'                                                       # a complex article carries a suffix, not a color code
        comps = (sap.ppks_under_generic.get(art) or {}).get('comps') or detail.get('comps') or []
        mine = [c for c in comps if c[0] == code]
        if not exists:
            color_check = ratio_check = sizes_check = '—'; ratio_note = ''
        elif comps and not mine:
            color_check = sizes_check = ratio_check = f'✗ NOT IN {art}'
            ratio_note = f'SAP prepack {art} holds colors ' + ', '.join(sorted({f"{c[0]} ({c[1]})" for c in comps})) + f'; {code} is not one of them.'
            color_note = ratio_note
        elif mine:
            sap_color = mine[0][1]
            color_check = '✓' if strip_color(buy_color) == strip_color(sap_color) else f'✗ SAP={sap_color}'
            color_note = '' if color_check == '✓' else f'In {art}, color {code} is "{sap_color}"; the prebuy says "{buy_color}".'
            sizes_check, ratio_check, ratio_note = compare_ratio(True, buy_sizes, buy_ratio, {'variants': [(sort, sz, q) for _, _, sz, q, sort in mine]})
            sap_sizes, sap_ratio = '-'.join(c[2] for c in mine), '-'.join(str(c[3]) for c in mine)
        else:
            color_check = sizes_check = ratio_check = '⚠ NO DATA'; ratio_note = f'SAP returned no components for {art}.'
        whole_sap = sum(c[3] for c in comps) if comps else None
        whole_buy = r.get('pc_per_pac')
        if whole_sap and whole_buy:
            pc_ppk_check = '✓' if int(whole_sap) == int(whole_buy) else f'✗ SAP={int(whole_sap)} PREBUY={int(whole_buy)} (whole prepack)'
        else:
            pc_ppk_check = '⚠ NO DATA' if exists else '—'
        # rounding: this color's pieces must be whole packs, and every color of the prepack must give the SAME number of packs
        row_pc = r.get('row_pc') or 0
        packs = (ppk_qty / row_pc) if (ppk_qty and row_pc) else None
        others = [(x['color_code'], x['ppk_qty'] / x['row_pc']) for x in (group or []) if x.get('row_pc') and x.get('ppk_qty')]
        if packs is None: round_check = '—'
        elif abs(packs - round(packs)) > 1e-9: round_check = f'✗ {ppk_qty}/{row_pc} = {packs:.2f} packs'
        elif any(abs(p - packs) > 1e-9 for _, p in others): round_check = '✗ PACKS DIFFER: ' + ', '.join(f'{c} {p:g}' for c, p in others)
        else:
            full = packs / float(pac_per_ctn) if pac_per_ctn else None
            round_check = '✓' if (full is None or abs(full - round(full)) < 1e-9) else f'✗ {packs:g} packs / {pac_per_ctn} per ctn = {full:.2f} cartons'
        if r.get('qty_conflict'):
            round_check = '⚠ QTY DIFFERS'          # the split used the first PPK QTY typed; the rows do not agree

    return {
        'ppk_article': art, 'prebuy_article': r['raw_style'], 'match_note': match_note, 'color_code': code, 'cc_check': cc_match, 'exists': '✓ EXISTS' if exists else '✗ MISSING',
        'complex': bool(r.get('complex')),
        'season': season_text(r.get('season_label'), style), 'season_src': 'header' if r.get('season_label') else 'style number',
        'pc_ppk_check': pc_ppk_check,
        'buy_price': price, 'sap_price': sap_price, 'price_check': price_check_text(price, sap_price), 'price_trace': price_trace, 'buy_msrp': msrp,
        'buy_color': buy_color, 'sap_color': sap_color, 'color_check': color_check, 'color_note': color_note,
        'buy_sizes': buy_sizes, 'sap_sizes': sap_sizes, 'sizes_check': sizes_check,
        'buy_ratio': buy_ratio, 'sap_ratio': sap_ratio, 'ratio_check': ratio_check, 'ratio_note': ratio_note,
        'buy_desc': desc, 'sap_desc': sap_desc, 'desc_check': desc_check, 'desc_note': desc_note, 'ppk_desc': ppk_full,
        'round_check': round_check,
        'components': sorted(((sap.ppks_under_generic.get(art) or {}).get('components') or detail.get('components') or []), key=lambda c: (c.get('code', ''), c.get('sort', ''))),
        'qty': sum(r['sizes'].values()) if r.get('sizes') else (ppk_qty or 0),
        'customer': r.get('customer', ''), 'segment': r.get('segment', ''), 'delivery': iso_date(r.get('delivery')), 'warehouse': r.get('warehouse', ''),
    }


GEN_CHECKS = (('status', 'Exists'), ('color_check', 'Color'), ('price_check', 'Price'), ('desc_check', 'Description'), ('pc_ctn_check', 'PC / CTN'), ('round_check', 'Rounding'))
PPK_CHECKS = (('exists', 'Exists'), ('cc_check', 'Color code'), ('color_check', 'Color'), ('sizes_check', 'Sizes'), ('ratio_check', 'Ratio'),
              ('pc_ppk_check', 'PC / PPK'), ('price_check', 'Price'), ('round_check', 'Rounding'))

def check_counts(results, checks):
    """For each check: how many rows passed (✓), failed (✗) and were only warned about (⚠). '—' (not applicable) is not counted."""
    out = []
    for key, label in checks:
        vals = [str(r.get(key, '')) for r in results]
        out.append({'key': key, 'label': label, 'ok': sum(v.startswith('✓') for v in vals), 'bad': sum(v.startswith('✗') for v in vals),
                    'warn': sum(v.startswith('⚠') for v in vals), 'info': sum(v.startswith('ℹ') for v in vals),
                    'na': sum(v.startswith('—') or v == '' for v in vals)})
    return out

VARY_PER_ROW = ('qty', 'customer', 'segment', 'delivery', 'warehouse', 'season', 'season_src')   # what legitimately differs between rows of one finding

def collapse_results(results, vary=VARY_PER_ROW):
    """One entry per DISTINCT finding. A buysheet has a DATA row per customer / delivery / color, so the same style or
    prepack is on many rows, and each row is checked against SAP. An identical result must not be listed again and again.
    Rows merge only when everything that was compared (the buy values, SAP's values and every check outcome) is the same;
    what differs between them (customer, delivery, quantity ...) is kept as a count and as lists."""
    groups = {}
    for r in results:
        key = json.dumps({k: v for k, v in r.items() if k not in vary}, sort_keys=True, default=str)
        g = groups.get(key)
        if g is None:
            g = groups[key] = dict(r, count=0, qty=0, customers=[], deliveries=[], segments=[])
        g['count'] += 1
        g['qty'] += r.get('qty') or 0
        for field, bucket in (('customer', 'customers'), ('delivery', 'deliveries'), ('segment', 'segments')):
            v = r.get(field)
            if v and v not in g[bucket]:
                g[bucket].append(v)
    return list(groups.values())

PPK_ISSUE_KEYS = ('cc_check', 'price_check', 'color_check', 'sizes_check', 'ratio_check', 'desc_check', 'round_check', 'pc_ppk_check')

def validate_all(rows):
    sap = SapView(rows)
    generic = [res for r in rows if r['type'] == 'OS' for res in check_os_row(r, sap)]
    def group_key(r):
        """The rows that make up ONE order of a complex prepack: same article, same customer, same delivery. The same article
        can be ordered several times in one file (two customers); each order is checked against itself, not against the others."""
        return (r['raw_style'], str(r.get('customer') or ''), str(r.get('delivery') or ''))
    by_art = defaultdict(list)
    for r in rows:
        if r['type'] == 'PPK': by_art[group_key(r)].append(r)
    ppk = [check_ppk_row(r, sap, by_art[group_key(r)]) for r in rows if r['type'] == 'PPK']

    diag = {
        'os_styles_in_file':           len(sap.generic_styles),
        'ppk_articles_in_file':        len(sap.ppk_articles),
        'sap_variants_returned':       len(sap.generic_exists),
        'sap_styles_with_variants':    len({g for (g, c, s) in sap.generic_exists}),
        'sap_styles_with_pc_per_ctn':  len(sap.style_pc_ctn),
        'sap_ppk_articles_found':      len(sap.ppk_exists),
        'price_keys_queried':          len(sap.price_keys),
        'price_rows_returned':         len(sap.prices),
        'os_size_rows':                len(generic),
        'os_size_rows_with_sap_price': sum(1 for r in generic if r['sap_price'] is not None),
        'ppk_rows':                    len(ppk),
        'ppk_rows_with_sap_price':     sum(1 for r in ppk if r['sap_price'] is not None),
        'os_color_codes_not_on_any_article': sorted({r['color_code'] for r in generic if r['color_check'] == '⚠ NOT ON ANY ARTICLE'})[:20],
        'os_missing_by_kind':          {k: sum(1 for r in generic if r['missing_kind'] == k) for k in ('style', 'color', 'size')},
        'desc_from_generic_record':    sum(1 for st in sap.generic_styles if st in sap.generic_desc),
        'desc_from_variants_fallback': sum(1 for st in sap.generic_styles if st not in sap.generic_desc and st in sap.style_desc),
        'rows_with_buy_price':         sum(1 for r in rows if num(r.get('price'))),
        'rows_with_buy_msrp':          sum(1 for r in rows if num(r.get('msrp'))),
        'os_styles_without_sap_price': sorted({r['style'] for r in generic if r['sap_price'] is None})[:20],
        'ppk_without_sap_price':       [r['ppk_article'] for r in ppk if r['sap_price'] is None][:20],
    }
    generic_u, ppk_u = collapse_results(generic), collapse_results(ppk)
    summary = {
        'gen_exists':  sum(1 for r in generic_u if r['status'] == '✓ EXISTS'),
        'gen_missing': sum(1 for r in generic_u if r['status'] == '✗ MISSING'),
        'rows_os_checked': len(generic), 'rows_ppk_checked': len(ppk), 'lines_os': len(generic_u), 'lines_ppk': len(ppk_u),
        'ppk_exists':  len({r['ppk_article'] for r in ppk if r['exists'] == '✓ EXISTS'}),
        'ppk_missing': len({r['ppk_article'] for r in ppk if r['exists'] == '✗ MISSING'}),
        'ppk_issues':  sum(1 for r in ppk_u if any('✗' in str(r.get(k, '')) for k in PPK_ISSUE_KEYS)),
        'total_os':    sum(1 for r in rows if r['type'] == 'OS'),
        'total_ppk':   len({r['ppk_article'] for r in ppk}),
        'checks_ppk':  check_counts(ppk_u, PPK_CHECKS),
        'checks_gen':  check_counts(generic_u, GEN_CHECKS),
    }
    return {'diag': diag, 'generic': generic, 'ppk': ppk, 'generic_unique': generic_u, 'ppk_unique': ppk_u, 'rows': rows, 'summary': summary}


# ═══════════════════════════════════════════════════════════════════════════════
# 6. OUTPUTS - buysheet DATA + HEADER OUT, creation templates, PR CSV
# ═══════════════════════════════════════════════════════════════════════════════
NAVY  = PatternFill('solid', start_color='1F4E79')
WBOLD = Font(bold=True, color='FFFFFF', name='Arial', size=10)
DFONT = Font(name='Arial', size=10)
BFONT = Font(bold=True, name='Arial', size=10)
THIN  = Side(style='thin', color='BFBFBF')
BDR   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR   = Alignment(horizontal='center', vertical='center')
LEFT  = Alignment(horizontal='left',   vertical='center', indent=1)

# the buysheet's DATA sheet: these columns, in this order, with the prebuy's size columns between them
DATA_HEAD = ['CUSTOMER','STYLE','SEGMENT','COLOR CODE','COLOR','DESCRIPTION','CRD','DELIVERY','PRICE','MSRP',
             'SIZE RATIO','SIZES','PC PER PAC','PAC PER CTN','PPK QTY']
DATA_TAIL = ['PC PER CTN OS','DATE UPDATED','COMMENTS']
HEADER_OUT = ['CUSTOMER','Tickets','Hang/Fold','PER BAG','Warehouse','Pre-Buy Rcvd','Buysheet Created','REQ#','TERMS','TERMS DESC']

def buysheet_customer(r):
    """The buysheet code reads OS / PPK / COMPLEX out of the customer name, so every customer carries one of them."""
    kind = 'OS' if r.get('type') == 'OS' else ('COMPLEX' if r.get('complex') or r.get('pack_kind') == 'COMPLEX' else 'PPK')
    name = str(r.get('customer') or '').strip().upper()
    if re.search(rf'\b{kind}\b', name):
        return name
    return f"{name} - {kind}" if name else kind

def data_size_columns(rows):
    """Size columns of the buysheet: only sizes the PREBUY has quantities for, in the prebuy's own column order.
    Nothing is assumed: no fixed list, no renaming."""
    used = {sz for r in rows if r.get('type') == 'OS' for sz in (r.get('sizes') or {})}
    order = []
    for r in rows:
        for sz in r.get('size_order') or []:
            if sz in used and sz not in order:
                order.append(sz)
    return order + sorted(used - set(order))

def make_data_sheet(rows, customers):
    sizes   = data_size_columns(rows)
    headers = DATA_HEAD + sizes + DATA_TAIL
    col     = {h: i for i, h in enumerate(headers, 1)}
    wb = Workbook(); ws = wb.active; ws.title = "DATA"
    if rows:
        style = rows[0]['style']
        ws.cell(1, 1).value = up(rows[0].get('brand') or '')
        label = next((r.get('season_label') for r in rows if r.get('season_label')), '')
        ws.cell(1, 2).value = label.replace(' ', '-') if label else (f"{SEASON_LABEL.get(style[5],'Q1')}-20{style[3:5]}" if len(style) > 5 else '')
    for ci, h in enumerate(headers, 1):
        c = ws.cell(3, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
    ws.row_dimensions[3].height = 22
    today = datetime.now().date()
    for ri, r in enumerate(rows, 4):
        is_os = r.get('type') == 'OS'
        d = r.get('delivery')
        vals = {
            'CUSTOMER': buysheet_customer(r), 'STYLE': up(r['style'] if is_os else r['raw_style']),
            'SEGMENT': up(r.get('segment') or ''), 'COLOR CODE': up(r.get('color_code') or ''), 'COLOR': up(r.get('color') or ''),
            'DESCRIPTION': up(r.get('desc') or ''), 'CRD': d, 'DELIVERY': d, 'PRICE': num(r.get('price')), 'MSRP': num(r.get('msrp')),
            'SIZE RATIO': None if is_os else up(r.get('ratio') or ''), 'SIZES': None if is_os else up(r.get('sizes_str') or ''),
            'PC PER PAC': None if is_os else r.get('pc_per_pac'), 'PAC PER CTN': None if is_os else num(r.get('pac_per_ctn')),
            'PPK QTY': None if is_os else r.get('ppk_qty'),
            'PC PER CTN OS': num(r.get('pc_ctn_os') or r.get('pc_ctn')) if is_os else None,
            'DATE UPDATED': today, 'COMMENTS': up(r.get('notes') or ''),
        }
        if is_os:
            for sz, qty in (r.get('sizes') or {}).items():
                vals[sz] = qty
        for h, v in vals.items():
            c = ws.cell(ri, col[h], value=v if v != '' else None)
            if h in ('CRD', 'DELIVERY', 'DATE UPDATED'): c.number_format = 'mm/dd/yyyy'
        for ci in range(1, len(headers) + 1):
            c = ws.cell(ri, ci); c.font = DFONT; c.border = BDR
            c.alignment = LEFT if headers[ci - 1] in ('CUSTOMER', 'COLOR', 'DESCRIPTION', 'COMMENTS') else CTR
        ws.row_dimensions[ri].height = 18
    wide = {'CUSTOMER': 26, 'STYLE': 18, 'SEGMENT': 11, 'COLOR CODE': 11, 'COLOR': 22, 'DESCRIPTION': 36, 'CRD': 11, 'DELIVERY': 11,
            'PRICE': 8, 'MSRP': 8, 'SIZE RATIO': 14, 'SIZES': 16, 'PC PER PAC': 10, 'PAC PER CTN': 10, 'PPK QTY': 9,
            'PC PER CTN OS': 12, 'DATE UPDATED': 12, 'COMMENTS': 24}
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = wide.get(h, 7)

    # HEADER OUT: one line per customer exactly as it appears on DATA; everything comes from the prebuy header
    wh = wb.create_sheet("HEADER OUT")
    for ci, h in enumerate(HEADER_OUT, 1):
        c = wh.cell(1, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
        wh.column_dimensions[get_column_letter(ci)].width = 30 if ci == 1 else 16
    seen = {}
    for r in rows:
        cust = buysheet_customer(r)
        if cust not in seen:
            terms = r.get('terms')
            seen[cust] = [cust, up(r.get('tickets') or ''), up(r.get('hang_fold') or ''), None, up(r.get('warehouse') or ''), None, today, None,
                          num(terms) if num(terms) is not None else up(terms or ''), up(r.get('terms_desc') or '')]
    for ri, vals in enumerate(seen.values(), 2):
        for ci, v in enumerate(vals, 1):
            c = wh.cell(ri, ci, value=v if v != '' else None); c.font = DFONT; c.border = BDR
            c.alignment = LEFT if ci == 1 else CTR
            if ci in (6, 7): c.number_format = 'mm/dd/yyyy'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def ppk_components(r):
    """[(size, units per pack)] for a prepack row, from the size grid (legacy prebuy / buysheet) or from
    the SIZES + SIZE RATIO text (new template). The creation template needs this per size."""
    if r.get('sizes'):
        return list(r['sizes'].items())
    sizes, ratio = dash_list(r.get('sizes_str')), dash_list(r.get('ratio'))
    if sizes and len(sizes) == len(ratio):
        try: return [(s, int(q)) for s, q in zip(sizes, ratio)]
        except ValueError: pass
    return []

def make_creation_templates(validation):
    """GENERIC CREATION (one row per missing style) and PPK CREATION (one row per size of each missing prepack),
    in the layout the MM41 automation scripts read."""
    rows = validation.get('rows', [])
    missing_generics = sorted({r['style'] for r in validation.get('generic', []) if r['status'] == '✗ MISSING'})
    missing_ppks     = [r for r in validation.get('ppk', []) if r['exists'] == '✗ MISSING']
    os_rows_by_style = defaultdict(list)
    for r in rows:
        if r['type'] == 'OS': os_rows_by_style[r['style']].append(r)
    ppk_rows_by_art = defaultdict(list)
    for r in rows:
        if r['type'] == 'PPK': ppk_rows_by_art[r['raw_style']].append(r)

    def header(ws, labels):
        for ci, h in enumerate(labels, 1):
            c = ws.cell(1, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
    def write(ws, ri, vals):
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, value=v); c.font = DFONT; c.alignment = CTR; c.border = BDR
    season_formulas = lambda ri: (f'=CHOOSE(MID(A{ri},6,1),"SPR","SUM","FALL","HOL")', f'="20"&MID(A{ri},4,2)')

    wb = Workbook()
    ws_g = wb.active; ws_g.title = "GENERIC CREATION"
    header(ws_g, ['Material','COLOR CODES','DESCRIPTION','Pcs per Carton','SIZES','WT2','WT','LB','1','2','3','4','VOL','FT3','Season','Year'])
    ri = 2
    for style in missing_generics:
        style_rows = os_rows_by_style.get(style, [])
        if not style_rows: continue
        r0 = style_rows[0]
        codes = ';'.join(sorted({r['color_code'] for r in style_rows}))
        sizes = ';'.join(sorted({sz for r in style_rows for sz in r['sizes']}, key=lambda s: SIZE_ORDER.index(s) if s in SIZE_ORDER else 99))
        wt, vol = get_weights(style, r0.get('desc', ''))
        write(ws_g, ri, [style, codes, r0.get('desc', ''), r0.get('pc_ctn_os', ''), sizes, wt, wt, 'LB', None, None, None, None, vol, 'FT3', *season_formulas(ri)])
        ri += 1

    ws_p = wb.create_sheet("PPK CREATION")
    header(ws_p, ['Material','Description','Batch','Qty','Unit','WT2','WT','LB','1','2','3','4','VOL','FT3','Season','Year','PC PER PPK','PPK PER CTN','COMPONENT'])
    ri = 2
    done = set()
    for res in missing_ppks:
        art = res['ppk_article']
        if art in done: continue                     # one block per prepack, however many color rows it has
        done.add(art)
        art_rows = ppk_rows_by_art.get(art, [])
        r = art_rows[0] if art_rows else {}
        comps = [(f"{x['color_code']} {sz}", q) for x in art_rows for sz, q in ppk_components(x)]      # color code + size for every line
        pc_ppk = sum(q for _, q in comps) or None
        wt, vol = get_weights(r.get('style', art), r.get('desc', ''))
        ppk_wt  = round(wt * pc_ppk, 4) if wt and pc_ppk else None
        ppk_vol = round(vol * pc_ppk, 4) if vol and pc_ppk else None
        for j, (sz, qty) in enumerate(comps):
            first = j == 0
            s_f, y_f = season_formulas(ri)
            write(ws_p, ri, [art if first else '', r.get('desc', '') if first else '', '', qty, 'H4',
                             ppk_wt if first else '', ppk_wt if first else '', 'LB', None, None, None, None,
                             ppk_vol if first else '', 'FT3', s_f if first else '', y_f if first else '',
                             pc_ppk if first else '', 1 if first else '', sz])
            ri += 1

    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf


# ── 6b. the buysheet TEMPLATE (.xlsm): fill its DATA and HEADER tables, keep everything else byte for byte ──
BUYSHEET_TEMPLATE = os.environ.get('BUYSHEET_TEMPLATE') or os.path.join(TOOLS_DIR, 'templates', 'Buysheet_Template.xlsm')
XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12'

def _xml_text(v):
    return (str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))

def _excel_serial(d):
    if hasattr(d, 'date'): d = d.date()
    return (d - datetime(1899, 12, 30).date()).days

def centric_style(style):
    """'RDZ272ZS01C' -> 'RDZ272ZS01': the style without its trailing variant letter(s)."""
    return str(style or '')[:10]

def template_cell_value(col, r, i, rows_in_customer):
    """What goes in DATA column `col` for prebuy row r (None = leave blank). Everything text is upper case."""
    ppk = r['type'] == 'PPK'
    style = r.get('raw_style') or r.get('style') or ''
    today = datetime.now()
    size_qty = r.get('sizes') or {}
    simple = {
        'CUSTOMER': buysheet_customer(r), 'STYLE': style, 'CENTRIC STYLE': centric_style(r.get('style') or style),
        'SEGMENT': r.get('segment'), 'COLOR CODE': r.get('color_code'), 'COLOR': r.get('color'), 'DESCRIPTION': r.get('desc'),
        'CRD': r.get('delivery'), 'DELIVERY': r.get('delivery'), 'PRICE': num(r.get('price')), 'MSRP': num(r.get('msrp')),
        'SIZE RATIO': r.get('ratio') if ppk else None, 'SIZES': r.get('sizes_str') if ppk else None,
        'PC PER PAC': r.get('pc_per_pac') if ppk else None, 'PAC PER CTN': r.get('pac_per_ctn') if ppk else None,
        'PPK QTY': r.get('ppk_qty') if ppk else None, 'PC PER CTN OS': None if ppk else r.get('pc_ctn_os'),
        'DATE UPDATED': today, 'COMMENTS': r.get('notes'), 'ROW_LINE': i + 1,
    }
    if col in simple:
        v = simple[col]
        return up(v) if isinstance(v, str) else v
    if col in KNOWN_SIZES or col in size_qty:
        return size_qty.get(col) if not ppk else None
    return None

def make_buysheet_from_template(rows, customers, template_bytes):
    """A copy of the buysheet template with DATA and HEADER filled from the prebuy rows. Only the two sheet parts,
    their two table parts and the calc settings are touched: VBA, Power Query, pivots and slicers stay as they are."""
    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    parts = {n: zin.read(n) for n in zin.namelist()}
    wbx = parts['xl/workbook.xml'].decode('utf-8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rid_target = {m.group(2): m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)}
    rid_target.update({m.group(1): m.group(2) for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)})
    def sheet_part(name):
        m = re.search(rf'<sheet [^>]*name="{re.escape(name)}"[^>]*r:id="([^"]+)"', wbx) or re.search(rf'<sheet [^>]*r:id="([^"]+)"[^>]*name="{re.escape(name)}"', wbx)
        if not m: raise ValueError(f'the buysheet template has no sheet named {name}')
        return 'xl/' + rid_target[m.group(1)].lstrip('/').replace('xl/', '', 1) if not rid_target[m.group(1)].startswith('xl/') else rid_target[m.group(1)]
    def table_part(sheet):
        srel = parts.get(sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels', b'').decode('utf-8')
        m = re.search(r'Target="\.\./tables/([^"]+)"', srel)
        if not m: raise ValueError(f'{sheet} has no table')
        return 'xl/tables/' + m.group(1)

    # ── DATA ──
    dsheet = sheet_part('DATA'); dtable = table_part(dsheet)
    sx, tx = parts[dsheet].decode('utf-8'), parts[dtable].decode('utf-8')
    tref = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tx)
    first_col, head_row, last_col = tref.group(1), int(tref.group(2)), tref.group(3)
    cols = re.findall(r'<tableColumn [^>]*name="([^"]+)"', tx)
    cols = [c.replace('&amp;', '&') for c in cols]
    # size columns the prebuy has that the table does not: insert after the table's last size column
    want_sizes = []
    for r in rows:
        if r['type'] == 'OS':
            for sz in (r.get('size_order') or list(r.get('sizes') or {})):
                if sz in (r.get('sizes') or {}) and sz not in cols and sz not in want_sizes: want_sizes.append(sz)
    if want_sizes:
        last_size_i = max(i for i, c in enumerate(cols) if c in KNOWN_SIZES)
        cols = cols[:last_size_i + 1] + want_sizes + cols[last_size_i + 1:]
    ncols = len(cols)
    rows_xml = re.findall(r'<row r="(\d+)"[^>]*>.*?</row>|<row r="(\d+)"[^>]*/>', sx, re.S)
    def row_xml(n): 
        m = re.search(rf'<row r="{n}"[^>]*?(?:/>|>.*?</row>)', sx, re.S); return m.group(0) if m else ''
    head_xml, ex_xml = row_xml(head_row), row_xml(head_row + 1)
    style_of = lambda xml, letter: (re.search(rf'<c r="{letter}\d+"[^>]*?\ss="(\d+)"', xml) or [None, None])[1] if xml else None
    head_style = style_of(head_xml, first_col) or '1'
    orig_cols = [c.replace('&amp;', '&') for c in re.findall(r'<tableColumn [^>]*name="([^"]+)"', tx)]
    orig_letter = {c: get_column_letter(i + 1) for i, c in enumerate(orig_cols)}
    ex_style = {c: style_of(ex_xml, orig_letter[c]) for c in orig_cols}
    ex_formula = {}
    for c in orig_cols:
        m = re.search(rf'<c r="{orig_letter[c]}{head_row + 1}"[^>]*>(?:<f[^>]*>(.*?)</f>)', ex_xml, re.S)
        if m: ex_formula[c] = m.group(1)
    size_style = next((ex_style[c] for c in orig_cols if c in KNOWN_SIZES and ex_style.get(c)), None)

    def cell(letter, rown, value, style, formula=None):
        s_attr = f' s="{style}"' if style else ''
        if formula is not None:
            return f'<c r="{letter}{rown}"{s_attr}><f>{formula}</f></c>'
        if value is None or value == '':
            return f'<c r="{letter}{rown}"{s_attr}/>'
        if hasattr(value, 'year'):
            return f'<c r="{letter}{rown}"{s_attr}><v>{_excel_serial(value)}</v></c>'
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            v = int(value) if float(value).is_integer() else value
            return f'<c r="{letter}{rown}"{s_attr}><v>{v}</v></c>'
        return f'<c r="{letter}{rown}"{s_attr} t="inlineStr"><is><t>{_xml_text(value)}</t></is></c>'

    new_rows = ['<row r="%d" spans="1:%d">%s</row>' % (head_row, ncols, ''.join(cell(get_column_letter(i + 1), head_row, c, head_style) for i, c in enumerate(cols)))]
    by_customer = defaultdict(int)
    for i, r in enumerate(rows):
        rown = head_row + 1 + i
        cells = []
        for ci, c in enumerate(cols):
            letter = get_column_letter(ci + 1)
            if c in ex_formula:
                cells.append(cell(letter, rown, None, ex_style.get(c), ex_formula[c]))
            else:
                cells.append(cell(letter, rown, template_cell_value(c, r, i, by_customer), ex_style.get(c) if c in ex_style else size_style))
        new_rows.append('<row r="%d" spans="1:%d">%s</row>' % (rown, ncols, ''.join(cells)))
    last_row = head_row + max(len(rows), 1)
    last_letter = get_column_letter(ncols)
    # brand + season in A1 / B1
    label = next((r.get('season_label') for r in rows if r.get('season_label')), '')
    brand = next((r.get('brand') for r in rows if r.get('brand')), '')
    sx = re.sub(r'<c r="A1"[^>]*?(?:/>|>.*?</c>)', cell('A', 1, up(brand), style_of(row_xml(1), 'A')), sx, count=1, flags=re.S)
    sx = re.sub(r'<c r="B1"[^>]*?(?:/>|>.*?</c>)', cell('B', 1, label.replace('-', ' ') if label else '', style_of(row_xml(1), 'B')), sx, count=1, flags=re.S)
    # replace every row from the header row down
    keep = re.search(r'<sheetData>(.*?)</sheetData>', sx, re.S).group(1)
    kept = ''.join(m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', keep, re.S) if int(m.group(1)) < head_row)
    sx = re.sub(r'<sheetData>.*?</sheetData>', lambda m: '<sheetData>' + kept + ''.join(new_rows) + '</sheetData>', sx, count=1, flags=re.S)
    sx = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:{last_letter}{last_row}"/>', sx, count=1)
    sx = re.sub(r'<selection [^>]*/>', '<selection activeCell="A4" sqref="A4"/>', sx, count=1)
    new_ref = f'{first_col}{head_row}:{last_letter}{last_row}'
    tx = re.sub(r' ref="[A-Z]+\d+:[A-Z]+\d+"', f' ref="{new_ref}"', tx)
    if want_sizes:
        max_id = max(int(x) for x in re.findall(r'<tableColumn [^>]*id="(\d+)"', tx))
        col_xml = {m.group(1).replace('&amp;', '&'): m.group(0) for m in re.finditer(r'<tableColumn [^>]*name="([^"]+)"[^>]*?(?:/>|>.*?</tableColumn>)', tx, re.S)}
        parts_cols = []
        for c in cols:
            if c in col_xml: parts_cols.append(col_xml[c])
            else:
                max_id += 1; parts_cols.append(f'<tableColumn id="{max_id}" name="{_xml_text(c)}"/>')
        tx = re.sub(r'<tableColumns [^>]*>.*?</tableColumns>', lambda m: f'<tableColumns count="{ncols}">' + ''.join(parts_cols) + '</tableColumns>', tx, count=1, flags=re.S)
    parts[dsheet], parts[dtable] = sx.encode('utf-8'), tx.encode('utf-8')

    # ── HEADER ──
    hsheet = sheet_part('HEADER'); htable = table_part(hsheet)
    hx, htx = parts[hsheet].decode('utf-8'), parts[htable].decode('utf-8')
    href = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', htx)
    hhead = int(href.group(2))
    hcols = [c.replace('&amp;', '&') for c in re.findall(r'<tableColumn [^>]*name="([^"]+)"', htx)]
    def hrow(n):
        m = re.search(rf'<row r="{n}"[^>]*?(?:/>|>.*?</row>)', hx, re.S); return m.group(0) if m else ''
    hex_xml = hrow(hhead + 1)
    hstyle = {c: style_of(hex_xml, get_column_letter(i + 1)) for i, c in enumerate(hcols)}
    seen = {}
    for r in rows:
        cust = buysheet_customer(r)
        if cust not in seen:
            seen[cust] = {'CUSTOMER': cust, 'TICKETS': up(r.get('tickets') or ''), 'HANG_FOLD': up(r.get('hang_fold') or ''), 'PER BAG': None,
                          'WHSE': up(r.get('warehouse') or ''), 'REQ#': None, 'TERMS': r.get('terms'), 'TERMS DESC': up(r.get('terms_desc') or '')}
    hrows = []
    for i, vals in enumerate(seen.values()):
        rown = hhead + 1 + i
        hrows.append('<row r="%d" spans="1:%d">%s</row>' % (rown, len(hcols), ''.join(cell(get_column_letter(ci + 1), rown, vals.get(c), hstyle.get(c)) for ci, c in enumerate(hcols))))
    hlast = hhead + max(len(seen), 1)
    hkeep = ''.join(m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.search(r'<sheetData>(.*?)</sheetData>', hx, re.S).group(1), re.S) if int(m.group(1)) <= hhead)
    hx = re.sub(r'<sheetData>.*?</sheetData>', lambda m: '<sheetData>' + hkeep + ''.join(hrows) + '</sheetData>', hx, count=1, flags=re.S)
    hx = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:{get_column_letter(len(hcols))}{hlast}"/>', hx, count=1)
    htx = re.sub(r' ref="[A-Z]+\d+:[A-Z]+\d+"', f' ref="A{hhead}:{get_column_letter(len(hcols))}{hlast}"', htx)
    parts[hsheet], parts[htable] = hx.encode('utf-8'), htx.encode('utf-8')

    # ── recalc on open; drop the stale calc chain ──
    wbx = re.sub(r'<calcPr([^>]*)/>', lambda m: '<calcPr' + re.sub(r' fullCalcOnLoad="\d"', '', m.group(1)) + ' fullCalcOnLoad="1"/>', wbx, count=1)
    parts['xl/workbook.xml'] = wbx.encode('utf-8')
    if 'xl/calcChain.xml' in parts:
        del parts['xl/calcChain.xml']
        parts['xl/_rels/workbook.xml.rels'] = re.sub(r'<Relationship [^>]*calcChain\.xml"[^>]*/>', '', rels).encode('utf-8')
        parts['[Content_Types].xml'] = re.sub(r'<Override [^>]*calcChain\.xml"[^>]*/>', '', parts['[Content_Types].xml'].decode('utf-8')).encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in parts:
                zout.writestr(info.filename, parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
    out.seek(0)
    return out

def make_buysheet(rows, customers):
    """The buysheet download: the .xlsm template filled when it is installed, else the plain DATA + HEADER OUT workbook."""
    if os.path.exists(BUYSHEET_TEMPLATE):
        with open(BUYSHEET_TEMPLATE, 'rb') as f:
            return make_buysheet_from_template(rows, customers, f.read()), 'xlsm'
    return make_data_sheet(rows, customers), 'xlsx'

# ═══════════════════════════════════════════════════════════════════════════════
# 6c. STOCK - the SAP "Cut and Sold" report, exported, joined to every article by article number
# ═══════════════════════════════════════════════════════════════════════════════
STOCK_HEADERS = {   # key -> the report's column headings it can appear under (lower case, single spaces)
    'site': ('site', 'plant'), 'merch': ('merchcat', 'merch cat'), 'season': ('yr-seas', 'season'), 'article': ('article', 'material'),
    'segment': ('segment', 'stock segment'), 'sloc': ('s.loc', 'sloc', 'storage location', 'storage loc'), 'desc': ('description',),
    'color': ('color desc', 'color'), 'size': ('size',),
    'on_po': ('on po',), 'in_txt': ('in txt', 'in transit'), 'inv': ('inv qty',), 'dlv': ('dlv qty',), 'order': ('order qty',),
    'ata': ('ata inv',), 'ats': ('ats inv',), 'oversold': ('oversold inv',)}
STOCK_NUMBERS = ('on_po', 'in_txt', 'inv', 'dlv', 'order', 'ata', 'ats', 'oversold')
STOCK_CHECK = (('stock_check', 'In stock'),)

def _norm_header(v):
    return re.sub(r'\s+', ' ', str(v or '').strip().lower())

def sap_number(v):
    """SAP quantities: 12,428  /  1.234,5  /  5-  (a trailing minus is negative). Blank or text -> 0."""
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    t = str(v).strip().replace('\xa0', '').replace(' ', '')
    neg = t.endswith('-') or t.startswith('-') or (t.startswith('(') and t.endswith(')'))
    t = t.strip('-()')
    t = t.replace('.', '').replace(',', '.') if re.fullmatch(r'\d{1,3}(\.\d{3})+,\d+', t) else t.replace(',', '')
    try: n = float(t)
    except ValueError: return 0.0
    return -n if neg else n

def read_stock_table(file_bytes, filename=''):
    """Rows of cells from an Excel export or from a delimited text export (CSV, SAP 'text with tabs')."""
    if (filename or '').lower().endswith(('.xlsx', '.xlsm')) or file_bytes[:2] == b'PK':
        ws = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    if file_bytes[:2] in (b'\xff\xfe', b'\xfe\xff'): text = file_bytes.decode('utf-16')
    else:
        try: text = file_bytes.decode('utf-8-sig')
        except UnicodeDecodeError: text = file_bytes.decode('cp1252', errors='replace')
    lines = [l for l in text.splitlines() if l.strip() and not re.fullmatch(r'[\s\-+=|_]+', l)]      # drop SAP's ruler lines
    sample = '\n'.join(lines[:40])
    delim = max(('\t', ';', '|', ','), key=lambda d: sample.count(d))
    out = []
    for r in csv.reader(lines, delimiter=delim):
        r = [c.strip() for c in r]
        if delim == '|':
            r = r[1:] if r and r[0] == '' else r
            r = r[:-1] if r and r[-1] == '' else r
        out.append(r)
    return out

def parse_cut_sold(file_bytes, filename=''):
    """The Cut and Sold export -> {'articles': {ARTICLE: totals + the split by site / segment / storage location}, ...}.
    The report's own total and sub-total rows (no article) are skipped, or every unit would be counted twice."""
    table = read_stock_table(file_bytes, filename)
    head = next((i for i, row in enumerate(table[:40])
                 if 'article' in [_norm_header(c) for c in row] and any(h in [_norm_header(c) for c in row] for h in ('inv qty', 'ats inv', 'ats total'))), None)
    if head is None:
        raise ValueError("This does not look like the Cut and Sold report: no row has an 'Article' column together with 'Inv Qty' or 'ATS Inv'. "
                         "Export the report to Excel or CSV with its columns, then try again.")
    cells = [_norm_header(c) for c in table[head]]
    col = {}
    for key, names in STOCK_HEADERS.items():                      # the first heading variant that is present wins
        for n in names:
            if n in cells:
                col[key] = cells.index(n); break
    articles, lines, skipped = {}, 0, 0
    for row in table[head + 1:]:
        get = lambda k: row[col[k]] if k in col and col[k] < len(row) else None
        art = str(get('article') or '').strip().upper()
        if not art:
            skipped += 1; continue
        lines += 1
        a = articles.setdefault(art, dict({k: 0.0 for k in STOCK_NUMBERS}, by={}, desc='', color='', size='', season=''))
        a['desc'] = a['desc'] or str(get('desc') or '').strip(); a['color'] = a['color'] or str(get('color') or '').strip()
        a['size'] = a['size'] or str(get('size') or '').strip(); a['season'] = a['season'] or str(get('season') or '').strip()
        where = (str(get('site') or '').strip(), str(get('segment') or '').strip(), str(get('sloc') or '').strip())
        b = a['by'].setdefault(where, {k: 0.0 for k in STOCK_NUMBERS})
        for k in STOCK_NUMBERS:
            n = sap_number(get(k)); a[k] += n; b[k] += n
    whole = lambda x: int(x) if float(x).is_integer() else round(x, 2)
    for a in articles.values():
        for k in STOCK_NUMBERS: a[k] = whole(a[k])
        a['by'] = sorted(({'site': w[0], 'segment': w[1], 'sloc': w[2], **{k: whole(v) for k, v in n.items()}} for w, n in a['by'].items()), key=lambda x: -x['inv'])
    m = re.search(r'(\d{4})(\d{2})(\d{2})[_\-](\d{2})(\d{2})(\d{2})', filename or '')
    return {'source': 'export', 'filename': filename, 'as_of': f"{m.group(1)}-{m.group(2)}-{m.group(3)} {m.group(4)}:{m.group(5)}" if m else '',
            'lines': lines, 'skipped_total_rows': skipped, 'articles': articles, 'article_count': len(articles),
            'total_inv': whole(sum(a['inv'] for a in articles.values())), 'total_ats': whole(sum(a['ats'] for a in articles.values())),
            'sites': sorted({b['site'] for a in articles.values() for b in a['by'] if b['site']}),
            'segments': sorted({b['segment'] for a in articles.values() for b in a['by'] if b['segment']}),
            'columns': sorted(col), 'missing_columns': [k for k in ('inv', 'ats', 'on_po', 'order') if k not in col]}

STOCK_TYPES = (('inv', 'CLABS', 'on hand'), ('transfer', 'CUMLM', 'in transfer'), ('quality', 'CINSM', 'in quality inspection'),
               ('restricted', 'CEINM', 'restricted'), ('blocked', 'CSPEM', 'blocked'), ('returns', 'CRETM', 'returns'))

def query_hana_stock(conn, articles):
    """On-hand stock straight from STOCK_VIEW (default YOUR_COMPANY_SCHEMA.YOUR_STOCK_VIEW, the Cut and Sold report's own stock view:
    checked line for line against the report, identical). One line per article / site / segment / storage location.
    'On hand' is unrestricted stock (CLABS), as on the report; returns, blocked and the other stock types are kept
    beside it per storage location, because that stock is owned too."""
    schema, view = STOCK_VIEW.split('.', 1)
    ident = lambda x: '"' + str(x).replace('"', '""') + '"'
    arts = sorted({str(a).strip().upper() for a in articles if a})
    sums = ', '.join(f'SUM({ident(col)}) AS {ident(key)}' for key, col, _ in STOCK_TYPES)
    cur, rows = conn.cursor(), []
    for i in range(0, len(arts), 500):
        part = arts[i:i + 500]
        cur.execute(f'SELECT "MATNR", "WERKS", "SGT_SCAT", "LGORT", {sums} FROM {ident(schema)}.{ident(view)} '
                    f'WHERE "MATNR" IN ({", ".join("?" for _ in part)}) GROUP BY "MATNR", "WERKS", "SGT_SCAT", "LGORT"', tuple(part))
        rows += cur.fetchall()
    whole = lambda x: int(x) if float(x).is_integer() else round(float(x), 2)
    out, lines = {}, 0
    for r in rows:
        art = str(r[0] or '').strip().upper()
        nums = {key: float(r[4 + i] or 0) for i, (key, _, _) in enumerate(STOCK_TYPES)}
        if not art or not any(nums.values()):
            continue
        lines += 1
        a = out.setdefault(art, {'inv': 0.0, 'ats': None, 'on_po': None, 'in_txt': None, 'dlv': None, 'order': None, 'ata': None,
                                 'oversold': None, 'by': [], 'other': {}, 'other_total': 0.0, 'desc': '', 'color': '', 'size': '', 'season': ''})
        a['inv'] += nums['inv']
        other = {k: whole(v) for k, v in nums.items() if k != 'inv' and v}
        for k, v in other.items():
            a['other'][k] = whole(a['other'].get(k, 0) + v); a['other_total'] += v
        a['by'].append({'site': str(r[1] or '').strip(), 'segment': str(r[2] or '').strip(), 'sloc': str(r[3] or '').strip(),
                        'inv': whole(nums['inv']), 'ats': None, 'other': other})
    for a in out.values():
        a['inv'] = whole(a['inv']); a['other_total'] = whole(a['other_total'])
        a['by'].sort(key=lambda b: (-b['inv'], b['sloc'], b['segment']))
    return {'source': 'hana', 'view': STOCK_VIEW, 'filename': STOCK_VIEW, 'as_of': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'lines': lines, 'skipped_total_rows': 0, 'articles': out, 'article_count': len(out),
            'total_inv': whole(sum(a['inv'] for a in out.values())), 'total_ats': None,
            'total_other': whole(sum(a['other_total'] for a in out.values())),
            'sites': sorted({b['site'] for a in out.values() for b in a['by'] if b['site']}),
            'segments': sorted({b['segment'] for a in out.values() for b in a['by'] if b['segment']}),
            'labels': {k: label for k, _, label in STOCK_TYPES}, 'columns': [], 'missing_columns': []}

def live_stock(validation):
    """Live stock for every article in this validation (each size's article, each prepack and its components).
    A failure never stops the validation: it is reported on the stock strip and an export can still be added."""
    S = validation.setdefault('summary', {})
    S.pop('stock_error', None)
    if not STOCK_VIEW:
        return None
    arts = ({r.get('sap_article') for r in validation.get('generic', [])} | {r.get('ppk_article') for r in validation.get('ppk', [])}
            | {c.get('article') for r in validation.get('ppk', []) for c in r.get('components') or []})
    arts = {str(a).strip().upper() for a in arts if a}
    if not arts:
        return None
    try:
        conn = connect_hana()
        try:
            stock = query_hana_stock(conn, arts)
        finally:
            conn.close()
    except Exception as exc:
        S['stock_error'] = f"Live stock could not be read from {STOCK_VIEW}: {str(exc).splitlines()[0][:200]}. You can still add a Cut and Sold export."
        return None
    _cache['live_stock'] = stock
    return stock

def stock_info(stock):
    """The stock report's summary for the portal (never the per-article detail)."""
    return {k: v for k, v in stock.items() if k != 'articles'} if stock else None

def _stock_numbers(agg, source='export'):
    return None if not agg else dict({k: agg[k] for k in ('inv', 'ats', 'on_po', 'in_txt', 'order', 'dlv', 'oversold')}, by=agg['by'],
                                     other=agg.get('other', {}), other_total=agg.get('other_total', 0), source=source)

def apply_stock(validation, stock):
    """Adds what is already in stock to every result: an open-stock line by its SAP variant article, a prepack by its own
    article and by each component article (and how many packs the components could make). No stock -> clears it again."""
    arts = (stock or {}).get('articles', {})
    src = (stock or {}).get('source', 'export')
    unknown = None if src == 'hana' else 0               # live stock has no ATS / PO figures (yet): unknown, not zero
    for lst in (validation.get('generic', []), validation.get('generic_unique', [])):
        for r in lst:
            if not stock:
                r.pop('stock', None); r.pop('stock_check', None); continue
            agg = arts.get((r.get('sap_article') or '').upper())
            r['stock'] = _stock_numbers(agg, src)
            r['stock_check'] = f"ℹ IN STOCK {int(agg['inv']):,}" if agg and agg['inv'] > 0 else '✓ NONE'
    for lst in (validation.get('ppk', []), validation.get('ppk_unique', [])):
        for r in lst:
            if not stock:
                r.pop('stock', None); r.pop('stock_check', None); continue
            comps = r.get('components') or []
            for c in comps:
                a = arts.get((c.get('article') or '').upper())
                c['inv'], c['ats'], c['on_po'] = (a['inv'], a['ats'], a['on_po']) if a else (0, unknown, unknown)
            mine = [c for c in comps if not r.get('complex') or c.get('code') == r.get('color_code')]
            have = [c for c in mine if c['inv'] > 0]
            sized = [c for c in mine if (c.get('qty') or 0) > 0]
            own = arts.get((r.get('ppk_article') or '').upper())
            usable = lambda c: c['ats'] if c['ats'] is not None else c['inv']        # ATS when known, else on hand
            r['stock'] = {'own': _stock_numbers(own, src), 'components': len(mine), 'components_in_stock': len(have), 'source': src,
                          'packs': min(int(usable(c) // c['qty']) for c in sized) if sized else None, 'units': sum(c['inv'] for c in mine)}
            r['stock_check'] = (f"ℹ IN STOCK {int(own['inv']):,}" if own and own['inv'] > 0
                                else f"ℹ COMPONENTS {len(have)}/{len(mine)}" if have else '✓ NONE')
    S = validation.setdefault('summary', {})
    S['checks_gen'] = check_counts(validation.get('generic_unique', []), GEN_CHECKS + (STOCK_CHECK if stock else ()))
    S['checks_ppk'] = check_counts(validation.get('ppk_unique', []), PPK_CHECKS + (STOCK_CHECK if stock else ()))
    if stock: S['stock'] = stock_info(stock)
    else: S.pop('stock', None)

def make_validation_report(validation):
    """GENERIC VALIDATION, PPK VALIDATION, PPK COMPONENTS and SUMMARY - the layout _sap_creation_validator.py produces."""
    GREEN, RED, AMBER = PatternFill('solid', start_color='E2EFDA'), PatternFill('solid', start_color='FCE4E4'), PatternFill('solid', start_color='FFF2CC')
    GF, RF, AF = Font(name='Arial', size=10, color='375623', bold=True), Font(name='Arial', size=10, color='C00000', bold=True), Font(name='Arial', size=10, color='7F6000', bold=True)
    def write(ws, headers, lines, widths, left=(1, 2)):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(1, ci, h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
        for ri, line in enumerate(lines, 2):
            for ci, v in enumerate(line, 1):
                c = ws.cell(ri, ci, v); c.border = BDR; c.alignment = LEFT if ci in left else CTR; c.font = DFONT
                if isinstance(v, str) and v and v[0] in '✓✗⚠':
                    c.fill, c.font = (GREEN, GF) if v[0] == '✓' else (RED, RF) if v[0] == '✗' else (AMBER, AF)
        for ci, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(lines), 1) + 1}"
    wb = Workbook(); ws = wb.active; ws.title = 'SUMMARY'
    S = validation.get('summary', {})
    lines = [('CREATION VALIDATION SUMMARY', None), (validation.get('filename', ''), None), ('', None),
             ('GENERIC (OPEN STOCK)', None), (f"✓ Size rows already in SAP : {S.get('gen_exists', 0)}", 'g'), (f"✗ Size rows to create       : {S.get('gen_missing', 0)}", 'r')]
    count_text = lambda c: f"{c['label']:<12} ✓ {c['ok']}   ✗ {c['bad']}   ⚠ {c['warn']}" + (f"   ℹ {c['info']}" if c.get('info') else '')
    for c in S.get('checks_gen', []): lines.append((count_text(c), 'r' if c['bad'] else 'g'))
    lines += [('', None), ('PREPACKS', None), (f"✓ Already in SAP : {S.get('ppk_exists', 0)}", 'g'), (f"✗ To create       : {S.get('ppk_missing', 0)}", 'r')]
    for c in S.get('checks_ppk', []): lines.append((count_text(c), 'r' if c['bad'] else 'g'))
    for i, (t, kind) in enumerate(lines, 1):
        c = ws.cell(i, 1, t); c.font = Font(name='Consolas', size=14 if i == 1 else 10, bold=(i == 1 or kind is None), color={'g': '375623', 'r': 'C00000'}.get(kind, '000000'))
    ws.column_dimensions['A'].width = 60
    wg = wb.create_sheet('GENERIC VALIDATION')
    write(wg, ['STYLE','COLOR CODE','SIZE','STATUS','WHY','BUY PRICE','SAP PRICE','PRICE ✓','BUY COLOR','SAP COLOR','COLOR ✓','BUY DESC','SAP DESC','DESC ✓','PC PER CTN ✓','ROUNDING ✓','ROWS IN FILE','CUSTOMERS'],
          [[r['style'], r['color_code'], r['size'], r['status'], r.get('missing_reason', ''), r.get('buy_price'), r.get('sap_price'), r['price_check'], r.get('buy_color'), r.get('sap_color'),
            r['color_check'], r.get('buy_desc'), r.get('sap_desc'), r['desc_check'], r['pc_ctn_check'], r['round_check'], r.get('count', 1), ', '.join(r.get('customers', []))] for r in validation.get('generic_unique', validation.get('generic', []))],
          [16, 12, 8, 14, 50, 10, 10, 18, 22, 22, 24, 36, 36, 12, 24, 14, 10, 40])
    p = validation.get('ppk_unique', validation.get('ppk', []))
    wp = wb.create_sheet('PPK VALIDATION')
    write(wp, ['PPK ARTICLE','PREBUY ARTICLE','COLOR CODE','CC ✓','EXISTS','PC PER PPK ✓','BUY PRICE','SAP PRICE','PRICE ✓','BUY COLOR','SAP COLOR','COLOR ✓',
               'BUY SIZES','SAP SIZES','SIZES ✓','BUY RATIO','SAP RATIO','RATIO ✓','RATIO DIFFERENCES','BUY DESC','SAP DESC','ROUNDING ✓','ROWS IN FILE','CUSTOMERS'],
          [[r['ppk_article'], r.get('prebuy_article', ''), r['color_code'], r['cc_check'], r['exists'], r['pc_ppk_check'], r.get('buy_price'), r.get('sap_price'), r['price_check'],
            r.get('buy_color'), r.get('sap_color'), r['color_check'], r.get('buy_sizes'), r.get('sap_sizes'), r['sizes_check'], r.get('buy_ratio'), r.get('sap_ratio'), r['ratio_check'],
            r.get('ratio_note', ''), r.get('buy_desc'), r.get('sap_desc'), r['round_check'], r.get('count', 1), ', '.join(r.get('customers', []))] for r in p],
          [22, 20, 12, 10, 12, 30, 10, 10, 22, 22, 22, 14, 16, 16, 10, 14, 14, 12, 40, 36, 36, 30, 10, 40])
    wc = wb.create_sheet('PPK COMPONENTS')
    write(wc, ['PPK ARTICLE','COMPONENT ARTICLE','COLOR CODE','COLOR','SIZE','QTY PER PACK'],
          [[r['ppk_article'], c.get('article'), c.get('code'), c.get('color'), c.get('size'), c.get('qty')] for r in p if r['exists'] == '✓ EXISTS' for c in r.get('components', [])],
          [22, 24, 12, 24, 8, 12])
    if S.get('stock'):
        wst = wb.create_sheet('IN STOCK')
        by_text = lambda s_: '; '.join(f"{b['segment'] or '-'}/{b['sloc'] or '-'}: {b['inv']:,}" for b in (s_ or {}).get('by', []))
        lines_st = []
        for r in validation.get('generic_unique', validation.get('generic', [])):
            st = r.get('stock')
            if st and st['inv'] > 0:
                lines_st.append(['OPEN STOCK', r.get('sap_article'), r['style'], r['color_code'], r['size'], st['inv'], st['ats'], st['on_po'], st['in_txt'], st['order'], st['oversold'], by_text(st)])
        for r in p:
            st = r.get('stock') or {}
            if st.get('own') and st['own']['inv'] > 0:
                o = st['own']; lines_st.append(['PREPACK', r['ppk_article'], r['ppk_article'].split('-')[0], r['color_code'], '', o['inv'], o['ats'], o['on_po'], o['in_txt'], o['order'], o['oversold'], by_text(o)])
            for c in r.get('components', []):
                if c.get('inv', 0) > 0 and (not r.get('complex') or c.get('code') == r.get('color_code')):
                    lines_st.append(['PREPACK COMPONENT', c.get('article'), r['ppk_article'], c.get('code'), c.get('size'), c['inv'], c['ats'], c['on_po'], '', '', '', f"{c.get('qty')} per pack"])
        write(wst, ['TYPE','ARTICLE','STYLE / PREPACK','COLOR CODE','SIZE','ON HAND','ATS','ON PO','IN TRANSIT','ORDER QTY','OVERSOLD','WHERE'], lines_st, [18, 20, 22, 12, 8, 10, 10, 10, 10, 10, 10, 60])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

PR_HEADER = ['Group By', 'Article', 'Purch req quantity', 'Delivery Date(MMDDYYYY)',
             'Site', 'Storage Location', 'Purch Group', 'Stock Segment', 'SEASON']

def make_req_csv(rows, customers, validation=None):
    """The purchase-requisition upload: exactly the nine columns, in this order. Plain UTF-8 (no BOM), CRLF line endings.
    Open stock: one line per size variant. Prepacks: one line per buysheet row; only a complex prepack's color rows
    (one order, one article) collapse into a single line. Warehouse, delivery date and segment come from each buysheet row.
    Prepack quantity is in PACKS: the order's pieces divided by PC PER PPK (60 pc at 5 pc per ppk = 12).
    Anything SAP does not have yet (size variants or prepacks) is left out: it has to be created first.
    The prepack's article is the one the validation found in SAP, which can differ from the name the prebuy typed."""
    os_rows = [r for r in rows if r.get('type') == 'OS']
    ppk_rows = [r for r in rows if r.get('type') == 'PPK']
    ppk_res = (validation or {}).get('ppk') or []
    ppk_pairs = list(zip(ppk_rows, ppk_res)) if len(ppk_rows) == len(ppk_res) else []     # results are in row order

    conn = connect_hana()
    try:
        variant_map = query_variant_articles(conn, list({r['style'] for r in os_rows}))
    finally:
        conn.close()

    out = io.StringIO()
    w = csv.writer(out, lineterminator='\r\n')
    w.writerow(PR_HEADER)
    groups = {}                                               # customer -> Group By number, in order of first appearance
    group_no = lambda r: groups.setdefault(str(r.get('customer') or '').strip(), len(groups) + 1)
    for r in os_rows:
        style, code = r['style'], str(r.get('color_code') or '')
        group = group_no(r)
        for sz, qty in (r.get('sizes') or {}).items():
            article = variant_map.get((style.upper(), code.upper(), sz.upper()), '')
            if not article:
                continue
            w.writerow([group, up(article), int(qty), mmddyyyy(r.get('delivery')), 'H001', up(r.get('warehouse') or ''),
                        'H01', up(r.get('segment') or ''), up(season_text(r.get('season_label'), style))])

    seen_complex = set()                                      # a complex prepack's color rows are ONE order: written once
    for r, res in ppk_pairs:
        if res.get('exists') != '✓ EXISTS':
            continue
        article = res['ppk_article']
        if r.get('complex'):
            key = (article, str(r.get('customer') or ''), str(r.get('delivery') or ''))
            if key in seen_complex:
                continue
            seen_complex.add(key)
        pieces = sum((r.get('sizes') or {}).values()) or num(r.get('ppk_qty'))      # the row in pieces (an old-layout prebuy types its grid in pieces)
        per_pack = num(r.get('row_pc') if r.get('complex') else r.get('pc_per_pac')) # complex: this color's pieces per pack; else pieces per pack
        if not pieces or not per_pack:
            continue                                          # no quantity or no pieces-per-pack: the pack count cannot be worked out
        qty = pieces / per_pack                               # PAC qty = pieces / PC PER PPK  (PAC PER CTN does not enter into it)
        w.writerow([group_no(r), up(article), int(round(qty)), mmddyyyy(r.get('delivery')), 'H001', up(r.get('warehouse') or ''),
                    'H01', up(r.get('segment') or ''), up(season_text(r.get('season_label'), r['style']))])
    return out.getvalue()

# ═══════════════════════════════════════════════════════════════════════════════
# 7. DIAGNOSTICS
# ═══════════════════════════════════════════════════════════════════════════════
def build_diagnostics(cache, connect=None):
    """Everything needed to tell 'SAP has no price' from 'the lookup missed it' in one paste.
    Reads the last validation from the cache, then asks HANA directly about the articles that
    came back without a price, and lists which condition types SAP holds for the first of them."""
    out = {'build': BUILD, 'time': datetime.now().isoformat(timespec='seconds')}
    v = cache.get('validation')
    if not v:
        out['note'] = 'Nothing validated yet - upload a file on the BUYSHEET tab and click Validate first.'
        return out
    out.update({'file': v.get('filename'), 'file_type': v.get('file_type'),
                'summary': v.get('summary'), 'counts': v.get('diag')})

    d = v.get('diag') or {}
    tokens = list(d.get('os_styles_without_sap_price', [])[:3]) + list(d.get('ppk_without_sap_price', [])[:3])
    if not tokens:                       # everything priced - probe the first row so the page is still useful
        first = (v.get('generic') or v.get('ppk') or [{}])[0]
        t = first.get('style') or first.get('ppk_article')
        if t: tokens.append(t)
    out['probed'] = tokens

    try:
        conn = (connect or connect_hana)()
    except Exception as exc:
        out['hana_error'] = str(exc)
        return out
    try:
        cur = conn.cursor()
        probes = {}
        for tok in tokens:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', tok):
                continue
            try:
                cur.execute(
                    'SELECT "ARTICLE" AS "Article", "F_AMOUNT_PER_PC" AS "Price", "SEGMENT" AS "Segment", '
                    '"VALIDITY_START_DATE" AS "Valid From", "VALIDITY_END_DATE" AS "Valid To" '
                    'FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_PRICE_VIEW" '
                    'WHERE "ARTICLE" LIKE ? ORDER BY "ARTICLE" LIMIT 12', (tok + '%',))
                cols = [c[0] for c in cur.description]
                probes[tok] = [dict(zip(cols, [None if x is None else str(x).strip() for x in row]))
                               for row in cur.fetchall()]
            except Exception as exc:
                probes[tok] = f'ERROR: {exc}'
        out['zsls_rows_under_each_probed_article'] = probes

        missing = {}
        for r in v.get('generic') or []:
            if r.get('status') == '✗ MISSING':
                missing.setdefault(r['style'], []).append(f"{r['color_code']}/{r['size']}")
        held = {}
        for style, asked in list(missing.items())[:3]:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', style):
                continue
            try:
                cur.execute(
                    'SELECT "MARA_ARTICLE" AS "Article", "F_GENERIC" AS "Generic", "MARA_COLOR_CODE" AS "Color Code", '
                    '"MARA_SIZE_CODE" AS "Size", "F_ARTICLE_TYPE" AS "Article Type", "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name" '
                    'FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW" '
                    'WHERE "MARA_ARTICLE" LIKE ? OR "F_GENERIC" = ? ORDER BY "MARA_COLOR_CODE", "MARA_SIZE_CODE" LIMIT 150',
                    (style + '%', style))
                cols = [c[0] for c in cur.description]
                held[style] = {
                    'portal_asked_for_but_found_missing (color/size)': asked[:40],
                    'every_article_sap_has_under_this_style': [
                        dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()],
                }
            except Exception as exc:
                held[style] = f'ERROR: {exc}'
        if held:
            out['sap_articles_under_styles_with_missing_rows'] = held

        # prepacks the portal called missing: everything SAP holds under that article, any article type
        held_ppk = {}
        for art in list({r['ppk_article'] for r in v.get('ppk') or [] if r.get('exists') == '✗ MISSING'})[:3]:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', art):
                continue
            try:
                cur.execute(
                    'SELECT "MAST_ARTICLE" AS "Prepack", "MARA_ARTICLE" AS "Component", "F_ARTICLE_TYPE" AS "Article Type", '
                    '"MARA_COLOR_CODE" AS "Color Code", "MARA_SIZE_CODE" AS "Size", "STPO_PPK_QTY" AS "PPK Qty", "MARM_PC_PER_PAC" AS "PC Per PAC" '
                    'FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_ARTICLE_MASTER_VIEW" '
                    'WHERE "MAST_ARTICLE" LIKE ? OR "MARA_ARTICLE" LIKE ? ORDER BY "MAST_ARTICLE", "MARA_SIZE_CODE" LIMIT 60',
                    (art + '%', art + '%'))
                cols = [c[0] for c in cur.description]
                held_ppk[art] = {'portal_queried': f'MAST_ARTICLE = {art!r} AND F_ARTICLE_TYPE = \'C\'',
                                 'every_row_sap_has_starting_with_this_article': [
                                     dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]}
            except Exception as exc:
                held_ppk[art] = f'ERROR: {exc}'
        if held_ppk:
            out['sap_rows_under_prepacks_called_missing'] = held_ppk

        if tokens and re.fullmatch(r'[A-Za-z0-9_\-]+', tokens[0]):
            style = base_style(tokens[0])
            try:
                cur.execute(
                    'SELECT "CONDITION_TYPE" AS "Type", "KONP_CURRENCY" AS "Currency", "SEGMENT" AS "Segment", '
                    'COUNT(*) AS "Rows", MIN("ARTICLE") AS "Sample article" '
                    'FROM "_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_PRICING_CONDITIONS_VIEW" '
                    'WHERE "ARTICLE" LIKE ? GROUP BY "CONDITION_TYPE", "KONP_CURRENCY", "SEGMENT" '
                    'ORDER BY "CONDITION_TYPE"', (style + '%',))
                cols = [c[0] for c in cur.description]
                out['condition_types_under_' + style] = [
                    dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]
            except Exception as exc:
                out['condition_types_error'] = str(exc)
    finally:
        conn.close()
    return out

# ═══════════════════════════════════════════════════════════════════════════════
# 7b. SAP EXPLORER  (read-only: finds where the PO / margin data lives in HANA; nothing is written)
# ═══════════════════════════════════════════════════════════════════════════════
EXPLORE_TERMS = ['YOUR_TXN_1', 'YOUR_TXN_2', 'YOUR_TXN_3', 'YOUR_TXN_4', 'YOUR_TXN_5', 'YOUR_TXN_6']       # SAP transactions the PO / margin tab has to match
EXPLORE_FIELDS = {                                                               # field -> column-name fragments that might carry it
    'HTS code':          ['HTS', 'STEUC', 'TARIFF'],
    'Comm code':         ['COMM%CODE', 'COMMOD'],
    'Set code':          ['SET%CODE'],
    'Chain code':        ['CHAIN'],
    'Fiber content':     ['FIBER', 'FIBRE', 'COMPOS'],
    'Duty rate':         ['DUTY', 'ZOLL'],
    'Country of origin': ['HERKL', 'ORIGIN'],
    'PO / req link':     ['BANFN', 'EBELN', 'PURCH%REQ', 'PO%NUMBER'],
}

def _ident(x):
    return '"' + str(x).replace('"', '""') + '"'

def explore_sap(conn, terms=None, view=None, col=None, like=None):
    """Read-only questions put to HANA, each in its own try/except so one refusal never hides the rest:
    1. which views and tables are NAMED after each transaction
    2. the dictionary route: transaction -> program (TSTC) -> every table that program uses (D010TAB), Z tables first
    3. which view/table columns look like HTS, comm code, set code, chain code, fiber content, duty rate, origin, req/PO number
    4. with view=NAME: the first rows of that view and how many of its columns are actually filled in"""
    sample_only = bool(view) and not terms          # ?view=X on its own: just that object, without the long catalog sections
    terms = terms or EXPLORE_TERMS
    out = {'build': BUILD, 'time': datetime.now().isoformat(timespec='seconds')}
    if not sample_only:
        out['asked_about'] = terms
    cur = conn.cursor()

    def run(sql, params=()):
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]

    def safe(fn, *args):
        try:
            return fn(*args)
        except Exception as exc:
            return f'ERROR: {exc}'

    named = {}
    for t in ([] if sample_only else terms):
        like = '%' + t + '%'
        named[t] = {
            'views': safe(run, 'SELECT "SCHEMA_NAME", "VIEW_NAME" FROM "SYS"."VIEWS" WHERE "SCHEMA_NAME" <> \'SYS\' '
                               'AND UPPER("VIEW_NAME") LIKE ? ORDER BY "SCHEMA_NAME", "VIEW_NAME" LIMIT 40', (like,)),
            'tables': safe(run, 'SELECT "SCHEMA_NAME", "TABLE_NAME" FROM "SYS"."TABLES" WHERE "SCHEMA_NAME" <> \'SYS\' '
                                'AND UPPER("TABLE_NAME") LIKE ? ORDER BY "SCHEMA_NAME", "TABLE_NAME" LIMIT 40', (like,)),
        }
    if not sample_only:
        out['1_objects_named_like_each_transaction'] = named

    def dictionary_route():
        homes = run('SELECT DISTINCT "SCHEMA_NAME", "TABLE_NAME" FROM "SYS"."TABLES" WHERE "TABLE_NAME" IN (\'TSTC\', \'D010TAB\', \'DD02T\')')
        schema_of = {h['TABLE_NAME']: h['SCHEMA_NAME'] for h in homes}
        if not {'TSTC', 'D010TAB'} <= set(schema_of):
            return ('The SAP dictionary tables (TSTC / D010TAB) are not visible to this HANA user, so a transaction cannot be '
                    'traced to its tables from here. Use sections 1 and 3.')
        tab = lambda n: f'{_ident(schema_of[n])}.{_ident(n)}'
        res = {}
        for t in terms:
            tc = run(f'SELECT "TCODE", "PGMNA" AS "Program", "DYPNO" AS "Screen" FROM {tab("TSTC")} WHERE "TCODE" = ?', (t,))
            if not tc:
                res[t] = 'Not a transaction code in this SAP system (TSTC has no row for it).'
                continue
            join = (f'LEFT JOIN {tab("DD02T")} t ON t."TABNAME" = d."TABNAME" AND t."DDLANGUAGE" = \'E\'') if 'DD02T' in schema_of else ''
            text = 't."DDTEXT" AS "Description"' if join else '\'\' AS "Description"'
            used = run(f'SELECT DISTINCT d."TABNAME" AS "Table", {text} FROM {tab("D010TAB")} d {join} '
                       f'WHERE d."MASTER" = ? ORDER BY CASE WHEN d."TABNAME" LIKE \'Z%\' THEN 0 ELSE 1 END, d."TABNAME" LIMIT 250',
                       (tc[0]['Program'],))
            res[t] = {'transaction': tc[0], 'tables_the_program_uses_z_tables_first': used}
        return res
    if not sample_only:
        out['2_dictionary_route_transaction_to_tables'] = safe(dictionary_route)

    def field_search():
        found = {}
        for field, frags in EXPLORE_FIELDS.items():
            conds = ' OR '.join(['UPPER("COLUMN_NAME") LIKE ?'] * len(frags))
            params = tuple('%' + f + '%' for f in frags)
            in_views = safe(run, f'SELECT "SCHEMA_NAME", "VIEW_NAME" AS "OBJECT", "COLUMN_NAME" FROM "SYS"."VIEW_COLUMNS" '
                                 f'WHERE "SCHEMA_NAME" = \'_SYS_BIC\' AND ({conds}) ORDER BY "VIEW_NAME", "POSITION" LIMIT 40', params)
            in_tables = safe(run, f'SELECT "SCHEMA_NAME", "TABLE_NAME" AS "OBJECT", "COLUMN_NAME" FROM "SYS"."TABLE_COLUMNS" '
                                  f'WHERE "SCHEMA_NAME" NOT LIKE \'%SYS%\' AND ({conds}) ORDER BY "TABLE_NAME", "POSITION" LIMIT 40', params)
            found[field] = {'in_views': in_views, 'in_tables': in_tables}
        return found
    if not sample_only:
        out['3_columns_that_could_hold_each_field'] = safe(field_search)

    if view:
        def sample():
            # resolve the name through HANA's own catalog (view, then calculation view, then table): only a name the catalog lists reaches the SQL text
            hit = run('SELECT "SCHEMA_NAME", "VIEW_NAME" AS "OBJ", \'view\' AS "KIND" FROM "SYS"."VIEWS" WHERE "VIEW_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                hit = run('SELECT DISTINCT "SCHEMA_NAME", "VIEW_NAME" AS "OBJ", \'view\' AS "KIND" FROM "SYS"."VIEW_COLUMNS" WHERE "VIEW_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                hit = run('SELECT "SCHEMA_NAME", "TABLE_NAME" AS "OBJ", \'table\' AS "KIND" FROM "SYS"."TABLES" WHERE "TABLE_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                return f'No view or table called {view!r} in the catalog. Use a name from section 1 or 3.'
            schema, name, kind = hit[0]['SCHEMA_NAME'], hit[0]['OBJ'], hit[0]['KIND']
            src = f'{_ident(schema)}.{_ident(name)}'
            cols_sql = ('SELECT "COLUMN_NAME" FROM "SYS"."VIEW_COLUMNS" WHERE "SCHEMA_NAME" = ? AND "VIEW_NAME" = ? ORDER BY "POSITION"' if kind == 'view'
                        else 'SELECT "COLUMN_NAME" FROM "SYS"."TABLE_COLUMNS" WHERE "SCHEMA_NAME" = ? AND "TABLE_NAME" = ? ORDER BY "POSITION"')
            where, params, shown_filter = '', (), None
            if col:
                catalog_cols = [r['COLUMN_NAME'] for r in run(cols_sql, (schema, name))]
                match = next((c for c in catalog_cols if c.upper() == col.upper()), None)
                if not match:
                    return f'{col!r} is not a column of {name}. Its columns are: {catalog_cols}'
                where, params = f' WHERE UPPER({_ident(match)}) LIKE ?', (str(like or '').upper() + '%',)   # column name from the catalog; the value is a bound parameter
                shown_filter = f'{match} starts with {str(like or "").upper()!r}'
            cols = [r['COLUMN_NAME'] for r in run(cols_sql, (schema, name))][:300]          # every column the catalog lists, not just the first 80
            rows = run(f'SELECT * FROM {src}{where} LIMIT {5 if col else 3}', params)
            cols = cols or (list(rows[0].keys()) if rows else [])
            fills = ', '.join(f'SUM(CASE WHEN {_ident(c)} IS NULL OR TRIM(TO_NVARCHAR({_ident(c)})) = \'\' THEN 0 ELSE 1 END) AS "F{i}"'
                              for i, c in enumerate(cols))
            stats = safe(run, f'SELECT COUNT(*) AS "N", {fills} FROM (SELECT * FROM {src}{where} LIMIT 5000) AS "T"', params) if cols else []
            res = {kind: f'{schema}.{name}'}
            if shown_filter:
                res['filter'] = shown_filter
            if isinstance(stats, list) and stats:
                counts = {c: int(stats[0][f'F{i}']) for i, c in enumerate(cols)}
                res['rows_in_sample'] = int(stats[0]['N'])                                   # capped at 5000; a 0 is counted as filled, only blank / NULL are not
                res['filled_counts'] = {c: v for c, v in counts.items() if v}
                res['always_empty'] = [c for c, v in counts.items() if not v]
            else:
                res.update({'columns': cols, 'stats': stats})
            res['sample_rows'] = [{k: v for k, v in r.items() if v not in ('', None)} for r in rows]
            return res
        out['4_view_sample'] = safe(sample)
    return out

# ═══════════════════════════════════════════════════════════════════════════════
# 7c. PO / REQ CHECK  (read-only: pulls a PO or requisition straight from SAP, compares the two, and checks the margin file)
# ═══════════════════════════════════════════════════════════════════════════════
PO_VIEW    = '"_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_PURCHASE_ORDERS_VIEW"'
REQ_VIEW   = '"_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_REQUISITIONS_VIEW"'
FIBER_VIEW = '"_SYS_BIC"."YOUR_COMPANY_SCHEMA/YOUR_FIBER_CONTENTS_VIEW"'
SET_TABLE  = '"YOUR_COMPANY_SCHEMA"."YOUR_SET_COMPONENTS_TABLE"'

PO_COLS = ['EKKO_PO_NUMBER', 'EKPO_PO_ITEM', 'EKKO_PO_TYPE_ID', 'T161T_PO_TYPE_TEXT', 'EKKO_VENDOR_ID', 'EKKO_PURCH_ORG_ID', 'EKKO_RELEASE_INDICATOR',
           'EKKO_PO_DATE', 'EKPA_FACTORY_ID', 'EKPA_TRADING_CO_ID', 'EKPO_ARTICLE', 'EKPO_ARTICLE_SHORT_TEXT', 'EKPO_STORAGE_LOCATION_ID',
           'EKPO_PREQ_NUMBER', 'EKPO_PREQ_ITEM_NO', 'EKPO_DELETION_INDICATOR', 'EBAN_DELIVERY_DATE', 'EKPO_STAT_EXFTY_DATE', 'EKPO_STAT_DLVY_DATE',
           'EKET_SCHEDULE_LINE', 'EKET_DELIVERY_DATE', 'DLHD_HANDOVER_DATE', 'DLIT_01_DATE', 'DLIT_02_DATE', 'DLIT_03_DATE', 'DLIT_05_DATE',
           'F_EKET_PO_QTY_PC', 'F_GENERIC', 'MARM_PC_PER_PAC', 'T006A_PO_COMMERCIAL_UNIT', 'MAW1_COMMODITY_CODE', 'EIPO_COMMODITY_CODE',
           'EIPO_DISPATCH_COUNTRY', 'EIPO_TRADING_COUNTRY', 'EKPO_PO_QTY_PUOM', 'F_PO_QTY_PC', 'F_REQ_QTY_PC', 'F_NET_PRICE_PER_PC', 'F_PB00_PER_PC',
           'F_ZFS1_PER_PC', 'F_ZCST_RATE', 'F_ZCSR_RATE', 'F_ZTRF_RATE']
REQ_COLS = ['EBAN_REQUISITION_NO', 'EBAN_REQUISITION_ITEM', 'EBAN_ARTICLE', 'EBAN_DELETION_INDICATOR', 'EBAN_DELIVERY_DATE', 'EBAN_STORAGE_LOCATION',
            'EBAN_PURCHASE_ORDER_NO', 'EBAN_PO_ITEM_NO', 'F_REQ_QTY_PC', 'EBAN_RELEASE_INDICATOR', 'MAW1_COMMODITY_CODE', 'F_GENERIC', 'EBAN_DOC_TYPE']

def pad10(v):
    """A PO or requisition number as SAP stores it: digits only, zero-padded to 10."""
    d = re.sub(r'\D', '', str(v or ''))
    return d.zfill(10) if d else ''

def _s(v):
    return '' if v is None else str(v).strip()

def blank_date(v):
    s = _s(v)
    return '' if s in ('', '0', '00000000') else s

def _chunks(items, n=500):
    items = list(items)
    for i in range(0, len(items), n):
        yield items[i:i + n]

def _select(conn, view, cols, where_col, values):
    """SELECT the named columns where one column is IN a list of bound values. Names are fixed constants above; values are parameters."""
    out = []
    cur = conn.cursor()
    for part in _chunks(sorted(set(values))):
        marks = ', '.join('?' for _ in part)
        cur.execute(f'SELECT {", ".join(chr(34) + c + chr(34) for c in cols)} FROM {view} WHERE "{where_col}" IN ({marks})', tuple(part))
        names = [d[0] for d in cur.description]
        out += [dict(zip(names, r)) for r in cur.fetchall()]
    return out

def fetch_lookup_data(conn, number, kind='auto'):
    """Everything SAP holds for one PO or requisition number. Returns (kind, po_rows, req_rows, fiber_rows, set_rows, notes)."""
    n = pad10(number)
    if not n: raise ValueError('Enter a PO or requisition number.')
    notes = []
    po_rows, req_rows = [], []
    if kind in ('auto', 'po'):
        po_rows = _select(conn, PO_VIEW, PO_COLS, 'EKKO_PO_NUMBER', [n])
        if po_rows: kind = 'po'
    if not po_rows and kind in ('auto', 'req'):
        kind = 'req'
        req_rows = _select(conn, REQ_VIEW, REQ_COLS, 'EBAN_REQUISITION_NO', [n])
        if not req_rows: raise LookupError(f'{n} was not found as a PO or as a requisition.')
    if kind == 'po' and not po_rows: raise LookupError(f'PO {n} was not found.')
    if kind == 'po':
        reqs = {pad10(r['EKPO_PREQ_NUMBER']) for r in po_rows if pad10(r['EKPO_PREQ_NUMBER'])}
        if reqs: req_rows = _select(conn, REQ_VIEW, REQ_COLS, 'EBAN_REQUISITION_NO', reqs)
        else: notes.append('This PO has no requisition on any line.')
        if reqs:                                                  # every PO that draws on those requisitions, to see the split across POs
            known = {(r['EKKO_PO_NUMBER'], r['EKPO_PO_ITEM'], r['EKET_SCHEDULE_LINE']) for r in po_rows}
            po_rows += [r for r in _select(conn, PO_VIEW, PO_COLS, 'EKPO_PREQ_NUMBER', reqs) if (r['EKKO_PO_NUMBER'], r['EKPO_PO_ITEM'], r['EKET_SCHEDULE_LINE']) not in known]
    else:
        po_rows = _select(conn, PO_VIEW, PO_COLS, 'EKPO_PREQ_NUMBER', [n])
    generics = sorted({_s(r.get('F_GENERIC')) for r in po_rows + req_rows if _s(r.get('F_GENERIC'))})
    articles = sorted({_s(r.get('EKPO_ARTICLE')) for r in po_rows if _s(r.get('EKPO_ARTICLE'))} | {_s(r.get('EBAN_ARTICLE')) for r in req_rows if _s(r.get('EBAN_ARTICLE'))})
    fiber_rows, set_rows = [], []
    try:
        fiber_rows = _select(conn, FIBER_VIEW, ['ARTICLE', 'VENDOR', 'PURCHORG', 'FIBERCONT', 'FROMDELIVERY', 'TODELIVERY'], 'ARTICLE', generics) if generics else []
    except Exception as e:
        notes.append(f'Fiber lookup unavailable: {e}')
    try:
        set_rows = _select(conn, SET_TABLE, ['MATNR', 'EKORG', 'COMP_COMM_CODE', 'COMP_COMM_DESC', 'FOB_PRICE', 'Z1ST_SALE_PRICE', 'QUANTITY', 'VALID_FROM', 'VALID_TO'], 'MATNR', articles) if articles else []
    except Exception as e:
        notes.append(f'Set code lookup unavailable: {e}')
    return kind, po_rows, req_rows, fiber_rows, set_rows, notes

# ── margin file ─────────────────────────────────────────────────────────────────
MARGIN_FIELDS = {'CUSTOMER': 'customer', 'DELIVERY': 'delivery', 'REQ#': 'req', 'PO#': 'po', 'SELL PRICE': 'sell', 'QTY': 'qty', 'GENERIC': 'generic',
                 'COLOR DESC': 'color', 'SPLIT BY COLOR': 'split_color', 'VENDOR': 'vendor', 'FACTORY': 'factory', 'FOB': 'fob', 'TERM': 'term',
                 'DUTY RATE': 'duty_rate', 'SIZE RANGE': 'size_range', 'FIRST SALE VARIABLE': 'first_sale_var', 'COO': 'coo',
                 'MAIN FABRICATION': 'fabrication', 'ADDITIONAL TARIFF': 'add_tariff', 'FINAL DUTY': 'final_duty', 'FINAL PO FOB': 'final_fob',
                 'FIRST SALE': 'first_sale'}

def generic_key(v):
    """The style code in a margin GENERIC cell ('RDM271SS06F -TABLE PROGRAM' -> 'RDM271SS06F')."""
    t = _s(v).upper().split()
    return t[0] if t else ''

def parse_margin(file_bytes):
    """Rows of the 'Margin data recap' sheet as dicts (values only; the file is read with its saved formula results)."""
    if file_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError('This file is password-protected (or an old .xls). In Excel: File > Info > Protect Workbook > Encrypt with Password, clear the password, and save as .xlsx. Then upload it again.')
    if file_bytes[:2] != b'PK':
        raise ValueError('This is not an Excel .xlsx/.xlsm file. Open it in Excel and save it as .xlsx.')
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    name = next((s for s in wb.sheetnames if 'MARGIN' in s.upper() and 'RECAP' in s.upper()), None) or wb.sheetnames[0]
    rows = list(wb[name].iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows[:15]) if r and 'GENERIC' in [_s(c).upper() for c in r] and 'FOB' in [_s(c).upper() for c in r]), None)
    if hi is None: raise ValueError(f"No header row with GENERIC and FOB found on sheet '{name}'.")
    cols = {}
    for i, h in enumerate(rows[hi]):
        key = MARGIN_FIELDS.get(_s(h).upper())
        if key and key not in cols: cols[key] = i
    for need in ('generic', 'fob'):
        if need not in cols: raise ValueError(f'Margin sheet has no {need.upper()} column.')
    out = []
    for ri, r in enumerate(rows[hi + 1:], hi + 2):
        g = generic_key(r[cols['generic']] if cols['generic'] < len(r) else '')
        if not g: continue
        d = {k: (r[i] if i < len(r) else None) for k, i in cols.items()}
        d.update(row=ri, key=g, req=pad10(d.get('req')), po=pad10(d.get('po')))
        for k in ('qty', 'fob', 'duty_rate', 'add_tariff', 'final_duty', 'final_fob', 'first_sale', 'sell'):
            d[k] = num(d.get(k))
        d['delivery'] = iso_date(d.get('delivery')) if d.get('delivery') else ''
        d['term'] = _s(d.get('term')); d['vendor'] = _s(d.get('vendor')).upper(); d['factory'] = _s(d.get('factory')).upper()
        d['color'] = _s(d.get('color') or d.get('split_color')); d['coo'] = _s(d.get('coo')).upper(); d['fabrication'] = _s(d.get('fabrication'))
        d['size_range'] = _s(d.get('size_range')); d['customer'] = _s(d.get('customer')); d.pop('split_color', None)
        out.append(d)
    return {'sheet': name, 'rows': out}

# ── fiber text ──────────────────────────────────────────────────────────────────
_FIBER_ALIAS = {'POLY': 'POLYESTER', 'POLYESTER': 'POLYESTER', 'ELASTANE': 'SPANDEX', 'LYCRA': 'SPANDEX', 'SPANDEX': 'SPANDEX', 'ELASTANE.': 'SPANDEX',
                'POLYAMIDE': 'NYLON', 'NYLON': 'NYLON', 'COTTON': 'COTTON', 'CTN': 'COTTON', 'VISCOSE': 'RAYON', 'RAYON': 'RAYON'}

def fiber_pairs(text):
    """'80% POLYESTER / 20% SPANDEX' -> {'POLYESTER': 80.0, 'SPANDEX': 20.0}. Percent then fiber name; anything after the name is ignored."""
    pairs = {}
    for pct, word in re.findall(r'(\d+(?:\.\d+)?)\s*%\s*(?:RECYCLED\s+|ORGANIC\s+)?([A-Z]+)', _s(text).upper()):
        name = _FIBER_ALIAS.get(word, word)
        pairs[name] = pairs.get(name, 0) + float(pct)
    return pairs

def fiber_compare(sap_text, margin_text):
    """('ok'|'diff'|'manual', note). More than one set of percentages (size-specific fibers) goes to a person."""
    a, b = fiber_pairs(sap_text), fiber_pairs(margin_text)
    if not a: return 'diff', 'SAP fiber text has no percentages: ' + _s(sap_text)
    if not b: return 'diff', 'Margin file fabrication has no percentages: ' + _s(margin_text)
    if abs(sum(a.values()) - 100) > 0.5 or abs(sum(b.values()) - 100) > 0.5:
        return 'manual', f'Fiber percentages do not add to 100 (SAP: {_s(sap_text)} | margin: {_s(margin_text)}), check by eye'
    return ('ok', '') if a == b else ('diff', f'SAP {_s(sap_text)} vs margin {_s(margin_text)}')

# ── the comparison ──────────────────────────────────────────────────────────────
TOL_PRICE = 0.0101      # one cent
TOL_RATE  = 0.051       # rates are compared as percent: 0.32 in the margin file = 32 in SAP

def build_lookup(kind, number, po_rows, req_rows, fiber_rows, set_rows, margin=None, notes=None):
    """Join PO lines to requisition lines and run every check. Pure: no database, no files. Returns the JSON the portal shows."""
    n = pad10(number)
    notes = list(notes or [])
    # PO item rows carry quantity and price; schedule-line rows carry delivery dates. Merge them per PO item.
    items = {}
    for r in po_rows:
        key = (_s(r['EKKO_PO_NUMBER']), _s(r['EKPO_PO_ITEM']))
        it = items.setdefault(key, {'item_row': None, 'sched': []})
        if _s(r.get('EKET_SCHEDULE_LINE')):
            it['sched'].append(r)
        else:
            it['item_row'] = r
    req_by_key = {(pad10(r['EBAN_REQUISITION_NO']), _s(r['EBAN_REQUISITION_ITEM'])): r for r in req_rows}
    # pieces drawn on each requisition item, across every PO (the one asked about and any others)
    drawn = defaultdict(float)
    for (po, item), it in items.items():
        base = it['item_row'] or (it['sched'][0] if it['sched'] else None)
        if base is None: continue
        q = num(it['item_row'].get('F_PO_QTY_PC')) if it['item_row'] else None
        if q is None: q = sum(num(s.get('F_EKET_PO_QTY_PC')) or 0 for s in it['sched'])
        it['qty_pc'] = q
        if pad10(base.get('EKPO_PREQ_NUMBER')) and _s(base.get('EKPO_DELETION_INDICATOR')) != 'L':
            drawn[(pad10(base['EKPO_PREQ_NUMBER']), _s(base['EKPO_PREQ_ITEM_NO']))] += q or 0

    fiber_by = defaultdict(list)
    for f in fiber_rows: fiber_by[_s(f['ARTICLE']).upper()].append(f)
    set_by = defaultdict(list)
    for s in set_rows: set_by[_s(s['MATNR']).upper()].append(s)

    mrows = (margin or {}).get('rows') or []
    my_reqs = {pad10(r['EBAN_REQUISITION_NO']) for r in req_rows} | ({n} if kind == 'req' else set())
    tagged = [m for m in mrows if (kind == 'po' and m['po'] == n) or (m['req'] and m['req'] in my_reqs)]
    pool = tagged if tagged else mrows
    margin_note = ''
    if not mrows: margin_note = 'No margin file loaded.'
    elif tagged: margin_note = f'{len(tagged)} margin rows carry this REQ#/PO#; only those are used.'
    else: margin_note = 'Margin rows are not tagged with this REQ#/PO#, so they are matched by style only and quantity is not compared.'

    out_items = []
    focus_po = n if kind == 'po' else None
    for (po, item), it in sorted(items.items()):
        base = it['item_row'] or (it['sched'][0] if it['sched'] else None)
        if base is None: continue
        if focus_po and po != focus_po: continue                        # other POs only feed the "other POs" quantity below
        out_items.append(_check_item(po, item, it, base, req_by_key, drawn, fiber_by, set_by, pool, bool(tagged)))
    # requisition lines nobody ordered (requisition lookups, or items of the requisition not on this PO)
    used = {(i['req'], i['req_item']) for i in out_items if i['req']}
    for (rq, ri), r in sorted(req_by_key.items()):
        if (rq, ri) in used: continue
        if kind == 'po' and drawn.get((rq, ri)): continue               # ordered on another PO: shown there
        if kind == 'req' and drawn.get((rq, ri)): continue
        out_items.append(_req_only_item(rq, ri, r))
    margin_only = []
    if tagged:                                                          # quantity per style: margin file vs everything the PO carries for that style
        po_sum, first_at = defaultdict(float), {}
        for idx, i in enumerate(out_items):
            po_sum[i['generic']] += i['po_qty_pc'] or 0
            first_at.setdefault(i['generic'], idx)
        for g, idx in first_at.items():
            mrs = [m for m in tagged if _style_match(m['key'], g)]
            if mrs:
                mq = sum(m['qty'] or 0 for m in mrs)
                out_items[idx]['margin']['qty'] = mq
                if abs(mq - po_sum[g]) > 0.5: _flag(out_items[idx]['flags'], 'warn', f'Style quantity: margin file {mq:g} pc vs {po_sum[g]:g} pc on the PO (all colors and sizes)')
        margin_only = sorted({m['key'] for m in tagged if not any(_style_match(m['key'], g) for g in first_at)})
    summary = {'items': len(out_items), 'error': sum(1 for i in out_items for f in i['flags'] if f['level'] == 'error'),
               'warn': sum(1 for i in out_items for f in i['flags'] if f['level'] == 'warn'),
               'clean': sum(1 for i in out_items if not any(f['level'] in ('error', 'warn') for f in i['flags']))}
    return {'build': BUILD, 'kind': kind, 'number': n, 'items': out_items, 'summary': summary, 'notes': notes, 'margin_note': margin_note,
            'margin_only': margin_only, 'pos': sorted({i['po'] for i in out_items if i['po']}), 'reqs': sorted({i['req'] for i in out_items if i['req']})}

def _flag(flags, level, msg):
    flags.append({'level': level, 'msg': msg})

def _style_match(margin_key, sap_generic):
    """Same style code; a margin code one or two characters longer (RDM271SS06F vs RDM271SS06) still counts."""
    a, b = margin_key.upper(), sap_generic.upper()
    return bool(a and b) and (a == b or (a.startswith(b) and len(a) - len(b) <= 2) or (b.startswith(a) and len(b) - len(a) <= 2))

def _same(a, b, tol=TOL_PRICE):
    return a is not None and b is not None and abs(a - b) <= tol

def _check_item(po, item, it, base, req_by_key, drawn, fiber_by, set_by, pool, tagged):
    flags = []
    ir = it['item_row'] or base
    article = _s(base.get('EKPO_ARTICLE')); generic = _s(base.get('F_GENERIC')).upper() or article.split('-')[0].upper()
    rq, ri = pad10(base.get('EKPO_PREQ_NUMBER')), _s(base.get('EKPO_PREQ_ITEM_NO'))
    req = req_by_key.get((rq, ri)) if rq else None
    qty = it.get('qty_pc')
    fob = num(ir.get('F_NET_PRICE_PER_PC'))
    if fob is None: fob = num(base.get('F_PB00_PER_PC'))
    zfs1, zcsr, zcst, ztrf = (num(base.get(k)) for k in ('F_ZFS1_PER_PC', 'F_ZCSR_RATE', 'F_ZCST_RATE', 'F_ZTRF_RATE'))
    comm_m, comm_p = _s(base.get('MAW1_COMMODITY_CODE')), _s(base.get('EIPO_COMMODITY_CODE'))
    sched = [{'line': _s(s.get('EKET_SCHEDULE_LINE')), 'delivery': blank_date(s.get('EKET_DELIVERY_DATE')), 'ex_factory': blank_date(s.get('DLHD_HANDOVER_DATE') or s.get('DLIT_01_DATE')),
              'etd': blank_date(s.get('DLIT_02_DATE')), 'eta': blank_date(s.get('DLIT_03_DATE')), 'qty_pc': num(s.get('F_EKET_PO_QTY_PC'))} for s in it['sched']]
    ex_fact = next((s['ex_factory'] for s in sched if s['ex_factory']), '') or blank_date(base.get('EKPO_STAT_EXFTY_DATE'))
    po_deliv = next((s['delivery'] for s in sched if s['delivery']), '') or blank_date(base.get('EKPO_STAT_DLVY_DATE'))
    is_fs = _s(base.get('EKKO_PO_TYPE_ID')) == 'Z4'

    # ── SAP: nothing blank, HTS agrees ──
    for label, val in (('Article master commodity code (HTS)', comm_m), ('PO commodity code (HTS)', comm_p), ('Vendor', _s(base.get('EKKO_VENDOR_ID'))),
                       ('Factory', _s(base.get('EKPA_FACTORY_ID'))), ('Storage location', _s(base.get('EKPO_STORAGE_LOCATION_ID')))):
        if not val: _flag(flags, 'error', f'{label} is blank in SAP')
    if comm_m and comm_p and comm_m != comm_p: _flag(flags, 'error', f'HTS differs: article master {comm_m} vs PO {comm_p}')
    if not fob: _flag(flags, 'error', 'PO price is blank or zero')
    if is_fs and not zfs1: _flag(flags, 'error', 'First sale PO (Z4) but first sale price (ZFS1) is blank or zero')
    if zcsr is None or zcst is None: _flag(flags, 'error', 'Duty rate is blank in SAP (ZCSR base duty or ZCST duty)')
    if ztrf is None: _flag(flags, 'warn', 'US import tariff rate (ZTRF) is blank in SAP')
    if _s(base.get('EKKO_RELEASE_INDICATOR')) not in ('Y', ''): _flag(flags, 'warn', f"PO release indicator is '{_s(base.get('EKKO_RELEASE_INDICATOR'))}' (not released)")
    if _s(base.get('EKPO_DELETION_INDICATOR')): _flag(flags, 'error', f"PO line deletion indicator is '{_s(base.get('EKPO_DELETION_INDICATOR'))}'")
    if not sched: _flag(flags, 'warn', 'PO line has no delivery schedule line')

    # ── PO vs requisition ──
    req_qty = num(req.get('F_REQ_QTY_PC')) if req else num(ir.get('F_REQ_QTY_PC'))
    other = None
    if not rq:
        _flag(flags, 'warn', 'PO line is not tied to a requisition')
    elif not req:
        _flag(flags, 'warn', f'Requisition {rq}/{ri} was not found in the requisition view')
    else:
        if _s(req.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'error', f"Requisition line {rq}/{ri} is flagged deleted ('{_s(req.get('EBAN_DELETION_INDICATOR'))}') but a PO draws on it")
        if _s(req.get('EBAN_ARTICLE')) and article and _s(req.get('EBAN_ARTICLE')) != article: _flag(flags, 'error', f"Article differs: requisition {_s(req.get('EBAN_ARTICLE'))} vs PO {article}")
        if _s(req.get('EBAN_STORAGE_LOCATION')) != _s(base.get('EKPO_STORAGE_LOCATION_ID')): _flag(flags, 'warn', f"Storage location differs: requisition {_s(req.get('EBAN_STORAGE_LOCATION'))} vs PO {_s(base.get('EKPO_STORAGE_LOCATION_ID'))}")
        total = drawn.get((rq, ri), 0)
        other = total - (qty or 0)
        if req_qty is not None and qty is not None:
            if abs(total - req_qty) > 0.5:
                _flag(flags, 'error' if total > req_qty else 'warn', f'Quantity: requisition {req_qty:g} pc vs {total:g} pc on POs' + (f' ({qty:g} on this PO, {other:g} on others)' if other else ''))
    # ── fiber ──
    fib = [f for f in fiber_by.get(generic, [])]
    d = po_deliv or blank_date(base.get('EBAN_DELIVERY_DATE'))
    vend = _s(base.get('EKKO_VENDOR_ID')).upper()
    vends = {vend, _s(base.get('EKPA_FACTORY_ID')).upper(), _s(base.get('EKPA_TRADING_CO_ID')).upper()} - {''}
    org = _s(base.get('EKKO_PURCH_ORG_ID')).upper()
    in_org = [f for f in fib if not org or not _s(f['PURCHORG']) or _s(f['PURCHORG']).upper() == org]      # fiber rows are per purchasing org
    def window(f): return not d or (_s(f['FROMDELIVERY']) <= d <= _s(f['TODELIVERY']))
    live = [f for f in in_org if _s(f['VENDOR']).upper() in vends and window(f)]            # this vendor's own row wins
    if not live: live = [f for f in in_org if not _s(f['VENDOR']) and window(f)]            # a row with no vendor applies to every vendor
    sap_fiber = next((_s(f['FIBERCONT']) for f in live), '')
    if not sap_fiber:
        mine = [f for f in in_org if _s(f['VENDOR']).upper() in vends or not _s(f['VENDOR'])]
        if mine: _flag(flags, 'error', f"Fiber content for {generic} / {vend} does not cover delivery {d or '(no date)'}: valid " + '; '.join(f"{_s(f['FROMDELIVERY'])}-{_s(f['TODELIVERY'])}" for f in mine))
        else: _flag(flags, 'error', f'No fiber content in SAP for {generic} / {vend}' + (f' (other vendors have it: {sorted({_s(f["VENDOR"]) for f in fib})})' if fib else ''))
    elif len({_s(f['FIBERCONT']) for f in live}) > 1:
        _flag(flags, 'warn', 'More than one fiber content is valid for this style, vendor and date')
    # ── set components (YOUR_SET_COMPONENTS_TABLE) ──
    comps = set_by.get(article.upper(), [])
    for c in comps:
        if not _s(c['COMP_COMM_CODE']): _flag(flags, 'error', f"Set component of {article} has no commodity code in YOUR_SET_COMPONENTS_TABLE")
    # ── margin file ──
    mr = [m for m in pool if _style_match(m['key'], generic)]
    if mr and any(m['key'] != generic for m in mr): _flag(flags, 'info', f"Margin file style {mr[0]['key']} matched to SAP style {generic} by prefix")
    msum = {'rows': len(mr)}
    if not pool:
        pass
    elif not mr:
        _flag(flags, 'warn', f'Style {generic} is not in the margin file')
    else:
        fobs = sorted({m['final_fob'] for m in mr if m['final_fob'] is not None})
        if fob is not None and fobs and not any(_same(fob, x) for x in fobs): _flag(flags, 'error', f'PO price {fob:g} is not a margin FINAL PO FOB ({", ".join(f"{x:g}" for x in fobs)})')
        terms = {m['term'] for m in mr if m['term']}
        fs_rows = [m for m in mr if m['term'].upper().startswith('FIRST')]
        if is_fs and terms and not fs_rows: _flag(flags, 'warn', f'SAP PO is first sale (Z4) but margin Term is {", ".join(sorted(terms))}')
        if not is_fs and fs_rows and zfs1: _flag(flags, 'info', 'Margin Term is First Sale; SAP PO type is not Z4')
        if fs_rows and zfs1 is not None:
            fsv = sorted({m['first_sale'] for m in fs_rows if m['first_sale'] is not None})
            if fsv and not any(_same(zfs1, x) for x in fsv): _flag(flags, 'error', f'First sale: SAP ZFS1 {zfs1:g} vs margin {", ".join(f"{x:g}" for x in fsv)}')
        for label, sap_v, key in (('Base duty (ZCSR)', zcsr, 'duty_rate'), ('US tariff (ZTRF) vs margin Additional Tariff', ztrf, 'add_tariff')):
            vals = sorted({round(m[key] * 100, 3) for m in mr if m[key] is not None})
            if sap_v is not None and vals and not any(_same(sap_v, x, TOL_RATE) for x in vals):
                _flag(flags, 'warn', f'{label}: SAP {sap_v:g}% vs margin {", ".join(f"{x:g}%" for x in vals)}')
        fin = sorted({round(m['final_duty'] * 100, 3) for m in mr if m['final_duty'] is not None})
        sap_fin = [v for v in (zcst, (zcsr or 0) + (ztrf or 0) if zcsr is not None and ztrf is not None else None, (zcst or 0) + (ztrf or 0) if zcst is not None and ztrf is not None else None) if v is not None]
        if fin and sap_fin and not any(_same(a, b, TOL_RATE) for a in sap_fin for b in fin):
            _flag(flags, 'warn', f'Final duty: margin {", ".join(f"{x:g}%" for x in fin)} matches none of SAP Duty ZCST {zcst:g}%, base+tariff {(zcsr or 0) + (ztrf or 0):g}%, Duty+tariff {(zcst or 0) + (ztrf or 0):g}%')
        mv = {m['vendor'] for m in mr if m['vendor']}
        if mv and vend and not any(vend in x or x in vend for x in mv): _flag(flags, 'warn', f'Vendor: SAP {vend} vs margin {", ".join(sorted(mv))}')
        fabs = sorted({m['fabrication'] for m in mr if m['fabrication']})
        if sap_fiber and fabs:
            res = [fiber_compare(sap_fiber, f) for f in fabs]
            if not any(r[0] == 'ok' for r in res):
                lvl = 'info' if all(r[0] == 'manual' for r in res) else 'warn'
                _flag(flags, lvl, 'Fiber: ' + res[0][1])
        msum.update(fob=fobs, term=sorted(terms), coo=sorted({m['coo'] for m in mr if m['coo']}), first_sale=sorted({m['first_sale'] for m in mr if m['first_sale'] is not None}),
                    duty_rate=sorted({m['duty_rate'] for m in mr if m['duty_rate'] is not None}), fabrication=fabs, rows_at=[m['row'] for m in mr][:20])
    return {'po': po, 'po_item': item, 'req': rq, 'req_item': ri, 'article': article, 'generic': generic, 'desc': _s(base.get('EKPO_ARTICLE_SHORT_TEXT')),
            'vendor': vend, 'factory': _s(base.get('EKPA_FACTORY_ID')), 'po_type': _s(base.get('EKKO_PO_TYPE_ID')), 'po_type_text': _s(base.get('T161T_PO_TYPE_TEXT')),
            'storage_po': _s(base.get('EKPO_STORAGE_LOCATION_ID')), 'storage_req': _s(req.get('EBAN_STORAGE_LOCATION')) if req else '',
            'req_qty_pc': req_qty, 'po_qty_pc': qty, 'other_po_qty_pc': other, 'unit': _s(base.get('T006A_PO_COMMERCIAL_UNIT')),
            'pc_per_pac': num(base.get('MARM_PC_PER_PAC')), 'fob_pc': fob, 'zfs1': zfs1, 'zcsr': zcsr, 'zcst': zcst, 'ztrf': ztrf,
            'hts_master': comm_m, 'hts_po': comm_p, 'dispatch': _s(base.get('EIPO_DISPATCH_COUNTRY')), 'trading': _s(base.get('EIPO_TRADING_COUNTRY')),
            'req_date': blank_date(req.get('EBAN_DELIVERY_DATE')) if req else blank_date(base.get('EBAN_DELIVERY_DATE')),
            'po_delivery': po_deliv, 'ex_factory': ex_fact, 'sched': sched, 'fiber': sap_fiber,
            'set_components': [{'code': _s(c['COMP_COMM_CODE']), 'desc': _s(c['COMP_COMM_DESC']), 'fob': num(c['FOB_PRICE']), 'first_sale': num(c['Z1ST_SALE_PRICE'])} for c in comps],
            'margin': msum, 'flags': flags}

def _req_only_item(rq, ri, r):
    flags = []
    if _s(r.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'info', 'Requisition line is deleted')
    else: _flag(flags, 'info', 'No PO draws on this requisition line yet')
    if not _s(r.get('MAW1_COMMODITY_CODE')): _flag(flags, 'error', 'Article master commodity code (HTS) is blank in SAP')
    po = pad10(r.get('EBAN_PURCHASE_ORDER_NO'))
    if po and not _s(r.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'warn', f'Requisition view says PO {po} item {_s(r.get("EBAN_PO_ITEM_NO"))} but no PO line was found')
    return {'po': '', 'po_item': '', 'req': rq, 'req_item': ri, 'article': _s(r.get('EBAN_ARTICLE')), 'generic': _s(r.get('F_GENERIC')).upper(), 'desc': '',
            'vendor': '', 'factory': '', 'po_type': '', 'po_type_text': '', 'storage_po': '', 'storage_req': _s(r.get('EBAN_STORAGE_LOCATION')),
            'req_qty_pc': num(r.get('F_REQ_QTY_PC')), 'po_qty_pc': None, 'other_po_qty_pc': None, 'unit': '', 'pc_per_pac': None, 'fob_pc': None, 'zfs1': None,
            'zcsr': None, 'zcst': None, 'ztrf': None, 'hts_master': _s(r.get('MAW1_COMMODITY_CODE')), 'hts_po': '', 'dispatch': '', 'trading': '',
            'req_date': blank_date(r.get('EBAN_DELIVERY_DATE')), 'po_delivery': '', 'ex_factory': '', 'sched': [], 'fiber': '', 'set_components': [],
            'margin': {'rows': 0}, 'flags': flags}

# ═══════════════════════════════════════════════════════════════════════════════
# 8. ROUTES
# ═══════════════════════════════════════════════════════════════════════════════
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

def _send_xlsx(buf, stem):
    return send_file(buf, as_attachment=True, download_name=f"{stem}_{datetime.now():%m-%d-%Y}.xlsx", mimetype=XLSX)

def _from_cache(build, stem):
    """Shared body of the three download routes: need a validation in the cache, build the file, send it."""
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        return _send_xlsx(build(), stem)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/')
def portal():
    if not os.path.exists(os.path.join(TOOLS_DIR, 'portal.html')):
        return Response(f"portal.html was not found in {TOOLS_DIR}. Put portal.html in the same folder as server.py.", status=404, mimetype='text/plain')
    return send_from_directory(TOOLS_DIR, 'portal.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'build': BUILD, 'cached': bool(_cache.get('validation')), 'time': datetime.now().isoformat(), 'stock_view': STOCK_VIEW or None,
                    'hana': {'configured': hana_configured(), 'host': HANA_HOST, 'port': HANA_PORT, 'user': HANA_USER, 'settings_from': SETTINGS_SOURCE}})

@app.route('/diagnostics')
def diagnostics():
    return Response(json.dumps(build_diagnostics(_cache), indent=2, default=str), mimetype='text/plain')

@app.route('/explore')
def explore():
    """Read-only look into HANA.  /explore (the six transactions) | /explore?q=YOUR_TXN_1,YOUR_TXN_2 | /explore?view=<name>  (that object only) | /explore?view=<name>&col=<column>&like=<starts with>  (rows for one article etc.)"""
    terms = [t.upper() for t in re.split(r'[,\s]+', request.args.get('q', '')) if t] or None
    view = request.args.get('view', '').strip()
    col = request.args.get('col', '').strip()
    like = request.args.get('like', '').strip()
    if (col or like) and not (view and col and like) or (col and not re.fullmatch(r'[A-Za-z0-9_]{1,60}', col)) or (like and not re.fullmatch(r'[A-Za-z0-9_.\- ]{1,60}', like)):
        return Response('Filtering needs all of view, col and like, e.g. /explore?view=SOMEVIEW&col=MARA_ARTICLE&like=RDM271SS06F  (like is a prefix; letters, digits, _ . - only).', status=400, mimetype='text/plain')
    if any(not re.fullmatch(r'[A-Z0-9_]{2,40}', t) for t in terms or []) or (view and not re.fullmatch(r'[A-Za-z0-9_.\-/]{2,120}', view)):
        return Response('q takes transaction names like YOUR_TXN_1,YOUR_TXN_2; view takes a view name from the first section.', status=400, mimetype='text/plain')
    try:
        conn = connect_hana()
    except Exception as exc:
        return Response(f'Could not connect to HANA: {exc}', status=502, mimetype='text/plain')
    try:
        return Response(json.dumps(explore_sap(conn, terms, view or None, col or None, like or None), indent=2, default=str), mimetype='text/plain')
    finally:
        conn.close()

@app.route('/lookup', methods=['GET'])
def lookup():
    number, kind = request.args.get('q', '').strip(), request.args.get('kind', 'auto')
    if kind not in ('auto', 'po', 'req'): return jsonify({'error': 'kind must be auto, po or req'}), 400
    if not pad10(number): return jsonify({'error': 'Enter a PO or requisition number.'}), 400
    try:
        conn = connect_hana()
    except Exception as e:
        return jsonify({'error': f'Could not connect to HANA: {e}'}), 502
    try:
        k, po_rows, req_rows, fiber_rows, set_rows, notes = fetch_lookup_data(conn, number, kind)
    except LookupError as e:
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        return jsonify({'error': f'SAP query failed: {e}'}), 502
    finally:
        try: conn.close()
        except Exception: pass
    return jsonify(build_lookup(k, pad10(number), po_rows, req_rows, fiber_rows, set_rows, _cache.get('margin'), notes))

def margin_info(m):
    if not m: return {'loaded': False}
    rows = m['rows']
    return {'loaded': True, 'filename': m.get('filename'), 'sheet': m['sheet'], 'rows': len(rows), 'styles': len({r['key'] for r in rows}),
            'with_req': sum(1 for r in rows if r['req']), 'with_po': sum(1 for r in rows if r['po']),
            'terms': {t: sum(1 for r in rows if r['term'] == t) for t in sorted({r['term'] for r in rows})}}

@app.route('/margin', methods=['GET'])
def margin_current():
    return jsonify(margin_info(_cache.get('margin')))

@app.route('/margin', methods=['POST'])
def margin_upload():
    if 'file' not in request.files: return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        m = parse_margin(f.read()); m['filename'] = f.filename
    except Exception as e:
        return jsonify({'error': str(e)}), 400
    _cache['margin'] = m
    return jsonify(margin_info(m))

@app.route('/margin/clear', methods=['POST'])
def margin_clear():
    _cache.pop('margin', None)
    return jsonify(margin_info(None))

@app.route('/validate', methods=['POST'])
def validate():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        rows, customers, file_type = parse_prebuy(f.read())
        validation = validate_all(rows)
        validation.update(customers=customers, filename=f.filename, file_type=file_type)
        apply_stock(validation, _cache.get('stock') or live_stock(validation))   # an export added earlier wins, else live from HANA
        _cache.update(validation=validation, rows=rows, customers=customers)
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    return jsonify(validation_payload(validation))

def validation_payload(validation):
    return {'build': BUILD, 'filename': validation.get('filename'), 'file_type': validation.get('file_type'), 'customers': validation.get('customers', []),
            'summary': validation['summary'], 'generic': validation['generic'], 'ppk': validation['ppk'],
            'generic_unique': validation['generic_unique'], 'ppk_unique': validation['ppk_unique']}

@app.route('/validation', methods=['GET'])
def validation_current():
    if not _cache.get('validation'):
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    return jsonify(validation_payload(_cache['validation']))

@app.route('/stock', methods=['GET'])
def stock_current():
    return jsonify({'loaded': bool(_cache.get('stock')), 'stock': stock_info(_cache.get('stock')), 'live': STOCK_VIEW or None})

@app.route('/stock', methods=['POST'])
def stock_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        stock = parse_cut_sold(f.read(), f.filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 400
    _cache['stock'] = stock
    refreshed = bool(_cache.get('validation'))
    if refreshed: apply_stock(_cache['validation'], stock)
    return jsonify({'stock': stock_info(stock), 'refreshed': refreshed})

@app.route('/stock/clear', methods=['POST'])
def stock_clear():
    _cache.pop('stock', None)
    refreshed = bool(_cache.get('validation'))
    if refreshed: apply_stock(_cache['validation'], live_stock(_cache['validation']))      # back to live stock
    return jsonify({'loaded': False, 'refreshed': refreshed, 'live': STOCK_VIEW or None})

@app.route('/generate', methods=['GET'])
def generate():
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        buf, kind = make_buysheet(_cache['rows'], _cache.get('customers', []))
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    if kind == 'xlsm':
        return send_file(buf, as_attachment=True, download_name=f"Buysheet_{datetime.now():%m-%d-%Y}.xlsm", mimetype=XLSM)
    return _send_xlsx(buf, 'Buysheet_DATA')

@app.route('/creation_templates', methods=['GET'])
def creation_templates():
    if not _cache.get('validation'):
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    return _from_cache(lambda: make_creation_templates(_cache['validation']), 'Creation_Templates')

@app.route('/validation_report', methods=['GET'])
def validation_report():
    if not _cache.get('validation'):
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    return _from_cache(lambda: make_validation_report(_cache['validation']), 'Validation')

@app.route('/req_csv', methods=['GET'])
def req_csv():
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        content = make_req_csv(_cache['rows'], _cache.get('customers', []), _cache.get('validation'))
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return send_file(io.BytesIO(content.encode('utf-8')), as_attachment=True,
                     download_name=f"PR_Upload_{datetime.now():%m-%d-%Y}.csv", mimetype='text/csv')

@app.route('/delivery_split', methods=['POST'])
def delivery_split():
    if 'file' not in request.files: return jsonify({'error':'No file'}), 400
    f=request.files['file']
    del1=int(request.form.get('del1',0)); del2=int(request.form.get('del2',0))
    del1_date=request.form.get('del1_date',''); del2_date=request.form.get('del2_date','')
    if not del1 or not del2: return jsonify({'error':'Missing delivery quantities'}), 400
    total=del1+del2; prop=del1/total

    wb_in=load_workbook(io.BytesIO(f.read()),data_only=True); ws_in=wb_in.active
    header_ri=None; headers=[]
    for ri,row in enumerate(ws_in.iter_rows(values_only=True),1):
        cells=[str(v).strip().upper() if v else '' for v in row]
        if 'COLOR' in cells or 'STYLE' in cells:
            header_ri=ri; headers=cells; break
    if not header_ri: return jsonify({'error':'Could not find header row'}),400

    color_ci=next((headers.index(h) for h in ['COLOR DESCRIPTION','COLOR'] if h in headers),3)
    ratio_ci=next((headers.index(h) for h in ['RATIO','RATIOS'] if h in headers),None)
    style_ci=headers.index('STYLE') if 'STYLE' in headers else 1
    all_sizes=[h for h in headers if h in KNOWN_SIZES]
    size_cis={sz:headers.index(sz) for sz in all_sizes}

    wb_out=Workbook(); ws_out=wb_out.active; ws_out.title="DELIVERY SPLIT"
    out_hdrs=['STYLE','COLOR','DELIVERY']+all_sizes+['TOTAL']
    for ci,h in enumerate(out_hdrs,1):
        c=ws_out.cell(1,ci,value=h); c.font=WBOLD; c.fill=NAVY; c.alignment=CTR

    ri_out=2
    for row in ws_in.iter_rows(min_row=header_ri+1,values_only=True):
        if not any(v for v in row): continue
        style=str(row[style_ci]).strip() if row[style_ci] else ''
        color=str(row[color_ci]).strip() if row[color_ci] else ''
        if not color: continue
        sizes={sz:int(float(row[headers.index(sz)])) for sz in all_sizes
               if sz in headers and row[headers.index(sz)] and str(row[headers.index(sz)]).strip() not in ('','0','None')}
        total_qty=sum(sizes.values())
        if not total_qty: continue
        d1_sizes={sz:round(qty*prop) for sz,qty in sizes.items()}
        d2_sizes={sz:sizes[sz]-d1_sizes[sz] for sz in sizes}
        for label,sz_data,fill in [
            ('TOTAL',sizes,PatternFill('solid',start_color='D9E1F2')),
            (f'DEL 1{" ("+del1_date+")" if del1_date else ""}',d1_sizes,PatternFill('solid',start_color='E2EFDA')),
            (f'DEL 2{" ("+del2_date+")" if del2_date else ""}',d2_sizes,None)]:
            row_total=sum(sz_data.values())
            vals=[style,color,label]+[sz_data.get(sz,'') for sz in all_sizes]+[row_total]
            for ci,val in enumerate(vals,1):
                c=ws_out.cell(ri_out,ci,value=val)
                c.font=BFONT if label=='TOTAL' else DFONT; c.alignment=CTR
                if fill: c.fill=fill
            ri_out+=1
        ri_out+=1

    buf=io.BytesIO(); wb_out.save(buf); buf.seek(0)
    today=datetime.now().strftime('%m-%d-%Y')
    return send_file(buf,as_attachment=True,download_name=f'Delivery_Split_{today}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def lan_ip():
    """This machine's address on the network: what other computers type to reach the portal."""
    import socket
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(('10.255.255.255', 1))        # no packet is sent; it just picks the outgoing interface
        ip = sock.getsockname()[0]; sock.close()
        return ip
    except Exception:
        return None

if __name__ == '__main__':
    print("Buysheet Portal starting on http://0.0.0.0:5050")
    ip = lan_ip()
    print("Open in browser:        http://localhost:5050")
    if ip: print(f"From other computers:   http://{ip}:5050   (this machine's address today)")
    print(f"Server build: {BUILD}")
    print(f"HANA connection:        {HANA_USER or '?'}@{HANA_HOST or '?'}:{HANA_PORT}   (settings: {SETTINGS_SOURCE})")
    if not hana_configured():
        print("  >> HANA IS NOT SET UP YET: copy config.example.json to config.json, fill in host, user and password, then restart.")
    if not os.path.exists(os.path.join(TOOLS_DIR, 'portal.html')):
        print(f"  >> portal.html was not found in {TOOLS_DIR}: put it in the same folder as server.py.")
    print("Press Ctrl+C to stop\n")
    app.run(host='0.0.0.0', port=5050, debug=False)
