"""
PHÚ MỸ LUMEN CONNECTOR - v3.0
Cài đặt: pip install streamlit pandas openpyxl python-dateutil
Chạy   : streamlit run phu_my_lumen_connector.py

Thay đổi v3.0:
  1. DON_GIA_BH = DON_GIA (tự động)
  2. NUOC_SX ánh xạ từ cột Xuất xứ
  3. Quản lý nhà thầu: filter tìm nhanh + điền ngày 1 lần cho nhiều NCC
  4. Đường dùng: chỉ khớp CHÍNH XÁC với bảng Cách dùng thuốc, không khớp → để trống tô vàng
  5. MA_THUOC + TEN_HOAT_CHAT: chỉ khớp 100% từ file tân dược, TEN_HOAT_CHAT xuất ra đúng
     theo tên chuẩn trong file tân dược (không lấy tên file thầu)
  6. MA_THUOC xuất ra dạng text, không bị Excel đổi thành số
  7. Mã có đuôi .1/.2: tô tím, hiển thị MIEUTA để người dùng chọn đúng mã
"""

import csv
import io
import os
import re
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import streamlit as st

# ================================================================
# FILE LƯU TRỮ CỤC BỘ
# ================================================================
LOG_FILE   = Path('mapping_log.csv')
ALIAS_FILE    = Path('alias_ten.csv')   # ánh xạ tên hoạt chất
ALIAS_DD_FILE = Path('alias_dd.csv')    # ánh xạ đường dùng
LOG_COLS      = ['NHA_THAU', 'NGAY_KY', 'THOI_HAN', 'DEN_NGAY']
ALIAS_COLS    = ['TEN_THAU', 'TEN_TANDUC', 'GHI_CHU']
ALIAS_DD_COLS = ['TEN_HOAT_CHAT', 'DD_GOC', 'DD_CHUAN', 'GHI_CHU']

# ================================================================
# MODULE LOG
# ================================================================
def load_log() -> dict:
    result = {}
    if not LOG_FILE.exists():
        return result
    try:
        with open(LOG_FILE, 'r', encoding='utf-8', newline='') as f:
            for row in csv.DictReader(f):
                nt = row.get('NHA_THAU', '').strip()
                if not nt: continue
                try:
                    result[nt] = {
                        'da_ky':   True,
                        'ngay_ky': datetime.strptime(row['NGAY_KY'], '%Y%m%d').date(),
                        'thoi_han':int(row['THOI_HAN']),
                        'den_ngay':datetime.strptime(row['DEN_NGAY'], '%Y%m%d').date(),
                    }
                except Exception: pass
    except Exception: pass
    return result

def save_log(nha_thau_info: dict):
    rows = [
        {'NHA_THAU': nt,
         'NGAY_KY':  info['ngay_ky'].strftime('%Y%m%d'),
         'THOI_HAN': str(info.get('thoi_han', 12)),
         'DEN_NGAY': info['den_ngay'].strftime('%Y%m%d') if info.get('den_ngay') else ''}
        for nt, info in nha_thau_info.items()
        if info.get('da_ky') and info.get('ngay_ky')
    ]
    try:
        with open(LOG_FILE, 'w', encoding='utf-8', newline='') as f:
            w = csv.DictWriter(f, fieldnames=LOG_COLS)
            w.writeheader(); w.writerows(rows)
    except Exception as e:
        st.warning(f'⚠️ Không ghi được mapping_log.csv: {e}')

# ================================================================
# MODULE ALIAS MAP
# ================================================================
def load_alias() -> pd.DataFrame:
    """Đọc alias tên hoạt chất: TEN_THAU → TEN_TANDUC."""
    if not ALIAS_FILE.exists():
        return pd.DataFrame(columns=ALIAS_COLS)
    try:
        df = pd.read_csv(ALIAS_FILE, dtype=str, encoding='utf-8').fillna('')
        for col in ALIAS_COLS:
            if col not in df.columns: df[col] = ''
        return df[ALIAS_COLS]
    except Exception:
        return pd.DataFrame(columns=ALIAS_COLS)

def save_alias(df_alias: pd.DataFrame):
    try:
        df_alias.to_csv(ALIAS_FILE, index=False, encoding='utf-8')
    except Exception as e:
        st.warning(f'⚠️ Không ghi được alias_ten.csv: {e}')

def build_alias_lookup(df_alias: pd.DataFrame) -> dict:
    """
    Từ điển alias TÊN: ten_thau_norm → ten_tanduc (chuỗi gốc, chưa normalize).
    Chỉ ánh xạ tên, KHÔNG liên quan đến đường dùng.
    """
    lk = {}
    for _, r in df_alias.iterrows():
        k = _norm(r.get('TEN_THAU', ''))
        v = r.get('TEN_TANDUC', '').strip()
        if k and v:
            lk[k] = v   # key = tên file thầu norm, value = tên chuẩn file TD
    return lk

def load_alias_dd() -> pd.DataFrame:
    """Đọc alias đường dùng: (TEN_HOAT_CHAT, DD_GOC) → DD_CHUAN."""
    if not ALIAS_DD_FILE.exists():
        return pd.DataFrame(columns=ALIAS_DD_COLS)
    try:
        df = pd.read_csv(ALIAS_DD_FILE, dtype=str, encoding='utf-8').fillna('')
        for col in ALIAS_DD_COLS:
            if col not in df.columns: df[col] = ''
        return df[ALIAS_DD_COLS]
    except Exception:
        return pd.DataFrame(columns=ALIAS_DD_COLS)

def save_alias_dd(df: pd.DataFrame):
    try:
        df.to_csv(ALIAS_DD_FILE, index=False, encoding='utf-8')
    except Exception as e:
        st.warning(f'⚠️ Không ghi được alias_dd.csv: {e}')

def build_alias_dd_lookup(df_dd: pd.DataFrame) -> dict:
    """
    Từ điển alias ĐD: (ten_hc_norm, dd_goc_norm) → dd_chuan (chuỗi gốc).
    Khi đường dùng trong file thầu không khớp bảng chuẩn,
    tra ở đây để lấy đường dùng chuẩn đã được user xác nhận.
    """
    lk = {}
    for _, r in df_dd.iterrows():
        hc  = _norm(r.get('TEN_HOAT_CHAT', ''))
        dd_g = _norm(r.get('DD_GOC', ''))
        dd_c = r.get('DD_CHUAN', '').strip()
        if hc and dd_g and dd_c:
            lk[(hc, dd_g)] = dd_c
    return lk

# ================================================================
# MODULE TIỀN XỬ LÝ
# ================================================================
def sc(val) -> str:
    if pd.isna(val): return ''
    return str(val).replace('\n', ' ').replace('\r', '').strip()

def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [sc(c) for c in df.columns]
    for col in df.columns:
        df[col] = df[col].apply(sc)
    return df

def find_header(df_raw, keywords=('STT', 'Tên hoạt chất', 'Số lượng', 'Nhà thầu')) -> int:
    for i, row in df_raw.iterrows():
        text = ' '.join(str(v) for v in row.values if str(v) != 'nan')
        if sum(1 for kw in keywords if kw in text) >= 2:
            return i
    return 0

def read_excel(file_bytes: bytes, keywords=None) -> pd.DataFrame:
    kw = keywords or ('STT', 'Tên hoạt chất', 'Số lượng')
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp.write(file_bytes); path = tmp.name
    df_raw = pd.read_excel(path, header=None, dtype=str)
    hrow   = find_header(df_raw, kw)
    df     = pd.read_excel(path, header=hrow, dtype=str)
    os.unlink(path)
    return sanitize_df(df).dropna(how='all').reset_index(drop=True)

def read_xls_compat(file_bytes: bytes) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    for engine in ('openpyxl', None):
        buf.seek(0)
        try:
            kw = {'engine': engine} if engine else {}
            return sanitize_df(pd.read_excel(buf, header=0, dtype=str, **kw))
        except Exception: pass
    raise RuntimeError("Không đọc được file. Thử: pip install xlrd==1.2.0")

# ================================================================
# MODULE REGEX - TÁCH SĐK
# ================================================================
RE_NEW    = re.compile(r'\b(\d{12})\b')
RE_OLD_VN = re.compile(r'\b(VN\d*-\d+-\d+)\b',    re.IGNORECASE)
RE_OLD_VD = re.compile(r'\b(VD-\d+-\d+)\b',        re.IGNORECASE)
RE_OLD_QL = re.compile(r'\b((?:QLSP|QLDB|QLĐB|GPNK|GC|DP|PB|TN|SP\d+)-[\d\w/-]+)\b', re.IGNORECASE)
RE_QLD    = re.compile(r'(\d{3,5}/QLD-\w+)', re.IGNORECASE)

def _csdk(s): return s.strip().rstrip('.,;) ')

def parse_sdk(sdk_raw: str) -> list:
    """Trả về list (sdk, is_primary, is_qld_kd)."""
    if not sdk_raw: return [('', True, False)]
    sdk_raw = str(sdk_raw).strip().strip('"')
    new_list = RE_NEW.findall(sdk_raw)
    old_list, seen = [], set()
    for pat in (RE_OLD_VN, RE_OLD_VD, RE_OLD_QL):
        for m in pat.findall(sdk_raw):
            v = _csdk(m)
            if v and v not in seen: seen.add(v); old_list.append(v)
    qld_list = [x.strip() for x in RE_QLD.findall(sdk_raw)]
    if new_list and old_list:
        return [(_csdk(new_list[0]), True, False)] + [(o, False, False) for o in old_list]
    if new_list: return [(_csdk(new_list[0]), True, False)]
    if old_list: return [(old_list[0], True, False)]
    if qld_list:
        return [(qld_list[0], True, True)] + [(q, False, True) for q in qld_list[1:]]
    cleaned = re.split(r'[\.(]', sdk_raw)[0].strip()
    return [(cleaned or sdk_raw, True, False)]

def expand_sdk_rows(df, col_sdk, col_sl):
    rows = []
    for _, row in df.iterrows():
        for sdk_val, is_primary, is_qld in parse_sdk(str(row.get(col_sdk, ''))):
            r = row.copy()
            r[col_sdk] = sdk_val
            r['IS_QLD_KD'] = 'CẦN RÀ SOÁT' if is_qld else ''
            if not is_primary: r[col_sl] = '0'
            rows.append(r)
    return pd.DataFrame(rows).reset_index(drop=True)

# ================================================================
# MODULE CHUẨN HÓA & TRA CỨU
# ================================================================
def _norm(s: str) -> str:
    """Chuẩn hóa chuỗi: lower, strip, chuẩn hóa dấu + và khoảng trắng."""
    s = str(s).lower().strip()
    s = re.sub(r'\s*\+\s*', ' + ', s)   # chuẩn hóa dấu + có khoảng trắng cố định
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def _norm_tight(s: str) -> str:
    # Nhu _norm nhung dau + KHONG co khoang trang -> dung de so sanh noi bo
    s = str(s).lower().strip()
    s = re.sub(r'\s*\+\s*', '+', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def _count_components(s: str) -> int:
    # Dem so hoat chat dua vao dau +. Vi du: 'A + B + C' -> 3
    return len(re.split(r'\s*\+\s*', s.strip()))

# Bảng alias đường dùng tích hợp (không cần user cấu hình)
# Chuẩn hóa các biến thể phổ biến về dạng canonical
_DD_ALIAS_BUILTIN: list[tuple[list[str], str]] = [
    (['tiêm bắp (im)', 'tiêm bắp im', 'tiêm bắp', 'im', 'intramuscular'], 'tiêm bắp'),
    (['tiêm tĩnh mạch (iv)', 'tiêm tĩnh mạch iv', 'tiêm tĩnh mạch', 'iv', 'intravenous', 'tiêm tm'], 'tiêm tĩnh mạch'),
    (['tiêm dưới da (sc)', 'tiêm dưới da sc', 'tiêm dưới da', 'sc', 'subcutaneous'], 'tiêm dưới da'),
    (['uống', 'viên nang', 'viên nén', 'oral', 'per os'], 'uống'),
    (['nhỏ mắt', 'tra mắt', 'ophthalmic'], 'nhỏ mắt'),
    (['dùng ngoài da', 'bôi ngoài da', 'bôi da', 'topical'], 'dùng ngoài da'),
    (['đặt âm đạo', 'đặt âm đạo (vaginal)', 'vaginal'], 'đặt âm đạo'),
    (['đặt hậu môn', 'rectal', 'đặt trực tràng'], 'đặt hậu môn'),
    (['hít', 'xịt', 'inhalation', 'inhaled'], 'hít'),
    (['truyền tĩnh mạch', 'truyền tm', 'iv infusion', 'infusion'], 'truyền tĩnh mạch'),
    (['màng bụng', 'dùng theo đường màng bụng', 'lọc màng bụng', 'thẩm phân', 'peritoneal', 'dialysis'], 'dùng theo đường màng bụng'),
]

def _dd_canonical(dd: str) -> str:
    """Chuẩn hóa đường dùng về dạng canonical dùng để so sánh."""
    n = _norm(dd)
    for variants, canonical in _DD_ALIAS_BUILTIN:
        for v in variants:
            if v in n or n in v:
                return canonical
    return n

def _dd_similarity(dd_thau: str, dd_td: str) -> float:
    """Tính độ tương đồng đường dùng (0.0 → 1.0).

    Thứ tự ưu tiên:
    1. Exact string match (sau _norm_tight) → 1.0 (tuyệt đối cao nhất)
       VD: dd_thau='Tiêm' vs dd_td='Tiêm' → 1.0
    2. dd_thau='tiêm' (1 từ, không phân biệt) VÀ dd_td bắt đầu bằng 'tiêm'
       → 0.9 (bao quát — 'Tiêm truyền', 'Tiêm bắp'... đều được đề xuất)
    3. Cùng canonical (alias group) → 0.85
    4. difflib ratio + từ chung bonus
    """
    import difflib as _dl
    n1 = _norm_tight(dd_thau)
    n2 = _norm_tight(dd_td)

    # Mức 1: Exact match → 1.0
    if n1 == n2:
        return 1.0

    # Mức 2: 'tiêm' bao quát mọi dạng tiêm (1 từ, không suffix)
    # Áp dụng khi file thầu chỉ ghi "Tiêm" — bắt cả "Tiêm truyền", "Tiêm bắp"...
    if n1 == 'tiêm' and n2.startswith('tiêm'):
        return 0.9
    if n1 == 'tiem' and n2.startswith('tiem'):
        return 0.9

    # Mức 2b: 'tiêm truyền' bao quát 'truyền tĩnh mạch', 'tiêm truyền tĩnh mạch'
    if n1 in ('tiêm truyền', 'tiem truyen') and ('truyền' in n2 or 'truyen' in n2):
        return 0.88

    # Mức 3: Cùng canonical alias group → 0.85
    c1 = _dd_canonical(dd_thau)
    c2 = _dd_canonical(dd_td)
    if c1 and c2 and c1 == c2:
        return 0.85

    # Mức 4: difflib + từ chung bonus
    n1f = _norm(dd_thau); n2f = _norm(dd_td)
    ratio = _dl.SequenceMatcher(None, n1f, n2f).ratio()
    words1 = set(n1f.split()); words2 = set(n2f.split())
    common = words1 & words2
    if common:
        ratio = min(0.8, ratio + 0.1 * len(common))
    return ratio

def build_thuoc_lookup(df_thuoc: pd.DataFrame):
    """
    Xây dựng từ điển tra cứu MA_THUOC (v3.2).

    Nguyên tắc:
    - Key dùng _norm_tight (dấu + không khoảng trắng) để so sánh nội bộ
    - TEN lưu NGUYÊN GỐC từ file, không thêm/sửa nội dung
    - goi_y: hc_norm_tight → {ma: (dd_raw, ten_raw)} — dedup theo MA
    - Mỗi MA chỉ xuất hiện 1 lần (lấy dòng đầu tiên gặp)
    """
    lk           = {}   # (hc_tight, dd_tight) → MA
    name_map     = {}   # (hc_tight, dd_tight) → TEN gốc
    goi_y        = {}   # hc_tight → {ma: (dd_raw, ten_raw)}
    conflict_map = {}

    for _, r in df_thuoc.iterrows():
        ten_raw = sc(r.get('TEN', ''))
        dd_raw  = sc(r.get('DUONGDUNG', ''))
        ma      = sc(r.get('MA', ''))
        ten_t   = _norm_tight(ten_raw)   # tight: dấu + không khoảng trắng
        dd_t    = _norm_tight(dd_raw)
        if not ten_t or not ma:
            continue

        key = (ten_t, dd_t)
        conflict_map.setdefault(key, set())
        conflict_map[key].add(ma)

        if key not in lk:
            lk[key]      = ma
            name_map[key] = ten_raw   # lưu TEN GỐC, không sửa

        # Gợi ý: dedup theo (MA, DD) — cùng mã nhưng khác đường dùng phải lấy đủ
        goi_y.setdefault(ten_t, {})
        pair_k = f"{ma}||{dd_raw}"
        if pair_k not in goi_y[ten_t]:
            goi_y[ten_t][pair_k] = (dd_raw, ma, ten_raw)   # key → (dd, ma, ten)

    # Chuyển sang list [(dd_raw, ma, ten_raw)]
    goi_y_list = {
        hc: [(dd, ma, ten) for dd, ma, ten in ma_dd.values()]
        for hc, ma_dd in goi_y.items()
    }

    conflict_map = {k: sorted(v) for k, v in conflict_map.items() if len(v) > 1}

    return lk, name_map, goi_y_list, conflict_map


def build_dd_lookup(df_cach: pd.DataFrame) -> dict:
    """
    Tra cứu MA_DUONG_DUNG. Khớp CHÍNH XÁC tên đường dùng.
    Trả về dict: ten_norm → MA_DUONG_DUNG
    """
    lk = {}
    cols = df_cach.columns.tolist()
    col_ma, col_ten = cols[0], cols[1]
    for _, r in df_cach.iterrows():
        ma  = sc(r.get(col_ma, ''))
        ten = _norm(sc(r.get(col_ten, '')))
        if ten and ma:
            lk[ten] = ma
    return lk


def lookup_thuoc(hoat_chat: str, duong_dung: str,
                 lk: dict, name_map: dict, goi_y: dict,
                 alias_ten_lk: dict, alias_dd_lk: dict):
    """
    Tra MA_THUOC + TEN_HOAT_CHAT chuẩn — v3.2 (Strict Component Matching).

    Quy tắc mới:
    - Kiểm tra số hoạt chất (dấu +): file thầu có 3 HC → chỉ khớp với TD có 3 HC.
    - Tầng 1: Exact match (tight norm) trên cả tên + đường dùng.
    - Tầng 2: Exact tên + canonical đường dùng.
    - Tầng 3: Trả hints — NHƯNG chỉ lấy entries có cùng số thành phần.

    Trả về: (MA_THUOC, TEN_CHUẨN, hints_list)
      hints_list: [(dd_raw, ma, ten_raw)]
    """
    if not hoat_chat:
        return '', '', []

    n_comp_thau = _count_components(hoat_chat)   # số hoạt chất từ file thầu
    hc_tight    = _norm_tight(hoat_chat)
    dd_tight    = _norm_tight(duong_dung) if duong_dung else ''

    # v3.6: Không dùng alias tự động (an toàn nghiệp vụ y tế — user tự xác nhận mỗi lần)
    dd_use = dd_tight
    hc_candidates = [hc_tight]

    # Tầng 1: Khớp chính xác tên + đường dùng
    for k_hc in hc_candidates:
        key = (k_hc, dd_use)
        if key in lk:
            return lk[key], name_map.get(key, hoat_chat), []

    # Tầng 2: Khớp chính xác tên + canonical đường dùng
    dd_canon = _dd_canonical(dd_use)
    for k_hc in hc_candidates:
        for (lk_hc, lk_dd), ma in lk.items():
            if lk_hc == k_hc and _dd_canonical(lk_dd) == dd_canon and dd_canon:
                ten_chuan = name_map.get((lk_hc, lk_dd), hoat_chat)
                return ma, ten_chuan, []

    # Tầng 3: Không khớp → hints chỉ từ key hc_tight (exact name match)
    # v3.6: không dùng alias, chỉ lấy từ goi_y[hc_tight]
    hints_all = list(goi_y.get(hc_tight, {}).values()) if isinstance(goi_y.get(hc_tight), dict) else goi_y.get(hc_tight, [])

    # Lọc: số thành phần phải bằng nhau
    dd_orig_n = _norm_tight(duong_dung) if duong_dung else ''
    hints_filtered = [e for e in hints_all if _count_components(e[2]) == n_comp_thau]
    hints_filtered.sort(key=lambda e: _dd_similarity(dd_orig_n, e[0]), reverse=True)
    return '', '', hints_filtered


def lookup_dd(duong_dung: str, lk: dict) -> str:
    """Khớp CHÍNH XÁC đường dùng với bảng chuẩn. Không khớp → trả ''."""
    if not duong_dung:
        return ''
    return lk.get(_norm(duong_dung), '')

# ================================================================
# MODULE NHÓM & NGÀY
# ================================================================
def convert_nhom(raw: str) -> str:
    if not raw: return ''
    m = re.search(r'[Nn]h[oóô]m\s*(\d+)', raw)
    if m: return f"N{m.group(1)}"
    m2 = re.match(r'^(N\d+)$', raw.strip(), re.IGNORECASE)
    if m2: return m2.group(1).upper()
    return raw.strip()

def to_yyyymmdd(d) -> str:
    if d is None: return ''
    if isinstance(d, (datetime, date)): return d.strftime('%Y%m%d')
    return ''

def add_months(d: date, n: int) -> date:
    return d + relativedelta(months=int(n))

# ================================================================
# MODULE XUẤT FILE MAU_03
# ================================================================
MAU03_COLS = [
    'STT','MA_THUOC','TEN_HOAT_CHAT','TEN_THUOC','DON_VI_TINH',
    'HAM_LUONG','DUONG_DUNG','MA_DUONG_DUNG','DANG_BAO_CHE',
    'SO_DANG_KY','SO_LUONG','DON_GIA','DON_GIA_BH','QUY_CACH',
    'NHA_SX','NUOC_SX','NHA_THAU','TT_THAU',
    'TU_NGAY_HD','DEN_NGAY_HD','MA_CSKCB','LOAI_THUOC',
    'LOAI_THAU','HT_THAU','MA_DVKT','TCCL','BO_PHAN_VT',
    'TEN_KHOA_HOC','NGUON_GOC','PP_CHEBIEN','MA_DL_NHAP',
    'MA_DL_CB','TLHH_CB','TLHH_BQ','MA_CSKCB_THUOC',
    'TU_NGAY','DEN_NGAY',
]

RED_FILL  = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
ORG_FILL  = PatternFill(start_color='FFE5CC', end_color='FFE5CC', fill_type='solid')
YEL_FILL  = PatternFill(start_color='FFFACC', end_color='FFFACC', fill_type='solid')
PUR_FILL  = PatternFill(start_color='E8D5FF', end_color='E8D5FF', fill_type='solid')
NORM_FONT = Font(name='Arial', size=10)
TEXT_FONT = Font(name='Arial', size=10)

def export_mau03(df: pd.DataFrame, template_bytes: bytes, filter_nt=None) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    hmap = {}
    for c in range(1, ws.max_column + 2):
        val = ws.cell(1, c).value
        if val: hmap[str(val).strip()] = c

    df_out = df.copy()
    if filter_nt is not None:
        df_out = df_out[df_out['NHA_THAU'].isin(filter_nt)]
    df_out = df_out.reset_index(drop=True)

    for row_i, (_, row) in enumerate(df_out.iterrows(), start=2):
        no_ma    = not str(row.get('MA_THUOC', '')).strip()
        is_qld   = str(row.get('IS_QLD_KD', '')) == 'CẦN RÀ SOÁT'
        is_no_dd = not str(row.get('MA_DUONG_DUNG', '')).strip() and str(row.get('DUONG_DUNG','')).strip()
        is_sub   = str(row.get('IS_MA_CON', '')) == '1'

        for col_name, col_idx in hmap.items():
            val = row.get(col_name, '')
            val = '' if pd.isna(val) else str(val)

            # MA_THUOC: luôn xuất dạng text để tránh Excel đổi thành số/float
            if col_name == 'MA_THUOC' and val:
                cell = ws.cell(row=row_i, column=col_idx, value=val)
                cell.number_format = '@'  # định dạng text
            else:
                cell = ws.cell(row=row_i, column=col_idx, value=val or None)

            cell.font = NORM_FONT
            # Tô màu theo ưu tiên: đỏ > cam > vàng > tím
            if no_ma:       cell.fill = RED_FILL
            elif is_qld:    cell.fill = ORG_FILL
            elif is_no_dd:  cell.fill = YEL_FILL
            elif is_sub:    cell.fill = PUR_FILL

    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()

# ================================================================
# ÁNH XẠ CỘT
# ================================================================
COL_KW = {
    'TEN_HOAT_CHAT':['hoạt chất mời thầu','hoạt chất','thành phần'],
    'TEN_THUOC':    ['tên thuốc','biệt dược'],
    'DON_VI_TINH':  ['đơn vị tính','đvt','đơn vị'],
    'HAM_LUONG':    ['nồng độ','hàm lượng'],
    'DUONG_DUNG':   ['đường dùng'],
    'DANG_BAO_CHE': ['dạng bào chế'],
    'SO_DANG_KY':   ['gđklh','gpnk','số đăng ký','định danh'],
    'SO_LUONG':     ['số lượng'],
    'DON_GIA':      ['đơn giá trúng thầu','đơn giá'],
    'QUY_CACH':     ['quy cách'],
    # NHA_SX: chỉ khớp cột "cơ sở sản xuất" / "nhà sản xuất" — KHÔNG lấy "Xuất xứ"
    'NHA_SX':       ['tên cơ sở sản xuất', 'cơ sở sản xuất', 'nhà sản xuất', 'manufacturer'],
    # NUOC_SX: ưu tiên tuyệt đối "xuất xứ" — keyword riêng, không chồng lấn NHA_SX
    'NUOC_SX':      ['xuất xứ', 'nước sản xuất', 'nuoc sx', 'nước sx',
                     'xuat xu', 'origin', 'country'],
    'NHA_THAU':     ['tên nhà thầu','nhà thầu'],
    'NHOM_THUOC':   ['nhóm thuốc','nhóm'],
    'GOI_THAU':     ['gói'],
}

# Fragment fuzzy cho NUOC_SX — chỉ dùng khi keyword chính không khớp
_FUZZY_FRAGMENTS = {
    'NUOC_SX': ['xuất xứ', 'xuất', 'xứ', 'origin', 'country'],
}

# Keyword bị cấm dùng cho NUOC_SX (để tránh nhầm với NHA_SX)
_NUOC_SX_BLACKLIST = ['sản xuất', 'nhà sản', 'cơ sở']

def suggest_col(source_cols, target):
    """
    Gợi ý cột nguồn tương ứng với target.

    Quy tắc đặc biệt cho NUOC_SX vs NHA_SX:
    - NUOC_SX: phải chứa 'xuất xứ' hoặc 'xuất' (nhưng KHÔNG phải 'sản xuất')
    - NHA_SX: phải chứa 'cơ sở sản xuất' hoặc 'nhà sản xuất'
    """
    kws = COL_KW.get(target, [target.lower()])
    sl  = [c.lower() for c in source_cols]

    # Bước 1: khớp cứng theo keyword
    for kw in kws:
        kw_norm = kw.lower()
        for i, sc_val in enumerate(sl):
            if kw_norm in sc_val:
                # Nếu là NUOC_SX, cần đảm bảo cột này KHÔNG phải NHA_SX
                if target == 'NUOC_SX':
                    if any(bl in sc_val for bl in _NUOC_SX_BLACKLIST):
                        continue  # Bỏ qua nếu tên cột thuộc nhóm "nhà sản xuất"
                return source_cols[i]

    # Bước 2: fuzzy fragment (chỉ target có _FUZZY_FRAGMENTS)
    for frag in _FUZZY_FRAGMENTS.get(target, []):
        for i, sc_val in enumerate(sl):
            if frag in sc_val:
                if target == 'NUOC_SX':
                    if any(bl in sc_val for bl in _NUOC_SX_BLACKLIST):
                        continue
                return source_cols[i]

    return ''

# ================================================================
# STREAMLIT - CẤU HÌNH
# ================================================================
st.set_page_config(page_title="Phú Mỹ Lumen Connector v3.1", page_icon="💊",
                   layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
.step-header{background:#f0f4f8;border-left:4px solid #2e6da4;
    padding:8px 16px;border-radius:4px;margin:20px 0 8px 0;
    font-weight:600;color:#1e3a5f;}
</style>""", unsafe_allow_html=True)
st.markdown("""
<div style="background:linear-gradient(90deg,#1e3a5f,#2e6da4);
     padding:18px 24px;border-radius:10px;margin-bottom:20px;color:white">
<h2 style="margin:0">💊 PHÚ MỸ LUMEN CONNECTOR <span style="font-size:13px;opacity:.7">v3.8</span></h2>
<p style="margin:4px 0 0 0;opacity:.85;font-size:14px">
Module xử lý & ánh xạ danh mục thuốc trúng thầu → Xuất Mẫu 03 BHYT</p>
</div>""", unsafe_allow_html=True)

# Session state
for k, v in {'df_result':None,'mau03_bytes':None,'nha_thau_info':{},
             'df_alias':None,'df_alias_dd':None}.items():
    if k not in st.session_state: st.session_state[k] = v
if st.session_state['df_alias'] is None:
    st.session_state['df_alias'] = load_alias()
if st.session_state['df_alias_dd'] is None:
    st.session_state['df_alias_dd'] = load_alias_dd()

tab_main, tab_alias = st.tabs(["🏠 Xử lý chính", "✏️ Quản lý Alias Hoạt chất"])

# ════════════════════════════════════════════════════
# TAB 1: XỬ LÝ CHÍNH
# ════════════════════════════════════════════════════
with tab_main:

    # ── BƯỚC 1: UPLOAD ──────────────────────────────
    st.markdown('<div class="step-header">📁 Bước 1 — Tải lên 4 file dữ liệu</div>',unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown("**📋 File Danh mục thầu**")
        file_thau  = st.file_uploader("QĐ phê duyệt (.xlsx/.xls)", type=['xlsx','xls'], key='up_thau')
    with c2:
        st.markdown("**💊 File Thuốc tân dược**")
        file_thuoc = st.file_uploader("Thuốc tân dược (.xlsx)",     type=['xlsx'],       key='up_thuoc')
    with c3:
        st.markdown("**🔬 File Cách dùng thuốc**")
        file_cach  = st.file_uploader("Cách dùng (.xls/.xlsx)",     type=['xlsx','xls'], key='up_cach')
    with c4:
        st.markdown("**📄 Template MAU_03**")
        file_mau03 = st.file_uploader("MAU_03.xlsx",                type=['xlsx'],       key='up_mau03')

    if not all([file_thau,file_thuoc,file_cach,file_mau03]):
        st.info("👆 Vui lòng tải lên đủ 4 file để bắt đầu.")
        st.stop()

    # ── BƯỚC 2: ĐỌC FILE ────────────────────────────
    st.markdown('<div class="step-header">⚙️ Bước 2 — Đọc & phân tích dữ liệu</div>',unsafe_allow_html=True)
    # Phát hiện file mới upload → reset col_map để suggest lại đúng
    current_file_id = getattr(file_thau, 'file_id', None) or file_thau.name
    if st.session_state.get('_last_file_id') != current_file_id:
        st.session_state['_last_file_id'] = current_file_id
        # Xóa tất cả key cm_* để force re-suggest khi file mới
        for k in list(st.session_state.keys()):
            if k.startswith('cm_'):
                del st.session_state[k]

    with st.spinner("Đang đọc các file..."):
        thau_bytes  = file_thau.read()
        thuoc_bytes = file_thuoc.read()
        cach_bytes  = file_cach.read()
        mau03_bytes = file_mau03.read()
        st.session_state['mau03_bytes'] = mau03_bytes

        df_thau  = read_excel(thau_bytes,  keywords=('STT','Tên hoạt chất','Số lượng','Nhà thầu'))
        df_thuoc = read_excel(thuoc_bytes, keywords=('MA','TEN','DUONGDUNG'))
        try:
            df_cach = sanitize_df(pd.read_excel(io.BytesIO(cach_bytes), header=0, dtype=str))
        except Exception:
            df_cach = read_xls_compat(cach_bytes)

    thuoc_lk, name_map, goi_y_dd, conflict_map = build_thuoc_lookup(df_thuoc)
    dd_lk                         = build_dd_lookup(df_cach)
    alias_ten_lk                  = build_alias_lookup(st.session_state['df_alias'])
    alias_dd_lk                   = build_alias_dd_lookup(st.session_state['df_alias_dd'])

    r1,r2,r3,r4 = st.columns(4)
    r1.success(f"✅ File thầu: **{len(df_thau)}** dòng")
    r2.success(f"✅ Tân dược: **{len(goi_y_dd)}** hoạt chất")
    r3.success(f"✅ Đường dùng: **{len(dd_lk)}** loại")
    r4.success(f"✅ Alias tên: **{len(st.session_state['df_alias'])}** | ĐD: **{len(st.session_state['df_alias_dd'])}** cặp")

    # ── BƯỚC 3: ÁNH XẠ CỘT ─────────────────────────
    st.markdown('<div class="step-header">🗂️ Bước 3 — Ánh xạ cột file thầu → Template MAU_03</div>',unsafe_allow_html=True)
    st.caption("Hệ thống tự động gợi ý. Điều chỉnh nếu cần.")

    source_cols = df_thau.columns.tolist()
    SKIP = '-- Bỏ qua --'; opts = [SKIP] + source_cols
    MAP_TARGETS = [
        ('TEN_HOAT_CHAT','Tên hoạt chất'),('TEN_THUOC','Tên thuốc'),
        ('DON_VI_TINH','Đơn vị tính'),('HAM_LUONG','Nồng độ/Hàm lượng'),
        ('DUONG_DUNG','Đường dùng'),('DANG_BAO_CHE','Dạng bào chế'),
        ('SO_DANG_KY','Số đăng ký (GĐKLH/GPNK)'),('SO_LUONG','Số lượng'),
        ('DON_GIA','Đơn giá trúng thầu'),('QUY_CACH','Quy cách'),
        ('NHA_SX','Tên cơ sở sản xuất'),('NUOC_SX','Xuất xứ'),
        ('NHA_THAU','Tên nhà thầu'),('NHOM_THUOC','Nhóm thuốc'),('GOI_THAU','Gói thầu'),
    ]

    # ── Hàm tìm cột mặc định (case-insensitive, không phân biệt hoa/thường) ──
    def _find_col(col_list, keywords):
        """Trả về tên cột đầu tiên khớp với bất kỳ keyword nào, hoặc '' nếu không thấy."""
        for kw in keywords:
            kw_l = kw.lower().strip()
            for col in col_list:
                if kw_l in col.lower().strip():
                    return col
        return ''

    # ── Bảng keyword riêng từng target ──
    _COL_PREFER = {
        'NUOC_SX': ['xuất xứ', 'nước sản xuất', 'nuoc_sx', 'origin', 'country'],
        'NHA_SX':  ['tên cơ sở sản xuất', 'cơ sở sản xuất', 'nhà sản xuất', 'manufacturer'],
    }

    # ── Khởi tạo session_state MỘT LẦN DUY NHẤT với GIÁ TRỊ ĐÚNG ──
    # Quy tắc: KHÔNG bao giờ truyền index= vào st.selectbox khi đã có key trong session_state
    # → Chỉ gán session_state, để Streamlit tự đọc từ đó. Không conflict.
    for tgt, _ in MAP_TARGETS:
        ss_key = f'cm_{tgt}'
        if ss_key not in st.session_state:
            # Ưu tiên keyword riêng, fallback sang suggest_col chung
            prefer_kws = _COL_PREFER.get(tgt)
            if prefer_kws:
                sug = _find_col(source_cols, prefer_kws)
            else:
                sug = suggest_col(source_cols, tgt)
            st.session_state[ss_key] = sug if sug else SKIP

    # ── Vẽ selectbox: KHÔNG truyền index= (session_state là nguồn duy nhất) ──
    col_map = {}
    grid = st.columns(4)
    for i, (tgt, label) in enumerate(MAP_TARGETS):
        with grid[i % 4]:
            chosen = st.selectbox(
                f"`{tgt}` ← {label}",
                opts,
                key=f'cm_{tgt}',          # Streamlit tự lấy giá trị từ session_state[key]
                # index= KHÔNG được dùng ở đây → tránh hoàn toàn lỗi conflict
            )
            if chosen != SKIP:
                col_map[tgt] = chosen

    # ── BƯỚC 4: CẤU HÌNH THẦU ───────────────────────
    st.markdown('<div class="step-header">📝 Bước 4 — Thông tin thầu chung</div>',unsafe_allow_html=True)
    cc1,cc2,cc3,cc4 = st.columns(4)
    with cc1:
        so_qd    = st.text_input("Số quyết định", value="449/QĐ-BVĐN")
        ma_cskcb = st.text_input("Mã CSKCB", value="48001")
    with cc2:
        nam_thau = st.text_input("Năm thầu", value="2026")
        goi_mac_dinh = st.text_input("Gói thầu mặc định", value="G1")
    with cc3:
        loai_thuoc = st.text_input("LOAI_THUOC", value="1")
        loai_thau  = st.text_input("LOAI_THAU", value="1")
    with cc4:
        ht_thau = st.text_input("HT_THAU", value="1")

    # ── BƯỚC 5: CHẠY XỬ LÝ ─────────────────────────
    st.markdown('<div class="step-header">🚀 Bước 5 — Chạy xử lý & ánh xạ dữ liệu</div>',unsafe_allow_html=True)

    if st.button("▶️  Chạy xử lý dữ liệu", type="primary", use_container_width=True):
        with st.spinner("Đang xử lý..."):
            df_work = pd.DataFrame()
            for tgt, src in col_map.items():
                if src in df_thau.columns:
                    df_work[tgt] = df_thau[src].apply(
                        lambda x: '' if str(x).strip() in ('nan','NaN','None','') else str(x))
                else:
                    df_work[tgt] = ''

            if 'TEN_HOAT_CHAT' in df_work.columns:
                df_work = df_work[df_work['TEN_HOAT_CHAT'].str.strip() != '']
            df_work = df_work.reset_index(drop=True)

            # Tách SĐK kép
            if 'SO_DANG_KY' in df_work.columns and 'SO_LUONG' in df_work.columns:
                df_work = expand_sdk_rows(df_work, 'SO_DANG_KY', 'SO_LUONG')

            df_work['STT'] = range(1, len(df_work)+1)

            # Tra MA_THUOC + TEN_HOAT_CHAT chuẩn (từ file tân dược)
            result_pairs = df_work.apply(
                lambda r: lookup_thuoc(r.get('TEN_HOAT_CHAT',''), r.get('DUONG_DUNG',''),
                                       thuoc_lk, name_map, goi_y_dd,
                                       alias_ten_lk, alias_dd_lk), axis=1)
            df_work['MA_THUOC']            = result_pairs.apply(lambda x: x[0])
            df_work['TEN_HOAT_CHAT_CHUAN'] = result_pairs.apply(lambda x: x[1])
            # Lưu gợi ý đường dùng dạng chuỗi để hiển thị
            df_work['GOI_Y_DD'] = result_pairs.apply(
                lambda x: ' | '.join([f"{d} → {m}" for d,m,_ in x[2]]) if x[2] else '')
            # TEN_HOAT_CHAT xuất ra là tên chuẩn trong file tân dược (nếu khớp)
            # TEN_HOAT_CHAT_XK: tên để XUẤT FILE
            # Nếu khớp MA_THUOC → dùng tên chuẩn từ file Tân dược
            # Nếu KHÔNG khớp → giữ nguyên tên gốc từ file thầu (KHÔNG tự ý sửa)
            df_work['TEN_HOAT_CHAT_XK'] = df_work.apply(
                lambda r: r['TEN_HOAT_CHAT_CHUAN'] if r['MA_THUOC'] and r['TEN_HOAT_CHAT_CHUAN']
                          else r['TEN_HOAT_CHAT'],
                axis=1)

            # Tra MA_DUONG_DUNG - chỉ khớp chính xác
            dd_series = df_work.get('DUONG_DUNG', pd.Series(['']*len(df_work)))
            df_work['MA_DUONG_DUNG'] = dd_series.apply(lambda x: lookup_dd(x, dd_lk))

            # Đánh dấu mã có đuôi .1/.2 (thanh toán theo tỷ lệ)
            import re as _re
            df_work['IS_MA_CON'] = df_work['MA_THUOC'].apply(
                lambda m: '1' if _re.search(r'\.\d+$', str(m)) and len(str(m).split('.')) >= 3 else '')

            # Đánh dấu hoạt chất có NHIỀU MÃ cùng đường dùng (VĐ4)
            def check_conflict(row):
                hc = _norm(row.get('TEN_HOAT_CHAT',''))
                dd = _norm(row.get('DUONG_DUNG',''))
                # Dùng alias_dd_lk để lấy đường dùng đã chuẩn hóa (nếu có)
                dd_use = _norm(alias_dd_lk.get((hc, dd), dd))
                key = (hc, dd_use)
                mas = conflict_map.get(key, [])
                return ' | '.join(mas) if mas else ''
            df_work['NHIEU_MA'] = df_work.apply(check_conflict, axis=1)

            # Nhóm thuốc + TT_THAU
            df_work['NHOM_MA'] = df_work.get('NHOM_THUOC', pd.Series(['']*len(df_work))).apply(convert_nhom)
            def build_tt(row):
                goi  = str(row.get('GOI_THAU','')).strip() or goi_mac_dinh
                return f"{so_qd};{goi};{row.get('NHOM_MA','')};{nam_thau}"
            df_work['TT_THAU'] = df_work.apply(build_tt, axis=1)

            # DON_GIA_BH = DON_GIA (yêu cầu 1)
            df_work['DON_GIA_BH'] = df_work.get('DON_GIA', pd.Series(['']*len(df_work)))

            # Hằng số
            df_work['MA_CSKCB']       = ma_cskcb
            df_work['MA_CSKCB_THUOC'] = ''   # để trống
            df_work['LOAI_THUOC']     = loai_thuoc
            df_work['LOAI_THAU']      = loai_thau
            df_work['HT_THAU']        = ht_thau

            for c in ('TU_NGAY_HD','DEN_NGAY_HD','TU_NGAY','DEN_NGAY'):
                df_work[c] = ''

            st.session_state['df_result']     = df_work
            st.session_state['nha_thau_info'] = {}

        n_qld  = (df_work.get('IS_QLD_KD', pd.Series()) == 'CẦN RÀ SOÁT').sum()
        n_nodd = (df_work['MA_DUONG_DUNG'] == '').sum()
        n_sub  = (df_work['IS_MA_CON'] == '1').sum()
        st.success(f"✅ Xử lý xong! Tổng **{len(df_work):,}** dòng.")
        if n_qld:  st.warning(f"🟠 **{n_qld}** dòng SĐK dạng QLD-KD — tô cam, cần rà soát thủ công.")
        if n_nodd: st.warning(f"🟡 **{n_nodd}** dòng đường dùng KHÔNG KHỚP bảng chuẩn — tô vàng, cần điều chỉnh thủ công.")
        if n_sub:  st.warning(f"🟣 **{n_sub}** dòng có mã dạng XX.XXX.1/.2 — tô tím, cần xem điều kiện thanh toán và chọn đúng mã.")

    # ── BƯỚC 6: NHÀ THẦU & NGÀY KÝ ─────────────────
    if st.session_state.get('df_result') is not None:
        df_result = st.session_state['df_result'].copy()

        st.markdown('<div class="step-header">🏢 Bước 6 — Quản lý nhà thầu & Ngày ký phụ lục</div>',unsafe_allow_html=True)

        nha_thau_list = sorted([
            x for x in df_result['NHA_THAU'].dropna().unique()
            if str(x).strip() not in ('','nan')
        ])
        st.info(f"📋 Tìm thấy **{len(nha_thau_list)}** nhà thầu.")

        nha_thau_info: dict = st.session_state['nha_thau_info']
        if not nha_thau_info:
            log_data = load_log()
            if log_data:
                n_loaded = sum(1 for nt in log_data if nt in nha_thau_list)
                st.info(f"📂 Đã tải **{n_loaded}** nhà thầu từ mapping_log.csv!")
            nha_thau_info = log_data
        for nt in nha_thau_list:
            if nt not in nha_thau_info:
                nha_thau_info[nt] = {'da_ky': False}

        # ── Điền ngày hàng loạt (ký 1 ngày nhiều công ty) ──
        st.subheader("⚡ Điền ngày nhanh cho nhiều nhà thầu")
        with st.expander("Mở để điền 1 ngày cho nhiều công ty", expanded=False):
            st.caption("Chọn nhà thầu và điền ngày ký + thời hạn một lần duy nhất → áp dụng cho tất cả đã chọn.")
            nt_options = [nt for nt in nha_thau_list if not nha_thau_info.get(nt,{}).get('da_ky')]
            selected_bulk = st.multiselect(
                "Chọn nhà thầu cần điền ngày (nhập tên để tìm nhanh)",
                options=nha_thau_list,
                default=[],
                key='bulk_select'
            )
            b1, b2 = st.columns(2)
            with b1:
                bulk_ngay    = st.date_input("Ngày ký chung", value=date.today(), format="DD/MM/YYYY", key='bulk_ngay')
            with b2:
                bulk_thoi_han = st.number_input("Thời hạn chung (tháng)", min_value=1, max_value=60, value=12, key='bulk_th')
            if st.button("✅ Áp dụng cho các nhà thầu đã chọn", type="primary", use_container_width=True):
                if selected_bulk:
                    den_ngay_bulk = add_months(bulk_ngay, bulk_thoi_han)
                    for nt in selected_bulk:
                        nha_thau_info[nt] = {
                            'da_ky': True, 'ngay_ky': bulk_ngay,
                            'thoi_han': int(bulk_thoi_han), 'den_ngay': den_ngay_bulk
                        }
                    st.session_state['nha_thau_info'] = nha_thau_info
                    st.success(f"✅ Đã cập nhật {len(selected_bulk)} nhà thầu: ngày {bulk_ngay.strftime('%d/%m/%Y')}, {bulk_thoi_han} tháng.")
                    st.rerun()
                else:
                    st.warning("Chưa chọn nhà thầu nào.")

        st.divider()
        # Filter tìm nhà thầu
        search_nt = st.text_input("🔍 Tìm nhà thầu", placeholder="Gõ tên để lọc...", key='search_nt')
        filtered_list = [nt for nt in nha_thau_list if search_nt.lower() in nt.lower()] if search_nt else nha_thau_list

        qa, qb = st.columns(2)
        if qa.button("☑️ Chọn tất cả", use_container_width=True):
            for nt in nha_thau_list: nha_thau_info[nt]['da_ky'] = True
            st.rerun()
        if qb.button("☐ Bỏ chọn tất cả", use_container_width=True):
            for nt in nha_thau_list: nha_thau_info[nt]['da_ky'] = False
            st.rerun()

        st.caption(f"Đang hiển thị {len(filtered_list)}/{len(nha_thau_list)} nhà thầu")
        left_col, right_col = st.columns(2)
        for idx, nt in enumerate(filtered_list):
            info = nha_thau_info.get(nt, {'da_ky': False})
            icon = '✅' if info.get('da_ky') else '⬜'
            with (left_col if idx%2==0 else right_col).expander(f"{icon} {nt}", expanded=False):
                da_ky = st.checkbox("Đã ký phụ lục", key=f"ck_{nt}", value=info.get('da_ky', False))
                if da_ky:
                    i1, i2 = st.columns(2)
                    with i1:
                        ngay_ky = st.date_input("Ngày ký", value=info.get('ngay_ky', date.today()),
                                                 format="DD/MM/YYYY", key=f"ngay_{nt}")
                    with i2:
                        thoi_han = st.number_input("Thời hạn (tháng)", min_value=1, max_value=60,
                                                    value=info.get('thoi_han', 12), key=f"th_{nt}")
                    den_ngay = add_months(ngay_ky, thoi_han)
                    st.caption(f"📅 {ngay_ky.strftime('%d/%m/%Y')} → {den_ngay.strftime('%d/%m/%Y')} | `{to_yyyymmdd(ngay_ky)}` → `{to_yyyymmdd(den_ngay)}`")
                    nha_thau_info[nt] = {'da_ky':True,'ngay_ky':ngay_ky,'thoi_han':int(thoi_han),'den_ngay':den_ngay}
                else:
                    nha_thau_info[nt] = {'da_ky': False}

        st.session_state['nha_thau_info'] = nha_thau_info

        def get_ngay(nt, col):
            info = nha_thau_info.get(nt, {})
            if not info.get('da_ky'): return ''
            return to_yyyymmdd(info.get('ngay_ky') if 'TU' in col else info.get('den_ngay'))

        # TU_NGAY_HD / DEN_NGAY_HD: điền ngày hợp đồng
        # TU_NGAY: = TU_NGAY_HD
        # DEN_NGAY: theo quy định BHYT = để TRỐNG
        for col_n in ('TU_NGAY_HD','DEN_NGAY_HD','TU_NGAY'):
            df_result[col_n] = df_result['NHA_THAU'].apply(lambda x, c=col_n: get_ngay(x, c))
        df_result['DEN_NGAY'] = ''   # Để trống theo quy định BHYT
        st.session_state['df_result'] = df_result

        da_ky_count = sum(1 for i in nha_thau_info.values() if i.get('da_ky'))
        st.info(f"✅ **{da_ky_count}/{len(nha_thau_list)}** nhà thầu đã ký phụ lục.")

        # ── BƯỚC 7: PREVIEW ─────────────────────────
        st.markdown('<div class="step-header">👁️ Bước 7 — Xem trước kết quả</div>',unsafe_allow_html=True)

        PREV = ['STT','MA_THUOC','IS_MA_CON','NHIEU_MA','TEN_HOAT_CHAT','TEN_HOAT_CHAT_XK',
                'TEN_THUOC','SO_DANG_KY','IS_QLD_KD',
                'DUONG_DUNG','MA_DUONG_DUNG','GOI_Y_DD','SO_LUONG','DON_GIA',
                'NHOM_MA','NHA_THAU','TT_THAU','TU_NGAY_HD','DEN_NGAY_HD']
        prev_cols = [c for c in PREV if c in df_result.columns]

        def highlight(row):
            no_ma   = not str(row.get('MA_THUOC','')).strip()
            is_qld  = str(row.get('IS_QLD_KD','')) == 'CẦN RÀ SOÁT'
            no_dd   = not str(row.get('MA_DUONG_DUNG','')).strip() and str(row.get('DUONG_DUNG','')).strip()
            is_sub  = str(row.get('IS_MA_CON','')) == '1'
            if no_ma:     c = '#FFDDDD'
            elif is_qld:  c = '#FFE5CC'
            elif no_dd:   c = '#FFFACC'
            elif is_sub:  c = '#E8D5FF'
            else:         c = ''
            return [f'background-color:{c}' for _ in row]

        st.dataframe(df_result[prev_cols].style.apply(highlight, axis=1),
                     use_container_width=True, height=420)

        total   = len(df_result)
        miss_ma = (df_result['MA_THUOC']=='').sum()
        n_qld   = (df_result.get('IS_QLD_KD', pd.Series())=='CẦN RÀ SOÁT').sum()
        n_nodd  = (df_result['MA_DUONG_DUNG']=='').sum() - miss_ma  # trừ dòng đã đỏ
        n_nodd  = max(0, (df_result['MA_DUONG_DUNG']=='').sum())
        n_sub   = (df_result.get('IS_MA_CON', pd.Series())=='1').sum()

        n_multi = (df_result.get('NHIEU_MA', pd.Series()) != '').sum()
        m1,m2,m3,m4,m5,m6 = st.columns(6)
        m1.metric("📊 Tổng",      f"{total:,}")
        m2.metric("✅ Có mã",     f"{total-miss_ma:,}")
        m3.metric("🔴 Thiếu mã",  f"{miss_ma:,}")
        m4.metric("🟠 QLD-KD",    f"{n_qld:,}")
        m5.metric("🟣 Mã .1/.2",  f"{n_sub:,}")
        m6.metric("⚠️ Nhiều mã",  f"{n_multi:,}")

        # Chọn mã cho hoạt chất có nhiều mã cùng đường dùng
        if n_multi > 0:
            # Lấy map MIEUTA 1 lần
            mieuta_map = {sc(r.get('MA','')): sc(r.get('MIEUTA',''))
                          for _,r in df_thuoc.iterrows()}
            with st.expander(
                f"⚠️ **{n_multi} dòng có NHIỀU MÃ cùng đường dùng** — Chọn mã đúng",
                expanded=True
            ):
                multi_rows = df_result[df_result.get('NHIEU_MA','')!=''][
                    ['TEN_HOAT_CHAT','DUONG_DUNG','MA_THUOC','NHIEU_MA']
                ].drop_duplicates(subset=['TEN_HOAT_CHAT','DUONG_DUNG']).reset_index(drop=True)

                if 'ma_chon_override' not in st.session_state:
                    st.session_state['ma_chon_override'] = {}

                # Build MA → list[(dd,ten)] từ df_thuoc — KHÔNG dedup theo MA
                # (cùng mã có thể nhiều đường dùng — giữ hết)
                ma_dd_list_map = {}   # MA → [(dd_raw, ten_raw), ...]
                for _, td_r in df_thuoc.iterrows():
                    ma_v  = sc(td_r.get('MA',''))
                    dd_v  = sc(td_r.get('DUONGDUNG',''))
                    ten_v = sc(td_r.get('TEN',''))
                    if ma_v:
                        ma_dd_list_map.setdefault(ma_v, [])
                        pair = (dd_v, ten_v)
                        if pair not in ma_dd_list_map[ma_v]:
                            ma_dd_list_map[ma_v].append(pair)

                for multi_idx, (_, mr) in enumerate(multi_rows.iterrows()):
                    hc_key  = f"{mr['TEN_HOAT_CHAT']}||{mr['DUONG_DUNG']}"
                    all_mas = mr['NHIEU_MA'].split(' | ')
                    dd_thau = mr['DUONG_DUNG']

                    # Xây danh sách (ma, dd, ten) — mỗi cặp (MA,DD) là 1 dòng riêng
                    # Sắp xếp: exact-name match lên đầu, rồi theo dd_score
                    hc_tight_mr = _norm_tight(mr['TEN_HOAT_CHAT'])
                    cand_rows = []   # [(dd_score, ma, dd, ten)]
                    seen_pairs = set()
                    for m in all_mas:
                        for dd_v, ten_v in ma_dd_list_map.get(m, [('','')]):
                            pk = (m, _norm_tight(dd_v))
                            if pk in seen_pairs:
                                continue
                            seen_pairs.add(pk)
                            score = _dd_similarity(dd_thau, dd_v)
                            cand_rows.append((score, m, dd_v, ten_v))
                    cand_rows.sort(key=lambda x: x[0], reverse=True)

                    # Xây label list — dùng INDEX để tránh trùng khi cùng MA
                    multi_labels = []
                    for score, m, dd_v, ten_v in cand_rows:
                        lbl = f"[{m}] — {ten_v} — {dd_v}" if dd_v else f"[{m}] — {ten_v}"
                        multi_labels.append(lbl)

                    xc1, xc2 = st.columns([3, 5])
                    with xc1:
                        st.markdown(f"**`{mr['TEN_HOAT_CHAT']}`** | ĐD file thầu: `{mr['DUONG_DUNG']}`")
                    with xc2:
                        if cand_rows:
                            chosen_i = st.selectbox(
                                "Chọn mã",
                                options=list(range(len(cand_rows))),
                                format_func=lambda i, lbs=multi_labels: lbs[i],
                                key=f"sel_ma_{multi_idx}_{mr['TEN_HOAT_CHAT'][:15]}",
                                label_visibility='collapsed'
                            )
                            if st.button("✅ Xác nhận mã này",
                                         key=f"multi_ok_{multi_idx}_{mr['TEN_HOAT_CHAT'][:15]}",
                                         use_container_width=True):
                                _, chosen_ma, chosen_dd, chosen_ten = cand_rows[chosen_i]
                                _df = st.session_state['df_result']
                                mask = (
                                    (_df['TEN_HOAT_CHAT'] == mr['TEN_HOAT_CHAT']) &
                                    (_df['DUONG_DUNG'] == mr['DUONG_DUNG'])
                                )
                                _df.loc[mask, 'MA_THUOC']          = chosen_ma
                                _df.loc[mask, 'TEN_HOAT_CHAT_XK']  = chosen_ten if chosen_ten else mr['TEN_HOAT_CHAT']
                                _df.loc[mask, 'DUONG_DUNG']         = chosen_dd if chosen_dd else mr['DUONG_DUNG']
                                _df.loc[mask, 'MA_DUONG_DUNG']      = lookup_dd(chosen_dd, dd_lk) if chosen_dd else ''
                                _df.loc[mask, 'NHIEU_MA']           = ''   # xóa flag nhiều mã
                                st.session_state['df_result'] = _df
                                st.session_state['ma_chon_override'][hc_key] = chosen_ma
                                st.success(f"✅ Đã gán **{chosen_ma}** — {chosen_ten} — {chosen_dd}")
                                st.rerun()
                        else:
                            st.caption("❌ Không tìm thấy thông tin trong file Tân dược")
                    st.divider()

        # Mã .1/.2 đã được cảnh báo qua metric 🟣, không cần block riêng

        # ── BẢNG XÁC NHẬN ALIAS & ĐƯỜNG DÙNG ──────────
        # Hiện thị TRỰC TIẾP trên tab chính để user xem và xác nhận
        if miss_ma > 0 or n_nodd > 0:
            st.markdown('<div class="step-header">✋ Cần xác nhận trước khi xuất file</div>',
                        unsafe_allow_html=True)

        if miss_ma > 0:
            with st.expander(f"🔴 **{miss_ma} dòng thiếu MA_THUOC** — Xem & xử lý ngay", expanded=True):
                import difflib as _dl_miss

                miss_rows = df_result[df_result['MA_THUOC']==''][
                    ['TEN_HOAT_CHAT','DUONG_DUNG']
                ].drop_duplicates(subset=['TEN_HOAT_CHAT','DUONG_DUNG']).reset_index(drop=True)

                if 'ma_chon_manual' not in st.session_state:
                    st.session_state['ma_chon_manual'] = {}

                df_alias_cur = st.session_state['df_alias'].copy()
                alias_changed = False

                def _find_td_candidates(hc_m: str, dd_m: str):
                    """
                    Tìm ứng viên từ file Tân dược — v3.7 (Exact + Contains).

                    Thứ tự ưu tiên:
                    1. Exact name match (tight-norm) → nhóm A (ưu tiên cao nhất)
                    2. Contains match: tên thầu (bỏ *) chứa trong tên TD, hoặc ngược lại
                       → nhóm B (hiển thị sau nhóm A)
                    Trong mỗi nhóm: sắp xếp theo dd_score giảm dần.
                    Dedup theo (MA, DD_tight).
                    ⭐ khi nhóm A VÀ dd_score >= 0.8.
                    """
                    hc_tight = _norm_tight(hc_m)
                    # Bỏ ký tự * để so sánh contains (Acid amin* → acid amin)
                    hc_clean = hc_tight.rstrip('*').strip()
                    n_comp   = _count_components(hc_m)

                    exact_rows   = []   # nhóm A: tên khớp 100%
                    contains_rows = []  # nhóm B: tên chứa keyword
                    seen_ma_dd   = set()

                    for _, td_r in df_thuoc.iterrows():
                        ten_td = sc(td_r.get('TEN', ''))
                        dd_td  = sc(td_r.get('DUONGDUNG', ''))
                        ma_td  = sc(td_r.get('MA', ''))
                        if not ma_td:
                            continue

                        ten_td_t = _norm_tight(ten_td)
                        ten_td_clean = ten_td_t.rstrip('*').strip()

                        # Số thành phần phải khớp
                        if _count_components(ten_td) != n_comp:
                            continue

                        # Dedup (MA, DD)
                        pair_key = (ma_td, _norm_tight(dd_td))
                        if pair_key in seen_ma_dd:
                            continue

                        dd_score = _dd_similarity(dd_m, dd_td)

                        if ten_td_t == hc_tight or ten_td_clean == hc_clean:
                            # Nhóm A: exact match (kể cả có/không có dấu *)
                            seen_ma_dd.add(pair_key)
                            star_eligible = dd_score >= 0.8
                            exact_rows.append((dd_score, ma_td, dd_td, ten_td, star_eligible))
                        elif hc_clean and (hc_clean in ten_td_t or ten_td_clean in hc_tight):
                            # Nhóm B: contains — 'acid amin' in 'acid amin*' hoặc ngược lại
                            seen_ma_dd.add(pair_key)
                            contains_rows.append((dd_score, ma_td, dd_td, ten_td, False))

                    exact_rows.sort(key=lambda x: x[0], reverse=True)
                    contains_rows.sort(key=lambda x: x[0], reverse=True)
                    return exact_rows + contains_rows

                for row_idx, (_, mrow) in enumerate(miss_rows.iterrows()):
                    hc_m   = mrow['TEN_HOAT_CHAT']
                    dd_m   = mrow['DUONG_DUNG']
                    hc_key = f"{hc_m}||{dd_m}"
                    # Key duy nhất cho mỗi dòng — tránh conflict session_state
                    sel_key     = f"miss_idx_{row_idx}"
                    sel_key_ss  = f"miss_sel_idx_{row_idx}"   # lưu index đã chọn

                    td_candidates = _find_td_candidates(hc_m, dd_m)
                    already_alias = (df_alias_cur['TEN_THAU'].str.strip().str.lower()==hc_m.strip().lower()).any()

                    with st.container():
                        rc1, rc2 = st.columns([4, 6])
                        with rc1:
                            st.markdown(f"**🔴 `{hc_m}`**")
                            st.caption(f"ĐD file thầu: `{dd_m or '—'}`")

                        with rc2:
                            if td_candidates:
                                # Xây label list — mỗi phần tử là chuỗi hiển thị đầy đủ
                                # Dùng INDEX (0,1,2...) làm option để tránh trùng khi
                                # nhiều dòng có cùng MA nhưng khác đường dùng
                                idx_labels = []
                                for i, (dd_score, ma_td, dd_td, ten_td, is_exact) in enumerate(td_candidates):
                                    star = "⭐ " if (is_exact and dd_score >= 0.8) else ""
                                    # Label: [Mã] — Tên gốc — Đường dùng gốc (không sửa)
                                    idx_labels.append(f"{star}[{ma_td}] — {ten_td} — {dd_td}")

                                # Khởi tạo session_state cho index đã chọn (chỉ lần đầu)
                                if sel_key_ss not in st.session_state:
                                    st.session_state[sel_key_ss] = 0

                                # Selectbox dùng index integer, format_func tra label từ list
                                chosen_idx = st.selectbox(
                                    "Chọn mã",
                                    options=list(range(len(td_candidates))),
                                    format_func=lambda i, lbs=idx_labels: lbs[i],
                                    key=sel_key,
                                    label_visibility='collapsed'
                                )

                                if st.button("✅ Gán mã này", key=f"miss_ok_{row_idx}",
                                             use_container_width=True):
                                    # Lấy thông tin từ td_candidates[chosen_idx]
                                    _, chosen_ma, chosen_dd_td, chosen_ten_td, _ = td_candidates[chosen_idx]
                                    # ten_td và dd_td đã là giá trị GỐC từ file Tân dược
                                    _df = st.session_state['df_result']
                                    mask = ((_df['TEN_HOAT_CHAT'] == hc_m) & (_df['DUONG_DUNG'] == dd_m))
                                    _df.loc[mask, 'MA_THUOC']          = chosen_ma
                                    _df.loc[mask, 'TEN_HOAT_CHAT_XK'] = chosen_ten_td
                                    _df.loc[mask, 'DUONG_DUNG']        = chosen_dd_td
                                    _df.loc[mask, 'MA_DUONG_DUNG']     = lookup_dd(chosen_dd_td, dd_lk)
                                    st.session_state['df_result'] = _df
                                    st.session_state['ma_chon_manual'][hc_key] = chosen_ma
                                    st.success(f"✅ Đã gán **{chosen_ma}** — {chosen_ten_td} — {chosen_dd_td}")
                                    st.rerun()
                            else:
                                st.caption("❌ Không tìm thấy trong file Tân dược")
                                if not already_alias:
                                    td_inp = st.text_input(
                                        "Tên đúng trong Tân dược",
                                        key=f"mis_td_{row_idx}_{hc_m[:15]}",
                                        label_visibility='collapsed',
                                        placeholder="Nhập tên chuẩn trong file Tân dược..."
                                    )
                                    if st.button("➕ Lưu alias tên", key=f"mis_add_{row_idx}_{hc_m[:15]}",
                                                 use_container_width=True):
                                        if td_inp.strip():
                                            new_row = pd.DataFrame([{
                                                'TEN_THAU':  hc_m.strip(),
                                                'TEN_TANDUC':td_inp.strip(),
                                                'GHI_CHU':   'Xác nhận trực tiếp'
                                            }])
                                            df_alias_cur = pd.concat(
                                                [df_alias_cur, new_row], ignore_index=True)
                                            alias_changed = True
                                        else:
                                            st.warning("Chưa nhập tên")
                                else:
                                    st.caption("✅ Đã có alias tên, chạy lại để áp dụng")
                        st.divider()

                if alias_changed:
                    st.session_state['df_alias'] = df_alias_cur
                    save_alias(df_alias_cur)
                    st.success("✅ Đã lưu alias. Bấm **'▶️ Chạy xử lý lại'** để áp dụng.")
                    st.rerun()

        if n_nodd > 0:
            nodd_mask = (df_result['MA_DUONG_DUNG']=='') & (df_result['DUONG_DUNG'].str.strip()!='')
            nodd_rows = df_result[nodd_mask][
                ['TEN_HOAT_CHAT','DUONG_DUNG']
            ].drop_duplicates().reset_index(drop=True)

            with st.expander(
                f"🟡 **{len(nodd_rows)} đường dùng không khớp bảng chuẩn** — Xem & xác nhận",
                expanded=True
            ):
                st.caption(
                    "Đường dùng file thầu **không có trong bảng BHYT**. "
                    "Chọn [Mã] - Tên - ĐD từ Tân dược → bấm 💾 để ánh xạ. "
                    "Lần sau sẽ tự điền."
                )
                df_dd_alias_cur = st.session_state['df_alias_dd'].copy()
                dd_changed = False

                for nodd_idx, (_, nrow) in enumerate(nodd_rows.iterrows()):
                    hc_n   = nrow['TEN_HOAT_CHAT']
                    dd_n   = nrow['DUONG_DUNG']
                    hc_tight = _norm_tight(hc_n)
                    dd_tight = _norm_tight(dd_n)
                    n_comp_n = _count_components(hc_n)

                    # Kiểm tra đã có trong alias_dd chưa
                    dd_already_mask = (
                        (df_dd_alias_cur['TEN_HOAT_CHAT'].apply(_norm_tight)==hc_tight) &
                        (df_dd_alias_cur['DD_GOC'].apply(_norm_tight)==dd_tight)
                    )
                    dd_already      = dd_already_mask.any()
                    dd_chuan_cur    = df_dd_alias_cur[dd_already_mask]['DD_CHUAN'].values[0] if dd_already else ''

                    # Tìm ứng viên từ Tân dược: exact tên (hoặc first-token)
                    # + cùng số thành phần + sắp xếp theo dd_score
                    td_cands_nodd = []   # [(dd_score, ma, dd_td, ten_td)]
                    seen_nodd_pairs = set()
                    hc_clean_n = hc_tight.rstrip('*').strip()
                    nodd_exact, nodd_contains = [], []
                    for _, td_r in df_thuoc.iterrows():
                        ten_td  = sc(td_r.get('TEN',''))
                        dd_td   = sc(td_r.get('DUONGDUNG',''))
                        ma_td   = sc(td_r.get('MA',''))
                        if not ma_td:
                            continue
                        ten_td_t = _norm_tight(ten_td)
                        ten_td_clean = ten_td_t.rstrip('*').strip()
                        if _count_components(ten_td) != n_comp_n:
                            continue
                        pair_key_nodd = (ma_td, _norm_tight(dd_td))
                        if pair_key_nodd in seen_nodd_pairs:
                            continue
                        dd_score = _dd_similarity(dd_n, dd_td)
                        if ten_td_t == hc_tight or ten_td_clean == hc_clean_n:
                            seen_nodd_pairs.add(pair_key_nodd)
                            nodd_exact.append((dd_score, ma_td, dd_td, ten_td))
                        elif hc_clean_n and (hc_clean_n in ten_td_t or ten_td_clean in hc_tight):
                            seen_nodd_pairs.add(pair_key_nodd)
                            nodd_contains.append((dd_score, ma_td, dd_td, ten_td))
                    nodd_exact.sort(key=lambda x: x[0], reverse=True)
                    nodd_contains.sort(key=lambda x: x[0], reverse=True)
                    td_cands_nodd = nodd_exact + nodd_contains

                    td_cands_nodd.sort(key=lambda x: x[0], reverse=True)

                    # Tạo options selectbox: [Mã] - Tên - ĐD
                    nodd_ma_opts, nodd_ma_labels = [], []
                    for dd_score, ma_td, dd_td, ten_td in td_cands_nodd:
                        star = "⭐ " if dd_score >= 0.8 else ""
                        lbl = f"{star}[{ma_td}] - {ten_td} - {dd_td}"
                        nodd_ma_opts.append(ma_td)
                        nodd_ma_labels.append(lbl)

                    # Fallback: nếu không tìm được gì → dùng list đường dùng chung
                    if not nodd_ma_opts:
                        dd_all = sorted(df_thuoc['DUONGDUNG'].dropna().unique().tolist())
                        dd_all_sel_key = f"dd_fallback_{nodd_idx}_{hc_n[:15]}"

                    with st.container():
                        dc1, dc2, dc3 = st.columns([3, 5, 1])
                        with dc1:
                            st.markdown(f"**🟡 `{hc_n}`**")
                            st.caption(f"ĐD file thầu: `{dd_n or '—'}`")
                        with dc2:
                            if dd_already:
                                st.success(f"✅ Đã ánh xạ → **`{dd_chuan_cur}`**")
                                if st.button("🔄 Chọn lại", key=f"dd_reset_{nodd_idx}_{hc_n[:15]}"):
                                    df_dd_alias_cur = df_dd_alias_cur[~dd_already_mask].reset_index(drop=True)
                                    st.session_state['df_alias_dd'] = df_dd_alias_cur
                                    save_alias_dd(df_dd_alias_cur)
                                    st.rerun()
                            elif nodd_ma_opts:
                                # Có ứng viên từ Tân dược → hiện selectbox [Mã]-Tên-ĐD
                                # Dùng index integer để tránh trùng label khi cùng MA
                                sel_nodd_chosen_idx = st.selectbox(
                                    "Chọn mã",
                                    options=list(range(len(nodd_ma_opts))),
                                    format_func=lambda i, lbs=nodd_ma_labels: lbs[i],
                                    key=f"dd_sel_ma_{nodd_idx}_{hc_n[:15]}",
                                    label_visibility='collapsed',
                                )
                                sel_nodd_ma = nodd_ma_opts[sel_nodd_chosen_idx]
                            else:
                                # Fallback: chọn đường dùng thô
                                dd_all = sorted(df_thuoc['DUONGDUNG'].dropna().unique().tolist())
                                sel_nodd_ma = None
                                sel_dd_raw = st.selectbox(
                                    "Chọn đường dùng chuẩn",
                                    ['-- Chọn đường dùng chuẩn --'] + dd_all,
                                    key=f"dd_fallback_{nodd_idx}_{hc_n[:15]}",
                                    label_visibility='collapsed',
                                )
                        with dc3:
                            if not dd_already:
                                if st.button("💾", key=f"dd_save_{nodd_idx}_{hc_n[:15]}",
                                             use_container_width=True,
                                             help="Lưu ánh xạ đường dùng"):
                                    if nodd_ma_opts:
                                        # Lấy DD chuẩn từ td_cands_nodd theo index đã chọn
                                        _cand = td_cands_nodd[sel_nodd_chosen_idx]
                                        dd_chuan_save = _cand[2]   # dd_td raw
                                        ma_save = _cand[1]         # ma_td
                                    else:
                                        dd_chuan_save = sel_dd_raw if sel_dd_raw != '-- Chọn đường dùng chuẩn --' else ''
                                        ma_save = None
                                    if dd_chuan_save:
                                        # Cập nhật NGAY vào df_result (không chờ rerun)
                                        _df = st.session_state['df_result']
                                        mask_n = (_df['TEN_HOAT_CHAT']==hc_n) & (_df['DUONG_DUNG']==dd_n)
                                        if mask_n.any():
                                            _df.loc[mask_n, 'DUONG_DUNG']    = dd_chuan_save.strip()
                                            _df.loc[mask_n, 'MA_DUONG_DUNG'] = lookup_dd(dd_chuan_save, dd_lk)
                                            # Nếu chọn từ Tân dược → cập nhật luôn MA_THUOC và TEN
                                            if nodd_ma_opts and ma_save:
                                                _cand_save = td_cands_nodd[sel_nodd_chosen_idx]
                                                ten_chuan_save = _cand_save[3]  # ten_td raw
                                                _df.loc[mask_n, 'MA_THUOC']          = ma_save
                                                _df.loc[mask_n, 'TEN_HOAT_CHAT_XK'] = ten_chuan_save
                                            st.session_state['df_result'] = _df
                                        # Ghi alias để lần sau tham khảo (không tự điền)
                                        new_row = pd.DataFrame([{
                                            'TEN_HOAT_CHAT': hc_n.strip(),
                                            'DD_GOC':        dd_n.strip(),
                                            'DD_CHUAN':      dd_chuan_save.strip(),
                                            'GHI_CHU':       'Xác nhận trực tiếp'
                                        }])
                                        df_dd_alias_cur = pd.concat([df_dd_alias_cur, new_row], ignore_index=True)
                                        dd_changed = True
                                        st.success(f"✅ Đã lưu: {hc_n} — {dd_chuan_save}")
                                    else:
                                        st.warning("Chưa chọn")
                        st.divider()

                if dd_changed:
                    st.session_state['df_alias_dd'] = df_dd_alias_cur
                    save_alias_dd(df_dd_alias_cur)
                    st.rerun()

        # ── BƯỚC 8: XUẤT FILE ───────────────────────
        st.markdown('<div class="step-header">📥 Bước 8 — Xuất file MAU_03</div>',unsafe_allow_html=True)

        def prepare_export(filter_nt=None):
            # LUÔN lấy từ session_state để đảm bảo mọi chỉnh sửa đã được áp dụng
            df_out = st.session_state['df_result'].copy()

            if filter_nt is not None:
                df_out = df_out[df_out['NHA_THAU'].isin(filter_nt)]
            # TEN_HOAT_CHAT xuất ra dùng tên chuẩn file tân dược
            if 'TEN_HOAT_CHAT_XK' in df_out.columns:
                df_out['TEN_HOAT_CHAT'] = df_out['TEN_HOAT_CHAT_XK']
            for col in MAU03_COLS:
                if col not in df_out.columns: df_out[col] = ''
            return df_out[MAU03_COLS + ['IS_QLD_KD','IS_MA_CON']].reset_index(drop=True)

        mau03_bytes_val = st.session_state['mau03_bytes']
        nha_thau_da_ky  = [nt for nt,info in nha_thau_info.items() if info.get('da_ky')]
        save_log(nha_thau_info)

        btn1, btn2 = st.columns(2)
        with btn1:
            st.markdown("""
**📌 File Ánh xạ** *(nạp lên cổng BHYT)*
- Chỉ nhà thầu đã ký ✅
- 🔴 Thiếu MA_THUOC | 🟠 QLD-KD | 🟡 Đường dùng lạ | 🟣 Mã .1/.2 cần chọn
            """)
            if nha_thau_da_ky:
                df_ax = prepare_export(filter_nt=nha_thau_da_ky)
                st.download_button(
                    f"⬇️ Tải file Ánh xạ ({len(df_ax):,} dòng / {len(nha_thau_da_ky)} NCC)",
                    export_mau03(df_ax, mau03_bytes_val),
                    f"AnhXa_MAU03_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary")
            else:
                st.warning("⚠️ Chưa có nhà thầu nào được tick 'Đã ký phụ lục'.")

        with btn2:
            st.markdown("""
**📊 File Theo dõi** *(báo cáo hối thúc)*
- Toàn bộ danh mục, kể cả chưa ký
            """)
            df_td = prepare_export(filter_nt=None)
            st.download_button(
                f"⬇️ Tải file Theo dõi ({len(df_td):,} dòng / {len(nha_thau_list)} NCC)",
                export_mau03(df_td, mau03_bytes_val),
                f"TheoDoi_MAU03_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        # Phụ lục
        with st.expander("📋 Bảng tổng hợp nhà thầu", expanded=False):
            rows_nt = [{'Nhà thầu':nt,'Đã ký':'✅' if i.get('da_ky') else '⬜',
                        'Ngày ký':i['ngay_ky'].strftime('%d/%m/%Y') if i.get('ngay_ky') else '',
                        'TU_NGAY_HD':to_yyyymmdd(i.get('ngay_ky')),
                        'Hết hạn':i['den_ngay'].strftime('%d/%m/%Y') if i.get('den_ngay') else '',
                        'DEN_NGAY_HD':to_yyyymmdd(i.get('den_ngay')),
                        'Thời hạn':i.get('thoi_han',''),
                        'Số dòng':int((df_result['NHA_THAU']==nt).sum())}
                       for nt,i in [(nt,nha_thau_info.get(nt,{})) for nt in nha_thau_list]]
            st.dataframe(pd.DataFrame(rows_nt), use_container_width=True, hide_index=True)

        with st.expander("🔬 Bảng mã đường dùng chuẩn", expanded=False):
            st.dataframe(pd.DataFrame(sorted(dd_lk.items()), columns=['Tên đường dùng','Mã']),
                         use_container_width=True, hide_index=True)


# ════════════════════════════════════════════════════
# TAB 2: ALIAS
# ════════════════════════════════════════════════════
with tab_alias:
    st.markdown("""
### ✏️ Quản lý Alias — Ánh xạ tên hoạt chất & đường dùng

Alias giúp hệ thống tự nhận ra những khác biệt giữa file thầu và file Tân dược.
Có **2 loại alias** lưu riêng biệt:
- **Alias Tên**: tên hoạt chất file thầu ≠ tên trong file Tân dược
- **Alias Đường dùng**: đường dùng file thầu không có trong bảng chuẩn BHYT
""")

    subtab_ten, subtab_dd = st.tabs(["📝 Alias Tên hoạt chất", "🔬 Alias Đường dùng"])

    # ── SUB-TAB 1: ALIAS TÊN ──────────────────────────
    with subtab_ten:
        st.caption("**Mục đích:** Tên file thầu khác tên Tân dược (dấu cách, viết tắt...) → thêm alias để tự khớp.")
        df_alias = st.session_state['df_alias'].copy()

    st.markdown("#### 📋 Alias hiện có")
    if df_alias.empty:
        st.info("Chưa có alias nào.")
    else:
        st.dataframe(df_alias, use_container_width=True, hide_index=True)
        with st.expander("🗑️ Xóa alias"):
            del_idx = st.number_input("Số thứ tự dòng cần xóa (từ 0)",
                                       min_value=0, max_value=max(0,len(df_alias)-1), value=0)
            if st.button("Xóa dòng"):
                df_alias = df_alias.drop(index=del_idx).reset_index(drop=True)
                st.session_state['df_alias'] = df_alias
                save_alias(df_alias); st.rerun()

    st.divider()
    st.markdown("#### ➕ Thêm alias mới")
    a1,a2,a3,a4 = st.columns([3,3,2,2])
    with a1: inp_thau = st.text_input("Tên trong file THẦU", placeholder="Magnesi aspartat + kali aspartat")
    with a2: inp_td   = st.text_input("Tên ĐÚNG trong file Tân dược", placeholder="Magnesi aspartat+ kali aspartat")
    with a3: inp_dd   = st.text_input("Đường dùng (trống=tất cả)", placeholder="Uống")
    with a4: inp_note = st.text_input("Ghi chú", placeholder="Lệch dấu cách")

    if st.button("➕ Thêm alias", type="primary", use_container_width=True):
        if not inp_thau.strip() or not inp_td.strip():
            st.error("❌ Phải nhập đủ cả 2 trường tên.")
        elif (df_alias['TEN_THAU'].str.strip().str.lower() == inp_thau.strip().lower()).any():
            st.warning("⚠️ Alias đã tồn tại.")
        else:
            new_row = pd.DataFrame([{'TEN_THAU':inp_thau.strip(),'TEN_TANDUC':inp_td.strip(),
                                      'DUONG_DUNG':inp_dd.strip(),'GHI_CHU':inp_note.strip()}])
            df_alias = pd.concat([df_alias, new_row], ignore_index=True)
            st.session_state['df_alias'] = df_alias
            save_alias(df_alias)
            st.success(f"✅ Đã thêm: **{inp_thau}** → **{inp_td}**. Bấm 'Chạy xử lý lại' để áp dụng.")
            st.rerun()

    # Gợi ý từ kết quả hiện tại
    if st.session_state.get('df_result') is not None:
        df_r = st.session_state['df_result']
        missing_hc = df_r[df_r['MA_THUOC']==''][['TEN_HOAT_CHAT','DUONG_DUNG']].drop_duplicates().reset_index(drop=True)
        if not missing_hc.empty:
            st.divider()
            st.markdown("#### 🔍 Hoạt chất chưa khớp từ lần xử lý gần nhất")
            st.caption("Điền tên đúng trong file Tân dược → bấm ➕ để thêm alias nhanh.")
            for _, miss_row in missing_hc.iterrows():
                hc_miss = miss_row['TEN_HOAT_CHAT']; dd_miss = miss_row['DUONG_DUNG']
                with st.container():
                    st.markdown(f"**🔴 `{hc_miss}`** | Đường dùng: `{dd_miss or '—'}`")
                    qa2,qb2,qc2 = st.columns([4,2,1])
                    with qa2:
                        td_val = st.text_input("Tên đúng trong Tân dược", value='',
                                                key=f"ias_td_{hc_miss}_{dd_miss}",
                                                label_visibility='collapsed',
                                                placeholder="Nhập tên chuẩn...")
                    with qb2:
                        dd_val = st.text_input("Đường dùng", value=dd_miss,
                                                key=f"ias_dd_{hc_miss}_{dd_miss}",
                                                label_visibility='collapsed')
                    with qc2:
                        if st.button("➕", key=f"btn_ias_{hc_miss}_{dd_miss}", use_container_width=True):
                            if td_val.strip():
                                exists = (df_alias['TEN_THAU'].str.strip().str.lower()==hc_miss.strip().lower()).any()
                                if not exists:
                                    new_row = pd.DataFrame([{'TEN_THAU':hc_miss.strip(),'TEN_TANDUC':td_val.strip(),
                                                              'DUONG_DUNG':dd_val.strip(),'GHI_CHU':'Từ danh sách chưa khớp'}])
                                    df_alias = pd.concat([df_alias, new_row], ignore_index=True)
                                    st.session_state['df_alias'] = df_alias
                                    save_alias(df_alias)
                                    st.success(f"✅ Đã thêm alias cho **{hc_miss}**"); st.rerun()
                                else: st.warning("Đã tồn tại.")
                    st.divider()
        else:
            st.success("🎉 Tất cả hoạt chất đã khớp MA_THUOC!")

        st.divider()
        st.markdown("##### 💾 Sao lưu alias_ten.csv")
        dl1, dl2 = st.columns(2)
        with dl1:
            if not df_alias.empty:
                st.download_button("⬇️ Tải alias_ten.csv",
                                   df_alias.to_csv(index=False).encode('utf-8'),
                                   "alias_ten.csv", "text/csv", use_container_width=True)
        with dl2:
            uploaded_ten = st.file_uploader("⬆️ Upload alias_ten.csv", type=['csv'], key='upload_alias_ten')
            if uploaded_ten:
                try:
                    df_up = pd.read_csv(uploaded_ten, dtype=str).fillna('')
                    for col in ALIAS_COLS:
                        if col not in df_up.columns: df_up[col] = ''
                    st.session_state['df_alias'] = df_up[ALIAS_COLS]
                    save_alias(df_up[ALIAS_COLS])
                    st.success(f"✅ Đã nạp {len(df_up)} alias tên."); st.rerun()
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}")

    # ── SUB-TAB 2: ALIAS ĐƯỜNG DÙNG ──────────────────
    with subtab_dd:
        st.caption(
            "**Mục đích:** Đường dùng trong file thầu không có trong bảng chuẩn BHYT. "
            "Sau khi bạn xác nhận đường dùng chuẩn, hệ thống tự nhớ cho lần sau.")
        df_alias_dd_tab = st.session_state['df_alias_dd'].copy()

        if df_alias_dd_tab.empty:
            st.info("Chưa có alias đường dùng nào. Xác nhận đường dùng ở tab Xử lý chính sẽ tự lưu vào đây.")
        else:
            st.dataframe(df_alias_dd_tab, use_container_width=True, hide_index=True)
            with st.expander("🗑️ Xóa alias đường dùng"):
                del_dd_idx = st.number_input("Số thứ tự dòng cần xóa (từ 0)",
                                              min_value=0, max_value=max(0,len(df_alias_dd_tab)-1), value=0,
                                              key='del_dd_idx')
                if st.button("Xóa dòng alias ĐD", key='del_dd_btn'):
                    df_alias_dd_tab = df_alias_dd_tab.drop(index=del_dd_idx).reset_index(drop=True)
                    st.session_state['df_alias_dd'] = df_alias_dd_tab
                    save_alias_dd(df_alias_dd_tab); st.rerun()

        st.divider()
        st.markdown("##### 💾 Sao lưu alias_dd.csv")
        dl3, dl4 = st.columns(2)
        with dl3:
            if not df_alias_dd_tab.empty:
                st.download_button("⬇️ Tải alias_dd.csv",
                                   df_alias_dd_tab.to_csv(index=False).encode('utf-8'),
                                   "alias_dd.csv", "text/csv", use_container_width=True)
        with dl4:
            uploaded_dd = st.file_uploader("⬆️ Upload alias_dd.csv", type=['csv'], key='upload_alias_dd')
            if uploaded_dd:
                try:
                    df_up_dd = pd.read_csv(uploaded_dd, dtype=str).fillna('')
                    for col in ALIAS_DD_COLS:
                        if col not in df_up_dd.columns: df_up_dd[col] = ''
                    st.session_state['df_alias_dd'] = df_up_dd[ALIAS_DD_COLS]
                    save_alias_dd(df_up_dd[ALIAS_DD_COLS])
                    st.success(f"✅ Đã nạp {len(df_up_dd)} alias đường dùng."); st.rerun()
                except Exception as e:
                    st.error(f"❌ Lỗi: {e}")
