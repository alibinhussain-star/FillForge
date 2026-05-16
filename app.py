from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd, re, io, tempfile, os, json, copy, random, string, time, zipfile
from datetime import datetime
from email.mime.text import MIMEText
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

app = Flask(__name__, template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ── Logging ────────────────────────────────────────────────────
LOG_PATH = os.path.join(os.path.dirname(__file__), 'activity.log')

def write_log(email, action, details=''):
    entry = {
        'ts': datetime.utcnow().isoformat() + 'Z',
        'email': email or 'anonymous',
        'action': action,
        'details': details,
    }
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(json.dumps(entry) + '\n')
    except Exception as e:
        print('log err', e)

def read_logs(limit=200):
    if not os.path.exists(LOG_PATH): return []
    out = []
    try:
        with open(LOG_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line: continue
                try: out.append(json.loads(line))
                except: pass
    except: pass
    return list(reversed(out))[:limit]

# ── Load embedded template file once at startup ────────────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Logic___Template_File.xlsx')

def _build_header_row_map():
    wb  = load_workbook(TEMPLATE_PATH)
    ws  = wb['PV Template']
    hdr_map    = {}
    static_map = {}
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == 'Category *':
            for r2 in range(r + 1, min(r + 5, ws.max_row + 1)):
                d2 = ws.cell(r2, 4).value
                if d2 and str(d2).strip() not in ('SubType','Category *','nan',''):
                    st = str(d2).strip()
                    if st not in hdr_map:
                        hdr_map[st] = r
                        hdrs = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                        entry = {}
                        for ci, col in enumerate(hdrs):
                            if col in ('Category *','SubCategory *','CategoryType *',
                                       'SubType','PVID *','discoveryCategoryIds'):
                                v = str(ws.cell(r2, ci + 1).value or '').strip()
                                if v and v not in ('nan','NaN','None'):
                                    entry[col] = v
                        static_map[st] = entry
                    break
    return hdr_map, static_map

try:
    SUBTYPE_HEADER_ROW, SUBTYPE_MAP = _build_header_row_map()
except Exception as e:
    print(f"Warning: Could not build header row map: {e}")
    SUBTYPE_HEADER_ROW, SUBTYPE_MAP = {}, {}

def load_pv_list():
    try:
        pv  = pd.read_excel(TEMPLATE_PATH, sheet_name='PV List')
        col = pv.columns[0]
        return [str(v).strip() for v in pv[col].dropna() if str(v).strip() not in ('nan','SubType')]
    except Exception as e:
        print(f"Warning: Could not load PV List: {e}")
        return []

try:
    PV_LIST = load_pv_list()
except Exception as e:
    print(f"Warning: Could not load PV_LIST: {e}")
    PV_LIST = []

def get_template_wb_for_subtype(subtype):
    try:
        wb_src  = load_workbook(TEMPLATE_PATH)
        ws_src  = wb_src['PV Template']
        hdr_row = SUBTYPE_HEADER_ROW.get(subtype, 1)
        headers = [ws_src.cell(hdr_row, c).value for c in range(1, ws_src.max_column + 1)]
        while headers and headers[-1] is None:
            headers.pop()
    except Exception as e:
        print(f"Warning: Could not load template for {subtype}: {e}")
        headers = []
    wb_new       = Workbook()
    ws_new       = wb_new.active
    ws_new.title = 'PV Template'
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h
    return wb_new, headers

# ── Default config ─────────────────────────────────────────────
DEFAULT_CONFIG = {
    "brands":              {},
    "biz_cat_id":          "BCAT-139461",
    "biz_cat_name":        "Footwear",
    "relationship":        "Parent",
    "catalog_status":      "ACTIVE",
    "status_remark":       "Ready to Launch",
    "tax_master_status":   "active",
    "gst_cgst":            50,
    "gst_sgst":            50,
    "gst_igst":            0,
    "country_of_origin":   "India",
    "product_condition":   "Fresh",
    "manufacturing_year":  "2026",
    "discovery_cat":       "DISCAT-135542",
}

CONFIG_PATH    = '/tmp/fillforge_config.json'
CE_CONFIG_PATH = '/tmp/fillforge_ce_config.json'
AP_CONFIG_PATH = '/tmp/fillforge_ap_config.json'

def _load_config(path, defaults):
    cfg = {k: v for k, v in defaults.items()}
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            cfg.update(saved)
        except Exception as e:
            print(f'Warning: could not load config from {path}: {e}')
    return cfg

def _save_config(path, cfg):
    try:
        tmp = path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except Exception as e:
        print(f'Warning: could not save config to {path}: {e}')

def get_config():
    return _load_config(CONFIG_PATH, DEFAULT_CONFIG)

def get_ce_config_from_disk():
    return _load_config(CE_CONFIG_PATH, CE_DEFAULT_CONFIG)

def get_ap_config_from_disk():
    return _load_config(AP_CONFIG_PATH, AP_DEFAULT_CONFIG)

config = get_config()

# ── Utility ────────────────────────────────────────────────────
def safe(val, default=''):
    try:
        if pd.isna(val): return default
    except: pass
    s = str(val).strip() if val is not None else default
    return default if s in ('nan','None','NaN') else s

def detect_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    for c in df.columns:
        for cand in candidates:
            if cand.lower() in c.lower(): return c
    return None

def build_col_map(df, hints):
    return {k: detect_col(df, v) for k, v in hints.items() if detect_col(df, v)}

def title_case_color(raw):
    if not raw: return raw
    return ' '.join(w.capitalize() for w in str(raw).strip().split())

def merge_colors(primary, secondary):
    p = title_case_color(safe(primary))
    s = title_case_color(safe(secondary))
    if not s or s in ('nan','None',''): return p
    if p.lower() == s.lower(): return p
    return f"{p} & {s}"

def extract_article(sku):
    s = str(sku).strip()
    if '_' in s:
        return s.split('_')[0].strip()
    COLOR_CODES = [
        'LBLUE','LGREY','LGRAY','LGREEN','LPINK','LBROWN',
        'DBLUE','DGREY','DGRAY','DGREEN','DPINK','DBROWN',
        'ONION','IVORY','ASSORTED','MULTI',
        'MRN','OLV','BGE','BRN','CRM','GLD','SLV','PPL','NVY',
        'MUL','AST','CLR','BLK','WHT','BLU','GRN','GRY','PNK',
        'ORG','PRP','BK','WH','BL','RD','GR','NV','PK','YL','OR',
    ]
    code_pattern = '(' + '|'.join(COLOR_CODES) + ')$'
    COLOR_WORDS = (r'BLACK|WHITE|BLUE|RED|GREEN|GREY|GRAY|BEIGE|BROWN|NAVY|PINK|YELLOW|'
                   r'ORANGE|PURPLE|PEACH|CREAM|MAROON|GOLD|SILVER|ASSORTED|MULTI|ONION|IVORY')
    cleaned = re.sub(r'[\s_]+[\dXx]+[-]+[\dXx]+[-]+[\dXx]*[\s\-]*(' + COLOR_WORDS + r')\b.*', '', s, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r'[\s\-]+(' + COLOR_WORDS + r')\b.*', '', cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r'\s*\(\d+.*\)$', '', cleaned).strip()
    cleaned = re.sub(r'[\s\-]+$', '', cleaned).strip()
    cleaned = re.sub(r'\s*-*\s*', '-', cleaned)
    if cleaned and cleaned != s:
        return cleaned
    m = re.search(code_pattern, s, re.IGNORECASE)
    if m:
        stripped = s[:m.start()].rstrip('-_').strip()
        if stripped:
            return stripped
    return s

def expand_size_range(size_str, size_type='UK'):
    s = str(size_str).strip()
    if ',' in s: return [x.strip() for x in s.split(',') if x.strip()]
    # Handle dash-separated sizes like "28-30-32-34-36"
    if '-' in s:
        parts = [x.strip() for x in s.split('-') if x.strip() and re.match(r'^\d+$', x.strip())]
        if len(parts) > 1:
            return parts
    m = re.match(r'^(\d+)[Xx](\d+)$', s)
    if m:
        start, end = int(m.group(1)), int(m.group(2))
        pfx = f"{size_type} " if size_type else ''
        if start > end:
            return [f"{pfx}{i}" for i in list(range(start, 14)) + list(range(1, end + 1))]
        return [f"{pfx}{i}" for i in range(start, end + 1)]
    if ' ' in s: return [x.strip() for x in s.split() if x.strip()]
    return [s] if s else []

def build_set_details(sizes_list, set_details_raw):
    raw = str(set_details_raw).strip() if set_details_raw else ''
    full = re.findall(r'((?:UK\s*)?\d+)\s*/\s*(\d+)', raw)
    if full:
        det   = [f'{s}/{q}' for s, q in full]
        desc  = [f'{q} pcs of {s}' for s, q in full]
        avail = ', '.join(s for s, _ in full)
        return ', '.join(det), ', '.join(desc), avail
    dash = re.findall(r'((?:UK\s*)?\d+)\s*[-–]+\s*(\d+)', raw)
    if dash:
        det   = [f'{s}/{q}' for s, q in dash]
        desc  = [f'{q} pcs of {s}' for s, q in dash]
        avail = ', '.join(s for s, _ in dash)
        return ', '.join(det), ', '.join(desc), avail
    qty_parts = [x.strip() for x in raw.split(',') if x.strip()]
    if qty_parts and sizes_list and all(q.isdigit() for q in qty_parts):
        if len(qty_parts) == len(sizes_list):
            pairs = list(zip(sizes_list, qty_parts))
        else:
            qty   = sum(int(q) for q in qty_parts) // max(len(sizes_list), 1)
            pairs = [(s, str(qty)) for s in sizes_list]
        det  = [f'{s}/{q}' for s, q in pairs]
        desc = [f'{q} pcs of {s}' for s, q in pairs]
        return ', '.join(det), ', '.join(desc), ', '.join(s for s, _ in pairs)
    if sizes_list:
        det  = [f'{s}/1' for s in sizes_list]
        desc = [f'1 pcs of {s}' for s in sizes_list]
        return ', '.join(det), ', '.join(desc), ', '.join(sizes_list)
    return raw, raw, ''

def parse_lbh(dim_str):
    parts = re.split(r'[Xx×]', str(dim_str).strip())
    if len(parts) == 3:
        try: return int(parts[0]), int(parts[1]), int(parts[2])
        except: pass
    return None, None, None

def derive_gender(subtype):
    st = str(subtype).lower()
    if "women" in st: return "Women's"
    if "men" in st:   return "Men's"
    if "girl" in st:  return "Girl's"
    if "boy" in st:   return "Boy's"
    return ""

def make_title(brand, gender, upper, closure, fw_type, color):
    parts = [p for p in [brand, gender, upper, closure, fw_type] if p]
    base  = ' '.join(parts)
    return f"{base}, {color}" if color else base

def make_internal_title(brand, article, gender, upper, closure, fw_type, color, set_count, set_details_tpl):
    parts = [p for p in [brand, article, gender, upper, closure, fw_type] if p]
    base  = ' '.join(parts)
    return f"{base}, {color}, Set of {set_count} ({set_details_tpl})"

def make_description(brand, article, gender, upper, closure, fw_type, sole, color, sizes, set_count):
    title_part = ' '.join(p for p in [brand, gender, upper, closure, fw_type] if p)
    upper_l    = upper.lower()   if upper   else 'upper'
    sole_l     = sole.lower()    if sole    else 'sole'
    closure_l  = closure.lower() if closure else 'slip-on'
    return (
        f"Step out in style with the {title_part} ({article}). "
        f"The {upper_l} upper offers a snug, comfortable fit, while the {sole_l} sole delivers "
        f"reliable grip and cushioned support. "
        f"The {closure_l} closure makes wearing easy. "
        f"Color: {color}. Available sizes: {sizes}. "
        f"Set of {set_count} bulk-pack — ideal for retailers and resellers."
    )

def normalize_brands(brands_data):
    if not brands_data:
        return {}
    if isinstance(brands_data, dict):
        return brands_data
    if isinstance(brands_data, list):
        result = {}
        for item in brands_data:
            if isinstance(item, dict):
                name = item.get('name', item.get('brandName', ''))
                bid  = item.get('id',   item.get('brandId',   ''))
                if name:
                    result[name] = bid
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                result[item[0]] = item[1]
        return result
    return {}

def get_brand_info(drow, col_map, brands_dict):
    brands_dict = normalize_brands(brands_dict)
    if not brands_dict:
        return '', ''
    fallback_brand, fallback_id = next(iter(brands_dict.items()))
    brand_col = col_map.get('brand')
    if not brand_col:
        return fallback_brand, fallback_id
    try:
        file_brand = str(drow.get(brand_col, '')).strip()
    except:
        file_brand = ''
    if not file_brand or file_brand.lower() in ('nan', 'none', '', 'null'):
        return fallback_brand, fallback_id
    file_brand_lower = file_brand.lower().strip()
    for b_name, b_id in brands_dict.items():
        if b_name.lower().strip() == file_brand_lower:
            return b_name, b_id
    for b_name, b_id in brands_dict.items():
        bn = b_name.lower().strip()
        if bn in file_brand_lower or file_brand_lower in bn:
            return b_name, b_id
    file_first_word = file_brand_lower.split()[0] if file_brand_lower.split() else ''
    if file_first_word:
        for b_name, b_id in brands_dict.items():
            cfg_first_word = b_name.lower().strip().split()[0] if b_name.strip().split() else ''
            if cfg_first_word and cfg_first_word == file_first_word:
                return b_name, b_id
    return fallback_brand, fallback_id


# ── Column hints ────────────────────────────────────────────────
DUMP_COL_HINTS = {
    'sku':            ['Seller SKU ID','Seller SKU_ID','ChildSKU *','ChildSKU','SKU'],
    'article':        ['Article Code','Article Number','ARTICLE_NUMBER'],
    'image':          ['Image Links','Image Link','ImageURL1','imageURL1 *'],
    'vertical':       ['Product Vertical','Subtype','SubType','PVName'],
    'gender':         ['Gender','GENDER *','GENDER'],
    'fw_type':        ['Foot Wear Type','FW Type','Footwear Type','FOOTWEAR_TYPE *'],
    'upper_material': ['Upper Material','UPPER_MATERIAL *'],
    'closure_type':   ['Closure Type','CLOSURE_TYPE *'],
    'sole_material':  ['Sole Material','SOLE_MATERIAL *'],
    'set_of':         ['Set of','Set Of','*Quantity','Set Count','SET_COUNT *','set count','Quantity'],
    'size_type':      ['Size Type','Unit of Measurement','Unit of Measuremen'],
    'color':          ['Primary Colour','Product Primary Color','Product Color','PRODUCT_COLOR *'],
    'color2':         ['Secondary Colour','Secondary Color'],
    'sizes':          ['Available Sizes','AVAILABLE_SIZES *'],
    'set_details':    ['Set Details','SET_DETAILS *'],
    'heel_height':    ['Heel Height','Heel_Height','HEEL_HEIGHT'],
    'heel_type':      ['Heel Type','Heel_Type','HEEL_TYPE'],
    'hsn':            ['HSN','*HSN Code','HSN Code','hsnCode *'],
    'gst':            ['GSTpercentage','*GST','GST','gstPercentage *'],
    'moq':            ['*MOQ','MOQ *','MOQ'],
    'mrp':            ['*MRP full Set','MRP *','MRP full Set','MRP'],
    'sp':             ['*Selling Price Full Set','SellingPrice *','*Selling Price per Pair'],
    'weight':         ['*Product Weight (In KG) Full Ste','PRODUCT_WEIGHT_IN_KG *'],
    'dims':           ['*Product Dimension (LXBXH) Full Set','Product Dimension'],
    'dim_uom':        ['*Product Dimension UOM','PRODUCT_DIMENSION_UOM *'],
    'packing':        ['Packing Type','PACKAGING_TYPE *'],
    'country':        ['Country of Origin','COUNTRY_OF_ORIGIN *'],
    'product_desc':   ['Product Description','productDescription *'],
    'brand':          ['Brand','Brand Name','brandName *','brand_name'],
}

BASE_COL_HINTS = {
    'article': ['Article Number','Article Code','ARTICLE_NUMBER'],
    'sku':     ['Seller SKU ID','Seller SKU_ID','ChildSKU'],
}

def fill_template(ws, headers, rows_df, col_map, subtype, existing_articles, existing_skus):
    tcol        = {h: i+1 for i, h in enumerate(headers) if h}
    _cfg        = get_config()
    brands_dict = normalize_brands(_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))
    gender  = derive_gender(subtype)
    st_data = SUBTYPE_MAP.get(subtype, {})
    skipped, filled = [], 0

    for _, drow in rows_df.iterrows():
        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand    = fallback_brand
            brand_id = fallback_id

        sku_raw = safe(drow.get(col_map.get('sku',''), ''))
        art_raw = safe(drow.get(col_map.get('article',''), ''))
        article = art_raw if (art_raw and '_' not in art_raw) else extract_article(sku_raw)

        if article.upper() in existing_articles or sku_raw.upper() in existing_skus:
            skipped.append({'sku': sku_raw, 'article': article, 'reason': 'Already exists in base data'})
            continue

        filled  += 1
        row_idx  = filled + 1

        size_type   = safe(drow.get(col_map.get('size_type',''), 'UK')) or 'UK'
        sizes_raw   = safe(drow.get(col_map.get('sizes',''), ''))
        set_det_raw = safe(drow.get(col_map.get('set_details',''), ''))
        _sc_raw     = safe(drow.get(col_map.get('set_of',''), '0'))
        _sc_nums    = re.findall(r'\d+', str(_sc_raw))
        set_count   = int(_sc_nums[0]) if _sc_nums else 0
        color       = merge_colors(
            drow.get(col_map.get('color',''), ''),
            drow.get(col_map.get('color2',''), '')
        )
        mrp         = drow.get(col_map.get('mrp',''), '')
        sp          = drow.get(col_map.get('sp',''), '')
        weight      = drow.get(col_map.get('weight',''), '')
        dim_raw     = safe(drow.get(col_map.get('dims',''), ''))
        hsn         = drow.get(col_map.get('hsn',''), '')
        gst         = drow.get(col_map.get('gst',''), 5)
        moq         = drow.get(col_map.get('moq',''), 1)
        upper_mat   = safe(drow.get(col_map.get('upper_material',''), ''))
        sole_mat    = safe(drow.get(col_map.get('sole_material',''), ''))
        closure     = safe(drow.get(col_map.get('closure_type',''), ''))
        heel_ht     = safe(drow.get(col_map.get('heel_height',''), ''))
        heel_type   = safe(drow.get(col_map.get('heel_type',''), ''))
        img_url     = safe(drow.get(col_map.get('image',''), ''))
        packing     = safe(drow.get(col_map.get('packing',''), '')) or 'Loose Packing'
        country     = safe(drow.get(col_map.get('country',''), '')) or _cfg['country_of_origin']
        dim_uom     = safe(drow.get(col_map.get('dim_uom',''), '')) or 'cm'
        fw_type     = safe(drow.get(col_map.get('fw_type',''), ''))
        prod_desc   = safe(drow.get(col_map.get('product_desc',''), ''))

        for v, field in [(packing,'packing'),(country,'country')]:
            if v in ('nan','None',''):
                if field == 'country': country = _cfg['country_of_origin']
                if field == 'packing': packing = 'Loose Packing'

        sizes_list = expand_size_range(sizes_raw, size_type)
        set_details_tpl, set_desc, avail_sizes = build_set_details(sizes_list, set_det_raw)
        if not avail_sizes: avail_sizes = ', '.join(sizes_list)

        title          = make_title(brand, gender, upper_mat, closure, fw_type, color)
        internal_title = make_internal_title(brand, article, gender, upper_mat, closure, fw_type, color, set_count, set_details_tpl)
        description    = prod_desc if prod_desc else make_description(brand, article, gender, upper_mat, closure, fw_type, sole_mat, color, avail_sizes, set_count)

        try:    mrp    = float(mrp)      if str(mrp).strip()    not in ('','nan') else ''
        except: mrp    = ''
        try:    sp     = float(sp)       if str(sp).strip()     not in ('','nan') else ''
        except: sp     = ''
        try:    hsn    = int(float(hsn)) if str(hsn).strip()    not in ('','nan') else ''
        except: hsn    = ''
        try:    gst    = int(float(gst))
        except: gst    = 5
        try:    moq    = int(float(moq))
        except: moq    = 1
        try:    weight = float(weight)   if str(weight).strip() not in ('','nan') else ''
        except: weight = ''
        L, B, H = parse_lbh(dim_raw)

        row_data = {
            'Category *':                                  st_data.get('Category *', 'Footwear'),
            'SubCategory *':                               st_data.get('SubCategory *', ''),
            'CategoryType *':                              st_data.get('CategoryType *', ''),
            'SubType':                                     subtype,
            'PVID *':                                      st_data.get('PVID *', ''),
            'BusinessCategoryId *':                        _cfg['biz_cat_id'],
            'BusinessCategoryName *':                      _cfg['biz_cat_name'],
            'Relationship *':                              _cfg['relationship'],
            'ParentProductId *':                           sku_raw,
            'ChildSKU *':                                  sku_raw,
            'MRP *':                                       mrp,
            'SellingPrice *':                              sp,
            'MOQ *':                                       moq,
            'title *':                                     title,
            'internalTitle *':                             internal_title,
            'brandId *':                                   brand_id,
            'brandName *':                                 brand,
            'imageURL1 *':                                 img_url,
            'catalogStatus *':                             _cfg['catalog_status'],
            'statusRemark':                                _cfg['status_remark'],
            'discoveryCategoryIds':                        st_data.get('discoveryCategoryIds', _cfg['discovery_cat']),
            'productDescription *':                        description,
            'PRODUCT_IDENTIFIER *':                        'Set',
            'SET_NAME *':                                  f'Set of {set_count}',
            'SET_COUNT *':                                 set_count,
            'PACK_NAME *':                                 'Pack of 1',
            'PACK_OF *':                                   1,
            'IS_COMBO *':                                  'yes',
            'AVAILABLE_SIZES *':                           avail_sizes,
            'SET_DETAILS *':                               set_details_tpl,
            'SET_DESCRIPTION *':                           set_desc,
            'PRODUCT_COLOR *':                             color,
            'ARTICLE_NUMBER *':                            article,
            'MODEL_NAME *':                                article,
            'PRODUCT_CONDITION *':                         _cfg['product_condition'],
            'UNIT_OF_MEASUREMENT_SINGULAR *':              'Pair',
            'UNIT_OF_MEASUREMENT_PLURAL *':                'Pairs',
            'UNIT_OF_MEASUREMENT_SINGULAR_ABBREVIATION *': 'Pair',
            'UNIT_OF_MEASUREMENT_PLURAL_ABBREVIATION *':   'Pairs',
            'SELLER_SKU_ID *':                             sku_raw,
            'PACKAGING_TYPE *':                            'Branded Box',
            'GENDER *':                                    gender,
            'AGE_GROUP *':                                 '4- 5 Years',
            'CLOSURE_TYPE *':                              closure,
            'COUNTRY_OF_ORIGIN *':                         country,
            'MANUFACTURING_YEAR':                          _cfg['manufacturing_year'],
            'PRODUCT_LENGTH *':                            L,
            'PRODUCT_BREADTH *':                           B,
            'PRODUCT_HEIGHT *':                            H,
            'PRODUCT_DIMENSION_UOM *':                     dim_uom,
            'PRODUCT_WEIGHT_IN_KG *':                      weight,
            'FOOTWEAR_TYPE *':                             fw_type,
            'HEEL_HEIGHT':                                 heel_ht,
            'HEEL_HEIGHT *':                               heel_ht,
            'HEEL_TYPE':                                   heel_type,
            'HEEL_TYPE *':                                 heel_type,
            'SOLE_MATERIAL *':                             sole_mat,
            'UPPER_MATERIAL *':                            upper_mat,
            'hsnCode *':                                   hsn,
            'gstPercentage *':                             gst,
            'cgstShare *':                                 _cfg['gst_cgst'],
            'sgstShare *':                                 _cfg['gst_sgst'],
            'igstShare *':                                 _cfg['gst_igst'],
            'taxMasterStatus':                             _cfg['tax_master_status'],
        }

        for col_name, val in row_data.items():
            if col_name in tcol and val is not None and str(val) not in ('None',''):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

    return filled, skipped


# ═══════════════════════════════════════════════════════════════
# CONSUMER ELECTRONICS MODULE
# ═══════════════════════════════════════════════════════════════

CE_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Consumer Electronic Mapping Logic & templates.xlsx')

def _build_ce_header_row_map():
    wb  = load_workbook(CE_TEMPLATE_PATH)
    ws  = wb['CE - PV Template']
    hdr_map    = {}
    static_map = {}
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == 'Category *':
            for r2 in range(r + 1, min(r + 5, ws.max_row + 1)):
                c3_val      = ws.cell(r2, 3).value
                c4_val      = ws.cell(r2, 4).value
                subtype_val = c3_val if c3_val else c4_val
                if subtype_val and str(subtype_val).strip() not in ('CategoryType *','SubType','Category *','nan',''):
                    st = str(subtype_val).strip()
                    if st not in hdr_map:
                        hdr_map[st] = r
                        hdrs  = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                        entry = {}
                        for ci, col in enumerate(hdrs):
                            if col in ('Category *','SubCategory *','CategoryType *',
                                       'SubType','PVID *','discoveryCategoryIds'):
                                v = str(ws.cell(r2, ci + 1).value or '').strip()
                                if v and v not in ('nan','NaN','None'):
                                    entry[col] = v
                        static_map[st] = entry
                    break
    return hdr_map, static_map

try:
    CE_SUBTYPE_HEADER_ROW, CE_SUBTYPE_MAP = _build_ce_header_row_map()
except Exception as e:
    print(f"Warning: Could not build CE header row map: {e}")
    CE_SUBTYPE_HEADER_ROW, CE_SUBTYPE_MAP = {}, {}

def load_ce_pv_list():
    try:
        pv  = pd.read_excel(CE_TEMPLATE_PATH, sheet_name='Category List')
        col = pv.columns[0]
        return [str(v).strip() for v in pv[col].dropna()
                if str(v).strip() not in ('nan','CategoryType *','')]
    except Exception as e:
        print(f"Warning: Could not load CE Category List: {e}")
        return []

try:
    CE_PV_LIST = load_ce_pv_list()
except Exception as e:
    print(f"Warning: Could not load CE_PV_LIST: {e}")
    CE_PV_LIST = []

CE_DEFAULT_CONFIG = {
    "brands":              {},
    "biz_cat_id":          "BCAT-139438",
    "biz_cat_name":        "Consumer Electronics",
    "relationship":        "Parent",
    "catalog_status":      "ACTIVE",
    "status_remark":       "Ready to Launch",
    "tax_master_status":   "active",
    "gst_cgst":            50,
    "gst_sgst":            50,
    "gst_igst":            0,
    "country_of_origin":   "India",
    "product_condition":   "Fresh",
    "manufacturing_year":  "2026",
    "discovery_cat":       "DISCAT-135528",
}

CE_DUMP_COL_HINTS = {
    'sku':            ['Child SKU','ChildSKU *','ChildSKU','SKU','Seller SKU ID'],
    'article':        ['Model Number','MODEL NUMBER','Model NUMBER',
                       'Article Number','Article Code','ARTICLE_NUMBER',
                       'Name of the model/Title name'],
    'image':          ['Main Image URL','Image Links','Image Link','ImageURL1','imageURL1 *'],
    'image2':         ['Other Image URL 1','Other Image URL1'],
    'image3':         ['Other Image URL 2','Other Image URL2'],
    'image4':         ['Other Image URL 3','Other Image URL3'],
    'image5':         ['Other Image URL 4','Other Image URL4'],
    'image6':         ['Other Image URL 5','Other Image URL5'],
    'vertical':       ['Product Type','Product Sub-type','CategoryType *','Subtype','SubType'],
    'brand':          ['Brand','Brand Name','brandName *','brand_name'],
    'mrp':            ['MRP','*MRP full Set','MRP *','MRP full Set'],
    'sp':             ['Selling Price','SellingPrice *','*Selling Price per Pair'],
    'moq':            ['*Minimum Order Quantity','*MOQ','MOQ *','MOQ'],
    'color':          ['Product Color','Product Colour','Primary Colour','PRODUCT_COLOR *'],
    'product_desc':   ['Product Description','productDescription *'],
    'hsn':            ['HSN Code','*HSN Code','hsnCode *'],
    'gst':            ['GST','*GST','gstPercentage *'],
    'weight':         ['Product Weight','*Product Weight (In KG) Full Ste','PRODUCT_WEIGHT_IN_KG *'],
    'dims':           ['*Product Dimension (LXBXH)','Product Dimension (LXBXH) Full Set','Product Dimension'],
    'dim_uom':        ['*Product Dimension UOM','PRODUCT_DIMENSION_UOM *'],
    'packing':        ['Packaging Type','PACKAGING_TYPE *'],
    'country':        ['Country/Region of Origin','Country of Origin','COUNTRY_OF_ORIGIN *'],
    'warranty':       ['Warranty Period','Warranty'],
    'battery':        ['Battery Capacity','BATTERY_CAPACITY_MAH *'],
    'charging_type':  ['Charging type supported','CHARGING_TYPE_SUPPORTED *'],
    'ram':            ['RAM','RAM *'],
    'storage':        ['Storage Capacity','INTERNAL_STORAGE *'],
    'sim_type':       ['Sim Type','SIM_TYPE *'],
    'os':             ['Operating System','OPERATING_SYSTEM_OS *'],
    'front_camera':   ['Front Camera','FRONT_CAMERA_RESOLUTION *'],
    'back_camera':    ['Back Camera','PRIMARY_CAMERA_RESOLUTION *'],
    'screen_size':    ['Screen Size','DISPLAY_SIZE *'],
    'display_type':   ['Display Type','DISPLAY_TYPE *'],
    'processor_core': ['Processor Core','NUMBER_OF_PROCESSOR_CORES *'],
    'network_support':['Network Support','Network'],
    'bluetooth':      ['Bluetooth Version','BLUETOOTH_VERSION *'],
    'product_type':   ['Product Type','Product Sub-type'],
}

CE_BASE_COL_HINTS = {
    'article': ['Model Number','MODEL NUMBER','Name of the model/Title name',
                'Article Number','Article Code','ARTICLE_NUMBER'],
    'sku':     ['Child SKU','ChildSKU'],
}

def extract_model_name(title_name):
    s = str(title_name).strip()
    if not s: return s
    cleaned = re.sub(r'\s*\([^)]*\)\s*', ' ', s)
    cleaned = re.sub(r'\s+(Fresh|Seal Open|Non Activated|Open Seal)\s*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+\d+\s*GB\s*RAM.*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+\d+\s*GB\s*ROM.*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+\d+G\s*.*$', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned if cleaned else s

def make_ce_title(brand, model_name, back_camera, category_type, ram_storage, color, condition):
    core_parts = []
    if brand:         core_parts.append(brand)
    if model_name:    core_parts.append(model_name)
    if back_camera:   core_parts.append(f'{back_camera} Camera')
    if category_type: core_parts.append(category_type)
    base = ' '.join(core_parts)
    suffix_parts = []
    if ram_storage: suffix_parts.append(ram_storage)
    color_condition = ' '.join(p for p in [color, f'({condition})' if condition else ''] if p)
    if color_condition: suffix_parts.append(color_condition)
    if suffix_parts:
        return f"{base}, {', '.join(suffix_parts)}"
    return base

def make_ce_description(brand, model_name, category_type, ram, storage, processor, battery,
                        screen_size, display_type, color, front_camera, back_camera, os):
    parts = []
    if brand and model_name:
        parts.append(f"Experience the {brand} {model_name}")
    if category_type:
        parts.append(f", a powerful {category_type}")
    if ram and storage:
        parts.append(f" featuring {ram} RAM and {storage} internal storage")
    if processor:
        parts.append(f", powered by a {processor} processor")
    if battery:
        parts.append(f". Equipped with a {battery} battery")
    if screen_size and display_type:
        parts.append(f", it boasts a {screen_size} {display_type} display")
    if front_camera and back_camera:
        parts.append(f". Capture stunning photos with {back_camera} rear and {front_camera} front cameras")
    if color:
        parts.append(f". Available in {color} color")
    if os:
        parts.append(f". Runs on {os}")
    return ''.join(parts) + ". Ideal for everyday use with reliable performance and modern features."

def extract_from_description(desc, field_type):
    if not desc: return ''
    desc_lower = str(desc).lower()
    if field_type == 'display_type':
        for dt in ['super amoled','amoled','pls lcd','lcd','ips','oled','tft']:
            if dt in desc_lower:
                if dt == 'super amoled': return 'Super AMOLED'
                if dt == 'pls lcd':      return 'PLS LCD'
                return dt.upper()
        return ''
    if field_type == 'charging_type':
        if 'usb type-c' in desc_lower or 'usb c' in desc_lower or 'type-c' in desc_lower:
            return 'USB Type-C'
        if 'micro usb' in desc_lower or 'micro-usb' in desc_lower:
            return 'Micro USB'
        if 'lightning' in desc_lower:
            return 'Lightning'
        return ''
    if field_type == 'bluetooth':
        m = re.search(r'bluetooth\s*(\d+\.?\d*)', desc_lower)
        return f'Bluetooth {m.group(1)}' if m else ''
    if field_type == 'processor':
        for proc in ['dimensity','snapdragon','helio','exynos','kirin','mediatek']:
            if proc in desc_lower:
                m = re.search(rf'{proc}\s+([a-z]*\d+\s*[a-z]*)', desc_lower, re.IGNORECASE)
                if m:
                    return f"{proc.capitalize()} {m.group(1).strip().upper()}"
                return proc.capitalize()
        return ''
    return ''

def get_ce_template_wb_for_subtype(subtype):
    try:
        wb_src  = load_workbook(CE_TEMPLATE_PATH)
        ws_src  = wb_src['CE - PV Template']
        hdr_row = CE_SUBTYPE_HEADER_ROW.get(subtype, 2)
        headers = [ws_src.cell(hdr_row, c).value for c in range(1, ws_src.max_column + 1)]
        while headers and headers[-1] is None:
            headers.pop()
    except Exception as e:
        print(f"Warning: Could not load CE template for {subtype}: {e}")
        headers = []
    wb_new       = Workbook()
    ws_new       = wb_new.active
    ws_new.title = 'CE - PV Template'
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h
    return wb_new, headers

def fill_ce_template(ws, headers, rows_df, col_map, subtype, existing_articles, existing_skus):
    tcol        = {h: i+1 for i, h in enumerate(headers) if h}
    _ce_cfg     = get_ce_config_from_disk()
    brands_dict = normalize_brands(_ce_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))
    st_data = CE_SUBTYPE_MAP.get(subtype, {})
    skipped, filled = [], 0

    for _, drow in rows_df.iterrows():
        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand    = fallback_brand
            brand_id = fallback_id

        sku_raw    = safe(drow.get(col_map.get('sku',''), ''))
        model_num  = safe(drow.get(col_map.get('article',''), ''))
        article    = model_num if model_num else sku_raw
        model_name = article

        if article.upper() in existing_articles or sku_raw.upper() in existing_skus:
            skipped.append({'sku': sku_raw, 'article': article, 'reason': 'Already exists in base data'})
            continue

        filled  += 1
        row_idx  = filled + 1

        mrp          = drow.get(col_map.get('mrp',''), '')
        sp           = drow.get(col_map.get('sp',''), '')
        moq          = drow.get(col_map.get('moq',''), 1)
        color        = safe(drow.get(col_map.get('color',''), ''))
        weight       = drow.get(col_map.get('weight',''), '')
        dim_raw      = safe(drow.get(col_map.get('dims',''), ''))
        hsn          = drow.get(col_map.get('hsn',''), '')
        gst          = drow.get(col_map.get('gst',''), 18)
        packing      = safe(drow.get(col_map.get('packing',''), '')) or 'BOX'
        country      = safe(drow.get(col_map.get('country',''), '')) or _ce_cfg['country_of_origin']
        dim_uom      = safe(drow.get(col_map.get('dim_uom',''), '')) or 'cm'
        prod_desc    = safe(drow.get(col_map.get('product_desc',''), ''))
        warranty     = safe(drow.get(col_map.get('warranty',''), ''))
        battery      = safe(drow.get(col_map.get('battery',''), ''))
        charging     = safe(drow.get(col_map.get('charging_type',''), ''))
        ram          = safe(drow.get(col_map.get('ram',''), ''))
        storage      = safe(drow.get(col_map.get('storage',''), ''))
        sim_type     = safe(drow.get(col_map.get('sim_type',''), ''))
        os           = safe(drow.get(col_map.get('os',''), ''))
        front_cam    = safe(drow.get(col_map.get('front_camera',''), ''))
        back_cam     = safe(drow.get(col_map.get('back_camera',''), ''))
        screen_size  = safe(drow.get(col_map.get('screen_size',''), ''))
        display_type = safe(drow.get(col_map.get('display_type',''), ''))
        proc_core    = safe(drow.get(col_map.get('processor_core',''), ''))
        network      = safe(drow.get(col_map.get('network_support',''), ''))
        bluetooth    = safe(drow.get(col_map.get('bluetooth',''), ''))
        prod_type    = safe(drow.get(col_map.get('product_type',''), ''))

        img_url  = safe(drow.get(col_map.get('image',''), ''))
        img2_url = safe(drow.get(col_map.get('image2',''), ''))
        img3_url = safe(drow.get(col_map.get('image3',''), ''))
        img4_url = safe(drow.get(col_map.get('image4',''), ''))
        img5_url = safe(drow.get(col_map.get('image5',''), ''))
        img6_url = safe(drow.get(col_map.get('image6',''), ''))

        desc = prod_desc if prod_desc else ''
        if not display_type:
            display_type = extract_from_description(desc, 'display_type')
        if not charging:
            charging = extract_from_description(desc, 'charging_type')
        if not bluetooth:
            bluetooth = extract_from_description(desc, 'bluetooth')

        ram_rom     = f"{ram} + {storage}" if (ram and storage) else ''
        ram_storage = ram_rom if ram_rom else (ram or storage or '')
        condition   = safe(drow.get(col_map.get('product_condition',''), '')) or _ce_cfg['product_condition']

        title          = make_ce_title(brand, model_name, back_cam, subtype, ram_storage, color, condition)
        internal_title = title

        description = prod_desc if prod_desc else make_ce_description(
            brand, model_name, subtype, ram, storage, proc_core, battery,
            screen_size, display_type, color, front_cam, back_cam, os
        )

        package_contents = 'Handset' if (prod_type and 'smart phone' in prod_type.lower()) else ''

        L, B, H = parse_lbh(dim_raw)

        weight_clean = ''
        if weight:
            m = re.search(r'([0-9.]+)', str(weight))
            if m: weight_clean = float(m.group(1))

        try:    mrp = float(mrp)      if str(mrp).strip()    not in ('','nan') else ''
        except: mrp = ''
        try:    sp  = float(sp)       if str(sp).strip()     not in ('','nan') else ''
        except: sp  = ''
        try:    hsn = int(float(hsn)) if str(hsn).strip()    not in ('','nan') else ''
        except: hsn = ''
        try:    gst = int(float(gst))
        except: gst = 18
        try:    moq = int(float(moq))
        except: moq = 1

        row_data = {
            'Category *':                                  st_data.get('Category *', 'Consumer Electronics'),
            'SubCategory *':                               st_data.get('SubCategory *', 'Mobile'),
            'CategoryType *':                              subtype,
            'SubType':                                     subtype,
            'PVID *':                                      st_data.get('PVID *', ''),
            'BusinessCategoryId *':                        _ce_cfg['biz_cat_id'],
            'BusinessCategoryName *':                      _ce_cfg['biz_cat_name'],
            'Relationship *':                              _ce_cfg['relationship'],
            'ParentProductId *':                           sku_raw,
            'ChildSKU *':                                  sku_raw,
            'MRP *':                                       mrp,
            'SellingPrice *':                              sp,
            'MOQ *':                                       moq,
            'title *':                                     title,
            'internalTitle *':                             internal_title,
            'brandId *':                                   brand_id,
            'brandName *':                                 brand,
            'imageURL1 *':                                 img_url,
            'imageURL2':                                   img2_url,
            'imageURL3':                                   img3_url,
            'imageURL4':                                   img4_url,
            'imageURL5':                                   img5_url,
            'imageURL6':                                   img6_url,
            'catalogStatus *':                             _ce_cfg['catalog_status'],
            'statusRemark':                                _ce_cfg['status_remark'],
            'discoveryCategoryIds':                        st_data.get('discoveryCategoryIds', _ce_cfg['discovery_cat']),
            'productDescription *':                        description,
            'PRODUCT_IDENTIFIER *':                        'Set',
            'SET_NAME *':                                  'Set of 1',
            'SET_COUNT *':                                 1,
            'PACK_NAME *':                                 'Pack of 1',
            'PACK_OF *':                                   1,
            'IS_COMBO *':                                  'yes',
            'AVAILABLE_SIZES *':                           '1',
            'SET_DETAILS *':                               '1pc',
            'SET_DESCRIPTION *':                           '1pc of Smartphones',
            'PRODUCT_COLOR *':                             color,
            'ARTICLE_NUMBER *':                            article,
            'MODEL_NAME *':                                model_name,
            'PRODUCT_CONDITION *':                         _ce_cfg['product_condition'],
            'UNIT_OF_MEASUREMENT_SINGULAR *':              'Piece',
            'UNIT_OF_MEASUREMENT_PLURAL *':                'Pieces',
            'UNIT_OF_MEASUREMENT_SINGULAR_ABBREVIATION *': 'Pc',
            'UNIT_OF_MEASUREMENT_PLURAL_ABBREVIATION *':   'Pcs',
            'SELLER_SKU_ID *':                             sku_raw,
            'PACKAGING_TYPE *':                            packing,
            'DESCRIPTION':                                 '',
            'BATTERY_CAPACITY_MAH *':                      battery,
            'CHARGING_TYPE_SUPPORTED *':                   charging,
            'COUNTRY_OF_ORIGIN *':                         country,
            'EAN *':                                       '',
            'IMPORTED_BY':                                 '',
            'KEY_FEATURES':                                '',
            'MANUFACTURING_YEAR':                          _ce_cfg['manufacturing_year'],
            'PACKAGE_CONTENTS *':                          package_contents,
            'PORT_TYPE':                                   '',
            'PRODUCT_BREADTH *':                           B,
            'PRODUCT_DIMENSION_UOM *':                     dim_uom,
            'PRODUCT_HEIGHT *':                            H,
            'PRODUCT_LENGTH *':                            L,
            'PRODUCT_WEIGHT_IN_KG *':                      weight_clean,
            'PRODUCT_MANUFACTURING_CITY':                  '',
            'PRODUCT_MANUFACTURING_STATE':                 '',
            'WARRANTY':                                    warranty,
            'MANUFACTURER':                                '',
            'OPERATING_SYSTEM_OS':                         os,
            'OS_VERSION':                                  '',
            'DISPLAY_SIZE *':                              screen_size,
            'DISPLAY_TYPE *':                              display_type,
            'DISPLAY_RESOLUTION':                          '',
            'RAM *':                                       ram,
            'INTERNAL_STORAGE *':                          storage,
            'EXPANDABLE_STORAGE *':                        '',
            'SIM_TYPE *':                                  sim_type,
            'BLUETOOTH_VERSION *':                         bluetooth,
            'EXPANDABLE_STORAGE_TYPE':                     '',
            'EXPANDABLE_STORAGE_CAPACITY_MAX':             '',
            'BATTERY_TYPE':                                '',
            'REMOVABLE_BATTERY':                           '',
            'HYBRID_SIM_SLOT':                             '',
            'NETWORK_TYPE_SUPPORTED':                      network,
            'AUDIO_JACK':                                  '',
            'FM_RADIO *':                                  '',
            'TORCH_OR_FLASHLIGHT *':                       '',
            'RAM_ROM *':                                   ram_rom,
            'PROCESSOR_BRAND_AND_MODEL_NAME':              extract_from_description(desc, 'processor'),
            'NUMBER_OF_PROCESSOR_CORES *':                 proc_core,
            'PRIMARY_CAMERA_RESOLUTION *':                 back_cam,
            'FRONT_CAMERA_RESOLUTION *':                   front_cam,
            'REAR_FLASH':                                  '',
            'SIM_SIZE':                                    '',
            'WIFI':                                        '',
            'FINGERPRINT_SENSOR':                          '',
            'CLOCK_SPEED':                                 '',
            'REFRESH_RATE':                                '',
            'TOUCHSCREEN_TYPE':                            '',
            'PRIMARY_CAMERA_SETUP':                        '',
            'FRONT_FLASH':                                 '',
            'VIDEO_RECORDING_RESOLUTION':                  '',
            'FAST_CHARGING_WATTAGE':                       '',
            'WIRELESS_CHARGING_SUPPORT':                   '',
            'GPS_SUPPORT':                                 '',
            'NFC_SUPPORT':                                 '',
            'INFRARED_IR_BLASTER':                         '',
            'FINGERPRINT_SENSOR_POSITION':                 '',
            'FACE_UNLOCK':                                 '',
            'WATER_RESISTANCE_RATING':                     '',
            'hsnCode *':                                   hsn,
            'gstPercentage *':                             gst,
            'cgstShare *':                                 _ce_cfg['gst_cgst'],
            'sgstShare *':                                 _ce_cfg['gst_sgst'],
            'igstShare *':                                 _ce_cfg['gst_igst'],
            'cess':                                        '',
            'sinTax':                                      '',
            'vatPercentage':                               '',
            'otherCess':                                   '',
            'validityPeriodStartDate':                     '',
            'validityPeriodEndDate':                       '',
            'declarationForm':                             '',
            'taxMasterStatus':                             _ce_cfg['tax_master_status'],
        }

        for col_name, val in row_data.items():
            if col_name in tcol and val is not None and str(val) not in ('None',''):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

    return filled, skipped


# ═══════════════════════════════════════════════════════════════
# APPAREL & FASHION MODULE - UPDATED FOR MASTER TEMPLATE
# Reads 196 PVs dynamically from Apparel & Fashion Master Template
# Auto-detects JT_PV type (Top/Bottom/Top & Bottom) from PV name
# ═══════════════════════════════════════════════════════════════

from openpyxl.styles import PatternFill, Font
from openpyxl import load_workbook
from copy import copy
import os

# ──────────────────────────────────────────────────────────────
# MASTER TEMPLATE PATH
# ──────────────────────────────────────────────────────────────

AP_MASTER_TEMPLATE = os.path.join(
    os.path.dirname(__file__),
    'Apparel & Fashion - Master Template.xlsx'
)

# ──────────────────────────────────────────────────────────────
# LOAD PV MAPPING FROM MASTER TEMPLATE AT STARTUP
# ──────────────────────────────────────────────────────────────

def _load_pv_mapping_from_master():
    """Load all 196 PVs from master template PV to L4 Mapping sheet"""
    try:
        pv_map = pd.read_excel(AP_MASTER_TEMPLATE, sheet_name='PV to L4 Mapping')
        mapping = {}
        for _, row in pv_map.iterrows():
            pv_name = str(row['Product Vertical']).strip()
            pv_key = pv_name.lower().replace(' ', '_')
            mapping[pv_key] = {
                'pv_name': pv_name,
                'type': row.get('Type', 'SET'),
                'industry_category': row.get('Industry Category', 'Apparels & Fashion'),
                'industry_sub_category': row.get('Sub Category', ''),
                'industry_product_type': row.get('Product Type', ''),
                'industry_sub_type': row.get('Product Sub-type', pv_name),
                'jt_pv': _determine_jt_pv(pv_name),  # Auto-determine type
            }
        return mapping
    except Exception as e:
        print(f"Warning: Could not load PV mapping from master template: {e}")
        return {}

# ──────────────────────────────────────────────────────────────
# DETERMINE JT_PV TYPE FROM PV NAME
# ──────────────────────────────────────────────────────────────

def _determine_jt_pv(pv_name):
    """
    Intelligently determine JT_PV (Top/Bottom/Top & Bottom) from PV name
    
    TOP: Shirts, Tops, Blouses, Kurtis, Sweaters, Hoodies, Jackets, Vests, Cardigans
    BOTTOM: Jeans, Pants, Trousers, Skirts, Leggings, Shorts, Capris, Joggers, Track Pants
    TOP & BOTTOM: Dresses, Gowns, Sarees, Jumpsuits, Co-Ords, Sets, Rompers, Lehengas
    """
    pv_lower = pv_name.lower()
    
    # TOP keywords
    top_keywords = ['shirt', 'top', 'blouse', 'kurti', 'tunic', 'sweater', 'pullover',
                    'hoodie', 'jacket', 'vest', 'camisole', 'tank', 'cardigan', 'shrug',
                    'thermal top', 'casual shirt', 'formal shirt', 'polo']
    
    # BOTTOM keywords  
    bottom_keywords = ['jean', 'pant', 'trouser', 'skirt', 'legging', 'short', 'capri',
                       'tight', 'jogger', 'track pant', 'thermal trouser', 'jogging pant',
                       'harem']
    
    # TOP & BOTTOM keywords
    topbottom_keywords = ['dress', 'gown', 'saree', 'jumpsuit', 'co-ord', 'co ord',
                          'set', 'romper', 'playsuit', 'suit', 'ethnic', 'lehenga',
                          'sharara', 'kurti set']
    
    # Check TOP & BOTTOM first (most specific)
    for kw in topbottom_keywords:
        if kw in pv_lower:
            return "Top & Bottom"
    
    # Then BOTTOM
    for kw in bottom_keywords:
        if kw in pv_lower:
            return "Bottom"
    
    # Then TOP
    for kw in top_keywords:
        if kw in pv_lower:
            return "Top"
    
    # Default: Top & Bottom
    return "Top & Bottom"

# Load mapping at startup
AP_PV_MAPPING = _load_pv_mapping_from_master()

# ──────────────────────────────────────────────────────────────
# DEFAULT CONFIG
# ──────────────────────────────────────────────────────────────

AP_DEFAULT_CONFIG = {
    "brands": {},
    "biz_cat_id": "BCAT-139439",
    "biz_cat_name": "Apparel & Fashion",
    "catalog_status": "ACTIVE",
    "status_remark": "Ready to Launch",
    "tax_master_status": "active",
    "gst_cgst": 50,
    "gst_sgst": 50,
    "gst_igst": 0,
    "country_of_origin": "India",
    "product_condition": "Fresh",
    "manufacturing_year": "2026",
    "discovery_cat": "DISCAT-135530",
    
    # Auto-loaded from master template (196 PVs)
    "pv_mapping": AP_PV_MAPPING,
}

# ──────────────────────────────────────────────────────────────
# MANDATORY PAV ATTRIBUTES BY JT_PV
# ──────────────────────────────────────────────────────────────

MANDATORY_PAV_ATTRS = {
    "Top": [
        "Jpin", "Title", "PvId", "PvName", "BrandId", "BrandName",
        "ImageURL1", "CatalogStatus", "StatusRemark", "USER_TYPE",
        "TOP_COLOR", "TOP_FABRIC_MATERIAL", "TOP_LENGTH", "TOP_SIZE",
        "TOP_TYPE", "TOP_CLOSURE", "NECK_TYPE", "SLEEVE_TYPE",
        "FIT", "OCCASION", "COUNTRY_OF_ORIGIN"
    ],
    "Bottom": [
        "Jpin", "Title", "PvId", "PvName", "BrandId", "BrandName",
        "ImageURL1", "CatalogStatus", "StatusRemark", "USER_TYPE",
        "BOTTOM_COLOR", "BOTTOM_FABRIC_MATERIAL", "BOTTOM_LENGTH",
        "BOTTOM_SIZE", "BOTTOM_TYPE", "BOTTOM_CLOSURE", "RISE", "FIT",
        "NUMBER_OF_POCKETS", "DISTRESS", "OCCASION", "COUNTRY_OF_ORIGIN"
    ],
    "Top & Bottom": [
        "Jpin", "Title", "PvId", "PvName", "BrandId", "BrandName",
        "ImageURL1", "CatalogStatus", "StatusRemark", "USER_TYPE",
        "TOP_COLOR", "TOP_FABRIC_MATERIAL", "TOP_TYPE",
        "BOTTOM_COLOR", "BOTTOM_FABRIC_MATERIAL", "BOTTOM_TYPE",
        "FIT", "OCCASION", "COUNTRY_OF_ORIGIN"
    ]
}

# ──────────────────────────────────────────────────────────────
# GREEN COLOR FOR MANDATORY FIELDS IN PAV
# ──────────────────────────────────────────────────────────────

MANDATORY_FILL = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
MANDATORY_FONT = Font(bold=True, color="FFFFFF")

# ──────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ──────────────────────────────────────────────────────────────

def _ap_mandatory(val):
    """Return value or '#' if empty"""
    v = safe(val)
    return v if v else '#'

def _ap_apply_pav_colors(ws, headers, mandatory_cols):
    """Apply green coloring to mandatory PAV columns"""
    for col_idx, header in enumerate(headers, 1):
        if header in mandatory_cols:
            cell = ws.cell(1, col_idx)
            cell.fill = MANDATORY_FILL
            cell.font = MANDATORY_FONT

def _ap_get_pav_headers(jt_pv):
    """Return PAV headers based on JT_PV type from Master Template"""
    base = [
        'Jpin', 'Title', 'PvId', 'PvName', 'BrandId', 'BrandName',
        'ImageURL1', 'ImageURL2', 'CatalogStatus', 'StatusRemark',
    ]
    
    if jt_pv == "Top":
        top_attrs = ['USER_TYPE', 'TOP_COLOR', 'TOP_FABRIC_MATERIAL', 'TOP_LENGTH', 
                     'TOP_SIZE', 'TOP_TYPE', 'TOP_CLOSURE', 'NECK_TYPE', 'SLEEVE_TYPE',
                     'FIT', 'STRETCHABILITY', 'GSM', 'PATTERN']
        shared = ['CLOSURE_TYPE', 'COUNTRY_OF_ORIGIN', 'EAN', 'IMPORTED_BY',
                  'MANUFACTURING_YEAR', 'MATERIAL', 'PRODUCT_WEIGHT_IN_KG',
                  'PRODUCT_DIMENSION_UOM', 'PRODUCT_LENGTH', 'PRODUCT_BREADTH',
                  'PRODUCT_HEIGHT', 'KEY_FEATURES', 'MANUFACTURER']
        return base + top_attrs + shared
    elif jt_pv == "Bottom":
        bottom_attrs = ['USER_TYPE', 'BOTTOM_COLOR', 'BOTTOM_FABRIC_MATERIAL', 
                        'BOTTOM_LENGTH', 'BOTTOM_SIZE', 'BOTTOM_TYPE', 'BOTTOM_CLOSURE',
                        'RISE', 'FIT', 'STRETCHABILITY', 'DISTRESS', 'NUMBER_OF_POCKETS',
                        'FADE']
        shared = ['CLOSURE_TYPE', 'COUNTRY_OF_ORIGIN', 'EAN', 'IMPORTED_BY',
                  'MANUFACTURING_YEAR', 'MATERIAL', 'PRODUCT_WEIGHT_IN_KG',
                  'PRODUCT_DIMENSION_UOM', 'PRODUCT_LENGTH', 'PRODUCT_BREADTH',
                  'PRODUCT_HEIGHT', 'KEY_FEATURES', 'MANUFACTURER']
        return base + bottom_attrs + shared
    else:  # Top & Bottom
        both_attrs = ['USER_TYPE', 'TOP_COLOR', 'TOP_FABRIC_MATERIAL', 'TOP_LENGTH',
                      'TOP_SIZE', 'TOP_TYPE', 'TOP_CLOSURE', 'NECK_TYPE', 'SLEEVE_TYPE',
                      'BOTTOM_COLOR', 'BOTTOM_FABRIC_MATERIAL', 'BOTTOM_LENGTH',
                      'BOTTOM_SIZE', 'BOTTOM_TYPE', 'BOTTOM_CLOSURE', 'RISE',
                      'FIT', 'STRETCHABILITY', 'DISTRESS', 'NUMBER_OF_POCKETS', 'FADE']
        shared = ['CLOSURE_TYPE', 'COUNTRY_OF_ORIGIN', 'EAN', 'IMPORTED_BY',
                  'MANUFACTURING_YEAR', 'MATERIAL', 'PRODUCT_WEIGHT_IN_KG',
                  'PRODUCT_DIMENSION_UOM', 'PRODUCT_LENGTH', 'PRODUCT_BREADTH',
                  'PRODUCT_HEIGHT', 'KEY_FEATURES', 'MANUFACTURER']
        return base + both_attrs + shared

def _ap_get_pv_by_subtype(subtype):
    """Get PV config by industry sub-type from master template mapping"""
    _ap_cfg = get_ap_config_from_disk()
    pv_mapping = _ap_cfg.get('pv_mapping', AP_PV_MAPPING)
    
    subtype_lower = subtype.lower().replace(' ', '_')
    
    # Exact match
    if subtype_lower in pv_mapping:
        pv_data = pv_mapping[subtype_lower]
        return {
            'pv_name': pv_data['pv_name'],
            'industry_category': pv_data['industry_category'],
            'industry_sub_category': pv_data['industry_sub_category'],
            'industry_product_type': pv_data['industry_product_type'],
            'industry_sub_type': pv_data['industry_sub_type'],
            'jt_pv': pv_data['jt_pv'],
        }
    
    # Partial match
    for key, pv_data in pv_mapping.items():
        if subtype_lower in key or key in subtype_lower:
            return {
                'pv_name': pv_data['pv_name'],
                'industry_category': pv_data['industry_category'],
                'industry_sub_category': pv_data['industry_sub_category'],
                'industry_product_type': pv_data['industry_product_type'],
                'industry_sub_type': pv_data['industry_sub_type'],
                'jt_pv': pv_data['jt_pv'],
            }
    
    # Default: auto-determine from name
    return {
        'pv_name': subtype,
        'industry_category': 'Apparels & Fashion',
        'industry_sub_category': '',
        'industry_product_type': '',
        'industry_sub_type': subtype,
        'jt_pv': _determine_jt_pv(subtype),
    }


# ═══════════════════════════════════════════════════════════════
# IN-MEMORY FILE STORAGE
FILE_STORE = {}

# ═══════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        try:
            html_path = os.path.join(os.path.dirname(__file__), 'templates', 'index.html')
            if not os.path.exists(html_path):
                html_path = os.path.join(os.path.dirname(__file__), 'index.html')
            with open(html_path, 'r', encoding='utf-8') as f:
                html = f.read()
            html = html.replace("{{ user_email|default('', true) }}", '')
            return html
        except Exception as e2:
            import traceback
            return f"<h1>Template Error</h1><pre>{traceback.format_exc()}</pre>", 500

@app.route('/subtypes')
def get_subtypes():
    return jsonify({'subtypes': PV_LIST})

@app.route('/ce_subtypes')
def get_ce_subtypes():
    return jsonify({'subtypes': CE_PV_LIST})

@app.route('/ap_categories')
def get_ap_categories():
    return jsonify({'categories': AP_CATEGORIES})

@app.route('/config', methods=['GET'])
def config_get_route():
    return jsonify(get_config())

@app.route('/ce_config', methods=['GET'])
def ce_config_get_route():
    return jsonify(get_ce_config_from_disk())

@app.route('/ap_config', methods=['GET'])
def ap_config_get_route():
    return jsonify(get_ap_config_from_disk())

@app.route('/config', methods=['POST'])
def update_config():
    cfg  = get_config()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(CONFIG_PATH, cfg)
    write_log('anonymous', 'config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})

@app.route('/ce_config', methods=['POST'])
def update_ce_config():
    cfg  = get_ce_config_from_disk()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(CE_CONFIG_PATH, cfg)
    write_log('anonymous', 'ce_config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})

@app.route('/ap_config', methods=['POST'])
def update_ap_config():
    cfg  = get_ap_config_from_disk()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(AP_CONFIG_PATH, cfg)
    write_log('anonymous', 'ap_config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})

@app.route('/logs')
def get_logs():
    return jsonify({'logs': read_logs(500)})

@app.route('/detect_verticals', methods=['POST'])
def detect_verticals():
    try:
        dump_file = request.files.get('dump')
        if not dump_file:
            return jsonify({'verticals': []})
        xl     = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map  = build_col_map(all_dump, DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')
        if vert_col and vert_col in all_dump.columns:
            found   = [str(v).strip() for v in all_dump[vert_col].dropna().unique()
                       if str(v).strip() not in ('nan','None','')]
            matched = [v for v in found if v in SUBTYPE_MAP]
            return jsonify({'verticals': matched, 'all_found': found})
        return jsonify({'verticals': [], 'all_found': []})
    except Exception as e:
        return jsonify({'verticals': [], 'error': str(e)})

@app.route('/detect_ce_verticals', methods=['POST'])
def detect_ce_verticals():
    try:
        dump_file = request.files.get('dump')
        if not dump_file:
            return jsonify({'verticals': []})
        xl     = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map  = build_col_map(all_dump, CE_DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')
        if vert_col and vert_col in all_dump.columns:
            found   = [str(v).strip() for v in all_dump[vert_col].dropna().unique()
                       if str(v).strip() not in ('nan','None','')]
            matched = [v for v in found if v in CE_SUBTYPE_MAP]
            return jsonify({'verticals': matched, 'all_found': found})
        return jsonify({'verticals': [], 'all_found': []})
    except Exception as e:
        return jsonify({'verticals': [], 'error': str(e)})

@app.route('/detect_ap_categories', methods=['POST'])
def detect_ap_categories():
    """Auto-detect apparel categories from an uploaded listing file."""
    try:
        dump_file = request.files.get('dump')
        if not dump_file:
            return jsonify({'categories': []})
        xl     = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map  = build_col_map(all_dump, AP_DUMP_COL_HINTS)
        sub_type_col = col_map.get('ind_sub_type')
        if sub_type_col and sub_type_col in all_dump.columns:
            found = [str(v).strip() for v in all_dump[sub_type_col].dropna().unique()
                     if str(v).strip() not in ('nan','None','')]
            matched = [v for v in found if any(
                cat.lower() in v.lower() or v.lower() in cat.lower() for cat in AP_CATEGORIES
            )]
            return jsonify({'categories': matched if matched else found, 'all_found': found})
        return jsonify({'categories': AP_CATEGORIES, 'all_found': []})
    except Exception as e:
        return jsonify({'categories': AP_CATEGORIES, 'error': str(e)})

@app.route('/process', methods=['POST'])
def process():
    try:
        subtypes_raw = request.form.get('subtypes', '')
        try:    subtypes = json.loads(subtypes_raw)
        except: subtypes = [s.strip() for s in subtypes_raw.split(',') if s.strip()]

        inline_cfg_raw = request.form.get('config', '')
        if inline_cfg_raw:
            try:
                inline_cfg = json.loads(inline_cfg_raw)
                if inline_cfg.get('brands'):
                    inline_cfg['brands'] = normalize_brands(inline_cfg['brands'])
                try:
                    disk_cfg = get_config()
                    disk_cfg.update(inline_cfg)
                    _save_config(CONFIG_PATH, disk_cfg)
                except: pass
            except Exception as e:
                print(f'inline config parse error: {e}')

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not subtypes:
            return jsonify({'error': 'Please select at least one SubType'}), 400
        if not dump_file:
            return jsonify({'error': 'Dump / listing file is required'}), 400
        for st in subtypes:
            if st not in SUBTYPE_MAP:
                return jsonify({'error': f'SubType "{st}" not found in template'}), 400

        dump_bytes = dump_file.read()
        xl         = pd.ExcelFile(io.BytesIO(dump_bytes))
        frames     = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if all_dump.empty:
            return jsonify({'error': 'Could not read any data from dump file'}), 400

        col_map  = build_col_map(all_dump, DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')

        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf  = bxl.parse(sname)
                    bcol = build_col_map(bdf, BASE_COL_HINTS)
                    if 'article' in bcol:
                        existing_articles |= set(bdf[bcol['article']].dropna().astype(str).str.strip().str.upper())
                    if 'sku' in bcol:
                        existing_skus |= set(bdf[bcol['sku']].dropna().astype(str).str.strip().str.upper())
                except: pass

        results, all_skipped, grand_filled = [], [], 0
        preview_rows, preview_cols = [], []

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for subtype in subtypes:
                if vert_col and vert_col in all_dump.columns:
                    mask     = all_dump[vert_col].astype(str).str.strip().str.lower() == subtype.lower()
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        mask2    = all_dump[vert_col].astype(str).str.lower().str.contains(re.escape(subtype.lower()), na=False)
                        filtered = all_dump[mask2].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                wb, headers = get_template_wb_for_subtype(subtype)
                ws = wb.active
                filled, skipped = fill_template(
                    ws, headers, filtered, col_map, subtype, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_st = re.sub(r"[^\w\s-]", "", subtype).replace(" ", "_")
                fname   = f'filled_{safe_st}.xlsx'
                xls_buf = io.BytesIO()
                wb.save(xls_buf)
                zout.writestr(fname, xls_buf.getvalue())
                results.append({'subtype': subtype, 'filled': filled,
                                 'skipped': len(skipped), 'filename': fname})

                if not preview_cols:
                    pcols = ['title *','ChildSKU *','ARTICLE_NUMBER *','MRP *','SellingPrice *',
                             'PRODUCT_COLOR *','AVAILABLE_SIZES *','SET_DETAILS *',
                             'UPPER_MATERIAL *','SOLE_MATERIAL *','hsnCode *','SET_COUNT *']
                    preview_cols = [c for c in pcols if c in headers]
                    for r in range(2, min(filled + 2, 52)):
                        rdata = {c: ws.cell(r, headers.index(c)+1).value for c in preview_cols}
                        if any(v for v in rdata.values()):
                            preview_rows.append({**rdata, '_subtype': subtype})

        zip_buf.seek(0)
        if len(subtypes) == 1:
            safe_st  = re.sub(r"[^\w\s-]", "", subtypes[0]).replace(" ", "_")
            out_name = f'filled_{safe_st}.xlsx'
            out_ext  = '.xlsx'
            with zipfile.ZipFile(io.BytesIO(zip_buf.getvalue())) as zin:
                out_bytes = zin.read(results[0]['filename'])
        else:
            out_name  = 'filled_templates.zip'
            out_ext   = '.zip'
            out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {'bytes': out_bytes, 'filename': out_name,
                                   'ext': out_ext, 'created': time.time()}

        write_log('anonymous', 'catalog_generated',
                  f'subtypes={subtypes} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status': 'ok', 'grand_filled': grand_filled,
            'grand_skipped': len(all_skipped), 'results': results,
            'skipped_details': all_skipped[:50], 'preview': preview_rows,
            'preview_cols': preview_cols, 'download_token': file_token,
            'filename': out_name, 'is_zip': len(subtypes) > 1,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/process_ce', methods=['POST'])
def process_ce():
    try:
        subtypes_raw = request.form.get('subtypes', '')
        try:    subtypes = json.loads(subtypes_raw)
        except: subtypes = [s.strip() for s in subtypes_raw.split(',') if s.strip()]

        inline_cfg_raw = request.form.get('ce_config', '')
        if inline_cfg_raw:
            try:
                inline_cfg = json.loads(inline_cfg_raw)
                if inline_cfg.get('brands'):
                    inline_cfg['brands'] = normalize_brands(inline_cfg['brands'])
                try:
                    disk_cfg = get_ce_config_from_disk()
                    disk_cfg.update(inline_cfg)
                    _save_config(CE_CONFIG_PATH, disk_cfg)
                except: pass
            except Exception as e:
                print(f'inline ce_config parse error: {e}')

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not subtypes:
            return jsonify({'error': 'Please select at least one SubType'}), 400
        if not dump_file:
            return jsonify({'error': 'Dump / listing file is required'}), 400
        for st in subtypes:
            if st not in CE_SUBTYPE_MAP:
                return jsonify({'error': f'SubType "{st}" not found in CE template'}), 400

        dump_bytes = dump_file.read()
        xl         = pd.ExcelFile(io.BytesIO(dump_bytes))
        frames     = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if all_dump.empty:
            return jsonify({'error': 'Could not read any data from dump file'}), 400

        col_map  = build_col_map(all_dump, CE_DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')

        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf  = bxl.parse(sname)
                    bcol = build_col_map(bdf, CE_BASE_COL_HINTS)
                    if 'article' in bcol:
                        existing_articles |= set(bdf[bcol['article']].dropna().astype(str).str.strip().str.upper())
                    if 'sku' in bcol:
                        existing_skus |= set(bdf[bcol['sku']].dropna().astype(str).str.strip().str.upper())
                except: pass

        results, all_skipped, grand_filled = [], [], 0
        preview_rows, preview_cols = [], []

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for subtype in subtypes:
                if vert_col and vert_col in all_dump.columns:
                    mask     = all_dump[vert_col].astype(str).str.strip().str.lower() == subtype.lower()
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        mask2    = all_dump[vert_col].astype(str).str.lower().str.contains(re.escape(subtype.lower()), na=False)
                        filtered = all_dump[mask2].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                wb, headers = get_ce_template_wb_for_subtype(subtype)
                ws = wb.active
                filled, skipped = fill_ce_template(
                    ws, headers, filtered, col_map, subtype, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_st = re.sub(r"[^\w\s-]", "", subtype).replace(" ", "_")
                fname   = f'ce_filled_{safe_st}.xlsx'
                xls_buf = io.BytesIO()
                wb.save(xls_buf)
                zout.writestr(fname, xls_buf.getvalue())
                results.append({'subtype': subtype, 'filled': filled,
                                 'skipped': len(skipped), 'filename': fname})

                if not preview_cols:
                    pcols = ['title *','ChildSKU *','ARTICLE_NUMBER *','MRP *','SellingPrice *',
                             'PRODUCT_COLOR *','RAM *','INTERNAL_STORAGE *',
                             'DISPLAY_SIZE *','DISPLAY_TYPE *','hsnCode *','SET_COUNT *']
                    preview_cols = [c for c in pcols if c in headers]
                    for r in range(2, min(filled + 2, 52)):
                        rdata = {c: ws.cell(r, headers.index(c)+1).value for c in preview_cols}
                        if any(v for v in rdata.values()):
                            preview_rows.append({**rdata, '_subtype': subtype})

        zip_buf.seek(0)
        if len(subtypes) == 1:
            safe_st  = re.sub(r"[^\w\s-]", "", subtypes[0]).replace(" ", "_")
            out_name = f'ce_filled_{safe_st}.xlsx'
            out_ext  = '.xlsx'
            with zipfile.ZipFile(io.BytesIO(zip_buf.getvalue())) as zin:
                out_bytes = zin.read(results[0]['filename'])
        else:
            out_name  = 'ce_filled_templates.zip'
            out_ext   = '.zip'
            out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {'bytes': out_bytes, 'filename': out_name,
                                   'ext': out_ext, 'created': time.time()}

        write_log('anonymous', 'ce_catalog_generated',
                  f'subtypes={subtypes} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status': 'ok', 'grand_filled': grand_filled,
            'grand_skipped': len(all_skipped), 'results': results,
            'skipped_details': all_skipped[:50], 'preview': preview_rows,
            'preview_cols': preview_cols, 'download_token': file_token,
            'filename': out_name, 'is_zip': len(subtypes) > 1,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/process_ap', methods=['POST'])
def process_ap():
    """
    Apparel & Fashion processor. Produces a ZIP containing 4 files per category:
      - ap_JPIN_<category>.xlsx
      - ap_TaxMaster_<category>.xlsx
      - ap_ProductAttributeValue_<category>.xlsx
      - ap_L4_<category>.xlsx
    """
    try:
        categories_raw = request.form.get('categories', '')
        try:    categories = json.loads(categories_raw)
        except: categories = [s.strip() for s in categories_raw.split(',') if s.strip()]

        inline_cfg_raw = request.form.get('ap_config', '')
        if inline_cfg_raw:
            try:
                inline_cfg = json.loads(inline_cfg_raw)
                if inline_cfg.get('brands'):
                    inline_cfg['brands'] = normalize_brands(inline_cfg['brands'])
                try:
                    disk_cfg = get_ap_config_from_disk()
                    disk_cfg.update(inline_cfg)
                    _save_config(AP_CONFIG_PATH, disk_cfg)
                except: pass
            except Exception as e:
                print(f'inline ap_config parse error: {e}')

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not categories:
            return jsonify({'error': 'Please select at least one category'}), 400
        if not dump_file:
            return jsonify({'error': 'Listing file is required'}), 400

        dump_bytes = dump_file.read()
        xl         = pd.ExcelFile(io.BytesIO(dump_bytes))
        frames     = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if all_dump.empty:
            return jsonify({'error': 'Could not read any data from listing file'}), 400

        col_map = build_col_map(all_dump, AP_DUMP_COL_HINTS)

        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf  = bxl.parse(sname)
                    bcol = build_col_map(bdf, AP_BASE_COL_HINTS)
                    if 'article' in bcol:
                        existing_articles |= set(bdf[bcol['article']].dropna().astype(str).str.strip().str.upper())
                    if 'sku' in bcol:
                        existing_skus |= set(bdf[bcol['sku']].dropna().astype(str).str.strip().str.upper())
                except: pass

        results, all_skipped, grand_filled = [], [], 0
        preview_rows = []
        preview_cols = ['Title','Seller SKU ID','Article Number','Product Color',
                        'Available Sizes','Set Details','Set Count']

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for category in categories:
                # Filter rows by ind_sub_type
                sub_type_col = col_map.get('ind_sub_type')
                if sub_type_col and sub_type_col in all_dump.columns:
                    mask = all_dump[sub_type_col].astype(str).str.lower().str.strip() == category.lower()
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        mask2 = all_dump[sub_type_col].astype(str).str.lower().str.contains(
                            re.escape(category.lower()), na=False)
                        filtered = all_dump[mask2].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                wb_jpin, wb_tax, wb_pav, wb_l4, filled, skipped = fill_ap_files(
                    filtered, col_map, category, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_cat = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")

                for wb_obj, label in [
                    (wb_jpin, 'JPIN'),
                    (wb_tax,  'TaxMaster'),
                    (wb_pav,  'ProductAttributeValue'),
                    (wb_l4,   'L4'),
                ]:
                    fname   = f'ap_{label}_{safe_cat}.xlsx'
                    xls_buf = io.BytesIO()
                    wb_obj.save(xls_buf)
                    zout.writestr(fname, xls_buf.getvalue())

                results.append({
                    'category': category,
                    'filled':   filled,
                    'skipped':  len(skipped),
                    'files': [
                        f'ap_JPIN_{safe_cat}.xlsx',
                        f'ap_TaxMaster_{safe_cat}.xlsx',
                        f'ap_ProductAttributeValue_{safe_cat}.xlsx',
                        f'ap_L4_{safe_cat}.xlsx',
                    ]
                })

                # Build preview from JPIN sheet
                ws_jpin = wb_jpin.active
                jpin_headers = [ws_jpin.cell(1, c).value for c in range(1, ws_jpin.max_column + 1)]
                for r in range(2, min(filled + 2, 52)):
                    rdata = {}
                    for pc in preview_cols:
                        if pc in jpin_headers:
                            rdata[pc] = ws_jpin.cell(r, jpin_headers.index(pc)+1).value
                    if any(v for v in rdata.values()):
                        preview_rows.append({**rdata, '_category': category})

        out_name  = 'ap_filled_templates.zip'
        out_ext   = '.zip'
        zip_buf.seek(0)
        out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {'bytes': out_bytes, 'filename': out_name,
                                   'ext': out_ext, 'created': time.time()}

        write_log('anonymous', 'ap_catalog_generated',
                  f'categories={categories} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status':         'ok',
            'grand_filled':   grand_filled,
            'grand_skipped':  len(all_skipped),
            'results':        results,
            'skipped_details':all_skipped[:50],
            'preview':        preview_rows,
            'preview_cols':   preview_cols,
            'download_token': file_token,
            'filename':       out_name,
            'is_zip':         True,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/download/<token>')
def download(token):
    if '..' in token or '/' in token or '\\' in token: return 'Invalid', 400
    if token in FILE_STORE:
        file_data = FILE_STORE[token]
        ext   = file_data['ext']
        fname = request.args.get('filename', file_data['filename'])
        mtype = 'application/zip' if ext == '.zip' else \
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        write_log('anonymous', 'file_downloaded', fname)
        return send_file(io.BytesIO(file_data['bytes']), as_attachment=True,
                         download_name=fname, mimetype=mtype)
    tmpdir = tempfile.gettempdir()
    for ext in ['', '.zip', '.xlsx']:
        path = os.path.join(tmpdir, token + ext)
        if os.path.exists(path):
            fname = request.args.get('filename', 'filled_template' + ext)
            mtype = 'application/zip' if ext == '.zip' else \
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            write_log('anonymous', 'file_downloaded', fname)
            return send_file(path, as_attachment=True, download_name=fname, mimetype=mtype)
    return 'File not found', 404

@app.route('/test_brand')
def test_brand():
    cfg         = get_config()
    brands_raw  = cfg.get('brands', {})
    brands_dict = normalize_brands(brands_raw)
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))
    return jsonify({
        'step1_raw_from_config_json': brands_raw,
        'step2_after_normalize':      brands_dict,
        'step3_fallback_brand':       fallback_brand,
        'step4_fallback_id':          fallback_id,
        'step5_config_json_exists':   os.path.exists(CONFIG_PATH),
        'step6_config_json_content':  open(CONFIG_PATH).read() if os.path.exists(CONFIG_PATH) else 'FILE NOT FOUND',
        'conclusion': 'Brand will appear in title' if fallback_brand else 'NO BRAND — brands dict is empty!',
    })

@app.route('/debug_config')
def debug_config():
    cfg    = get_config()
    ce_cfg = get_ce_config_from_disk()
    ap_cfg = get_ap_config_from_disk()
    return jsonify({
        'config_path':           CONFIG_PATH,
        'ce_config_path':        CE_CONFIG_PATH,
        'ap_config_path':        AP_CONFIG_PATH,
        'config_file_exists':    os.path.exists(CONFIG_PATH),
        'ce_config_file_exists': os.path.exists(CE_CONFIG_PATH),
        'ap_config_file_exists': os.path.exists(AP_CONFIG_PATH),
        'footwear_brands':       cfg.get('brands', {}),
        'ce_brands':             ce_cfg.get('brands', {}),
        'ap_brands':             ap_cfg.get('brands', {}),
        'footwear_config':       cfg,
        'ce_config':             ce_cfg,
        'ap_config':             ap_cfg,
    })

if __name__ == '__main__':
    app.run(debug=False, port=5050)
