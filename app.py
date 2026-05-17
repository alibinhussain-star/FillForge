from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd, re, io, tempfile, os, json, copy, random, string, time, zipfile
from datetime import datetime
from email.mime.text import MIMEText
from openpyxl import load_workbook, Workbook

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
            'PACKAGING_TYPE *':                            packing,
            'GENDER *':                                    gender,
            'AGE_GROUP *':                                 '',
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
# APPAREL & FASHION MODULE — MULTI-PV SUPPORT
# ═══════════════════════════════════════════════════════════════

import re
from openpyxl import Workbook

# ── Default Config ──────────────────────────────────────────
AP_DEFAULT_CONFIG = {
    "brands":              {},
    "biz_cat_id":          "BCAT-139439",
    "biz_cat_name":        "Apparel & Fashion",
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
    "pv_config": {
        # ── JEANS ───────────────────────────────────────────
        "jeans": {
            "pv_id":   "PV-1914272259",
            "pv_name": "Men's Jeans",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Menswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Jeans",
            "super_category":        "Jeans",  # ← ADD THIS
        },
        "women's jeans": {
            "pv_id":   "PV-1914272940",
            "pv_name": "Women's Jeans",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Womenswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Jeans",
            "super_category":        "Jeans",  # ← ADD THIS
        },
        "boy's jeans": {
            "pv_id":   "PV-1914272259",
            "pv_name": "Boy's Jeans",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Boyswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Jeans",
            "super_category":        "Jeans",  # ← ADD THIS
        },
        # ── SHIRTS ──────────────────────────────────────────
        "men's casual shirts": {
            "pv_id":   "PV-1914273102",
            "pv_name": "Men's Casual Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Menswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Shirts",
            "super_category":        "Shirts",  # ← ADD THIS
        },
        "men's formal shirts": {
            "pv_id":   "PV-1914273102",
            "pv_name": "Men's Formal Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Menswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Shirts",
            "super_category":        "Shirts",  # ← ADD THIS
        },
        "boy's casual shirts": {
            "pv_id":   "PV-1914272626",
            "pv_name": "Boy's Casual Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Boyswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Shirts",
            "super_category":        "Shirts",  # ← ADD THIS
        },
        "boy's formal shirts": {
            "pv_id":   "PV-1914272626",
            "pv_name": "Boy's Formal Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Boyswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Shirts",
            "super_category":        "Shirts",  # ← ADD THIS
        },
        "women's shirts": {
            "pv_id":   "PV-1914273102",
            "pv_name": "Women's Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Womenswear",
            "industry_product_type": "Westernwear",
            "industry_sub_type":     "Shirts",
            "super_category":        "Shirts",  # ← ADD THIS
        },
        # ── T-SHIRTS ────────────────────────────────────────
        "men's casual t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Men's Casual T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Menswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "men's polo t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Men's Polo T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Menswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "women's t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Women's T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Womenswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "women's polo t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Women's Polo T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Womenswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "boy's casual t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Boy's Casual T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Boyswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "boy's polo t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Boy's Polo T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Boyswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "girl's t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Girl's T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Girlswear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        "baby casual t-shirts": {
            "pv_id":   "PV-1914273100",
            "pv_name": "Baby Casual T-Shirts",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Babywear",
            "industry_product_type": "Casual & Sports",
            "industry_sub_type":     "T-Shirts",
            "super_category":        "T-Shirt",  # ← ADD THIS
        },
        # ── SAREES ──────────────────────────────────────────
        "sarees & blouses": {
            "pv_id":   "PV-1914272903",
            "pv_name": "Women's Sarees",
            "industry_category":     "Apparels & Fashion",
            "industry_sub_category": "Womenswear",
            "industry_product_type": "Ethnicwear",
            "industry_sub_type":     "Sarees & Blouses",
            "super_category":        "Sarees",  # ← ADD THIS
        },
    },
}
# ── Listing-file column hints for Apparel (L4-style input files) ──
# These are COMMON across all super-categories. Per-category extras are handled dynamically.
AP_DUMP_COL_HINTS = {
    'type':              ['*Type','Type'],
    'ind_category':      ['*Industry Category','Industry Category'],
    'ind_sub_category':  ['*Industry Sub Category','Industry Sub Category'],
    'ind_product_type':  ['*Industry Product Type','Industry Product Type'],
    'ind_sub_type':      ['*Industry Product Sub-type','Industry Product Sub-type'],
    'product_name':      ['*Product Name','Product Name'],
    'product_desc':      ['*Product Description','Product Description'],
    'seller_sku':        ['*Seller SKU','Seller SKU'],
    'product_code':      ['*Product Code','Product Code'],
    'relationship':      ['*Relationship','Relationship'],
    'parent_product_id': ['*Parent Product Id','Parent Product Id'],
    'child_sku':         ['*Child SKU','Child SKU'],
    'quantity':          ['*Quantity','Quantity'],
    'set_name':          ['*Set Name','Set Name'],
    'hsn':               ['*HSN Code','HSN Code'],
    'gst':               ['*GST','GST'],
    'marketed_by':       ['Marketed By'],
    'country':           ['*Country Of Origin','Country Of Origin','Country of Origin'],
    'imported_by':       ['Imported By'],
    'ean':               ['EAN'],
    'moq':               ['*MOQ','MOQ'],
    'mrp':               ['*MRP','MRP'],
    'sp':                ['*Selling Price','Selling Price'],
    'weight':            ['*Product Weight (In KG)','Product Weight (In KG)'],
    'dims':              ['*Product Dimension (LXBXH)','Product Dimension (LXBXH)'],
    'mfg_year':          ['Manufacturing Year'],
    'unit_of_measure':   ['*Unit Of Measure','Unit Of Measure'],
    'dim_uom':           ['*Product Dimension UOM','Product Dimension UOM'],
    'gender':            ['*Gender','Gender'],
    'fabric':            ['*Select Fabric','Select Fabric'],
    'distress':          ['Distress'],
    'num_pockets':       ['Number of Pockets'],
    'trend':             ['Trend'],
    'fabric_composition':['Fabric Composition'],
    'fade':              ['Fade'],
    'fit':               ['Fit'],
    'stretch':           ['Stretch'],
    'waist_rise':        ['Waist Rise'],
    'waist_band':        ['Waist Band'],
    'closure':           ['*Closure','Closure'],
    'packing':           ['Packaging Type','*Packaging Type'],
    'length':            ['*Length','Length'],
    'pattern':           ['*Pattern','Pattern'],
    'color':             ['*Select color','Select color','*Select Color','Select Color'],
    'size':              ['*Size','Size'],
    'image':             ['*Main Image URL','Main Image URL'],
    'image2':            ['Other Image URL1','Other Image URL 1'],
    'image3':            ['Other Image URL2','Other Image URL 2'],
    'image4':            ['Other Image URL3','Other Image URL 3'],
    'image5':            ['Other Image URL4','Other Image URL 4'],
    'image6':            ['Other Image URL5','Other Image URL 5'],
    'brand':             ['*Brand Name','Brand Name','Brand'],
    'new_brand':         ['New Brand'],
    # ── T-Shirt / Shirt specific ──
    'neck_type':         ['Neck','Neck Type','*Neck Type'],
    'sleeve_length':     ['Sleeve Length','*Sleeve Length'],
    'multipack_set':     ['*Multipack Set','Multipack Set'],
    'occasion':          ['Occasion'],
    'hemline':           ['Hemline'],
    'shape':             ['Shape'],
    'set_includes':      ['Set Includes'],
    'bottom_type':       ['Bottom Type'],
    'work_type':         ['Work Type'],
    'stitch_type':       ['Stitch Type'],
    'border':            ['Border'],
    'collar':            ['Collar'],
    # ── Saree specific ──
    'blouse_fabric':     ['Blouse Fabric Material'],
    'blouse_included':   ['Blouse Included'],
    'blouse_neck':       ['Blouse Neck'],
    'blouse_sleeve':     ['Blouse Sleeve'],
    'blouse_type':       ['Blouse Type'],
    'saree_type':        ['Saree Type'],
    'saree_length':      ['Saree Length'],
    'fabric_type':       ['*Fabric Type','Fabric Type'],
}

AP_BASE_COL_HINTS = {
    'article': ['Product Code','*Product Code','Article Number'],
    'sku':     ['*Seller SKU','Seller SKU','Child SKU'],
}

# Apparel super-categories currently supported
AP_CATEGORIES = ['Jeans', 'Shirts', 'T-Shirt', 'Sarees']


# ═══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def _ap_normalize_gender(raw):
    """Convert MALE / Man / Men etc → Men's; FEMALE / Woman / Women → Women's; etc."""
    if not raw: return raw
    r = str(raw).strip().upper()
    if r in ('MALE','MAN','MEN','MEN\'S','MENS'): return "Men's"
    if r in ('FEMALE','WOMAN','WOMEN','WOMEN\'S','WOMENS'): return "Women's"
    if r in ('BOY','BOYS','BOY\'S'): return "Boy's"
    if r in ('GIRL','GIRLS','GIRL\'S'): return "Girl's"
    if r in ('BABY','BABIES','BABY\'S'): return "Baby's"
    return str(raw).strip()


def _ap_parse_set_count(qty_raw):
    """
    Handles qty formats:
      '5', '5pcs', '5 pcs', '5pc' → 5
      '1, 1, 1, 1, 1'             → 5 (sum)
      'Set of 5'                  → 5
    Returns (int count, '1, 1, ...' L4 quantity string)
    """
    s = str(qty_raw).strip()
    # Already 'Set of N'
    m = re.match(r'Set\s+of\s+(\d+)', s, re.IGNORECASE)
    if m:
        n = int(m.group(1))
        return n, ', '.join(['1'] * n)
    # comma-separated list like '1, 1, 1, 1, 1'
    parts = [x.strip() for x in s.split(',') if x.strip()]
    if len(parts) > 1 and all(re.match(r'^\d+$', p) for p in parts):
        total = sum(int(p) for p in parts)
        return total, ', '.join(parts)
    # plain number possibly with suffix
    m2 = re.match(r'^(\d+)', s)
    if m2:
        n = int(m2.group(1))
        return n, ', '.join(['1'] * n)
    return 0, ''


def _ap_parse_sizes(size_raw, super_category='Jeans'):
    """
    Parse *Size column. Handles:
      '28, 30, 32, 34, 36'  → numeric sizes (Jeans)
      'S, M, L, XL, 2XL'    → alpha sizes (Shirts, T-Shirts)
      'Free Size'           → free size (Sarees)
    Returns list of size strings.
    """
    s = str(size_raw).strip()
    if not s or s.lower() in ('nan', 'none', ''):
        return []
    if ',' in s:
        return [x.strip() for x in s.split(',') if x.strip()]
    if '-' in s and super_category in ('Jeans',):
        parts = [x.strip() for x in s.split('-') if x.strip() and re.match(r'^\d+$', x.strip())]
        if len(parts) > 1:
            return parts
    if ' ' in s:
        return [x.strip() for x in s.split() if x.strip()]
    return [s] if s else []


def _ap_build_set_fields(sizes_list, set_count):
    """
    Build set_details, set_description, set_composition for apparel.
    Each size gets qty = set_count // len(sizes), remainder distributed.
    """
    if not sizes_list:
        return '', '', ''
    base_qty = set_count // len(sizes_list) if sizes_list else 1
    remainder = set_count % len(sizes_list) if sizes_list else 0
    qtys = [base_qty + (1 if i < remainder else 0) for i in range(len(sizes_list))]
    pairs = list(zip(sizes_list, qtys))
    set_details  = ', '.join(f'{s}/{q}' for s, q in pairs)
    set_desc     = ', '.join(f'{q}pcs of {s}' for s, q in pairs)
    set_comp     = ' | '.join(f'Size {s} :- {q}' for s, q in pairs)
    return set_details, set_desc, set_comp


def _ap_pv_name_for_title(pv_name):
    """Remove gender prefix from PV name for title usage.
    e.g. 'Men's Casual Shirts' → 'Casual Shirts'
         'Men's Polo T-Shirts' → 'Polo T-Shirts'
    """
    if not pv_name:
        return pv_name
    s = str(pv_name).strip()
    for prefix in ("Men's ", "Women's ", "Boy's ", "Girl's ", "Baby's "):
        if s.startswith(prefix):
            return s[len(prefix):]
    return s


def _ap_make_title(super_category, brand, gender, fabric, length, pattern, product_type, color,
                     neck_type='', sleeve_length='', collar='', fit=''):
    """
    Form title based on super-category rules.

    JEANS:     Brand Gender Fabric Length Pattern ProductType, Color
    SHIRTS:    Brand Gender Fabric Collar Fit SleeveLength Pattern PVName(no gender), Color
    T-SHIRT:   Brand Gender Fabric Neck SleeveLength Pattern PVName(no gender), Color
    SAREES:    Brand Gender Fabric Size Pattern PVName(no gender), Color
    """
    sc = (super_category or 'Jeans').strip()

    if sc == 'Jeans':
        parts = [p for p in [brand, gender, fabric, length, pattern, product_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif sc == 'Shirts':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, gender, fabric, collar, fit, sleeve_length, pattern, pv_short] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif sc == 'T-Shirt':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, gender, fabric, neck_type, sleeve_length, pattern, pv_short] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif sc == 'Sarees':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, gender, fabric, length, pattern, pv_short] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    else:
        parts = [p for p in [brand, gender, fabric, length, pattern, product_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base


def _ap_make_internal_title(super_category, brand, product_code, gender, fabric, length, pattern,
                            product_type, color, set_name, set_details,
                            neck_type='', sleeve_length='', collar='', fit=''):
    """
    Form internal title based on super-category rules.

    JEANS:     Brand ProductCode Gender Fabric Length Pattern ProductType, Color, SetName (SetDetails)
    SHIRTS:    Brand ProductCode Gender Fabric Collar Fit SleeveLength Pattern PVName, Color, SetName (SetDetails)
    T-SHIRT:   Brand ProductCode Gender Fabric Neck SleeveLength Pattern PVName, Color, SetName (SetDetails)
    SAREES:    Brand ProductCode Gender Fabric Size Pattern PVName, Color, SetName (SetDetails)
    """
    sc = (super_category or 'Jeans').strip()

    if sc == 'Jeans':
        parts = [p for p in [brand, product_code, gender, fabric, length, pattern, product_type] if p]
        base = ' '.join(parts)
        suffix = f"{color}, {set_name} ({set_details})" if color else f"{set_name} ({set_details})"
        return f"{base}, {suffix}"

    elif sc == 'Shirts':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, product_code, gender, fabric, collar, fit, sleeve_length, pattern, pv_short] if p]
        base = ' '.join(parts)
        suffix = f"{color}, {set_name} ({set_details})" if color else f"{set_name} ({set_details})"
        return f"{base}, {suffix}"

    elif sc == 'T-Shirt':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, product_code, gender, fabric, neck_type, sleeve_length, pattern, pv_short] if p]
        base = ' '.join(parts)
        suffix = f"{color}, {set_name} ({set_details})" if color else f"{set_name} ({set_details})"
        return f"{base}, {suffix}"

    elif sc == 'Sarees':
        pv_short = _ap_pv_name_for_title(product_type)
        parts = [p for p in [brand, product_code, gender, fabric, length, pattern, pv_short] if p]
        base = ' '.join(parts)
        suffix = f"{color}, {set_name} ({set_details})" if color else f"{set_name} ({set_details})"
        return f"{base}, {suffix}"

    else:
        parts = [p for p in [brand, product_code, gender, fabric, length, pattern, product_type] if p]
        base = ' '.join(parts)
        suffix = f"{color}, {set_name} ({set_details})" if color else f"{set_name} ({set_details})"
        return f"{base}, {suffix}"


def _ap_get_pv_config(category_key, ap_cfg):
    """Get PV config dict for a given category keyword (case-insensitive)."""
    pv_cfg = ap_cfg.get('pv_config', AP_DEFAULT_CONFIG['pv_config'])
    for k, v in pv_cfg.items():
        if k.lower() == category_key.lower():
            return v
    for k, v in pv_cfg.items():
        if category_key.lower() in k.lower() or k.lower() in category_key.lower():
            return v
    return {}

def _ap_detect_pv_from_row(drow, col_map, ap_cfg, category_key=None):
    """
    Auto-detect PV config from the row's *Industry Product Sub-type column.
    Falls back to category_key match, then first available PV if not found.
    Returns (pv_cfg_dict, super_category, detected_key)
    """
    pv_cfg_map = ap_cfg.get('pv_config', AP_DEFAULT_CONFIG['pv_config'])

    # ── Helper: clean raw values ──
    def _clean(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        s = str(val).strip()
        return '' if s.lower() in ('nan', 'none', '') else s

    sub_type_raw = ''
    sub_type_col = col_map.get('ind_sub_type')
    if sub_type_col:
        try:
            sub_type_raw = _clean(drow.get(sub_type_col, ''))
        except:
            sub_type_raw = ''

    # Fallback hint columns
    if not sub_type_raw:
        for hint_col in ['product_name', 'ind_product_type', 'type']:
            c = col_map.get(hint_col)
            if c:
                try:
                    val = _clean(drow.get(c, ''))
                    if val:
                        sub_type_raw = val
                        break
                except:
                    pass

    # ── 1) Match by sub_type_raw against PV config keys ──
    if sub_type_raw:
        sub_type_lower = sub_type_raw.lower()
        for k, v in pv_cfg_map.items():
            k_lower = k.lower()
            if k_lower == sub_type_lower or sub_type_lower in k_lower or k_lower in sub_type_lower:
                return v, v.get('super_category', 'Jeans'), k

    # ── 2) Match category_key against PV config keys OR super_category ──
    if category_key and pv_cfg_map:
        cat_lower = category_key.lower().strip()
        for k, v in pv_cfg_map.items():
            k_lower = k.lower()
            super_cat = v.get('super_category', '').lower()

            # Match against the config KEY itself (e.g. "T-Shirts" matches "men's casual t-shirts")
            if k_lower == cat_lower or cat_lower in k_lower or k_lower in cat_lower:
                return v, v.get('super_category', 'Jeans'), k

            # Match against super_category with normalization
            if super_cat:
                # Normalize: remove trailing 's', spaces, hyphens for fuzzy match
                def _norm(s):
                    return s.lower().replace('-', '').replace(' ', '').rstrip('s')
                if _norm(super_cat) == _norm(cat_lower):
                    return v, v.get('super_category', 'Jeans'), k

    # ── 3) Only fall back to first PV if no category_key was provided ──
    if not category_key and pv_cfg_map:
        first_k = next(iter(pv_cfg_map))
        first_v = pv_cfg_map[first_k]
        return first_v, first_v.get('super_category', 'Jeans'), first_k

    # Return empty if truly nothing found
    return {}, 'Jeans', ''


def _ap_derive_product_type(pv_name):
    """Derive PRODUCT_TYPE from PV name (e.g. 'Men's Jeans' → 'Jeans')."""
    if not pv_name:
        return ''
    pv_lower = str(pv_name).lower()
    for pt in ['jeans','track pants','shirts','camisole','slips','cargo',
               'casual shirts','formal shirts','polo t-shirts','casual t-shirts',
               't-shirts','sarees','blouses']:
        if pt in pv_lower:
            return pt.title()
    return _ap_pv_name_for_title(pv_name) or pv_name


# ═══════════════════════════════════════════════════════════════
# PER-SUPER-CATEGORY PAV HEADERS
# ═══════════════════════════════════════════════════════════════

def _ap_get_pav_headers(super_category):
    """Return PAV headers for a given super-category."""
    sc = (super_category or 'Jeans').strip()

    base = [
        'Jpin','Title','PvId','PvName','BrandId','BrandName',
        'ImageURL1','ImageURL2','CatalogStatus','StatusRemark',
        'USER_TYPE','DESCRIPTION','COUNTRY_OF_ORIGIN','EAN','IMPORTED_BY',
        'KEY_FEATURES','MANUFACTURING_YEAR',
        'PRODUCT_BREADTH','PRODUCT_DIMENSION_UOM','PRODUCT_HEIGHT','PRODUCT_LENGTH',
        'PRODUCT_TYPE','PRODUCT_WEIGHT_IN_KG',
        'PRODUCT_MANUFACTURING_CITY','PRODUCT_MANUFACTURING_STATE',
    ]

    if sc == 'Jeans':
        specific = [
            'CLOSURE_TYPE','DISTRESS','FABRIC_MATERIAL','FIT','LENGTH',
            'MANUFACTURER','NUMBER_OF_POCKETS','OCCASION','PATTERN','RISE','STRETCHABILITY',
        ]
    elif sc == 'Shirts':
        specific = [
            'CLOSURE_TYPE','FABRIC_MATERIAL','FIT','GSM','HEMLINE','LENGTH',
            'MANUFACTURER','NECK_TYPE','NUMBER_OF_POCKETS','OCCASION','PATTERN','SLEEVE_TYPE',
        ]
    elif sc == 'T-Shirt':
        specific = [
            'CLOSURE_TYPE','FABRIC_MATERIAL','FIT','GSM','HEMLINE','LENGTH',
            'MANUFACTURER','NECK_TYPE','NUMBER_OF_POCKETS','OCCASION','PATTERN','SLEEVE_TYPE',
        ]
    elif sc == 'Sarees':
        specific = [
            'BLOUSE_FABRIC_MATERIAL','BLOUSE_INCLUDED','BLOUSE_NECK','BLOUSE_SLEEVE',
            'BLOUSE_TYPE','FABRIC_MATERIAL','LENGTH','MANUFACTURER','OCCASION',
            'PATTERN','SAREE_TYPE',
        ]
    else:
        specific = ['PATTERN']

    return base + specific


# ═══════════════════════════════════════════════════════════════
# PER-SUPER-CATEGORY L4 HEADERS
# ═══════════════════════════════════════════════════════════════

def _ap_get_l4_headers(super_category):
    """Return L4 headers for a given super-category."""
    sc = (super_category or 'Jeans').strip()

    base = [
        '*Type','*Industry Category','*Industry Sub Category',
        '*Industry Product Type','*Industry Product Sub-type',
        '*Product Name','*Product Description','*Seller SKU','*Product Code',
        '*Relationship','*Parent Product Id','*Child SKU',
        '*Quantity','*Set Name','*HSN Code','*GST',
        'Marketed By','*Country Of Origin','Imported By','EAN',
        '*MOQ','*MRP','*Selling Price',
        '*Product Weight (In KG)','*Product Dimension (LXBXH)',
        'Manufacturing Year','*Unit Of Measure','*Product Dimension UOM',
    ]

    if sc == 'Jeans':
        specific = [
            '*Gender','*Select Fabric','Distress','Number of Pockets','Trend',
            'Fabric Composition','Fade','Fit','Stretch','Waist Rise','Waist Band',
            'Manufacturing Year','Closure','Packaging Type','Length',
            '*Select color','*Size',
            '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
            'Other Image URL4','Other Image URL5',
            '*Brand Name','New Brand','*Pattern',
        ]
    elif sc == 'Shirts':
        specific = [
            '*Gender','*Select Fabric','Distress','Number of Pockets','Trend',
            'Fabric Composition','Fade','Fit','Stretch','Waist Rise','Waist Band',
            'Manufacturing Year','Closure','Packaging Type','Length',
            '*Select color','*Size',
            '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
            'Other Image URL4','Other Image URL5',
            '*Brand Name','New Brand','*Pattern','Collar',
        ]
    elif sc == 'T-Shirt':
        specific = [
            '*Gender','*Select Fabric','*Pattern','*Multipack Set','Number of Pockets',
            'Fabric Composition','Fit','Occasion','Manufacturing Year',
            'Neck','Sleeve Length','Closure','Packaging Type',
            '*Select color','*Size',
            '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
            'Other Image URL4','Other Image URL5',
            '*Brand Name','New Brand',
        ]
    elif sc == 'Sarees':
        specific = [
            '*Select Fabric','*Fabric Type','*Gender','*Pattern','Occasion',
            'Hemline','Shape','Set Includes','Bottom Type','Fabric Composition',
            'Number of Pockets','Work Type','Stitch Type','Border',
            'Manufacturing Year','Neck','Sleeve Length','Closure','Product Type',
            'Packaging Type','Length',
            '*Select color','*Size',
            '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
            'Other Image URL4','Other Image URL5',
            '*Brand Name','New Brand',
        ]
    else:
        specific = [
            '*Gender','*Select Fabric',
            '*Select color','*Size',
            '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
            'Other Image URL4','Other Image URL5',
            '*Brand Name','New Brand','*Pattern',
        ]

    return base + specific


# ═══════════════════════════════════════════════════════════════
# MAIN GENERATOR
# ═══════════════════════════════════════════════════════════════

def fill_ap_files(rows_df, col_map, category_key, existing_articles, existing_skus):
    """
    Generate all 4 apparel output workbooks for one category.
    Auto-detects PV from row's *Industry Product Sub-type.
    Returns: (wb_jpin, wb_tax, wb_pav_dict, wb_l4_dict, filled_count, skipped_list)
    wb_pav_dict and wb_l4_dict are keyed by super_category.
    """
    _ap_cfg     = get_ap_config_from_disk()
    # Ensure all required keys exist with safe defaults
    _ap_cfg.setdefault('gst_cgst', 50)
    _ap_cfg.setdefault('gst_sgst', 50)
    _ap_cfg.setdefault('gst_igst', 0)
    _ap_cfg.setdefault('catalog_status', 'ACTIVE')
    _ap_cfg.setdefault('status_remark', 'Ready to Launch')
    _ap_cfg.setdefault('tax_master_status', 'active')
    _ap_cfg.setdefault('biz_cat_id', 'BCAT-139439')
    _ap_cfg.setdefault('biz_cat_name', 'Apparel & Fashion')
    _ap_cfg.setdefault('country_of_origin', 'India')
    _ap_cfg.setdefault('product_condition', 'Fresh')
    _ap_cfg.setdefault('manufacturing_year', '2026')
    _ap_cfg.setdefault('discovery_cat', 'DISCAT-135542')
    brands_dict = normalize_brands(_ap_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))

    # ── JPIN headers (same for all categories) ────────────────
    JPIN_HEADERS = [
        'JPIN','Title','Internal_Title','BrandID','BrandName','PVID','PVName',
        'Business Category Id','Business Category Name',
        'Product Identifier','Set Name','Set Count','Pack Name','Pack of','is Combo',
        'Available Sizes','Set Details','Set Description','Set Composition',
        'Product Color','Article Number','Model Name','Product Condition',
        'ImageURL1','ImageURL2','ImageURL3','ImageURL4','ImageURL5','ImageURL6',
        'VideoURL1','VideoURL2','SizeChartURL',
        'CatalogStatus','StatusRemark','CustomerDiscoveryCategories',
        'Singular Unit Of Measurement','Plural Unit Of Measurement',
        'Singular Unit Of Measurement Abbreviation','Plural Unit Of Measurement Abbreviation',
        'Seller SKU ID','Product Description',
        'CreatedTime','LastUpdatedTime','LastUpdatedBy','Ingestion Row Status','Exception',
    ]

    # ── TaxMaster headers (same for all categories) ───────────
    TAX_HEADERS = [
        'TaxMasterID','Jpin','Title','ProductVerticalId','ProductVerticalName',
        'hsnCode','sinTax','cess','vatPercentage','gstPercentage',
        'cgstComponentShare','sgstComponentShare','IgstComponentShare',
        'Validity_Period_Start','Validity_Period_End','declarationForm','otherCess','status',
    ]

    # ── SCM (Supply Chain Management) headers ─────────────────
    SCM_HEADERS = [
        'JPIN','Title','Net_Weight','Net_Weight_Measuring_Unit','DeadWeight',
        'VolumetricWeight','ShippingCalculationType',
        'L1-caseSize','L2-caseSize','L3-caseSize','L4-caseSize',
        'L1-packagingType','L2-packagingType','L3-packagingType','L4-packagingType',
        'L0-UnitShippingContainerType','L1-UnitShippingContainerType',
        'L2-UnitShippingContainerType','L3-UnitShippingContainerType',
        'L4-UnitShippingContainerType',
        'Fragile','Brittle',
        'length_l0','width_l0','height_l0',
        'length_l1','width_l1','height_l1',
        'length_l2','width_l2','height_l2',
        'length_l3','width_l3','height_l3',
        'length_l4','width_l4','height_l4',
        'volumetricweight_l1','volumetricweight_l2','volumetricweight_l3','volumetricweight_l4',
        'APMC Notified Commodity',
        'L1-deadWeight','L2-deadWeight','L3-deadWeight','L4-deadWeight',
        'Net_Quantity','Net_Quantity_Measuring_Unit',
        'CreatedTime','LastUpdatedTime','LastUpdatedBy',
    ]


    def _make_wb(headers, sheet_name):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        for ci, h in enumerate(headers, 1):
            ws.cell(1, ci).value = h
        return wb, ws

    wb_jpin, ws_jpin = _make_wb(JPIN_HEADERS, 'JPIN Template')
    wb_tax,  ws_tax  = _make_wb(TAX_HEADERS,  'TaxMaster')
    wb_scm,  ws_scm  = _make_wb(SCM_HEADERS,  'SCM')

    # PAV and L4 workbooks per super-category
    wb_pav_map = {}
    wb_l4_map  = {}
    row_idx_pav = {}
    row_idx_l4  = {}

    def _col(headers):
        return {h: i+1 for i, h in enumerate(headers) if h}

    tcol_jpin = _col(JPIN_HEADERS)
    tcol_tax  = _col(TAX_HEADERS)
    tcol_scm  = _col(SCM_HEADERS)

    def _write(ws, tcol, data, row_idx):
        for col_name, val in data.items():
            if col_name in tcol and val is not None and str(val) not in ('None',):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

    skipped, filled = [], 0

    # Filter to only "Parent" rows if Relationship column exists
    rel_col = col_map.get('relationship')
    if rel_col and rel_col in rows_df.columns:
        parent_mask = rows_df[rel_col].astype(str).str.strip().str.lower().isin(['parent','set'])
        work_df = rows_df[parent_mask].copy()
        if work_df.empty:
            work_df = rows_df.copy()
    else:
        work_df = rows_df.copy()

    for _, drow in work_df.iterrows():
        # ── Auto-detect PV and super-category from row ────────
        pv_cfg, super_category, detected_key = _ap_detect_pv_from_row(drow, col_map, _ap_cfg, category_key)

        if not pv_cfg:
            skipped.append({'sku': '', 'article': '', 'reason': f'No PV config found for sub-type: {detected_key}'})
            continue

        pv_id   = pv_cfg.get('pv_id', '')
        pv_name = pv_cfg.get('pv_name', category_key)
        ind_cat      = pv_cfg.get('industry_category', 'Apparels & Fashion')
        ind_sub_cat  = pv_cfg.get('industry_sub_category', '')
        ind_prod_type = pv_cfg.get('industry_product_type', '')
        ind_sub_type  = pv_cfg.get('industry_sub_type', category_key)

        # Lazy-init PAV and L4 workbooks per super-category
        if super_category not in wb_pav_map:
            pav_headers = _ap_get_pav_headers(super_category)
            wb_pav_map[super_category], _ = _make_wb(pav_headers, 'ProductAttributeValue')
            row_idx_pav[super_category] = 1
        if super_category not in wb_l4_map:
            l4_headers = _ap_get_l4_headers(super_category)
            wb_l4_map[super_category], _ = _make_wb(l4_headers, 'L4')
            row_idx_l4[super_category] = 1

        ws_pav = wb_pav_map[super_category].active
        ws_l4  = wb_l4_map[super_category].active
        tcol_pav = _col(_ap_get_pav_headers(super_category))
        tcol_l4  = _col(_ap_get_l4_headers(super_category))

        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand    = fallback_brand
            brand_id = fallback_id

        seller_sku  = safe(drow.get(col_map.get('seller_sku',''), ''))
        product_code= safe(drow.get(col_map.get('product_code',''), ''))
        article     = product_code if product_code else seller_sku

        if article.upper() in existing_articles or seller_sku.upper() in existing_skus:
            skipped.append({'sku': seller_sku, 'article': article, 'reason': 'Already exists in base data'})
            continue

        # ── Extract fields ────────────────────────────────────
        qty_raw     = safe(drow.get(col_map.get('quantity',''), ''))
        set_count, l4_qty_str = _ap_parse_set_count(qty_raw)
        if set_count == 0:
            sn_raw = safe(drow.get(col_map.get('set_name',''), ''))
            set_count, l4_qty_str = _ap_parse_set_count(sn_raw)

        size_raw    = safe(drow.get(col_map.get('size',''), ''))
        sizes_list  = _ap_parse_sizes(size_raw, super_category)
        avail_sizes = ', '.join(sizes_list)

        set_details, set_desc, set_comp = _ap_build_set_fields(sizes_list, set_count)
        set_name_str = f'Set of {set_count}'

        gender_raw  = safe(drow.get(col_map.get('gender',''), ''))
        gender      = _ap_normalize_gender(gender_raw)
        fabric      = safe(drow.get(col_map.get('fabric',''), ''))
        length      = safe(drow.get(col_map.get('length',''), ''))
        closure     = safe(drow.get(col_map.get('closure',''), ''))
        distress    = safe(drow.get(col_map.get('distress',''), '')) or '#'
        fit         = safe(drow.get(col_map.get('fit',''), '')) or '#'
        pattern_val = safe(drow.get(col_map.get('pattern',''), '')) or '#'
        color       = title_case_color(safe(drow.get(col_map.get('color',''), '')))
        product_type_val = _ap_derive_product_type(pv_name)

        # Category-specific fields
        neck_type_val    = safe(drow.get(col_map.get('neck_type',''), '')) or '#'
        sleeve_len_val   = safe(drow.get(col_map.get('sleeve_length',''), '')) or '#'
        collar_val       = safe(drow.get(col_map.get('collar',''), '')) or '#'
        multipack_val    = safe(drow.get(col_map.get('multipack_set',''), '')) or '#'
        occasion_val     = safe(drow.get(col_map.get('occasion',''), '')) or '#'
        hemline_val      = safe(drow.get(col_map.get('hemline',''), '')) or '#'
        shape_val        = safe(drow.get(col_map.get('shape',''), '')) or '#'
        set_includes_val = safe(drow.get(col_map.get('set_includes',''), '')) or '#'
        bottom_type_val  = safe(drow.get(col_map.get('bottom_type',''), '')) or '#'
        work_type_val    = safe(drow.get(col_map.get('work_type',''), '')) or '#'
        stitch_type_val  = safe(drow.get(col_map.get('stitch_type',''), '')) or '#'
        border_val       = safe(drow.get(col_map.get('border',''), '')) or '#'
        blouse_fabric_val= safe(drow.get(col_map.get('blouse_fabric',''), '')) or '#'
        blouse_incl_val  = safe(drow.get(col_map.get('blouse_included',''), '')) or '#'
        blouse_neck_val  = safe(drow.get(col_map.get('blouse_neck',''), '')) or '#'
        blouse_sleeve_val= safe(drow.get(col_map.get('blouse_sleeve',''), '')) or '#'
        blouse_type_val  = safe(drow.get(col_map.get('blouse_type',''), '')) or '#'
        saree_type_val   = safe(drow.get(col_map.get('saree_type',''), '')) or '#'
        fabric_type_val  = safe(drow.get(col_map.get('fabric_type',''), '')) or '#'

        img_url  = safe(drow.get(col_map.get('image',''), ''))
        img2_url = safe(drow.get(col_map.get('image2',''), ''))
        img3_url = safe(drow.get(col_map.get('image3',''), ''))
        img4_url = safe(drow.get(col_map.get('image4',''), ''))
        img5_url = safe(drow.get(col_map.get('image5',''), ''))

        product_desc = safe(drow.get(col_map.get('product_desc',''), ''))
        hsn_raw      = drow.get(col_map.get('hsn',''), '')
        gst_raw      = drow.get(col_map.get('gst',''), 5)
        mrp_raw      = drow.get(col_map.get('mrp',''), '')
        sp_raw       = drow.get(col_map.get('sp',''), '')
        moq_raw      = drow.get(col_map.get('moq',''), 1)
        weight_raw   = drow.get(col_map.get('weight',''), '')
        dim_raw      = safe(drow.get(col_map.get('dims',''), ''))
        dim_uom      = safe(drow.get(col_map.get('dim_uom',''), '')) or 'Cms'
        country      = safe(drow.get(col_map.get('country',''), '')) or _ap_cfg['country_of_origin']
        packing      = safe(drow.get(col_map.get('packing',''), '')) or 'Bundles'
        num_pockets  = safe(drow.get(col_map.get('num_pockets',''), ''))
        fabric_comp  = safe(drow.get(col_map.get('fabric_composition',''), ''))
        fade         = safe(drow.get(col_map.get('fade',''), ''))
        stretch      = safe(drow.get(col_map.get('stretch',''), ''))
        waist_rise   = safe(drow.get(col_map.get('waist_rise',''), ''))
        waist_band   = safe(drow.get(col_map.get('waist_band',''), ''))
        mfg_year     = safe(drow.get(col_map.get('mfg_year',''), '')) or _ap_cfg['manufacturing_year']

        try:    hsn = int(float(hsn_raw)) if str(hsn_raw).strip() not in ('','nan') else ''
        except: hsn = ''
        try:    gst = int(float(gst_raw))
        except: gst = 5
        try:    mrp = float(mrp_raw) if str(mrp_raw).strip() not in ('','nan') else ''
        except: mrp = ''
        try:    sp  = float(sp_raw)  if str(sp_raw).strip()  not in ('','nan') else ''
        except: sp  = ''
        try:    moq = int(float(moq_raw))
        except: moq = 1
        try:    weight = float(weight_raw) if str(weight_raw).strip() not in ('','nan') else ''
        except: weight = ''

        # ── Derived title fields (per super-category) ───────────
        title = _ap_make_title(
            super_category, brand, gender, fabric, length, pattern_val, product_type_val, color,
            neck_type=neck_type_val, sleeve_length=sleeve_len_val, collar=collar_val, fit=fit
        )
        internal_title = _ap_make_internal_title(
            super_category, brand, article, gender, fabric, length, pattern_val,
            product_type_val, color, set_name_str, set_details,
            neck_type=neck_type_val, sleeve_length=sleeve_len_val, collar=collar_val, fit=fit
        )

        filled  += 1
        row_idx  = filled + 1
        row_idx_pav[super_category] += 1
        row_idx_l4[super_category]  += 1
        pav_row_idx = row_idx_pav[super_category]
        l4_row_idx  = row_idx_l4[super_category]

        # ── JPIN row (same for all) ───────────────────────────
        jpin_row = {
            'JPIN':                                    '',
            'Title':                                   title,
            'Internal_Title':                          internal_title,
            'BrandID':                                 brand_id,
            'BrandName':                               brand,
            'PVID':                                    pv_id,
            'PVName':                                  pv_name,
            'Business Category Id':                    _ap_cfg['biz_cat_id'],
            'Business Category Name':                  _ap_cfg['biz_cat_name'],
            'Product Identifier':                      'Set',
            'Set Name':                                set_name_str,
            'Set Count':                               set_count,
            'Pack Name':                               'Pack of 1',
            'Pack of':                                 1,
            'is Combo':                                'yes',
            'Available Sizes':                         avail_sizes,
            'Set Details':                             set_details,
            'Set Description':                         set_desc,
            'Set Composition':                         set_comp,
            'Product Color':                           color,
            'Article Number':                          article,
            'Model Name':                              article,
            'Product Condition':                       _ap_cfg['product_condition'],
            'ImageURL1':                               img_url,
            'ImageURL2':                               img2_url,
            'ImageURL3':                               img3_url,
            'ImageURL4':                               img4_url,
            'ImageURL5':                               img5_url,
            'ImageURL6':                               '',
            'VideoURL1':                               '',
            'VideoURL2':                               '',
            'SizeChartURL':                            '',
            'CatalogStatus':                           _ap_cfg['catalog_status'],
            'StatusRemark':                            _ap_cfg['status_remark'],
            'CustomerDiscoveryCategories':             _ap_cfg['discovery_cat'],
            'Singular Unit Of Measurement':            'Piece',
            'Plural Unit Of Measurement':              'Pieces',
            'Singular Unit Of Measurement Abbreviation': 'Pc',
            'Plural Unit Of Measurement Abbreviation': 'Pcs',
            'Seller SKU ID':                           seller_sku,
            'Product Description':                     product_desc,
            'CreatedTime':                             '',
            'LastUpdatedTime':                         '',
            'LastUpdatedBy':                           '',
            'Ingestion Row Status':                    '',
            'Exception':                               '',
        }
        _write(ws_jpin, tcol_jpin, jpin_row, row_idx)

        # ── TaxMaster row (same for all) ──────────────────────
        tax_row = {
            'TaxMasterID':          '',
            'Jpin':                 '',
            'Title':                title,
            'ProductVerticalId':    pv_id,
            'ProductVerticalName':  pv_name,
            'hsnCode':              hsn,
            'sinTax':               '',
            'cess':                 '',
            'vatPercentage':        '',
            'gstPercentage':        gst,
            'cgstComponentShare':   _ap_cfg['gst_cgst'],
            'sgstComponentShare':   _ap_cfg['gst_sgst'],
            'IgstComponentShare':   _ap_cfg['gst_igst'],
            'Validity_Period_Start':'',
            'Validity_Period_End':  '',
            'declarationForm':      '',
            'otherCess':            '',
            'status':               _ap_cfg['tax_master_status'],
        }
        _write(ws_tax, tcol_tax, tax_row, row_idx)

        # ── SCM row (Supply Chain Management) ─────────────────
        scm_row = {
            'JPIN':                          '',
            'Title':                         title,
            'Net_Weight':                    0,
            'Net_Weight_Measuring_Unit':     'g',
            'DeadWeight':                    0.25,
            'VolumetricWeight':              0,
            'ShippingCalculationType':       'Dead Weight',
            'L1-caseSize':                   1,
            'L2-caseSize':                   0,
            'L3-caseSize':                   0,
            'L4-caseSize':                   0,
            'L1-packagingType':              'Bag',
            'L2-packagingType':              '',
            'L3-packagingType':              '',
            'L4-packagingType':              '',
            'L0-UnitShippingContainerType':  'Crate - Medium',
            'L1-UnitShippingContainerType':  'Bag',
            'L2-UnitShippingContainerType':  'Bag',
            'L3-UnitShippingContainerType':  '',
            'L4-UnitShippingContainerType':  '',
            'Fragile':                       'No',
            'Brittle':                       'No',
            'length_l0':                     0,
            'width_l0':                      0,
            'height_l0':                     0,
            'length_l1':                     24,
            'width_l1':                      24,
            'height_l1':                     33,
            'length_l2':                     0,
            'width_l2':                      0,
            'height_l2':                     0,
            'length_l3':                     0,
            'width_l3':                      0,
            'height_l3':                     0,
            'length_l4':                     0,
            'width_l4':                      0,
            'height_l4':                     0,
            'volumetricweight_l1':           4.021675,
            'volumetricweight_l2':           0,
            'volumetricweight_l3':           0,
            'volumetricweight_l4':           0,
            'APMC Notified Commodity':       'No',
            'L1-deadWeight':                 0,
            'L2-deadWeight':                 0,
            'L3-deadWeight':                 0,
            'L4-deadWeight':                 0,
            'Net_Quantity':                  1,
            'Net_Quantity_Measuring_Unit':   'Pc',
            'CreatedTime':                   '',
            'LastUpdatedTime':               '',
            'LastUpdatedBy':                 '',
        }
        _write(ws_scm, tcol_scm, scm_row, row_idx)

        # ── ProductAttributeValue row (per super-category) ────
        if super_category == 'Jeans':
            pav_row = {
                'Jpin':                          '',
                'Title':                         title,
                'PvId':                          pv_id,
                'PvName':                        pv_name,
                'BrandId':                       brand_id,
                'BrandName':                     brand,
                'ImageURL1':                     img_url,
                'ImageURL2':                     img2_url,
                'CatalogStatus':                 _ap_cfg['catalog_status'],
                'StatusRemark':                  _ap_cfg['status_remark'],
                'USER_TYPE':                     gender,
                'DESCRIPTION':                   '',
                'COUNTRY_OF_ORIGIN':             country,
                'EAN':                           '',
                'IMPORTED_BY':                   '',
                'KEY_FEATURES':                  '',
                'MANUFACTURING_YEAR':            '',
                'PRODUCT_BREADTH':               0,
                'PRODUCT_DIMENSION_UOM':         0,
                'PRODUCT_HEIGHT':                0,
                'PRODUCT_LENGTH':                0,
                'PRODUCT_TYPE':                  product_type_val,
                'PRODUCT_WEIGHT_IN_KG':          0,
                'PRODUCT_MANUFACTURING_CITY':    '',
                'PRODUCT_MANUFACTURING_STATE':   '',
                'CLOSURE_TYPE':                  closure,
                'DISTRESS':                      distress,
                'FABRIC_MATERIAL':               fabric if fabric else '#',
                'FIT':                           fit,
                'LENGTH':                        length if length else '#',
                'MANUFACTURER':                  '',
                'NUMBER_OF_POCKETS':             '',
                'OCCASION':                      '',
                'PATTERN':                       pattern_val,
                'RISE':                          '',
                'STRETCHABILITY':                '',
            }
        elif super_category == 'Shirts':
            pav_row = {
                'Jpin':                          '',
                'Title':                         title,
                'PvId':                          pv_id,
                'PvName':                        pv_name,
                'BrandId':                       brand_id,
                'BrandName':                     brand,
                'ImageURL1':                     img_url,
                'ImageURL2':                     img2_url,
                'CatalogStatus':                 _ap_cfg['catalog_status'],
                'StatusRemark':                  _ap_cfg['status_remark'],
                'USER_TYPE':                     gender,
                'DESCRIPTION':                   '',
                'COUNTRY_OF_ORIGIN':             country,
                'EAN':                           '',
                'IMPORTED_BY':                   '',
                'KEY_FEATURES':                  '',
                'MANUFACTURING_YEAR':            '',
                'PRODUCT_BREADTH':               0,
                'PRODUCT_DIMENSION_UOM':         0,
                'PRODUCT_HEIGHT':                0,
                'PRODUCT_LENGTH':                0,
                'PRODUCT_TYPE':                  product_type_val,
                'PRODUCT_WEIGHT_IN_KG':          0,
                'PRODUCT_MANUFACTURING_CITY':    '',
                'PRODUCT_MANUFACTURING_STATE':   '',
                'CLOSURE_TYPE':                  closure,
                'FABRIC_MATERIAL':               fabric if fabric else '#',
                'FIT':                           fit,
                'GSM':                           '',
                'HEMLINE':                       '',
                'LENGTH':                        length if length else '#',
                'MANUFACTURER':                  '',
                'NECK_TYPE':                     collar_val,
                'NUMBER_OF_POCKETS':             '',
                'OCCASION':                      '',
                'PATTERN':                       pattern_val,
                'SLEEVE_TYPE':                   sleeve_len_val,
            }
        elif super_category == 'T-Shirt':
            pav_row = {
                'Jpin':                          '',
                'Title':                         title,
                'PvId':                          pv_id,
                'PvName':                        pv_name,
                'BrandId':                       brand_id,
                'BrandName':                     brand,
                'ImageURL1':                     img_url,
                'ImageURL2':                     img2_url,
                'CatalogStatus':                 _ap_cfg['catalog_status'],
                'StatusRemark':                  _ap_cfg['status_remark'],
                'USER_TYPE':                     gender,
                'DESCRIPTION':                   '',
                'COUNTRY_OF_ORIGIN':             country,
                'EAN':                           '',
                'IMPORTED_BY':                   '',
                'KEY_FEATURES':                  '',
                'MANUFACTURING_YEAR':            '',
                'PRODUCT_BREADTH':               0,
                'PRODUCT_DIMENSION_UOM':         0,
                'PRODUCT_HEIGHT':                0,
                'PRODUCT_LENGTH':                0,
                'PRODUCT_TYPE':                  product_type_val,
                'PRODUCT_WEIGHT_IN_KG':          0,
                'PRODUCT_MANUFACTURING_CITY':    '',
                'PRODUCT_MANUFACTURING_STATE':   '',
                'CLOSURE_TYPE':                  closure,
                'FABRIC_MATERIAL':               fabric if fabric else '#',
                'FIT':                           fit,
                'GSM':                           '',
                'HEMLINE':                       '',
                'LENGTH':                        length if length else '#',
                'MANUFACTURER':                  '',
                'NECK_TYPE':                     neck_type_val,
                'NUMBER_OF_POCKETS':             '',
                'OCCASION':                      '',
                'PATTERN':                       pattern_val,
                'SLEEVE_TYPE':                   sleeve_len_val,
            }
        elif super_category == 'Sarees':
            pav_row = {
                'Jpin':                          '',
                'Title':                         title,
                'PvId':                          pv_id,
                'PvName':                        pv_name,
                'BrandId':                       brand_id,
                'BrandName':                     brand,
                'ImageURL1':                     img_url,
                'ImageURL2':                     img2_url,
                'CatalogStatus':                 _ap_cfg['catalog_status'],
                'StatusRemark':                  _ap_cfg['status_remark'],
                'USER_TYPE':                     gender,
                'DESCRIPTION':                   '',
                'COUNTRY_OF_ORIGIN':             country,
                'EAN':                           '',
                'IMPORTED_BY':                   '',
                'KEY_FEATURES':                  '',
                'MANUFACTURING_YEAR':            '',
                'PRODUCT_BREADTH':               0,
                'PRODUCT_DIMENSION_UOM':         0,
                'PRODUCT_HEIGHT':                0,
                'PRODUCT_LENGTH':                0,
                'PRODUCT_TYPE':                  product_type_val,
                'PRODUCT_WEIGHT_IN_KG':          0,
                'PRODUCT_MANUFACTURING_CITY':    '',
                'PRODUCT_MANUFACTURING_STATE':   '',
                'BLOUSE_FABRIC_MATERIAL':        blouse_fabric_val,
                'BLOUSE_INCLUDED':               blouse_incl_val,
                'BLOUSE_NECK':                   blouse_neck_val,
                'BLOUSE_SLEEVE':                 blouse_sleeve_val,
                'BLOUSE_TYPE':                   blouse_type_val,
                'FABRIC_MATERIAL':               fabric if fabric else '#',
                'LENGTH':                        length if length else '#',
                'MANUFACTURER':                  '',
                'OCCASION':                      '',
                'PATTERN':                       pattern_val,
                'SAREE_TYPE':                    saree_type_val,
            }
        else:
            pav_row = {
                'Jpin':                          '',
                'Title':                         title,
                'PvId':                          pv_id,
                'PvName':                        pv_name,
                'BrandId':                       brand_id,
                'BrandName':                     brand,
                'ImageURL1':                     img_url,
                'ImageURL2':                     img2_url,
                'CatalogStatus':                 _ap_cfg['catalog_status'],
                'StatusRemark':                  _ap_cfg['status_remark'],
                'USER_TYPE':                     gender,
                'DESCRIPTION':                   '',
                'COUNTRY_OF_ORIGIN':             country,
                'EAN':                           '',
                'IMPORTED_BY':                   '',
                'KEY_FEATURES':                  '',
                'MANUFACTURING_YEAR':            '',
                'PRODUCT_BREADTH':               0,
                'PRODUCT_DIMENSION_UOM':         0,
                'PRODUCT_HEIGHT':                0,
                'PRODUCT_LENGTH':                0,
                'PRODUCT_TYPE':                  product_type_val,
                'PRODUCT_WEIGHT_IN_KG':          0,
                'PRODUCT_MANUFACTURING_CITY':    '',
                'PRODUCT_MANUFACTURING_STATE':   '',
                'PATTERN':                       pattern_val,
            }
        _write(ws_pav, tcol_pav, pav_row, pav_row_idx)

        # ── L4 row (per super-category) ───────────────────────
        if super_category == 'Jeans':
            l4_row = {
                '*Type':                         'SET',
                '*Industry Category':            ind_cat,
                '*Industry Sub Category':        ind_sub_cat,
                '*Industry Product Type':        ind_prod_type,
                '*Industry Product Sub-type':    ind_sub_type,
                '*Product Name':                 title,
                '*Product Description':          product_desc,
                '*Seller SKU':                   seller_sku,
                '*Product Code':                 article,
                '*Relationship':                 'Parent',
                '*Parent Product Id':            seller_sku,
                '*Child SKU':                    seller_sku,
                '*Quantity':                     l4_qty_str,
                '*Set Name':                     set_name_str,
                '*HSN Code':                     hsn,
                '*GST':                          gst,
                'Marketed By':                   '',
                '*Country Of Origin':            country,
                'Imported By':                   '',
                'EAN':                           '',
                '*MOQ':                          moq,
                '*MRP':                          mrp,
                '*Selling Price':                sp,
                '*Product Weight (In KG)':       weight,
                '*Product Dimension (LXBXH)':    dim_raw,
                'Manufacturing Year':            '',
                '*Unit Of Measure':              'Set',
                '*Product Dimension UOM':        dim_uom,
                '*Gender':                       gender,
                '*Select Fabric':                fabric,
                'Distress':                      distress,
                'Number of Pockets':             num_pockets,
                'Trend':                         '',
                'Fabric Composition':            fabric_comp,
                'Fade':                          fade,
                'Fit':                           fit,
                'Stretch':                       stretch,
                'Waist Rise':                    waist_rise,
                'Waist Band':                    waist_band,
                'Manufacturing Year':            '',
                'Closure':                       closure,
                'Packaging Type':                packing,
                'Length':                        length,
                '*Select color':                 color,
                '*Size':                         set_details,
                '*Main Image URL':               img_url,
                'Other Image URL1':              img2_url,
                'Other Image URL2':              img3_url,
                'Other Image URL3':              img4_url,
                'Other Image URL4':              img5_url,
                'Other Image URL5':              '',
                '*Brand Name':                   brand,
                'New Brand':                     '',
                '*Pattern':                      pattern_val,
            }
        elif super_category == 'Shirts':
            l4_row = {
                '*Type':                         'SET',
                '*Industry Category':            ind_cat,
                '*Industry Sub Category':        ind_sub_cat,
                '*Industry Product Type':        ind_prod_type,
                '*Industry Product Sub-type':    ind_sub_type,
                '*Product Name':                 title,
                '*Product Description':          product_desc,
                '*Seller SKU':                   seller_sku,
                '*Product Code':                 article,
                '*Relationship':                 'Parent',
                '*Parent Product Id':            seller_sku,
                '*Child SKU':                    seller_sku,
                '*Quantity':                     l4_qty_str,
                '*Set Name':                     set_name_str,
                '*HSN Code':                     hsn,
                '*GST':                          gst,
                'Marketed By':                   '',
                '*Country Of Origin':            country,
                'Imported By':                   '',
                'EAN':                           '',
                '*MOQ':                          moq,
                '*MRP':                          mrp,
                '*Selling Price':                sp,
                '*Product Weight (In KG)':       weight,
                '*Product Dimension (LXBXH)':    dim_raw,
                'Manufacturing Year':            '',
                '*Unit Of Measure':              'Set',
                '*Product Dimension UOM':        dim_uom,
                '*Gender':                       gender,
                '*Select Fabric':                fabric,
                'Distress':                      distress,
                'Number of Pockets':             num_pockets,
                'Trend':                         '',
                'Fabric Composition':            fabric_comp,
                'Fade':                          fade,
                'Fit':                           fit,
                'Stretch':                       stretch,
                'Waist Rise':                    waist_rise,
                'Waist Band':                    waist_band,
                'Manufacturing Year':            '',
                'Closure':                       closure,
                'Packaging Type':                packing,
                'Length':                        length,
                '*Select color':                 color,
                '*Size':                         set_details,
                '*Main Image URL':               img_url,
                'Other Image URL1':              img2_url,
                'Other Image URL2':              img3_url,
                'Other Image URL3':              img4_url,
                'Other Image URL4':              img5_url,
                'Other Image URL5':              '',
                '*Brand Name':                   brand,
                'New Brand':                     '',
                '*Pattern':                      pattern_val,
                'Collar':                        collar_val,
            }
        elif super_category == 'T-Shirt':
            l4_row = {
                '*Type':                         'SET',
                '*Industry Category':            ind_cat,
                '*Industry Sub Category':        ind_sub_cat,
                '*Industry Product Type':        ind_prod_type,
                '*Industry Product Sub-type':    ind_sub_type,
                '*Product Name':                 title,
                '*Product Description':          product_desc,
                '*Seller SKU':                   seller_sku,
                '*Product Code':                 article,
                '*Relationship':                 'Parent',
                '*Parent Product Id':            seller_sku,
                '*Child SKU':                    seller_sku,
                '*Quantity':                     l4_qty_str,
                '*Set Name':                     set_name_str,
                '*HSN Code':                     hsn,
                '*GST':                          gst,
                'Marketed By':                   '',
                '*Country Of Origin':            country,
                'Imported By':                   '',
                'EAN':                           '',
                '*MOQ':                          moq,
                '*MRP':                          mrp,
                '*Selling Price':                sp,
                '*Product Weight (In KG)':       weight,
                '*Product Dimension (LXBXH)':    dim_raw,
                'Manufacturing Year':            '',
                '*Unit Of Measure':              'Set',
                '*Product Dimension UOM':        dim_uom,
                '*Gender':                       gender,
                '*Select Fabric':                fabric,
                '*Pattern':                      pattern_val,
                '*Multipack Set':                multipack_val,
                'Number of Pockets':             num_pockets,
                'Fabric Composition':            fabric_comp,
                'Fit':                           fit,
                'Occasion':                      occasion_val,
                'Manufacturing Year':            '',
                'Neck':                          neck_type_val,
                'Sleeve Length':                 sleeve_len_val,
                'Closure':                       closure,
                'Packaging Type':                packing,
                '*Select color':                 color,
                '*Size':                         set_details,
                '*Main Image URL':               img_url,
                'Other Image URL1':              img2_url,
                'Other Image URL2':              img3_url,
                'Other Image URL3':              img4_url,
                'Other Image URL4':              img5_url,
                'Other Image URL5':              '',
                '*Brand Name':                   brand,
                'New Brand':                     '',
            }
        elif super_category == 'Sarees':
            l4_row = {
                '*Type':                         'SET',
                '*Industry Category':            ind_cat,
                '*Industry Sub Category':        ind_sub_cat,
                '*Industry Product Type':        ind_prod_type,
                '*Industry Product Sub-type':    ind_sub_type,
                '*Product Name':                 title,
                '*Product Description':          product_desc,
                '*Seller SKU':                   seller_sku,
                '*Product Code':                 article,
                '*Relationship':                 'Parent',
                '*Parent Product Id':            seller_sku,
                '*Child SKU':                    seller_sku,
                '*Quantity':                     l4_qty_str,
                '*Set Name':                     set_name_str,
                '*HSN Code':                     hsn,
                '*GST':                          gst,
                'Marketed By':                   '',
                '*Country Of Origin':            country,
                'Imported By':                   '',
                'EAN':                           '',
                '*MOQ':                          moq,
                '*MRP':                          mrp,
                '*Selling Price':                sp,
                '*Product Weight (In KG)':       weight,
                '*Product Dimension (LXBXH)':    dim_raw,
                'Manufacturing Year':            '',
                '*Unit Of Measure':              'Set',
                '*Product Dimension UOM':        dim_uom,
                '*Select Fabric':                fabric,
                '*Fabric Type':                  fabric_type_val,
                '*Gender':                       gender,
                '*Pattern':                      pattern_val,
                'Occasion':                      occasion_val,
                'Hemline':                       hemline_val,
                'Shape':                         shape_val,
                'Set Includes':                  set_includes_val,
                'Bottom Type':                   bottom_type_val,
                'Fabric Composition':            fabric_comp,
                'Number of Pockets':             num_pockets,
                'Work Type':                     work_type_val,
                'Stitch Type':                   stitch_type_val,
                'Border':                        border_val,
                'Manufacturing Year':            '',
                'Neck':                          neck_type_val,
                'Sleeve Length':                 sleeve_len_val,
                'Closure':                       closure,
                'Product Type':                    product_type_val,
                'Packaging Type':                packing,
                'Length':                        length,
                '*Select color':                 color,
                '*Size':                         set_details,
                '*Main Image URL':               img_url,
                'Other Image URL1':              img2_url,
                'Other Image URL2':              img3_url,
                'Other Image URL3':              img4_url,
                'Other Image URL4':              img5_url,
                'Other Image URL5':              '',
                '*Brand Name':                   brand,
                'New Brand':                     '',
            }
        else:
            l4_row = {
                '*Type':                         'SET',
                '*Industry Category':            ind_cat,
                '*Industry Sub Category':        ind_sub_cat,
                '*Industry Product Type':        ind_prod_type,
                '*Industry Product Sub-type':    ind_sub_type,
                '*Product Name':                 title,
                '*Product Description':          product_desc,
                '*Seller SKU':                   seller_sku,
                '*Product Code':                 article,
                '*Relationship':                 'Parent',
                '*Parent Product Id':            seller_sku,
                '*Child SKU':                    seller_sku,
                '*Quantity':                     l4_qty_str,
                '*Set Name':                     set_name_str,
                '*HSN Code':                     hsn,
                '*GST':                          gst,
                'Marketed By':                   '',
                '*Country Of Origin':            country,
                'Imported By':                   '',
                'EAN':                           '',
                '*MOQ':                          moq,
                '*MRP':                          mrp,
                '*Selling Price':                sp,
                '*Product Weight (In KG)':       weight,
                '*Product Dimension (LXBXH)':    dim_raw,
                'Manufacturing Year':            '',
                '*Unit Of Measure':              'Set',
                '*Product Dimension UOM':        dim_uom,
                '*Gender':                       gender,
                '*Select Fabric':                fabric,
                '*Select color':                 color,
                '*Size':                         set_details,
                '*Main Image URL':               img_url,
                'Other Image URL1':              img2_url,
                'Other Image URL2':              img3_url,
                'Other Image URL3':              img4_url,
                'Other Image URL4':              img5_url,
                'Other Image URL5':              '',
                '*Brand Name':                   brand,
                'New Brand':                     '',
                '*Pattern':                      pattern_val,
            }
        _write(ws_l4, tcol_l4, l4_row, l4_row_idx)

    return wb_jpin, wb_tax, wb_scm, wb_pav_map, wb_l4_map, filled, skipped

# ═══════════════════════════════════════════════════════════════
# IN-MEMORY FILE STORAGE
FILE_STORE = {}

# ═══════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════


# ── Global error handlers — always return JSON, never HTML ────
@app.errorhandler(400)
def bad_request(e):
    return jsonify({'error': 'Bad request', 'details': str(e)}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Not found', 'details': str(e)}), 404

@app.errorhandler(500)
def server_error(e):
    import traceback
    return jsonify({'error': 'Internal server error', 'details': str(e), 'trace': traceback.format_exc()}), 500

@app.errorhandler(Exception)
def unhandled_exception(e):
    import traceback
    return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

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
    """Footwear catalog processor. Generates a filled PV Template .xlsx."""
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
                disk_cfg = get_config()
                disk_cfg.update(inline_cfg)
                _save_config(CONFIG_PATH, disk_cfg)
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
                             'SET_COUNT *','FOOTWEAR_TYPE *','hsnCode *']
                    preview_cols = [c for c in pcols if c in headers]
                for r in range(2, min(filled + 2, 52)):
                    rdata = {}
                    for c in preview_cols:
                        if c in headers:
                            rdata[c] = ws.cell(r, headers.index(c)+1).value
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
            out_name  = 'filled_footwear_templates.zip'
            out_ext   = '.zip'
            out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {'bytes': out_bytes, 'filename': out_name,
                                   'ext': out_ext, 'created': time.time()}

        write_log('anonymous', 'fw_catalog_generated',
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
    Apparel & Fashion processor. Produces a ZIP containing 5 files per category:
      - ap_JPIN_<category>.xlsx
      - ap_TaxMaster_<category>.xlsx
      - ap_SCM_<category>.xlsx
      - ap_ProductAttributeValue_<category>_<super>.xlsx  (per super-category)
      - ap_L4_<category>_<super>.xlsx                     (per super-category)
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
                                                                # Filter rows by ind_sub_type OR by super_category match
                sub_type_col = col_map.get('ind_sub_type')
                if sub_type_col and sub_type_col in all_dump.columns:
                    col_lower = all_dump[sub_type_col].astype(str).str.lower().str.strip()
                    cat_lower = category.lower()
                    
                    # Exact match
                    mask = col_lower == cat_lower
                    filtered = all_dump[mask].copy()
                    
                    # Partial match on category string
                    if filtered.empty:
                        mask2 = col_lower.str.contains(re.escape(cat_lower), na=False)
                        filtered = all_dump[mask2].copy()
                    
                    # Match any PV key belonging to this super_category
                    if filtered.empty:
                        # Load config here since _ap_cfg is not in this scope
                        ap_cfg_local = get_ap_config_from_disk()
                        sc_keys = [k for k, v in ap_cfg_local.get('pv_config', {}).items()
                                   if v.get('super_category', '').lower().replace('-', '').rstrip('s') == cat_lower.replace('-', '').rstrip('s')]
                        if sc_keys:
                            pattern = '|'.join(re.escape(k.lower()) for k in sc_keys)
                            mask3 = col_lower.str.contains(pattern, na=False, regex=True)
                            filtered = all_dump[mask3].copy()
                    
                    # Last resort: broad keyword match
                    if filtered.empty:
                        broad_keywords = {
                            't-shirts': ['t-shirt', 'tshirt', 'tee'],
                            'shirts': ['shirt'],
                            'jeans': ['jeans'],
                            'sarees': ['saree'],
                        }
                        keywords = broad_keywords.get(cat_lower, [cat_lower])
                        for kw in keywords:
                            mask_broad = col_lower.str.contains(kw, na=False)
                            if mask_broad.any():
                                filtered = all_dump[mask_broad].copy()
                                break
                    
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                # ── NOW UNPACKS 7 VALUES (was 6) ─────────────────
                wb_jpin, wb_tax, wb_scm, wb_pav_map, wb_l4_map, filled, skipped = fill_ap_files(
                    filtered, col_map, category, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_cat = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
                files_written = []

                # Single workbooks (one per category)
                for wb_obj, label in [
                    (wb_jpin, 'JPIN'),
                    (wb_tax,  'TaxMaster'),
                    (wb_scm,  'SCM'),
                ]:
                    fname   = f'ap_{label}_{safe_cat}.xlsx'
                    xls_buf = io.BytesIO()
                    wb_obj.save(xls_buf)
                    zout.writestr(fname, xls_buf.getvalue())
                    files_written.append(fname)

                # PAV workbooks (one per super-category)
                for super_cat, wb_obj in wb_pav_map.items():
                    safe_super = re.sub(r"[^\w\s-]", "", super_cat).replace(" ", "_")
                    fname   = f'ap_ProductAttributeValue_{safe_cat}_{safe_super}.xlsx'
                    xls_buf = io.BytesIO()
                    wb_obj.save(xls_buf)
                    zout.writestr(fname, xls_buf.getvalue())
                    files_written.append(fname)

                # L4 workbooks (one per super-category)
                for super_cat, wb_obj in wb_l4_map.items():
                    safe_super = re.sub(r"[^\w\s-]", "", super_cat).replace(" ", "_")
                    fname   = f'ap_L4_{safe_cat}_{safe_super}.xlsx'
                    xls_buf = io.BytesIO()
                    wb_obj.save(xls_buf)
                    zout.writestr(fname, xls_buf.getvalue())
                    files_written.append(fname)

                results.append({
                    'category': category,
                    'filled':   filled,
                    'skipped':  len(skipped),
                    'files':    files_written,
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



@app.route('/download_template/<path:category>')
def download_template(category):
    """Serve the master template file for a given vertical."""
    category_lower = category.lower().strip()
    if 'electronic' in category_lower or category_lower == 'ce':
        path  = CE_TEMPLATE_PATH
        fname = 'Consumer_Electronics_Template.xlsx'
    elif 'apparel' in category_lower or category_lower in ('ap', 'fashion'):
        path  = os.path.join(os.path.dirname(__file__), 'Apparel Mapping & Logic Template.xlsx')
        fname = 'Apparels_Fashion_Template.xlsx'
    else:
        path  = TEMPLATE_PATH
        fname = 'Footwear_Template.xlsx'

    if not os.path.exists(path):
        return jsonify({'error': f'Template file not found: {fname}'}), 404

    return send_file(path, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

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
