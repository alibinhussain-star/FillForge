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
                                       'SubType','PVID *','discoveryCategoryIds','ProductCode *'):
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
    cfg = _load_config(AP_CONFIG_PATH, AP_DEFAULT_CONFIG)
    # NEVER let a saved (possibly empty/wrong) pv_config overwrite the hardcoded defaults
    # The frontend saves pv_config with wrong keys; always use the code defaults
    cfg['pv_config'] = AP_DEFAULT_CONFIG['pv_config']
    return cfg

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
    if "infant" in st: return "Infant's"
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
    skipped, filled = [], []

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

        filled_count  = len(filled) + 1
        row_idx  = filled_count + 1

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

        # ProductCode * = Same as Article Number
        product_code = article

        row_data = {
            'Category *':                                  st_data.get('Category *', 'Footwear'),
            'SubCategory *':                               st_data.get('SubCategory *', ''),
            'CategoryType *':                              st_data.get('CategoryType *', ''),
            'SubType':                                     subtype,
            'PVID *':                                      st_data.get('PVID *', ''),
            'BusinessCategoryId *':                        _cfg['biz_cat_id'],
            'BusinessCategoryName *':                      _cfg['biz_cat_name'],
            'ProductCode *':                               product_code,
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

        filled.append({'sku': sku_raw, 'article': article})

    return len(filled), skipped

"""
Consumer Electronics — Unified Template Filler Module
Based on Footwear pattern with Multi-PV/brand config, auto PV detection,
templates in single tab with PV in col D, output SubType blank,
title conventions from PV Level sheet, mapping from ProductVerticle<>Mapping Column.
"""

from flask import Flask, request, jsonify, send_file
import pandas as pd, re, io, os, json, random, string, time, zipfile
from datetime import datetime
from openpyxl import load_workbook, Workbook

# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════

CE_CONFIG_PATH = '/tmp/fillforge_ce_unified_config.json'

CE_DEFAULT_CONFIG = {
    "brands": {},
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

# ═══════════════════════════════════════════════════════════════
# LOAD TEMPLATE FILE (Unified Template Creation.xlsx)
# ═══════════════════════════════════════════════════════════════

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Unified Template Creation.xlsx')


def _build_ce_unified_header_row_map():
    """Build mapping: PV Name -> (header_row, static_data) from Templates sheet."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb['Templates']
    hdr_map = {}
    static_map = {}
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == 'Category *':
            for r2 in range(r + 1, min(r + 5, ws.max_row + 1)):
                c4_val = ws.cell(r2, 4).value
                if c4_val and str(c4_val).strip() not in ('SubType', 'Category *', 'nan', ''):
                    pv_name = str(c4_val).strip()
                    if pv_name not in hdr_map:
                        hdr_map[pv_name] = r
                        hdrs = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                        entry = {}
                        for ci, col in enumerate(hdrs):
                            if col in ('Category *', 'SubCategory *', 'CategoryType *',
                                       'SubType', 'PVID *', 'discoveryCategoryIds', 'ProductCode *'):
                                v = str(ws.cell(r2, ci + 1).value or '').strip()
                                if v and v not in ('nan', 'NaN', 'None'):
                                    entry[col] = v
                        static_map[pv_name] = entry
                    break
    return hdr_map, static_map


def _load_pv_list():
    """Load supported PV list from 'Supported Product Verticle' sheet."""
    try:
        pv = pd.read_excel(TEMPLATE_PATH, sheet_name='Supported Product Verticle')
        result = []
        for _, row in pv.iterrows():
            name = str(row.get('Product Verticle', '')).strip()
            pid = str(row.get('Product Verticle ID', '')).strip()
            family = str(row.get('Unified Template Family', '')).strip()
            if name and name.lower() != 'product verticle':
                result.append({'name': name, 'id': pid, 'family': family})
        return result
    except Exception as e:
        print(f"Warning: Could not load PV List: {e}")
        return []


def _load_title_conventions():
    """Load title conventions from 'PV Level Title Conventions' sheet."""
    try:
        df = pd.read_excel(TEMPLATE_PATH, sheet_name='PV Level Title Conventions', header=None)
        conventions = {}
        i = 0
        while i < len(df):
            pv_name = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
            if pv_name and pv_name != 'Product Verticle' and pv_name in CE_PV_LIST_NAMES:
                title_example = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
                internal_example = str(df.iloc[i, 2]).strip() if pd.notna(df.iloc[i, 2]) else ''
                title_conv = ''
                internal_conv = ''
                if i + 1 < len(df):
                    title_conv = str(df.iloc[i + 1, 1]).strip() if pd.notna(df.iloc[i + 1, 1]) else ''
                    internal_conv = str(df.iloc[i + 1, 2]).strip() if pd.notna(df.iloc[i + 1, 2]) else ''
                conventions[pv_name] = {
                    'title_example': title_example,
                    'internal_example': internal_example,
                    'title_convention': title_conv,
                    'internal_convention': internal_conv,
                }
                i += 2
            else:
                i += 1
        return conventions
    except Exception as e:
        print(f"Warning: Could not load title conventions: {e}")
        return {}


def _load_mapping_logic():
    """Load mapping logic from 'ProductVerticle<>Mapping Column' sheet."""
    try:
        df = pd.read_excel(TEMPLATE_PATH, sheet_name='ProductVerticle<>Mapping Column')
        logic = {}
        for _, row in df.iterrows():
            metric = str(row.get('Metric', '')).strip()
            attr_type = str(row.get('Attribute Type', '')).strip()
            mapping = str(row.get('Mapping Logic', '')).strip()
            subtypes = str(row.get('Subtype ( Comma Seprated)', '')).strip()
            pv_list = [s.strip() for s in subtypes.split(',')]
            for pv in pv_list:
                if pv not in logic:
                    logic[pv] = {}
                logic[pv][metric] = {
                    'attr_type': attr_type,
                    'logic': mapping
                }
        return logic
    except Exception as e:
        print(f"Warning: Could not load mapping logic: {e}")
        return {}


# Initialize at module load
try:
    CE_SUBTYPE_HEADER_ROW, CE_SUBTYPE_MAP = _build_ce_unified_header_row_map()
except Exception as e:
    print(f"Warning: Could not build header row map: {e}")
    CE_SUBTYPE_HEADER_ROW, CE_SUBTYPE_MAP = {}, {}

try:
    CE_PV_LIST_RAW = _load_pv_list()
    CE_PV_LIST = [p['name'] for p in CE_PV_LIST_RAW]
    CE_PV_LIST_NAMES = set(CE_PV_LIST)
    CE_PV_ID_MAP = {p['name']: p['id'] for p in CE_PV_LIST_RAW}
    CE_PV_FAMILY_MAP = {p['name']: p['family'] for p in CE_PV_LIST_RAW}
except Exception as e:
    print(f"Warning: Could not load PV_LIST: {e}")
    CE_PV_LIST_RAW, CE_PV_LIST, CE_PV_LIST_NAMES = [], [], set()
    CE_PV_ID_MAP, CE_PV_FAMILY_MAP = {}, {}

try:
    CE_TITLE_CONVENTIONS = _load_title_conventions()
except Exception as e:
    print(f"Warning: Could not load title conventions: {e}")
    CE_TITLE_CONVENTIONS = {}

try:
    CE_MAPPING_LOGIC = _load_mapping_logic()
except Exception as e:
    print(f"Warning: Could not load mapping logic: {e}")
    CE_MAPPING_LOGIC = {}


# ═══════════════════════════════════════════════════════════════
# CONFIG HELPERS
# ═══════════════════════════════════════════════════════════════

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


def get_ce_unified_config():
    return _load_config(CE_CONFIG_PATH, CE_DEFAULT_CONFIG)


# ═══════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════

def safe(val, default=''):
    try:
        if pd.isna(val):
            return default
    except:
        pass
    s = str(val).strip() if val is not None else default
    return default if s in ('nan', 'None', 'NaN', '') else s


def detect_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        for cand in candidates:
            if cand.lower() in c.lower():
                return c
    return None


def build_col_map(df, hints):
    return {k: detect_col(df, v) for k, v in hints.items() if detect_col(df, v)}


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
                bid = item.get('id', item.get('brandId', ''))
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


def parse_lbh(dim_str):
    parts = re.split(r'[Xx×]', str(dim_str).strip())
    if len(parts) == 3:
        try:
            return int(parts[0]), int(parts[1]), int(parts[2])
        except:
            pass
    return None, None, None


# ═══════════════════════════════════════════════════════════════
# COLUMN HINTS FOR INPUT FILE
# ═══════════════════════════════════════════════════════════════

CE_DUMP_COL_HINTS = {
    'pv':              ['Product Verticle', 'Product Vertical', 'ProductVerticle', 'PV', 'Product Type', 'Product Sub-type', 'CategoryType *', 'Subtype', 'SubType'],
    'sku':             ['Seller SKU ID', 'ChildSKU *', 'ChildSKU', 'SKU', 'Child SKU'],
    'model_name':      ['Model Name', 'Name of the model/Title name', 'MODEL_NAME *', 'Model Name'],
    'mrp':             ['MRP', '*MRP', 'MRP *', 'MRP full Set'],
    'sp':              ['Selling Price', 'SellingPrice *', '*Selling Price', '*Selling Price per Pair'],
    'moq':             ['*MOQ', 'MOQ *', 'MOQ', '*Minimum Order Quantity'],
    'brand':           ['Brand', 'Brand Name', 'brandName *', 'brand_name'],
    'image':           ['imageURL1', 'ImageURL1', 'Main Image URL', 'Image Link', 'Image Links'],
    'image2':          ['imageURL2', 'ImageURL2', 'Other Image URL 1', 'Other Image URL1'],
    'image3':          ['imageURL3', 'ImageURL3', 'Other Image URL 2', 'Other Image URL2'],
    'image4':          ['imageURL4', 'ImageURL4', 'Other Image URL 3', 'Other Image URL3'],
    'image5':          ['imageURL5', 'ImageURL5', 'Other Image URL 4', 'Other Image URL4'],
    'image6':          ['imageURL6', 'ImageURL6', 'Other Image URL 5', 'Other Image URL5'],
    'product_desc':    ['Product Description', 'productDescription *', 'Description'],
    'color':           ['Colour', 'Color', 'PRODUCT_COLOR *', 'Product Color', 'Product Colour', 'Primary Colour'],
    'product_condition': ['Product Condition', 'Condition'],
    'packing':         ['Packaging Type', 'PACKAGING_TYPE *', 'Packing Type'],
    'material':        ['Product Material', 'Material', 'MATERIAL *'],
    'adapter_connector': ['Adapter Connector Type', 'ADAPTER_CONNECTOR_TYPE *'],
    'country':         ['Country of Origin', 'COUNTRY_OF_ORIGIN *', 'Country/Region of Origin'],
    'num_ports':       ['Number of Ports', 'NO_OF_ADAPTER_PORTS *'],
    'connector_type':  ['Connector Type', 'Connector type'],
    'output_voltage':  ['Output Voltage', 'OUTPUT_CURRENT_OR_VOLTAGE *'],
    'port_type':       ['Port type', 'Port Type', 'PORT_TYPE *'],
    'dims':            ['Product Dimension (LXBXH)', '*Product Dimension (LXBXH)', 'Product Dimension'],
    'dim_uom':         ['Unit of Measurement', 'PRODUCT_DIMENSION_UOM *', '*Product Dimension UOM'],
    'no_of_connectors': ['No of Connetors', 'No Of Connectors'],
    'product_name':    ['Product Name', 'PRODUCT_TYPE *', 'Product Type'],
    'weight':          ['Product Weight (KG)', '*Product Weight (In KG)', 'PRODUCT_WEIGHT_IN_KG *', 'Product Weight'],
    'hsn':             ['HSN', 'HSN Code', '*HSN Code', 'hsnCode *'],
    'gst':             ['GST', '*GST', 'gstPercentage *', 'GSTpercentage'],
    'bluetooth':       ['Bluetooth Version', 'BLUETOOTH_VERSION *'],
    'speaker_type':    ['Speaker Type', 'SPEAKER_TYPE *'],
    'wired_wireless':  ['Wired or Unwired', 'WIRED_OR_WIRELESS *'],
    'mic_type':        ['Mic Type', 'MIC_TYPE *'],
    'compatible_brand_model': ['Compatible Brand + Model Name', 'COMPATIBLE_BRAND_MODEL *'],
    'case_cover_type': ['Case & Cover Type', 'CASE_COVER_TYPE *'],
    'closure_type':    ['Case & Cover Closure', 'CLOSURE_TYPE *'],
    'pattern':         ['Pattern', 'DESIGN *'],
    'compatible_brand': ['Compatible Brand', 'COMPATIBLE_BRAND *'],
    'coverage':        ['Coverage', 'COVERAGE *'],
    'screen_guard_type': ['Screen Guard type', 'SCREEN_GUARD_OR_PROTECTOR_TYPE *'],
    'screen_thickness': ['Screen Card Thickness', 'THICKNESS *'],
    'battery_capacity': ['Battery Capacity', 'BATTERY_CAPACITY_MAH *'],
    'charging_type':   ['Charging type supported', 'CHARGING_TYPE_SUPPORTED'],
    'display_size':    ['Display Size', 'DISPLAY_SIZE *', 'Screen Size'],
    'ram':             ['RAM', 'RAM *'],
    'storage':         ['Internal Storage', 'INTERNAL_STORAGE *', 'Storage Capacity'],
    'os':              ['Operating System', 'OPERATING_SYSTEM_OS *'],
    'os_version':      ['Operating System Version', 'OS_VERSION *'],
    'processor_core':  ['Processor Core', 'NUMBER_OF_PROCESSOR_CORES *'],
    'back_camera':     ['Back Camera', 'PRIMARY_CAMERA_RESOLUTION *'],
    'front_camera':    ['Front Camera', 'FRONT_CAMERA_RESOLUTION *'],
    'sim_type':        ['Sim Type', 'SIM_TYPE *'],
    'network_support': ['Network Support', 'Network', 'NETWORK_TYPE_SUPPORTED'],
    'card_type':       ['Card type', 'MEMORY_CARD_TYPE *'],
    'speed_class':     ['Speed Class', 'SPEED_CLASS *'],
    'storage_capacity': ['Storage Capacity', 'STORAGE_CAPACITY *'],
    'holder_type':     ['Holder Type', 'HOLDER_TYPE *'],
    'lock_mechanism':  ['Lock Mechanism', 'LOCK_MECHANISM *'],
    'rotation_type':   ['Rotation type', 'ROTATION_OR_ADJUSTABILITY *'],
    'no_of_output_ports': ['No of Output Ports', 'Number of Output Ports'],
    'output_ports_type': ['Output Ports Type', 'Output ports type'],
    'display_type':    ['Display Type', 'DISPLAY_TYPE *'],
    'display_shape':   ['Display Shape', 'DISPLAY_SHAPE *'],
    'display_resolution': ['Display Resolution', 'DISPLAY_RESOLUTION'],
    'expandable_storage': ['Expandable Storage', 'EXPANDABLE_STORAGE'],
    'bluetooth_version': ['Bluetooth Version', 'BLUETOOTH_VERSION'],
    'expandable_storage_type': ['Expandable Storage Type', 'EXPANDABLE_STORAGE_TYPE'],
    'expandable_storage_max': ['Expandable Storage Capacity Max', 'EXPANDABLE_STORAGE_CAPACITY_MAX'],
    'battery_type':    ['Battery Type', 'BATTERY_TYPE'],
    'removable_battery': ['Removable Battery', 'REMOVABLE_BATTERY'],
    'hybrid_sim':      ['Hybrid Sim Slot', 'HYBRID_SIM_SLOT'],
    'audio_jack':      ['Audio Jack', 'AUDIO_JACK'],
    'fm_radio':        ['FM Radio', 'FM_RADIO'],
    'torch':           ['Torch or Flashlight', 'TORCH_OR_FLASHLIGHT'],
    'rear_flash':      ['Rear Flash', 'REAR_FLASH'],
    'wifi':            ['WIFI', 'Wifi'],
    'fingerprint':     ['Fingerprint Sensor', 'FINGERPRINT_SENSOR'],
    'clock_speed':     ['Clock Speed', 'CLOCK_SPEED'],
    'refresh_rate':    ['Refresh Rate', 'REFRESH_RATE'],
    'touchscreen':     ['Touchscreen Type', 'TOUCHSCREEN_TYPE'],
    'primary_camera_setup': ['Primary Camera Setup', 'PRIMARY_CAMERA_SETUP'],
    'front_flash':     ['Front Flash', 'FRONT_FLASH'],
    'video_resolution': ['Video Recording Resolution', 'VIDEO_RECORDING_RESOLUTION'],
    'fast_charging':   ['Fast Charging Wattage', 'FAST_CHARGING_WATTAGE'],
    'wireless_charging': ['Wireless Charging Support', 'WIRELESS_CHARGING_SUPPORT'],
    'gps':             ['GPS Support', 'GPS_SUPPORT'],
    'nfc':             ['NFC Support', 'NFC_SUPPORT'],
    'infrared':        ['Infrared IR Blaster', 'INFRARED_IR_BLASTER'],
    'fingerprint_pos': ['Fingerprint Sensor Position', 'FINGERPRINT_SENSOR_POSITION'],
    'face_unlock':     ['Face Unlock', 'FACE_UNLOCK'],
    'water_resistance': ['Water Resistance Rating', 'WATER_RESISTANCE_RATING'],
    'strap_color':     ['Strap Color', 'STRAP_COLOR *'],
    'core_brand':      ['Core Brand', 'CORE_BRAND *'],
    'cable_length':    ['Cable Length', 'CABLE_LENGTH_IN_METER *'],
    'cable_type':      ['Cable Type', 'CABLE_TYPE'],
    'cable_material':  ['Cable Material', 'CABLE_MATERIAL *'],
}

CE_BASE_COL_HINTS = {
    'article': ['Model Name', 'Name of the model/Title name', 'Article Number', 'ARTICLE_NUMBER'],
    'sku':     ['Seller SKU ID', 'ChildSKU', 'Child SKU', 'SKU'],
}


# ═══════════════════════════════════════════════════════════════
# TITLE BUILDERS (per PV from conventions)
# ═══════════════════════════════════════════════════════════════

def _extract_condition(condition_raw, cfg):
    c = safe(condition_raw)
    if not c:
        c = cfg.get('product_condition', 'Fresh')
    return c


def _title_case_color(raw):
    if not raw:
        return raw
    return ' '.join(w.capitalize() for w in str(raw).strip().split())


def make_ce_title(pv_name, brand, model_name, color, condition, **kwargs):
    """Build title based on PV-specific convention from 'PV Level Title Conventions' sheet."""
    color = _title_case_color(color)
    condition = _extract_condition(condition, kwargs.get('cfg', {}))

    if pv_name == 'Mobile Adapters & Cables':
        # Brand + Product Condition + Output Voltage + Adapter Connector Type + , Colour
        output_voltage = kwargs.get('output_voltage', '')
        adapter_connector = kwargs.get('adapter_connector', '')
        parts = [p for p in [brand, condition, output_voltage, adapter_connector] if p]
        base = ' '.join(parts)
        # Add "Adapter" if not already in the base
        if 'Adapter' not in base:
            base = f"{base} Adapter"
        return f"{base}, {color}" if color else base

    elif pv_name == 'Speakers':
        # Brand + Model Name + Product Condition + Speaker Type + , Colour
        speaker_type = kwargs.get('speaker_type', '')
        parts = [p for p in [brand, model_name, condition, speaker_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Mobile Case & Covers':
        # Brand + Product Condition + Compatible Brand + Model Name + Product Material + Case & Cover Type + , Colour
        compatible_brand_model = kwargs.get('compatible_brand_model', '')
        material = kwargs.get('material', '')
        case_type = kwargs.get('case_cover_type', '')
        parts = [p for p in [brand, condition, compatible_brand_model, material, case_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Feature Phones':
        # Brand + Model Name + Display Size + "Display" + Product Verticle + , Colour (Product Condition)
        display_size = kwargs.get('display_size', '')
        ds_part = f'{display_size}" Display' if display_size else ''
        parts = [p for p in [brand, model_name, ds_part, pv_name] if p]
        base = ' '.join(parts)
        suffix = f"{color} ({condition})" if color else f"({condition})"
        return f"{base}, {suffix}"

    elif pv_name == 'Headsets':
        # Brand + Model Name + Product Condition + Product Verticle + , Colour
        parts = [p for p in [brand, model_name, condition, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'TWS Ear Buds':
        # Brand + Model Name + Product Condition + "Earbuds" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Earbuds'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Earphones':
        # Brand + Model Name + Product Condition + "Wired Earphones" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Wired Earphones'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Neck Bands':
        # Brand + Model Name + Product Condition + "Neckband" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Neckband'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Memory Cards':
        # Brand + Storage Capacity + Card type + Product Verticle + , Speed Class
        # NOTE: model_name is NOT used here; storage_capacity and card_type are used instead
        storage_capacity = kwargs.get('storage_capacity', '')
        card_type = kwargs.get('card_type', '')
        speed_class = kwargs.get('speed_class', '')
        parts = [p for p in [brand, storage_capacity, card_type, pv_name] if p]
        base = ' '.join(parts)
        suffix = f", {speed_class}" if speed_class else ''
        return f"{base}{suffix}"

    elif pv_name == 'Mobile Cables':
        # Brand + Product Condition + No of Connectors + Connector Type + Product Verticle + , Colour
        no_of_connectors = kwargs.get('no_of_connectors', '')
        connector_type = kwargs.get('connector_type', '')
        parts = [p for p in [brand, condition, no_of_connectors, connector_type, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Mobile Holders':
        # Brand + Product Condition + Product Verticle + Rotation type + , Colour
        rotation_type = kwargs.get('rotation_type', '')
        parts = [p for p in [brand, condition, pv_name, rotation_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Screen Guards / Protectors':
        # Brand + Product Condition + Screen Guard type + Coverage + "for" Compatible Brand + Model Name
        screen_guard_type = kwargs.get('screen_guard_type', '')
        coverage = kwargs.get('coverage', '')
        compatible_brand_model = kwargs.get('compatible_brand_model', '')
        parts = [p for p in [brand, condition, screen_guard_type, coverage] if p]
        base = ' '.join(parts)
        if compatible_brand_model:
            base = f"{base} for {compatible_brand_model}"
        return base

    elif pv_name == 'Power Bank':
        # Brand + Product Condition + Battery Capacity + Product Verticle + No of Output Ports + Output Ports Type + , Colour
        battery_capacity = kwargs.get('battery_capacity', '')
        no_of_output_ports = kwargs.get('no_of_output_ports', '')
        output_ports_type = kwargs.get('output_ports_type', '')
        parts = [p for p in [brand, condition, battery_capacity, pv_name, no_of_output_ports, output_ports_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Smartphones':
        # Brand + Model Name + Back Camera + "Camera" + Product Verticle + , RAM + Internal Storage + , Colour (Product Condition)
        back_camera = kwargs.get('back_camera', '')
        ram = kwargs.get('ram', '')
        storage = kwargs.get('storage', '')
        cam_part = f'{back_camera} Camera' if back_camera else ''
        ram_storage = f'{ram} + {storage}' if (ram and storage) else (ram or storage)
        parts = [p for p in [brand, model_name, cam_part, pv_name] if p]
        base = ' '.join(parts)
        suffix_parts = []
        if ram_storage:
            suffix_parts.append(ram_storage)
        if color:
            suffix_parts.append(f'{color} ({condition})')
        if suffix_parts:
            return f"{base}, {', '.join(suffix_parts)}"
        return base

    elif pv_name == 'Smart Watches':
        # Brand + Model Name + Product Condition + Display Size + "Display" + Product Verticle + , Colour
        display_size = kwargs.get('display_size', '')
        ds_part = f'{display_size}" Display' if display_size else ''
        parts = [p for p in [brand, model_name, condition, ds_part, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    # Fallback
    parts = [p for p in [brand, model_name, pv_name] if p]
    base = ' '.join(parts)
    return f"{base}, {color}" if color else base


def make_ce_internal_title(pv_name, brand, model_name, color, condition, **kwargs):
    """Build internal title based on PV-specific convention from 'PV Level Title Conventions' sheet."""
    color = _title_case_color(color)
    condition = _extract_condition(condition, kwargs.get('cfg', {}))

    if pv_name == 'Mobile Adapters & Cables':
        # Brand + Product Condition + Model Name + Output Voltage + Adapter Connector Type + , Colour
        output_voltage = kwargs.get('output_voltage', '')
        adapter_connector = kwargs.get('adapter_connector', '')
        parts = [p for p in [brand, condition, model_name, output_voltage, adapter_connector] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Speakers':
        # Brand + Model Name + Product Condition + Speaker Type + , Colour
        speaker_type = kwargs.get('speaker_type', '')
        parts = [p for p in [brand, model_name, condition, speaker_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Mobile Case & Covers':
        # Brand + Product Condition + Model Name + Compatible Brand + Model Name + Product Material + Case & Cover Type + , Colour
        compatible_brand_model = kwargs.get('compatible_brand_model', '')
        material = kwargs.get('material', '')
        case_type = kwargs.get('case_cover_type', '')
        parts = [p for p in [brand, condition, model_name, compatible_brand_model, material, case_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Feature Phones':
        # Brand + Model Name + Display Size + "Display" + Product Verticle + , Colour (Product Condition)
        display_size = kwargs.get('display_size', '')
        ds_part = f'{display_size}" Display' if display_size else ''
        parts = [p for p in [brand, model_name, ds_part, pv_name] if p]
        base = ' '.join(parts)
        suffix = f"{color} ({condition})" if color else f"({condition})"
        return f"{base}, {suffix}"

    elif pv_name == 'Headsets':
        # Brand + Model Name + Product Condition + Product Verticle + , Colour
        parts = [p for p in [brand, model_name, condition, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'TWS Ear Buds':
        # Brand + Model Name + Product Condition + "Earbuds" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Earbuds'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Earphones':
        # Brand + Model Name + Product Condition + "Wired Earphones" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Wired Earphones'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Neck Bands':
        # Brand + Model Name + Product Condition + "Neckband" + , Colour
        parts = [p for p in [brand, model_name, condition, 'Neckband'] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Memory Cards':
        # Brand + Storage Capacity + Card type + Product Verticle + , Speed Class
        # NOTE: model_name is NOT used here; storage_capacity and card_type are used instead
        storage_capacity = kwargs.get('storage_capacity', '')
        card_type = kwargs.get('card_type', '')
        speed_class = kwargs.get('speed_class', '')
        parts = [p for p in [brand, storage_capacity, card_type, pv_name] if p]
        base = ' '.join(parts)
        suffix = f", {speed_class}" if speed_class else ''
        return f"{base}{suffix}"

    elif pv_name == 'Mobile Cables':
        # Brand + Product Condition + Model Name + No of Connectors + Connector Type + Product Verticle + , Colour
        no_of_connectors = kwargs.get('no_of_connectors', '')
        connector_type = kwargs.get('connector_type', '')
        parts = [p for p in [brand, condition, model_name, no_of_connectors, connector_type, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Mobile Holders':
        # Brand + Product Condition + Model Name + Product Verticle + Rotation type + , Colour
        rotation_type = kwargs.get('rotation_type', '')
        parts = [p for p in [brand, condition, model_name, pv_name, rotation_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Screen Guards / Protectors':
        # Brand + Product Condition + Model Name + Screen Guard type + Coverage + "for" Compatible Brand + Model Name
        screen_guard_type = kwargs.get('screen_guard_type', '')
        coverage = kwargs.get('coverage', '')
        compatible_brand_model = kwargs.get('compatible_brand_model', '')
        parts = [p for p in [brand, condition, model_name, screen_guard_type, coverage] if p]
        base = ' '.join(parts)
        if compatible_brand_model:
            base = f"{base} for {compatible_brand_model}"
        return base

    elif pv_name == 'Power Bank':
        # Brand + Product Condition + Model Name + Battery Capacity + Product Verticle + No of Output Ports + Output Ports Type + , Colour
        battery_capacity = kwargs.get('battery_capacity', '')
        no_of_output_ports = kwargs.get('no_of_output_ports', '')
        output_ports_type = kwargs.get('output_ports_type', '')
        parts = [p for p in [brand, condition, model_name, battery_capacity, pv_name, no_of_output_ports, output_ports_type] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    elif pv_name == 'Smartphones':
        # Brand + Model Name + Back Camera + "Camera" + Product Verticle + , RAM + Internal Storage + , Colour (Product Condition)
        back_camera = kwargs.get('back_camera', '')
        ram = kwargs.get('ram', '')
        storage = kwargs.get('storage', '')
        cam_part = f'{back_camera} Camera' if back_camera else ''
        ram_storage = f'{ram} + {storage}' if (ram and storage) else (ram or storage)
        parts = [p for p in [brand, model_name, cam_part, pv_name] if p]
        base = ' '.join(parts)
        suffix_parts = []
        if ram_storage:
            suffix_parts.append(ram_storage)
        if color:
            suffix_parts.append(f'{color} ({condition})')
        if suffix_parts:
            return f"{base}, {', '.join(suffix_parts)}"
        return base

    elif pv_name == 'Smart Watches':
        # Brand + Model Name + Product Condition + Display Size + "Display" + Product Verticle + , Colour
        display_size = kwargs.get('display_size', '')
        ds_part = f'{display_size}" Display' if display_size else ''
        parts = [p for p in [brand, model_name, condition, ds_part, pv_name] if p]
        base = ' '.join(parts)
        return f"{base}, {color}" if color else base

    # Fallback
    parts = [p for p in [brand, model_name, pv_name] if p]
    base = ' '.join(parts)
    return f"{base}, {color}" if color else base


# ═══════════════════════════════════════════════════════════════
# SET DETAILS BUILDER
# ═══════════════════════════════════════════════════════════════

def _build_set_details_ce(pv_name, **kwargs):
    """Build SET_DETAILS based on PV rules from ProductVerticle<>Mapping Column."""
    if pv_name == 'Mobile Case & Covers':
        cbm = kwargs.get('compatible_brand_model', '')
        return cbm if cbm else '1pc'
    elif pv_name in ('Smartphones', 'Feature Phones'):
        ram = kwargs.get('ram', '')
        storage = kwargs.get('storage', '')
        if ram and storage:
            return f'{ram} + {storage}'
        return ram or storage or '1pc'
    else:
        return '1pc'


# ═══════════════════════════════════════════════════════════════
# TEMPLATE WORKBOOK BUILDER
# ═══════════════════════════════════════════════════════════════

def get_ce_unified_template_wb_for_pv(pv_name):
    """Get template workbook for a specific PV from the unified template."""
    try:
        wb_src = load_workbook(TEMPLATE_PATH)
        ws_src = wb_src['Templates']
        hdr_row = CE_SUBTYPE_HEADER_ROW.get(pv_name, 1)
        headers = [ws_src.cell(hdr_row, c).value for c in range(1, ws_src.max_column + 1)]
        while headers and headers[-1] is None:
            headers.pop()
    except Exception as e:
        print(f"Warning: Could not load template for {pv_name}: {e}")
        headers = []

    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.title = 'CE Unified Template'
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h
    return wb_new, headers


# ═══════════════════════════════════════════════════════════════
# MAIN TEMPLATE FILLER
# ═══════════════════════════════════════════════════════════════

def fill_ce_unified_template(ws, headers, rows_df, col_map, pv_name, existing_articles, existing_skus):
    """Fill template for Consumer Electronics unified template."""
    tcol = {h: i + 1 for i, h in enumerate(headers) if h}
    _cfg = get_ce_unified_config()
    brands_dict = normalize_brands(_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))

    st_data = CE_SUBTYPE_MAP.get(pv_name, {})
    skipped, filled = [], 0

    for _, drow in rows_df.iterrows():
        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand = fallback_brand
            brand_id = fallback_id

        sku_raw = safe(drow.get(col_map.get('sku', ''), ''))
        model_name = safe(drow.get(col_map.get('model_name', ''), ''))
        article = model_name if model_name else sku_raw

        if article.upper() in existing_articles or sku_raw.upper() in existing_skus:
            skipped.append({'sku': sku_raw, 'article': article, 'reason': 'Already exists in base data'})
            continue

        filled += 1
        row_idx = filled + 1

        # Extract all fields from input
        mrp = drow.get(col_map.get('mrp', ''), '')
        sp = drow.get(col_map.get('sp', ''), '')
        moq = drow.get(col_map.get('moq', ''), 1)
        color = safe(drow.get(col_map.get('color', ''), ''))
        weight = drow.get(col_map.get('weight', ''), '')
        dim_raw = safe(drow.get(col_map.get('dims', ''), ''))
        hsn = drow.get(col_map.get('hsn', ''), '')
        gst = drow.get(col_map.get('gst', ''), 18)
        packing = safe(drow.get(col_map.get('packing', ''), '')) or 'BOX'
        country = safe(drow.get(col_map.get('country', ''), '')) or _cfg['country_of_origin']
        dim_uom = safe(drow.get(col_map.get('dim_uom', ''), '')) or 'cm'
        prod_desc = safe(drow.get(col_map.get('product_desc', ''), ''))
        condition_raw = safe(drow.get(col_map.get('product_condition', ''), ''))
        condition = condition_raw if condition_raw else _cfg['product_condition']

        img_url = safe(drow.get(col_map.get('image', ''), ''))
        img2_url = safe(drow.get(col_map.get('image2', ''), ''))
        img3_url = safe(drow.get(col_map.get('image3', ''), ''))
        img4_url = safe(drow.get(col_map.get('image4', ''), ''))
        img5_url = safe(drow.get(col_map.get('image5', ''), ''))
        img6_url = safe(drow.get(col_map.get('image6', ''), ''))

        # PV-specific fields
        kwargs = {
            'output_voltage': safe(drow.get(col_map.get('output_voltage', ''), '')),
            'adapter_connector': safe(drow.get(col_map.get('adapter_connector', ''), '')),
            'speaker_type': safe(drow.get(col_map.get('speaker_type', ''), '')),
            'compatible_brand_model': safe(drow.get(col_map.get('compatible_brand_model', ''), '')),
            'material': safe(drow.get(col_map.get('material', ''), '')),
            'case_cover_type': safe(drow.get(col_map.get('case_cover_type', ''), '')),
            'display_size': safe(drow.get(col_map.get('display_size', ''), '')),
            'storage_capacity': safe(drow.get(col_map.get('storage_capacity', ''), '')),
            'card_type': safe(drow.get(col_map.get('card_type', ''), '')),
            'speed_class': safe(drow.get(col_map.get('speed_class', ''), '')),
            'no_of_connectors': safe(drow.get(col_map.get('no_of_connectors', ''), '')),
            'connector_type': safe(drow.get(col_map.get('connector_type', ''), '')),
            'rotation_type': safe(drow.get(col_map.get('rotation_type', ''), '')),
            'screen_guard_type': safe(drow.get(col_map.get('screen_guard_type', ''), '')),
            'coverage': safe(drow.get(col_map.get('coverage', ''), '')),
            'battery_capacity': safe(drow.get(col_map.get('battery_capacity', ''), '')),
            'no_of_output_ports': safe(drow.get(col_map.get('no_of_output_ports', ''), '')),
            'output_ports_type': safe(drow.get(col_map.get('output_ports_type', ''), '')),
            'back_camera': safe(drow.get(col_map.get('back_camera', ''), '')),
            'ram': safe(drow.get(col_map.get('ram', ''), '')),
            'storage': safe(drow.get(col_map.get('storage', ''), '')),
            'bluetooth': safe(drow.get(col_map.get('bluetooth', ''), '')),
            'wired_wireless': safe(drow.get(col_map.get('wired_wireless', ''), '')),
            'mic_type': safe(drow.get(col_map.get('mic_type', ''), '')),
            'num_ports': safe(drow.get(col_map.get('num_ports', ''), '')),
            'cable_length': safe(drow.get(col_map.get('cable_length', ''), '')),
            'cable_type': safe(drow.get(col_map.get('cable_type', ''), '')),
            'cable_material': safe(drow.get(col_map.get('cable_material', ''), '')),
            'closure_type': safe(drow.get(col_map.get('closure_type', ''), '')),
            'pattern': safe(drow.get(col_map.get('pattern', ''), '')),
            'compatible_brand': safe(drow.get(col_map.get('compatible_brand', ''), '')),
            'screen_thickness': safe(drow.get(col_map.get('screen_thickness', ''), '')),
            'charging_type': safe(drow.get(col_map.get('charging_type', ''), '')),
            'os': safe(drow.get(col_map.get('os', ''), '')),
            'os_version': safe(drow.get(col_map.get('os_version', ''), '')),
            'processor_core': safe(drow.get(col_map.get('processor_core', ''), '')),
            'front_camera': safe(drow.get(col_map.get('front_camera', ''), '')),
            'sim_type': safe(drow.get(col_map.get('sim_type', ''), '')),
            'network_support': safe(drow.get(col_map.get('network_support', ''), '')),
            'display_type': safe(drow.get(col_map.get('display_type', ''), '')),
            'display_resolution': safe(drow.get(col_map.get('display_resolution', ''), '')),
            'expandable_storage': safe(drow.get(col_map.get('expandable_storage', ''), '')),
            'bluetooth_version': safe(drow.get(col_map.get('bluetooth_version', ''), '')),
            'expandable_storage_type': safe(drow.get(col_map.get('expandable_storage_type', ''), '')),
            'expandable_storage_max': safe(drow.get(col_map.get('expandable_storage_max', ''), '')),
            'battery_type': safe(drow.get(col_map.get('battery_type', ''), '')),
            'removable_battery': safe(drow.get(col_map.get('removable_battery', ''), '')),
            'hybrid_sim': safe(drow.get(col_map.get('hybrid_sim', ''), '')),
            'audio_jack': safe(drow.get(col_map.get('audio_jack', ''), '')),
            'fm_radio': safe(drow.get(col_map.get('fm_radio', ''), '')),
            'torch': safe(drow.get(col_map.get('torch', ''), '')),
            'rear_flash': safe(drow.get(col_map.get('rear_flash', ''), '')),
            'wifi': safe(drow.get(col_map.get('wifi', ''), '')),
            'fingerprint': safe(drow.get(col_map.get('fingerprint', ''), '')),
            'clock_speed': safe(drow.get(col_map.get('clock_speed', ''), '')),
            'refresh_rate': safe(drow.get(col_map.get('refresh_rate', ''), '')),
            'touchscreen': safe(drow.get(col_map.get('touchscreen', ''), '')),
            'primary_camera_setup': safe(drow.get(col_map.get('primary_camera_setup', ''), '')),
            'front_flash': safe(drow.get(col_map.get('front_flash', ''), '')),
            'video_resolution': safe(drow.get(col_map.get('video_resolution', ''), '')),
            'fast_charging': safe(drow.get(col_map.get('fast_charging', ''), '')),
            'wireless_charging': safe(drow.get(col_map.get('wireless_charging', ''), '')),
            'gps': safe(drow.get(col_map.get('gps', ''), '')),
            'nfc': safe(drow.get(col_map.get('nfc', ''), '')),
            'infrared': safe(drow.get(col_map.get('infrared', ''), '')),
            'fingerprint_pos': safe(drow.get(col_map.get('fingerprint_pos', ''), '')),
            'face_unlock': safe(drow.get(col_map.get('face_unlock', ''), '')),
            'water_resistance': safe(drow.get(col_map.get('water_resistance', ''), '')),
            'strap_color': safe(drow.get(col_map.get('strap_color', ''), '')),
            'core_brand': safe(drow.get(col_map.get('core_brand', ''), '')),
            'holder_type': safe(drow.get(col_map.get('holder_type', ''), '')),
            'lock_mechanism': safe(drow.get(col_map.get('lock_mechanism', ''), '')),
            'cfg': _cfg,
        }

        # Build titles
        title = make_ce_title(pv_name, brand, model_name, color, condition, **kwargs)
        internal_title = make_ce_internal_title(pv_name, brand, model_name, color, condition, **kwargs)

        # Build set details
        set_details = _build_set_details_ce(pv_name, **kwargs)
        set_desc = f'1pc of {pv_name}'

        # Parse dimensions
        L, B, H = parse_lbh(dim_raw)

        # Clean weight
        weight_clean = ''
        if weight:
            m = re.search(r'([0-9.]+)', str(weight))
            if m:
                weight_clean = float(m.group(1))

        # Numeric conversions
        try:
            mrp = float(mrp) if str(mrp).strip() not in ('', 'nan') else ''
        except:
            mrp = ''
        try:
            sp = float(sp) if str(sp).strip() not in ('', 'nan') else ''
        except:
            sp = ''
        try:
            hsn = int(float(hsn)) if str(hsn).strip() not in ('', 'nan') else ''
        except:
            hsn = ''
        try:
            gst = int(float(gst))
        except:
            gst = 18
        try:
            moq = int(float(moq))
        except:
            moq = 1

        # Build row data - SubType is BLANK
        row_data = {
            'Category *': st_data.get('Category *', 'Consumer Electronics'),
            'SubCategory *': st_data.get('SubCategory *', ''),
            'CategoryType *': st_data.get('CategoryType *', ''),
            'SubType': '',  # BLANK as per requirement
            'PVID *': st_data.get('PVID *', CE_PV_ID_MAP.get(pv_name, '')),
            'BusinessCategoryId *': _cfg['biz_cat_id'],
            'BusinessCategoryName *': _cfg['biz_cat_name'],
            'ProductCode *': article,
            'Relationship *': _cfg['relationship'],
            'ParentProductId *': sku_raw,
            'ChildSKU *': sku_raw,
            'MRP *': mrp,
            'SellingPrice *': sp,
            'MOQ *': moq,
            'title *': title,
            'internalTitle *': internal_title,
            'brandId *': brand_id,
            'brandName *': brand,
            'imageURL1 *': img_url,
            'imageURL2': img2_url,
            'imageURL3': img3_url,
            'imageURL4': img4_url,
            'imageURL5': img5_url,
            'imageURL6': img6_url,
            'videoURL1': '',
            'videoURL2': '',
            'sizeChartURLImage': '',
            'catalogStatus *': _cfg['catalog_status'],
            'statusRemark': _cfg['status_remark'],
            'discoveryCategoryIds': st_data.get('discoveryCategoryIds', _cfg['discovery_cat']),
            'productDescription *': prod_desc,
            'PRODUCT_IDENTIFIER *': 'Set',
            'SET_NAME *': 'Set of 1',
            'SET_COUNT *': 1,
            'PACK_NAME *': 'Pack of 1',
            'PACK_OF *': 1,
            'IS_COMBO *': 'yes',
            'AVAILABLE_SIZES *': '1',
            'SET_DETAILS *': set_details,
            'SET_DESCRIPTION *': set_desc,
            'PRODUCT_COLOR *': color,
            'ARTICLE_NUMBER *': article,
            'MODEL_NAME *': model_name,
            'PRODUCT_CONDITION *': condition,
            'UNIT_OF_MEASUREMENT_SINGULAR *': 'Piece',
            'UNIT_OF_MEASUREMENT_PLURAL *': 'Pieces',
            'UNIT_OF_MEASUREMENT_SINGULAR_ABBREVIATION *': 'Pc',
            'UNIT_OF_MEASUREMENT_PLURAL_ABBREVIATION *': 'Pcs',
            'SELLER_SKU_ID *': sku_raw,
            'PACKAGING_TYPE *': packing,
            'DESCRIPTION': '',
            'MATERIAL *': kwargs.get('material', ''),
            'ADAPTER_CONNECTOR_TYPE *': kwargs.get('adapter_connector', ''),
            'CABLE_LENGTH_IN_METER *': kwargs.get('cable_length', ''),
            'CABLE_MATERIAL *': kwargs.get('cable_material', ''),
            'CABLE_TYPE': kwargs.get('cable_type', ''),
            'COUNTRY_OF_ORIGIN *': country,
            'EAN': '',
            'IMPORTED_BY': '',
            'KEY_FEATURES': '',
            'MANUFACTURING_YEAR': _cfg['manufacturing_year'],
            'NO_OF_ADAPTER_PORTS *': kwargs.get('num_ports', ''),
            'OUTPUT_CURRENT_OR_VOLTAGE *': kwargs.get('output_voltage', ''),
            'PORT_TYPE *': kwargs.get('port_type', ''),
            'PRODUCT_BREADTH *': B,
            'PRODUCT_DIMENSION_UOM *': dim_uom,
            'PRODUCT_HEIGHT *': H,
            'PRODUCT_LENGTH *': L,
            'PRODUCT_TYPE *': kwargs.get('product_name', ''),
            'PRODUCT_WEIGHT_IN_KG *': weight_clean,
            'PRODUCT_MANUFACTURING_CITY': '',
            'PRODUCT_MANUFACTURING_STATE': '',
            'SUITABLE_FOR': '',
            'WARRANTY': '',
            'MANUFACTURER': '',
            'hsnCode *': hsn,
            'gstPercentage *': gst,
            'cgstShare *': _cfg['gst_cgst'],
            'sgstShare *': _cfg['gst_sgst'],
            'igstShare *': _cfg['gst_igst'],
            'cess': '',
            'sinTax': '',
            'vatPercentage': '',
            'otherCess': '',
            'validityPeriodStartDate': '',
            'validityPeriodEndDate': '',
            'declarationForm': '',
            'taxMasterStatus': _cfg['tax_master_status'],
            'BATTERY_LIFE_FOR_WIRELESS': '',
            'BLUETOOTH_VERSION_FOR_WIRELESS *': kwargs.get('bluetooth', ''),
            'CHANNELS': '',
            'CHARGING_TIME': '',
            'CONNECTOR_TYPE_FOR_WIRED *': kwargs.get('connector_type', ''),
            'CONTROL_OPTIONS': '',
            'MIC_TYPE': kwargs.get('mic_type', ''),
            'MOUNTING_OR_PLACEMENT_TYPE': '',
            'SPEAKER_TYPE *': kwargs.get('speaker_type', ''),
            'WATER_RESISTANCE': '',
            'WIRED_OR_WIRELESS *': kwargs.get('wired_wireless', ''),
            'COMPATIBLE_BRAND_MODEL *': kwargs.get('compatible_brand_model', ''),
            'CASE_COVER_TYPE *': kwargs.get('case_cover_type', ''),
            'CLOSURE_TYPE *': kwargs.get('closure_type', ''),
            'DESIGN *': kwargs.get('pattern', ''),
            'THEME': '',
            'COMPATIBLE_BRAND *': kwargs.get('compatible_brand', ''),
            'BATTERY_CAPACITY_MAH *': kwargs.get('battery_capacity', ''),
            'CHARGING_TYPE_SUPPORTED': kwargs.get('charging_type', ''),
            'OPERATING_SYSTEM_OS': kwargs.get('os', ''),
            'OS_VERSION': kwargs.get('os_version', ''),
            'DISPLAY_SIZE *': kwargs.get('display_size', ''),
            'DISPLAY_TYPE *': kwargs.get('display_type', ''),
            'DISPLAY_RESOLUTION': kwargs.get('display_resolution', ''),
            'RAM *': kwargs.get('ram', ''),
            'INTERNAL_STORAGE *': kwargs.get('storage', ''),
            'EXPANDABLE_STORAGE': kwargs.get('expandable_storage', ''),
            'SIM_TYPE': kwargs.get('sim_type', ''),
            'BLUETOOTH_VERSION': kwargs.get('bluetooth_version', ''),
            'EXPANDABLE_STORAGE_TYPE': kwargs.get('expandable_storage_type', ''),
            'EXPANDABLE_STORAGE_CAPACITY_MAX': kwargs.get('expandable_storage_max', ''),
            'BATTERY_TYPE': kwargs.get('battery_type', ''),
            'REMOVABLE_BATTERY': kwargs.get('removable_battery', ''),
            'HYBRID_SIM_SLOT': kwargs.get('hybrid_sim', ''),
            'NETWORK_TYPE_SUPPORTED': kwargs.get('network_support', ''),
            'AUDIO_JACK': kwargs.get('audio_jack', ''),
            'FM_RADIO': kwargs.get('fm_radio', ''),
            'TORCH_OR_FLASHLIGHT': kwargs.get('torch', ''),
            'RAM_ROM *': f"{kwargs.get('ram', '')} + {kwargs.get('storage', '')}" if (kwargs.get('ram') and kwargs.get('storage')) else (kwargs.get('ram', '') or kwargs.get('storage', '')),
            'PROCESSOR_BRAND_AND_MODEL_NAME': '',
            'NUMBER_OF_PROCESSOR_CORES *': kwargs.get('processor_core', ''),
            'PRIMARY_CAMERA_RESOLUTION *': kwargs.get('back_camera', ''),
            'FRONT_CAMERA_RESOLUTION *': kwargs.get('front_camera', ''),
            'REAR_FLASH': kwargs.get('rear_flash', ''),
            'SIM_SIZE': '',
            'WIFI': kwargs.get('wifi', ''),
            'FINGERPRINT_SENSOR': kwargs.get('fingerprint', ''),
            'CLOCK_SPEED': kwargs.get('clock_speed', ''),
            'REFRESH_RATE': kwargs.get('refresh_rate', ''),
            'TOUCHSCREEN_TYPE': kwargs.get('touchscreen', ''),
            'PRIMARY_CAMERA_SETUP': kwargs.get('primary_camera_setup', ''),
            'FRONT_FLASH': kwargs.get('front_flash', ''),
            'VIDEO_RECORDING_RESOLUTION': kwargs.get('video_resolution', ''),
            'FAST_CHARGING_WATTAGE': kwargs.get('fast_charging', ''),
            'WIRELESS_CHARGING_SUPPORT': kwargs.get('wireless_charging', ''),
            'GPS_SUPPORT': kwargs.get('gps', ''),
            'NFC_SUPPORT': kwargs.get('nfc', ''),
            'INFRARED_IR_BLASTER': kwargs.get('infrared', ''),
            'FINGERPRINT_SENSOR_POSITION': kwargs.get('fingerprint_pos', ''),
            'FACE_UNLOCK': kwargs.get('face_unlock', ''),
            'WATER_RESISTANCE_RATING': kwargs.get('water_resistance', ''),
            'ACTIVE_NOISE_CANCELLATION_ANC': '',
            'ADJUSTABLE_OR_FOLDABLE': '',
            'WEARING_STYLE': '',
            'MEMORY_CARD_TYPE *': kwargs.get('card_type', ''),
            'SPEED_CLASS *': kwargs.get('speed_class', ''),
            'STORAGE_CAPACITY *': kwargs.get('storage_capacity', ''),
            'COMPATIBLE_BRAND': '',
            'ADAPTER_INCLUDED': '',
            'CONNECTION_INTERFACE': '',
            'HOLDER_TYPE *': kwargs.get('holder_type', ''),
            'LOCK_MECHANISM *': kwargs.get('lock_mechanism', ''),
            'ROTATION_OR_ADJUSTABILITY *': kwargs.get('rotation_type', ''),
            'COVERAGE *': kwargs.get('coverage', ''),
            'EDGE_TYPE': '',
            'SCREEN_GUARD_OR_PROTECTOR_TYPE *': kwargs.get('screen_guard_type', ''),
            'THICKNESS *': kwargs.get('screen_thickness', ''),
            'PACKAGING_CLASSIFICATION': '',
            'SUB_BRAND': '',
            'COLOR *': color,
            'FOOD_NON-FOOD': '',
            'STRAP_COLOR *': kwargs.get('strap_color', color),
            'DISPLAY_SHAPE *': kwargs.get('display_shape', ''),
            'CORE_BRAND *': kwargs.get('core_brand', ''),
        }

        for col_name, val in row_data.items():
            if col_name in tcol and val is not None and str(val) not in ('None', ''):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

    return filled, skipped


# ═══════════════════════════════════════════════════════════════
# PV AUTO-DETECTION
# ═══════════════════════════════════════════════════════════════

def detect_pv_from_value(val):
    """Auto-detect PV from input column A value."""
    if not val:
        return None
    s = str(val).strip().lower()

    # Direct match
    for pv in CE_PV_LIST:
        if pv.lower() == s:
            return pv

    # Contains match (longer first to avoid partial matches)
    for pv in sorted(CE_PV_LIST, key=len, reverse=True):
        if pv.lower() in s or s in pv.lower():
            return pv

    # Fuzzy aliases
    aliases = {
        'adapter': 'Mobile Adapters & Cables',
        'charger': 'Mobile Adapters & Cables',
        'charging adapter': 'Mobile Adapters & Cables',
        'speaker': 'Speakers',
        'bluetooth speaker': 'Speakers',
        'case': 'Mobile Case & Covers',
        'cover': 'Mobile Case & Covers',
        'mobile case': 'Mobile Case & Covers',
        'feature phone': 'Feature Phones',
        'featurephone': 'Feature Phones',
        'headset': 'Headsets',
        'headphones': 'Headsets',
        'earphone': 'Earphones',
        'earphones': 'Earphones',
        'tws': 'TWS Ear Buds',
        'earbuds': 'TWS Ear Buds',
        'neckband': 'Neck Bands',
        'memory card': 'Memory Cards',
        'sd card': 'Memory Cards',
        'micro sd': 'Memory Cards',
        'cable': 'Mobile Cables',
        'charging cable': 'Mobile Cables',
        'data cable': 'Mobile Cables',
        'holder': 'Mobile Holders',
        'mobile holder': 'Mobile Holders',
        'screen guard': 'Screen Guards / Protectors',
        'screen protector': 'Screen Guards / Protectors',
        'tempered glass': 'Screen Guards / Protectors',
        'power bank': 'Power Bank',
        'powerbank': 'Power Bank',
        'smartphone': 'Smartphones',
        'smart phone': 'Smartphones',
        'smartwatch': 'Smart Watches',
        'smart watch': 'Smart Watches',
    }

    for alias, pv in aliases.items():
        if alias in s:
            return pv

    return None


def detect_pvs_in_dump(all_dump, col_map):
    """Detect all PVs present in the input file."""
    pv_col = col_map.get('pv')
    if not pv_col or pv_col not in all_dump.columns:
        return []

    found = []
    for val in all_dump[pv_col].dropna().unique():
        v = str(val).strip()
        if v and v.lower() not in ('nan', 'none', ''):
            detected = detect_pv_from_value(v)
            if detected and detected not in found:
                found.append(detected)
    return found


# ═══════════════════════════════════════════════════════════════
# FLASK ROUTES
# ═══════════════════════════════════════════════════════════════

# In-memory file store
FILE_STORE = {}


def write_log(email, action, details=''):
    LOG_PATH = os.path.join(os.path.dirname(__file__), 'activity.log')
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


@app.route('/ce_unified_config', methods=['GET'])
def ce_unified_config_get():
    return jsonify(get_ce_unified_config())


@app.route('/ce_unified_config', methods=['POST'])
def ce_unified_config_post():
    cfg = get_ce_unified_config()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(CE_CONFIG_PATH, cfg)
    write_log('anonymous', 'ce_unified_config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})


@app.route('/ce_unified_pvs')
def ce_unified_pvs():
    return jsonify({
        'pvs': CE_PV_LIST_RAW,
        'pv_names': CE_PV_LIST,
        'title_conventions': {k: v['title_convention'] for k, v in CE_TITLE_CONVENTIONS.items()},
    })


@app.route('/detect_ce_unified_pvs', methods=['POST'])
def detect_ce_unified_pvs():
    try:
        dump_file = request.files.get('dump')
        if not dump_file:
            return jsonify({'pvs': []})
        xl = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try:
                frames.append(xl.parse(sname))
            except:
                pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map = build_col_map(all_dump, CE_DUMP_COL_HINTS)
        detected = detect_pvs_in_dump(all_dump, col_map)
        return jsonify({
            'pvs': detected,
            'pv_count': len(detected),
        })
    except Exception as e:
        return jsonify({'pvs': [], 'error': str(e)})


@app.route('/process_ce_unified', methods=['POST'])
def process_ce_unified():
    try:
        pvs_raw = request.form.get('pvs', '')
        try:
            selected_pvs = json.loads(pvs_raw)
        except:
            selected_pvs = [s.strip() for s in pvs_raw.split(',') if s.strip()]

        inline_cfg_raw = request.form.get('config', '')
        if inline_cfg_raw:
            try:
                inline_cfg = json.loads(inline_cfg_raw)
                if inline_cfg.get('brands'):
                    inline_cfg['brands'] = normalize_brands(inline_cfg['brands'])
                disk_cfg = get_ce_unified_config()
                disk_cfg.update(inline_cfg)
                _save_config(CE_CONFIG_PATH, disk_cfg)
            except Exception as e:
                print(f'inline config parse error: {e}')

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not dump_file:
            return jsonify({'error': 'Dump / listing file is required'}), 400

        dump_bytes = dump_file.read()
        xl = pd.ExcelFile(io.BytesIO(dump_bytes))
        frames = []
        for sname in xl.sheet_names:
            try:
                frames.append(xl.parse(sname))
            except:
                pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if all_dump.empty:
            return jsonify({'error': 'Could not read any data from dump file'}), 400

        col_map = build_col_map(all_dump, CE_DUMP_COL_HINTS)
        pv_col = col_map.get('pv')

        # Auto-detect PVs if none selected
        if not selected_pvs:
            selected_pvs = detect_pvs_in_dump(all_dump, col_map)
            if not selected_pvs:
                return jsonify({'error': 'No PVs detected in input file. Please select PVs manually or check Column A.'}), 400

        # Validate selected PVs
        for pv in selected_pvs:
            if pv not in CE_SUBTYPE_MAP:
                return jsonify({'error': f'PV "{pv}" not found in template'}), 400

        # Build existing articles/skus set
        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf = bxl.parse(sname)
                    bcol = build_col_map(bdf, CE_BASE_COL_HINTS)
                    if 'article' in bcol:
                        existing_articles |= set(bdf[bcol['article']].dropna().astype(str).str.strip().str.upper())
                    if 'sku' in bcol:
                        existing_skus |= set(bdf[bcol['sku']].dropna().astype(str).str.strip().str.upper())
                except:
                    pass

        results, all_skipped, grand_filled = [], [], 0
        preview_rows, preview_cols = [], []

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for pv_name in selected_pvs:
                # Filter rows for this PV
                if pv_col and pv_col in all_dump.columns:
                    mask = all_dump[pv_col].astype(str).apply(
                        lambda x: detect_pv_from_value(x) == pv_name
                    )
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        mask2 = all_dump[pv_col].astype(str).str.lower().str.contains(
                            re.escape(pv_name.lower()), na=False
                        )
                        filtered = all_dump[mask2].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                wb, headers = get_ce_unified_template_wb_for_pv(pv_name)
                ws = wb.active
                filled, skipped = fill_ce_unified_template(
                    ws, headers, filtered, col_map, pv_name, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_pv = re.sub(r"[^\w\s-]", "", pv_name).replace(" ", "_")
                fname = f'ce_unified_{safe_pv}.xlsx'
                xls_buf = io.BytesIO()
                wb.save(xls_buf)
                zout.writestr(fname, xls_buf.getvalue())
                results.append({
                    'pv': pv_name,
                    'pv_id': CE_PV_ID_MAP.get(pv_name, ''),
                    'filled': filled,
                    'skipped': len(skipped),
                    'filename': fname,
                })

                if not preview_cols:
                    pcols = [
                        'title *', 'internalTitle *', 'ChildSKU *', 'ARTICLE_NUMBER *',
                        'MRP *', 'SellingPrice *', 'PRODUCT_COLOR *', 'PVID *',
                        'SubType', 'SET_DETAILS *', 'hsnCode *', 'gstPercentage *'
                    ]
                    preview_cols = [c for c in pcols if c in headers]

                for r in range(2, min(filled + 2, 52)):
                    rdata = {}
                    for c in preview_cols:
                        if c in headers:
                            rdata[c] = ws.cell(r, headers.index(c) + 1).value
                    if any(v for v in rdata.values()):
                        preview_rows.append({**rdata, '_pv': pv_name})

        zip_buf.seek(0)
        if len(selected_pvs) == 1:
            safe_pv = re.sub(r"[^\w\s-]", "", selected_pvs[0]).replace(" ", "_")
            out_name = f'ce_unified_{safe_pv}.xlsx'
            out_ext = '.xlsx'
            with zipfile.ZipFile(io.BytesIO(zip_buf.getvalue())) as zin:
                out_bytes = zin.read(results[0]['filename'])
        else:
            out_name = 'ce_unified_templates.zip'
            out_ext = '.zip'
            out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {
            'bytes': out_bytes,
            'filename': out_name,
            'ext': out_ext,
            'created': time.time(),
        }

        write_log('anonymous', 'ce_unified_catalog_generated',
                  f'pvs={selected_pvs} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status': 'ok',
            'grand_filled': grand_filled,
            'grand_skipped': len(all_skipped),
            'results': results,
            'skipped_details': all_skipped[:50],
            'preview': preview_rows,
            'preview_cols': preview_cols,
            'download_token': file_token,
            'filename': out_name,
            'is_zip': len(selected_pvs) > 1,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/download_ce_unified/<token>')
def download_ce_unified(token):
    if '..' in token or '/' in token or '\\' in token:
        return 'Invalid', 400
    if token in FILE_STORE:
        file_data = FILE_STORE[token]
        ext = file_data['ext']
        fname = request.args.get('filename', file_data['filename'])
        mtype = 'application/zip' if ext == '.zip' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        write_log('anonymous', 'ce_unified_file_downloaded', fname)
        return send_file(io.BytesIO(file_data['bytes']), as_attachment=True,
                         download_name=fname, mimetype=mtype)
    return 'File not found', 404


@app.route('/download_ce_unified_template/<pv_name>')
def download_ce_unified_template(pv_name):
    if pv_name not in CE_SUBTYPE_HEADER_ROW:
        return jsonify({'error': f'PV "{pv_name}" not found in template'}), 404

    wb, headers = get_ce_unified_template_wb_for_pv(pv_name)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    safe_name = re.sub(r"[^\w\s-]", "", pv_name).replace(" ", "_")
    fname = f'CE_Unified_Template_{safe_name}.xlsx'
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



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
    pv_cfg = ap_cfg.get('pv_config') or AP_DEFAULT_CONFIG['pv_config']
    for k, v in pv_cfg.items():
        if k.lower() == category_key.lower():
            return v
    for k, v in pv_cfg.items():
        if category_key.lower() in k.lower() or k.lower() in category_key.lower():
            return v
    return {}

def _ap_detect_pv_from_row(drow, col_map, ap_cfg, category_key=None):
    """
    Auto-detect PV config from row data. Uses gender + sub-type + row clues
    (Neck, Sleeve, Pattern) to disambiguate when multiple PVs share the same
    industry_sub_type (e.g., Casual T-Shirts vs Polo T-Shirts both = "T-Shirts").
    Returns (pv_cfg_dict, super_category, detected_key)
    """
    pv_cfg_map = ap_cfg.get('pv_config') or AP_DEFAULT_CONFIG['pv_config']
    
    def _clean(val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        s = str(val).strip()
        return '' if s.lower() in ('nan', 'none', '') else s

    def _norm_gender(g):
        g = g.lower().strip()
        if g in ("men's", "male", "men", "mens"): return "men"
        if g in ("women's", "female", "women", "womens", "woman"): return "women"
        if g in ("boy's", "boys", "boy"): return "boy"
        if g in ("girl's", "girls", "girl"): return "girl"
        if g in ("baby's", "baby", "babies"): return "baby"
        return g

    def _pv_gender(pv_sub_category):
        sc = pv_sub_category.lower().strip()
        if sc.startswith("menswear") or sc == "men": return "men"
        if sc.startswith("womenswear") or sc == "women": return "women"
        if sc.startswith("boyswear") or sc == "boy": return "boy"
        if sc.startswith("girlswear") or sc == "girl": return "girl"
        if sc.startswith("babywear") or sc == "baby": return "baby"
        return ""

    def _type_matches(sub_type, pv_key, pv_sub_type):
        """Strict type matching: 'shirts' must NOT match 't-shirts'"""
        st = sub_type.lower()
        pk = pv_key.lower()
        pst = pv_sub_type.lower()
        if pst == st or pk == st: return True
        if st in pk:
            # Prevent "shirts" from matching "t-shirts"
            if st == "shirts" and "t-shirts" in pk:
                return False
            return True
        if pk in st: return True
        return False

    # ── Read sub_type from row ──
    sub_type_raw = ''
    sub_type_col = col_map.get('ind_sub_type')
    if sub_type_col:
        try: sub_type_raw = _clean(drow.get(sub_type_col, ''))
        except: pass
    if not sub_type_raw:
        for hint_col in ['product_name', 'ind_product_type', 'type']:
            c = col_map.get(hint_col)
            if c:
                try:
                    val = _clean(drow.get(c, ''))
                    if val: sub_type_raw = val; break
                except: pass

    # ── Read gender from row ──
    gender_raw = ''
    gender_col = col_map.get('gender')
    if gender_col:
        try: gender_raw = _clean(drow.get(gender_col, ''))
        except: pass
    target_gender = _norm_gender(gender_raw)

    # ── Collect disambiguation clues ──
    clues = {}
    for clue_key, col_key in [('neck', 'neck_type'), ('sleeve', 'sleeve_length'), ('pattern', 'pattern')]:
        c = col_map.get(col_key)
        if c:
            try: clues[clue_key] = _clean(drow.get(c, '')).lower()
            except: pass

    # ── Step 1: Filter candidates by gender + sub_type ──
    candidates = []
    for k, v in pv_cfg_map.items():
        pv_gender = _pv_gender(v.get('industry_sub_category', ''))
        pv_sub_type = v.get('industry_sub_type', '').lower().strip()
        if pv_gender == target_gender and _type_matches(sub_type_raw, k, pv_sub_type):
            candidates.append((k, v))

    if len(candidates) == 1:
        k, v = candidates[0]
        return v, v.get('super_category', 'Jeans'), k

    # ── Step 2: Disambiguate with clues ──
    if len(candidates) > 1:
        best_match = None
        best_score = -999
        for k, v in candidates:
            score = 0
            kl = k.lower()
            neck = clues.get('neck', '')
            sleeve = clues.get('sleeve', '')
            pattern = clues.get('pattern', '')

            # Neck type scoring
            if neck:
                if 'polo' in neck and 'polo' in kl: score += 20
                elif 'polo' in neck and 'polo' not in kl: score -= 20
                if 'round' in neck:
                    if 'casual' in kl and 't-shirt' in kl: score += 10
                    if 'polo' in kl: score -= 10
                if 'v-' in neck or 'vneck' in neck:
                    if 'casual' in kl: score += 8
                    if 'polo' in kl: score -= 8
                if 'button' in neck or ('collar' in neck and 'polo' not in neck):
                    if 'shirt' in kl and 't-shirt' not in kl: score += 15
                    if 't-shirt' in kl: score -= 10
            else:
                # No neck clue: prefer base variant (non-polo, non-formal)
                if 'polo' not in kl and 'formal' not in kl: score += 3

            # Sleeve scoring
            if sleeve:
                if 'full' in sleeve:
                    if 'formal' in kl: score += 10
                    if 'casual' in kl and 'shirt' in kl and 't-shirt' not in kl: score += 5
                if 'half' in sleeve or 'short' in sleeve:
                    if 't-shirt' in kl: score += 5
                    if 'formal' in kl: score -= 5

            # Pattern scoring
            if pattern:
                if 'embroidery' in pattern and 'polo' in kl: score += 5
                if 'checked' in pattern and 'casual' in kl and 'shirt' in kl: score += 5
                if 'solid' in pattern:
                    if 'formal' in kl: score += 8
                    if 'casual' in kl and 't-shirt' in kl and 'polo' not in kl: score -= 3

            score += len(k) * 0.01  # Tie-breaker
            if score > best_score:
                best_score = score
                best_match = (k, v)

        if best_match:
            k, v = best_match
            return v, v.get('super_category', 'Jeans'), k

    if candidates:
        k, v = candidates[0]
        return v, v.get('super_category', 'Jeans'), k

    # ── Step 3: Fallback to category_key ──
    if category_key and pv_cfg_map:
        cat_lower = category_key.lower().strip()
        for k, v in pv_cfg_map.items():
            kl = k.lower()
            super_cat = v.get('super_category', '').lower()
            if kl == cat_lower or cat_lower in kl or kl in cat_lower:
                return v, v.get('super_category', 'Jeans'), k
            if super_cat:
                def _norm(s): return s.lower().replace('-', '').replace(' ', '').rstrip('s')
                if _norm(super_cat) == _norm(cat_lower):
                    return v, v.get('super_category', 'Jeans'), k

    # ── Step 4: Ultimate fallback ──
    if not category_key and pv_cfg_map:
        first_k = next(iter(pv_cfg_map))
        first_v = pv_cfg_map[first_k]
        return first_v, first_v.get('super_category', 'Jeans'), first_k

    return {}, 'Jeans', ''


def _ap_derive_product_type(pv_name):
    """Derive PRODUCT_TYPE from PV name. More specific terms checked FIRST."""
    if not pv_name:
        return ''
    pv_lower = str(pv_name).lower()

    # Order CRITICAL: specific variants BEFORE generic terms
    # e.g., 'polo t-shirts' MUST be checked BEFORE 'shirts'
    product_types = [
        # T-Shirt variants (check BEFORE 'shirts')
        'polo t-shirts',
        'casual t-shirts',
        't-shirts',
        # Shirt variants
        'casual shirts',
        'formal shirts',
        'shirts',
        # Other apparel
        'track pants',
        'cargo',
        'jeans',
        'camisole',
        'slips',
        'sarees',
        'blouses',
    ]

    for pt in product_types:
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


@app.route('/ap_categories')
def get_ap_categories():
    return jsonify({'categories': AP_CATEGORIES})

@app.route('/config', methods=['GET'])
def config_get_route():
    return jsonify(get_config())


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


def _generate_blank_template(subtype):
    """Generate a blank Excel template containing only the header row for a specific subtype."""
    from openpyxl import Workbook

    if subtype not in SUBTYPE_HEADER_ROW:
        return None, f'SubType "{subtype}" not found in template'

    wb_src = load_workbook(TEMPLATE_PATH)
    ws_src = wb_src['PV Template']
    hdr_row = SUBTYPE_HEADER_ROW.get(subtype, 1)

    # Extract headers
    headers = []
    for c in range(1, ws_src.max_column + 1):
        val = ws_src.cell(hdr_row, c).value
        if val is not None:
            headers.append(str(val).strip())

    # Create new workbook with single sheet
    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.title = 'PV Template'

    # Write header row
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h

    # Auto-adjust column widths
    for ci, h in enumerate(headers, 1):
        ws_new.column_dimensions[ws_new.cell(1, ci).column_letter].width = max(len(h) + 2, 15)

    buf = io.BytesIO()
    wb_new.save(buf)
    buf.seek(0)
    return buf, None

@app.route('/download_template/<path:category>')
def download_template(category):
    """Serve template file for a given vertical.

    Query params:
      - subtype: specific SubType to download (optional). If provided, returns blank template
                 with only that subtype's headers.
    """
    category_lower = category.lower().strip()
    subtype = request.args.get('subtype', '').strip()

    # If subtype is specified, generate blank template for that subtype
    if subtype:
        if category_lower not in ('footwear', 'fw', ''):
            return jsonify({'error': 'SubType-specific download only supported for Footwear'}), 400

        buf, err = _generate_blank_template(subtype)
        if err:
            return jsonify({'error': err}), 404

        safe_name = re.sub(r"[^\w\s-]", "", subtype).replace(" ", "_")
        fname = f'Footwear_Template_{safe_name}.xlsx'
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Otherwise, serve the full master template file
    if 'electronic' in category_lower or category_lower == 'ce':
        path  = os.path.join(os.path.dirname(__file__), 'Unified Template Creation.xlsx')
        fname = 'Unified_Template_Creation.xlsx'
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

def get_ts_config_from_disk():
    cfg = _load_config(TS_CONFIG_PATH, TS_DEFAULT_CONFIG)
    cfg['pv_config'] = TS_DEFAULT_CONFIG['pv_config']
    return cfg


def _ts_get_pv_config(pv_name, ts_cfg):
    pv_cfg_map = ts_cfg.get('pv_config') or TS_DEFAULT_CONFIG['pv_config']
    key = pv_name.lower().strip()
    # Exact match first
    if key in pv_cfg_map:
        return pv_cfg_map[key]
    # Then check for exact equality or containment (longer keys first to avoid partial matches)
    for k in sorted(pv_cfg_map.keys(), key=len, reverse=True):
        v = pv_cfg_map[k]
        kl = k.lower()
        if kl == key or key == kl or key in kl or kl in key:
            return v
    return next(iter(pv_cfg_map.values()))


def _ts_yes_no(val, yes_text, no_text):
    s = safe(val).strip().lower()
    if s in ('yes', 'y', 'true', '1'): return yes_text
    return no_text


def _ts_extract_tyre_size_short(tyre_size):
    """26x2.40 -> 26 ; 700C / 700X35C -> 700C"""
    s = safe(tyre_size).strip()
    if not s: return ''
    if '700' in s.upper():
        return '700C'
    m = re.match(r'^(\d+)', s)
    return m.group(1) if m else s


def _ts_make_title(brand, sub_brand, tyre_size, pv, skd_ckd, cycle_type, ibc,
                   branded_tyre, tyre_specs, brake_type, color):
    """
    Title format: BrandName + Sub Brand + CYCLE_TYRE_SIZE + Product Verticle + SKD CKD
                  + CYCLE TYPE + IBC (with/without IBC) + Branded/Non Branded
                  + Tyre Specs + With + BRAKE_TYPE, + Product Primary Colour
    """
    ibc_text = _ts_yes_no(ibc, 'With IBC', 'Without IBC')
    branded_text = _ts_yes_no(branded_tyre, 'Branded', 'Non Branded')
    parts = [p for p in [brand, sub_brand, tyre_size, pv, skd_ckd, cycle_type,
                          ibc_text, branded_text, tyre_specs] if p]
    base = ' '.join(parts)
    with_brake = f"With {brake_type}" if brake_type else ''
    if with_brake:
        base = f"{base} {with_brake}"
    return f"{base}, {color}" if color else base


def _ts_make_internal_title(brand, article, tyre_size, pv, skd_ckd, cycle_type, ibc,
                             branded_tyre, tyre_specs, brake_type, color, set_count):
    ibc_text = _ts_yes_no(ibc, 'With IBC', 'Without IBC')
    branded_text = _ts_yes_no(branded_tyre, 'Branded', 'Non Branded')
    parts = [p for p in [brand, article, tyre_size, pv, skd_ckd, cycle_type,
                          ibc_text, branded_text, tyre_specs] if p]
    base = ' '.join(parts)
    with_brake = f"With {brake_type}" if brake_type else ''
    if with_brake:
        base = f"{base} {with_brake}"
    suffix = f"{color}, (Set of {set_count})" if color else f"(Set of {set_count})"
    return f"{base}, {suffix}"


def _ts_make_article_model(sub_brand, branded_tyre, tyre_specs, ibc):
    branded_text = _ts_yes_no(branded_tyre, 'Branded', 'Non-Branded')
    ibc_text = _ts_yes_no(ibc, 'With IBC', 'Without IBC')
    parts = [p for p in [sub_brand, branded_text, tyre_specs, ibc_text] if p]
    return '_'.join(parts)


def fill_ts_files(rows_df, col_map, pv_category, existing_articles, existing_skus):
    """
    Generate 5 output workbooks for one Toys & Sports PV (Gear / Non Gear / Battery Operated).
    Returns: (wb_jpin, wb_tax, wb_pav, wb_scm, wb_l4, filled, skipped)
    """
    _ts_cfg = get_ts_config_from_disk()
    brands_dict = normalize_brands(_ts_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))

    pv_cfg = _ts_get_pv_config(pv_category, _ts_cfg)
    pv_id   = pv_cfg.get('pv_id', '')
    pv_name = pv_cfg.get('pv_name', pv_category)
    ind_cat       = pv_cfg.get('industry_category', 'Toys and Sports')
    ind_sub_cat   = pv_cfg.get('industry_sub_category', 'Cycle')
    ind_prod_type = pv_cfg.get('industry_product_type', pv_category)

    # ── JPIN headers ─────────────────────────────────────────────
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

    # ── TaxMaster headers ────────────────────────────────────────
    TAX_HEADERS = [
        'TaxMasterID','Jpin','Title','ProductVerticalId','ProductVerticalName',
        'hsnCode','sinTax','cess','vatPercentage','gstPercentage',
        'cgstComponentShare','sgstComponentShare','IgstComponentShare',
        'Validity_Period_Start','Validity_Period_End','declarationForm','otherCess','status',
    ]

    # ── PAV headers ──────────────────────────────────────────────
    PAV_HEADERS = [
        'Jpin','Title','PvId','PvName','BrandId','BrandName',
        'ImageURL1','ImageURL2','CatalogStatus','StatusRemark',
        'SUB_BRAND','USER_TYPE','DESCRIPTION','MATERIAL','CERTIFICATION',
        'COUNTRY_OF_ORIGIN','EAN','IMPORTED_BY','KEY_FEATURES','MANUFACTURING_YEAR',
        'PRODUCT_BREADTH','PRODUCT_DIMENSION_UOM','PRODUCT_HEIGHT','PRODUCT_LENGTH',
        'PRODUCT_WEIGHT_IN_KG','PRODUCT_MANUFACTURING_CITY','PRODUCT_MANUFACTURING_STATE',
        'WARRANTY','ASSEMBLY_REQUIRED','BOTTLE_HOLDER_INCLUDED','BRAKE_TYPE',
        'BUILT-IN_MUSIC_AND_LIGHTS','CARRIER_OR_BASKET','CHAIN_GUARD',
        'CYCLE_TYRE_SIZE','FRAME_SIZE','FRAME_SIZE_UOM','GEAR_TYPE',
        'HANDLE_TYPE','HORN_OR_BELL','KICKSTAND','LIGHT_INCLUDED','MANUFACTURER',
        'MUDGUARD','NUMBER_OF_GEARS','NUMBER_OF_WHEELS','PEDAL_TYPE','PORTABILITY',
        'RECOMMENDED_AGE','REFLECTORS','RIM_MATERIAL','SADDLE_TYPE','SAFETY_FEATURES',
        'SUPPORTED_WHEELS','SUSPENSION','TIRE_TYPE','TYPE','WHEEL_MATERIAL',
        'WHEEL_SIZE','WHEEL_SIZE_UOM','TYRE_BRANDED_OR_NOT','CKD_SKD','TYRE_SPECS',
        'FORK_TYPE','IBC','REAR_SUSPENSION','SUSPENSION_TYPE','MODE_OF_OPERATION',
        'BRAKE_LEVER_MATERIAL','NUMBER_OF_SPOKES','WATER_BOTTLE_HOLDER',
        'FRAME_TYPE','BASKET',
    ]

    # ── SCM headers ──────────────────────────────────────────────
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

    # ── L4 headers ───────────────────────────────────────────────
    L4_HEADERS = [
        '*Type','*Industry Category','*Industry Sub Category','*Product type',
        'Product Sub-type','*Seller SKU','Product ID','*Product Name','*Product Description',
        '*Brand','Model Name','*Material Type','Tyre Size','Colour','*Select Age Group',
        'Variation Theme Id','*HSN Code','*GST','Tags','Manufacturer','Warranty',
        '*Gender','*Manufacturing Year','*Country/Region of Origin','Return Policy',
        '*Certification','*License','*Relationship','*Parent Product Id','*Child SKU',
        'Assembly Required','EAN Number','*Minimum Order Quantity','*MRP','*Selling Price',
        'Product Weight (In KG)','*Unit Of Measure','Geared','Rim Material','Frame Material',
        'Support Wheels','Operation Mode','Estimated Dispatch Time',
        'Product Dimension (LXBXH)','Product Dimension UOM','Packaging weight(KG)',
        '*Main Image URL','Other Image URL1','Other Image URL2','Other Image URL3',
        'Other Image URL4','Other Image URL5','Other Image URL6','Solv Commission','JPIN',
    ]

    def _make_wb(headers, sheet_name):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        for ci, h in enumerate(headers, 1):
            ws.cell(1, ci).value = h
        return wb, ws

    wb_jpin, ws_jpin = _make_wb(JPIN_HEADERS, 'JPIN Template')
    wb_tax,  ws_tax  = _make_wb(TAX_HEADERS, 'TaxMaster')
    wb_pav,  ws_pav  = _make_wb(PAV_HEADERS, 'ProductAttributeValue')
    wb_scm,  ws_scm  = _make_wb(SCM_HEADERS, 'SCM')
    wb_l4,   ws_l4   = _make_wb(L4_HEADERS,  'L4')

    def _col(headers):
        return {h: i+1 for i, h in enumerate(headers) if h}

    tcol_jpin = _col(JPIN_HEADERS)
    tcol_tax  = _col(TAX_HEADERS)
    tcol_pav  = _col(PAV_HEADERS)
    tcol_scm  = _col(SCM_HEADERS)
    tcol_l4   = _col(L4_HEADERS)

    def _write(ws, tcol, data, row_idx):
        for col_name, val in data.items():
            if col_name in tcol and val is not None and str(val) not in ('None',):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

    skipped, filled = [], 0

    for _, drow in rows_df.iterrows():
        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand = fallback_brand
            brand_id = fallback_id

        sku_raw      = safe(drow.get(col_map.get('sku',''), ''))
        product_code = safe(drow.get(col_map.get('product_code',''), ''))
        sub_brand    = safe(drow.get(col_map.get('sub_brand',''), ''))
        branded_tyre = safe(drow.get(col_map.get('branded_tyre',''), ''))
        tyre_specs   = safe(drow.get(col_map.get('tyre_specs',''), ''))
        ibc          = safe(drow.get(col_map.get('ibc',''), ''))

        article = _ts_make_article_model(sub_brand, branded_tyre, tyre_specs, ibc)
        if not article:
            article = product_code or sku_raw

        if article.upper() in existing_articles or sku_raw.upper() in existing_skus:
            skipped.append({'sku': sku_raw, 'article': article, 'reason': 'Already exists in base data'})
            continue

        # Read all fields
        product_desc  = safe(drow.get(col_map.get('product_desc',''), ''))
        color_raw     = safe(drow.get(col_map.get('color',''), ''))
        color         = title_case_color(color_raw)
        set_or_unit   = safe(drow.get(col_map.get('set_or_unit',''), 'Unit'))
        sc_raw        = safe(drow.get(col_map.get('set_count',''), '1'))
        sc_nums       = re.findall(r'\d+', sc_raw)
        set_count     = int(sc_nums[0]) if sc_nums else 1
        gender_raw    = safe(drow.get(col_map.get('gender',''), ''))
        material      = safe(drow.get(col_map.get('material',''), ''))
        license_raw   = safe(drow.get(col_map.get('license',''), 'No'))
        certification = safe(drow.get(col_map.get('certification',''), ''))
        warranty      = safe(drow.get(col_map.get('warranty',''), 'No'))
        assembly      = safe(drow.get(col_map.get('assembly',''), 'No'))
        brake_type    = safe(drow.get(col_map.get('brake_type',''), ''))
        tyre_size_raw = safe(drow.get(col_map.get('tyre_size',''), ''))
        tyre_size_short = _ts_extract_tyre_size_short(tyre_size_raw)
        tyre_size_uom = safe(drow.get(col_map.get('tyre_size_uom',''), ''))
        frame_size    = safe(drow.get(col_map.get('frame_size',''), ''))
        frame_size_uom= safe(drow.get(col_map.get('frame_size_uom',''), 'Inch'))
        num_gears     = safe(drow.get(col_map.get('num_gears',''), ''))
        rec_age       = safe(drow.get(col_map.get('recommended_age',''), ''))
        rim_material  = safe(drow.get(col_map.get('rim_material',''), ''))
        tire_type     = safe(drow.get(col_map.get('tire_type',''), ''))
        cycle_type    = safe(drow.get(col_map.get('cycle_type',''), ''))
        wheel_size    = safe(drow.get(col_map.get('wheel_size',''), ''))
        wheel_size_uom= safe(drow.get(col_map.get('wheel_size_uom',''), 'inch'))
        skd_ckd       = safe(drow.get(col_map.get('skd_ckd',''), ''))
        basket        = safe(drow.get(col_map.get('basket',''), ''))
        suspension    = safe(drow.get(col_map.get('suspension_type',''), ''))
        fork_type     = safe(drow.get(col_map.get('fork_type',''), ''))
        rear_susp     = safe(drow.get(col_map.get('rear_suspension',''), ''))
        frame_type    = safe(drow.get(col_map.get('frame_type',''), ''))
        mode_op       = safe(drow.get(col_map.get('mode_of_operation',''), ''))
        brake_lever   = safe(drow.get(col_map.get('brake_lever_mat',''), ''))
        num_spokes    = safe(drow.get(col_map.get('num_spokes',''), ''))
        chain_guard   = safe(drow.get(col_map.get('chain_guard',''), ''))
        water_bottle  = safe(drow.get(col_map.get('water_bottle',''), ''))

        img_url  = safe(drow.get(col_map.get('image',''), ''))
        img2_url = safe(drow.get(col_map.get('image2',''), ''))
        img3_url = safe(drow.get(col_map.get('image3',''), ''))
        img4_url = safe(drow.get(col_map.get('image4',''), ''))
        img5_url = safe(drow.get(col_map.get('image5',''), ''))
        img6_url = safe(drow.get(col_map.get('image6',''), ''))

        mrp_raw    = drow.get(col_map.get('set_mrp',''), '')
        sp_raw     = drow.get(col_map.get('set_sp',''), '')
        moq_raw    = drow.get(col_map.get('moq',''), 1)
        hsn_raw    = drow.get(col_map.get('hsn',''), '')
        gst_raw    = drow.get(col_map.get('gst',''), 0.05)
        weight_raw = safe(drow.get(col_map.get('weight',''), ''))
        dim_raw    = safe(drow.get(col_map.get('dims',''), ''))
        dim_uom    = safe(drow.get(col_map.get('unit_measure',''), 'INCH'))
        country    = safe(drow.get(col_map.get('country',''), '')) or _ts_cfg['country_of_origin']
        mfg_year   = safe(drow.get(col_map.get('mfg_year',''), '')) or _ts_cfg['manufacturing_year']

        try:    hsn = int(float(hsn_raw)) if str(hsn_raw).strip() not in ('','nan') else ''
        except: hsn = ''
        try:
            gst_f = float(gst_raw)
            gst = int(gst_f * 100) if gst_f < 1 else int(gst_f)
        except: gst = 5
        try:    mrp = float(mrp_raw) if str(mrp_raw).strip() not in ('','nan') else ''
        except: mrp = ''
        try:    sp  = float(sp_raw)  if str(sp_raw).strip()  not in ('','nan') else ''
        except: sp  = ''
        try:    moq = int(float(moq_raw))
        except: moq = 1

        # Weight: strip "kg" etc
        weight_clean = ''
        if weight_raw:
            m = re.search(r'([0-9.]+)', weight_raw)
            if m: weight_clean = float(m.group(1))

        L, B, H = parse_lbh(dim_raw)

        # Titles
        title = _ts_make_title(
            brand, sub_brand, tyre_size_short, pv_name, skd_ckd, cycle_type,
            ibc, branded_tyre, tyre_specs, brake_type, color
        )
        internal_title = _ts_make_internal_title(
            brand, article, tyre_size_short, pv_name, skd_ckd, cycle_type,
            ibc, branded_tyre, tyre_specs, brake_type, color, set_count
        )

        # Set fields
        set_name_str = f'Set of {set_count}' if set_count > 1 else 'Set of 1'
        set_details  = f'{color}/{set_count}' if color else f'Assorted/{set_count}'
        set_desc     = f'{set_count}pc of {color}' if color else f'{set_count}pc of Assorted'
        set_comp     = f'{color} :- {set_count}' if color else f'Assorted :- {set_count}'

        # Derived
        cert_text    = _ts_yes_no(certification, 'BIS', 'Non BIS')
        license_text = _ts_yes_no(license_raw, 'Licensed', 'Non Licensed')
        ibc_text     = _ts_yes_no(ibc, 'With IBC', 'Without IBC')
        branded_text = _ts_yes_no(branded_tyre, 'Branded', 'Non Branded')
        gear_type    = 'Geared' if 'gear' in pv_name.lower() and 'non' not in pv_name.lower() else ''
        num_gears_pav = num_gears if gear_type == 'Geared' else ''

        filled += 1
        row_idx = filled + 1

        # ── JPIN ──
        jpin_row = {
            'JPIN':                                    '',
            'Title':                                   title,
            'Internal_Title':                          internal_title,
            'BrandID':                                 brand_id,
            'BrandName':                               brand,
            'PVID':                                    pv_id,
            'PVName':                                  pv_name,
            'Business Category Id':                    _ts_cfg['biz_cat_id'],
            'Business Category Name':                  _ts_cfg['biz_cat_name'],
            'Product Identifier':                      'Set',
            'Set Name':                                set_name_str,
            'Set Count':                               set_count,
            'Pack Name':                               'Pack of 1',
            'Pack of':                                 1,
            'is Combo':                                'yes',
            'Available Sizes':                         tyre_size_short,
            'Set Details':                             set_details,
            'Set Description':                         set_desc,
            'Set Composition':                         set_comp,
            'Product Color':                           color,
            'Article Number':                          article,
            'Model Name':                              article,
            'Product Condition':                       _ts_cfg['product_condition'],
            'ImageURL1':                               img_url,
            'ImageURL2':                               img2_url,
            'ImageURL3':                               img3_url,
            'ImageURL4':                               img4_url,
            'ImageURL5':                               img5_url,
            'ImageURL6':                               img6_url,
            'CatalogStatus':                           _ts_cfg['catalog_status'],
            'StatusRemark':                            _ts_cfg['status_remark'],
            'CustomerDiscoveryCategories':             _ts_cfg['discovery_cat'],
            'Singular Unit Of Measurement':            'Piece',
            'Plural Unit Of Measurement':              'Pieces',
            'Singular Unit Of Measurement Abbreviation': 'Pc',
            'Plural Unit Of Measurement Abbreviation': 'Pcs',
            'Seller SKU ID':                           sku_raw,
            'Product Description':                     product_desc,
        }
        _write(ws_jpin, tcol_jpin, jpin_row, row_idx)

        # ── TaxMaster ──
        tax_row = {
            'Title':                title,
            'ProductVerticalId':    pv_id,
            'ProductVerticalName':  pv_name,
            'hsnCode':              hsn,
            'gstPercentage':        gst,
            'cgstComponentShare':   _ts_cfg['gst_cgst'],
            'sgstComponentShare':   _ts_cfg['gst_sgst'],
            'IgstComponentShare':   _ts_cfg['gst_igst'],
            'status':               _ts_cfg['tax_master_status'],
        }
        _write(ws_tax, tcol_tax, tax_row, row_idx)

        # ── PAV ──
        pav_row = {
            'Title':                  title,
            'PvId':                   pv_id,
            'PvName':                 pv_name,
            'BrandId':                brand_id,
            'BrandName':              brand,
            'ImageURL1':              img_url,
            'ImageURL2':              img2_url,
            'CatalogStatus':          _ts_cfg['catalog_status'],
            'StatusRemark':           _ts_cfg['status_remark'],
            'USER_TYPE':              gender_raw,
            'MATERIAL':               material,
            'CERTIFICATION':          cert_text,
            'COUNTRY_OF_ORIGIN':      country,
            'PRODUCT_BREADTH':        0,
            'PRODUCT_DIMENSION_UOM':  '#',
            'PRODUCT_HEIGHT':         0,
            'PRODUCT_LENGTH':         0,
            'PRODUCT_WEIGHT_IN_KG':   0,
            'WARRANTY':               warranty,
            'ASSEMBLY_REQUIRED':      assembly,
            'BRAKE_TYPE':             brake_type,
            'CYCLE_TYRE_SIZE':        tyre_size_short,
            'FRAME_SIZE':             tyre_size_short,
            'FRAME_SIZE_UOM':         frame_size_uom,
            'GEAR_TYPE':              gear_type,
            'NUMBER_OF_GEARS':        num_gears_pav,
            'RECOMMENDED_AGE':        rec_age,
            'RIM_MATERIAL':           rim_material,
            'TIRE_TYPE':              tyre_specs,
            'TYPE':                   cycle_type,
            'WHEEL_SIZE':             frame_size_uom,
            'WHEEL_SIZE_UOM':         frame_size,
            'TYRE_BRANDED_OR_NOT':    branded_text if branded_text == 'Branded' else 'Non Branded',
            'CKD_SKD':                skd_ckd,
            'TYRE_SPECS':             tyre_size_raw,
            'IBC':                    ibc_text,
        }
        _write(ws_pav, tcol_pav, pav_row, row_idx)

        # ── SCM ──
        scm_row = {
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
            'L0-UnitShippingContainerType':  'Crate - Medium',
            'L1-UnitShippingContainerType':  'Bag',
            'L2-UnitShippingContainerType':  'Bag',
            'Fragile':                       'No',
            'Brittle':                       'No',
            'length_l0':                     0,
            'width_l0':                      0,
            'height_l0':                     0,
            'length_l1':                     24.5,
            'width_l1':                      24.5,
            'height_l1':                     33.5,
            'volumetricweight_l1':           4.021675,
            'APMC Notified Commodity':       'No',
            'L1-deadWeight':                 0,
            'Net_Quantity':                  1,
            'Net_Quantity_Measuring_Unit':   'Pc',
        }
        _write(ws_scm, tcol_scm, scm_row, row_idx)

        # ── L4 ──
        l4_row = {
            '*Type':                         'Unit',
            '*Industry Category':            ind_cat,
            '*Industry Sub Category':        ind_sub_cat,
            '*Product type':                 ind_prod_type,
            '*Seller SKU':                   sku_raw,
            '*Product Name':                 internal_title,
            '*Product Description':          product_desc,
            '*Brand':                        brand,
            '*Material Type':                material,
            'Colour':                        color,
            '*Select Age Group':             rec_age,
            'Variation Theme Id':            'Unit',
            '*HSN Code':                     hsn,
            '*GST':                          gst,
            '*Gender':                       gender_raw,
            '*Manufacturing Year':           mfg_year,
            '*Country/Region of Origin':     country,
            '*Certification':                cert_text,
            '*License':                      license_text,
            '*Relationship':                 'Parent',
            '*Parent Product Id':            sku_raw,
            '*Child SKU':                    sku_raw,
            '*Minimum Order Quantity':       moq,
            '*MRP':                          mrp,
            '*Selling Price':                sp,
            'Product Weight (In KG)':        weight_clean,
            '*Unit Of Measure':              'Pcs',
            'Product Dimension (LXBXH)':     dim_raw,
            'Product Dimension UOM':         dim_uom,
            '*Main Image URL':               img_url,
        }
        _write(ws_l4, tcol_l4, l4_row, row_idx)

    return wb_jpin, wb_tax, wb_pav, wb_scm, wb_l4, filled, skipped



@app.route('/debug_config')
def debug_config():
    cfg    = get_config()
    ce_cfg = get_ce_unified_config()
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
        'ts_config_path':        TS_CONFIG_PATH,
        'ts_config_file_exists': os.path.exists(TS_CONFIG_PATH),
        'ts_brands':             get_ts_config_from_disk().get('brands', {}),
        'ts_config':             get_ts_config_from_disk(),
    })


# ═══════════════════════════════════════════════════════════════
# TOYS & SPORTS MODULE — CYCLES (Gear / Non Gear / Battery Operated)
# ═══════════════════════════════════════════════════════════════

TS_CONFIG_PATH = '/tmp/fillforge_ts_config.json'

TS_DEFAULT_CONFIG = {
    "brands":              {"Avitree": "BR-1190299999"},
    "biz_cat_id":          "BCAT-139427",
    "biz_cat_name":        "Toys & Sports",
    "relationship":        "Parent",
    "catalog_status":      "ACTIVE",
    "status_remark":       "Ready to Launch",
    "tax_master_status":   "active",
    "gst_cgst":            50,
    "gst_sgst":            50,
    "gst_igst":            0,
    "country_of_origin":   "India",
    "product_condition":   "Fresh",
    "manufacturing_year":  "2025",
    "discovery_cat":       "DISCAT-135529",
    "pv_config": {
        "gear": {
            "pv_id":   "PV-1914272830",
            "pv_name": "Gear",
            "industry_category":     "Toys and Sports",
            "industry_sub_category": "Cycle",
            "industry_product_type": "Gear",
        },
        "non gear": {
            "pv_id":   "PV-1914272829",
            "pv_name": "Non Gear",
            "industry_category":     "Toys and Sports",
            "industry_sub_category": "Cycle",
            "industry_product_type": "Non Gear",
        },
        "battery operated": {
            "pv_id":   "PV-1914272831",
            "pv_name": "Battery Operated",
            "industry_category":     "Toys and Sports",
            "industry_sub_category": "Cycle",
            "industry_product_type": "Battery Operated",
        },
    },
}

TS_CATEGORIES = ['Gear', 'Non Gear', 'Battery Operated']

TS_DUMP_COL_HINTS = {
    'pv':                ['Product Verticle','Product Vertical','PV'],
    'variant_id':        ['Variant ID'],
    'seller_name':       ['Seller Name'],
    'product_name':      ['Product Name'],
    'product_desc':      ['Product Description'],
    'brand':             ['BrandName','Brand Name','Brand'],
    'sub_brand':         ['Sub brand name','Sub Brand Name','Sub Brand'],
    'color':             ['Product Primary Colour','Product Primary Color','Product Color'],
    'set_or_unit':       ['Set or Unit'],
    'set_count':         ['Number Of Pcs in Set','Number oF Pcs in Set'],
    'sku':               ['Seller SKU ID'],
    'product_code':      ['Product Code'],
    'gender':            ['Gender'],
    'material':          ['Material'],
    'license':           ['License'],
    'certification':     ['CERTIFICATION'],
    'warranty':          ['WARRANTY'],
    'assembly':          ['ASSEMBLY_REQUIRED'],
    'brake_type':        ['BRAKE_TYPE'],
    'tyre_size':         ['CYCLE_TYRE_SIZE'],
    'tyre_size_uom':     ['CYCLE_TYRE_SIZE UOM'],
    'frame_size':        ['FRAME_SIZE'],
    'frame_size_uom':    ['FRAME_SIZE_UOM'],
    'num_gears':         ['NUMBER_OF_GEARS'],
    'recommended_age':   ['RECOMMENDED_AGE'],
    'rim_material':      ['RIM_MATERIAL'],
    'tire_type':         ['TIRE_TYPE'],
    'cycle_type':        ['CYCLE TYPE','CYCLE_TYPE'],
    'wheel_size':        ['WHEEL_SIZE'],
    'weight_capacity':   ['Weight Capacity'],
    'wheel_size_uom':    ['WHEEL_SIZE_UOM'],
    'ibc':               ['IBC'],
    'skd_ckd':           ['SKD CKD','SKD_CKD'],
    'branded_tyre':      ['Branded Tyre Or Not'],
    'tyre_specs':        ['Tyre Specs'],
    'basket':            ['Basket'],
    'suspension_type':   ['Suspension Type (Front, Rear or Dual)','Suspension Type'],
    'fork_type':         ['Fork Type (Rigid Fork or For Suspension)','Fork Type'],
    'rear_suspension':   ['Rear Suspension (No Suspension, Shocker)','Rear Suspension'],
    'frame_type':        ['Frame Type (Folding, Rigid)','Frame Type'],
    'mode_of_operation': ['Mode Of operation (Manual, Battery)','Mode Of Operation'],
    'battery_wattage':   ['Battery Wattage Power'],
    'brake_lever_mat':   ['Brake Lever Material'],
    'num_spokes':        ['Number Of Spokes'],
    'chain_guard':       ['Chain Guard'],
    'seat_type':         ['Seat Type'],
    'water_bottle':      ['Water Bottle Holder'],
    'image':             ['ImageURL1'],
    'image2':            ['ImageURL2'],
    'image3':            ['ImageURL3'],
    'image4':            ['ImageURL4'],
    'image5':            ['ImageURL5'],
    'image6':            ['ImageURL6'],
    'per_pc_sp':         ['Per Pc SP'],
    'per_pc_mrp':        ['Per Pc MRP'],
    'set_sp':            ['Set SP'],
    'set_mrp':           ['Set MRP'],
    'moq':               ['MOQ'],
    'country':           ['Country Of Origin'],
    'weight':            ['Weight of Product in KG'],
    'product_color2':    ['Product Color'],
    'dims':              ['Product Dimension (LXBXH)'],
    'unit_measure':      ['*Unit Of Measure'],
    'unit_measure2':     ['*Unit Of Measure.1'],
    'mfg_year':          ['*Manufacturing Year'],
    'hsn':               ['*HSN Code'],
    'gst':               ['*GST'],
}

TS_BASE_COL_HINTS = {
    'article': ['Product Code','Article Number'],
    'sku':     ['Seller SKU ID','Child SKU'],
}



# ── Toys & Sports routes ────────────────────────────────────────
@app.route('/ts_categories')
def get_ts_categories():
    return jsonify({'categories': TS_CATEGORIES})


@app.route('/ts_config', methods=['GET'])
def ts_config_get_route():
    return jsonify(get_ts_config_from_disk())


@app.route('/ts_config', methods=['POST'])
def update_ts_config():
    cfg  = get_ts_config_from_disk()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(TS_CONFIG_PATH, cfg)
    write_log('anonymous', 'ts_config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})


@app.route('/detect_ts_categories', methods=['POST'])
def detect_ts_categories():
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
        col_map  = build_col_map(all_dump, TS_DUMP_COL_HINTS)
        pv_col   = col_map.get('pv')
        if pv_col and pv_col in all_dump.columns:
            found = [str(v).strip() for v in all_dump[pv_col].dropna().unique()
                     if str(v).strip() not in ('nan','None','')]
            matched = []
            for v in found:
                v_lower = v.lower().strip()
                # EXACT match only - prevents "Gear" from matching inside "Non Gear"
                for cat in TS_CATEGORIES:
                    if cat.lower() == v_lower:
                        if cat not in matched:
                            matched.append(cat)
                        break
            return jsonify({'categories': matched if matched else found, 'all_found': found})
        return jsonify({'categories': [], 'all_found': []})
    except Exception as e:
        return jsonify({'categories': [], 'error': str(e)})


@app.route('/process_ts', methods=['POST'])
def process_ts():
    """Toys & Sports processor. Produces 5 files per PV (Gear / Non Gear / Battery Operated)."""
    try:
        categories_raw = request.form.get('categories', '')
        try:    categories = json.loads(categories_raw)
        except: categories = [s.strip() for s in categories_raw.split(',') if s.strip()]

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not categories:
            return jsonify({'error': 'Please select at least one Product Vertical'}), 400
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

        col_map = build_col_map(all_dump, TS_DUMP_COL_HINTS)
        pv_col  = col_map.get('pv')

        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf  = bxl.parse(sname)
                    bcol = build_col_map(bdf, TS_BASE_COL_HINTS)
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
                if pv_col and pv_col in all_dump.columns:
                    col_vals = all_dump[pv_col].astype(str).str.strip().str.lower()
                    # EXACT match only - prevents "Non Gear" from matching "Gear"
                    mask = col_vals == category.lower()
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                wb_jpin, wb_tax, wb_pav, wb_scm, wb_l4, filled, skipped = fill_ts_files(
                    filtered, col_map, category, existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_cat = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
                files_written = []
                for wb_obj, label in [
                    (wb_jpin, 'JPIN'),
                    (wb_tax,  'TaxMaster'),
                    (wb_pav,  'ProductAttributeValue'),
                    (wb_scm,  'SupplyChainAttribute'),
                    (wb_l4,   'L4'),
                ]:
                    fname   = f'ts_{label}_{safe_cat}.xlsx'
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

                ws_jpin = wb_jpin.active
                jpin_headers = [ws_jpin.cell(1, c).value for c in range(1, ws_jpin.max_column + 1)]
                for r in range(2, min(filled + 2, 52)):
                    rdata = {}
                    for pc in preview_cols:
                        if pc in jpin_headers:
                            rdata[pc] = ws_jpin.cell(r, jpin_headers.index(pc)+1).value
                    if any(v for v in rdata.values()):
                        preview_rows.append({**rdata, '_category': category})

        out_name  = 'ts_filled_templates.zip'
        out_ext   = '.zip'
        zip_buf.seek(0)
        out_bytes = zip_buf.getvalue()

        file_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
        FILE_STORE[file_token] = {'bytes': out_bytes, 'filename': out_name,
                                   'ext': out_ext, 'created': time.time()}

        write_log('anonymous', 'ts_catalog_generated',
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


if __name__ == '__main__':
    app.run(debug=False, port=5050)
