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


# ═══════════════════════════════════════════════════════════════
# CONSUMER ELECTRONICS MODULE — UNIFIED TEMPLATE EDITION
# ═══════════════════════════════════════════════════════════════

CE_UNIFIED_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'CE_Unified_Template_v2.xlsx')
CE_LOGIC_PATH            = os.path.join(os.path.dirname(__file__), 'Unified_Template_Creation.xlsx')
CE_UNI_CONFIG_PATH       = '/tmp/fillforge_ce_uni_config.json'

# ── Which CE_Unified tab covers which PV names ─────────────────
CE_PV_TO_TAB = {
    'Mobile Adapters & Cables':   'Mobile Accessories',
    'Mobile Case & Covers':       'Mobile Accessories',
    'Screen Guards / Protectors': 'Mobile Accessories',
    'Mobile Cables':              'Mobile Accessories',
    'Speakers':                   'Audio Devices',
    'Headsets':                   'Audio Devices',
    'TWS Ear Buds':               'Audio Devices',
    'Earphones':                  'Audio Devices',
    'Neck Bands':                 'Audio Devices',
    'Feature Phones':             'Mobile Phones',
    'Smartphones':                'Mobile Phones',
    'Memory Cards':               'Other Accessories',
    'Mobile Holders':             'Other Accessories',
    'Power Bank':                 'Other Accessories',
    'Smart Watches':              'Other Accessories',
}


def _load_ce_uni_pv_list():
    pv_map = {}
    pv_list = []
    try:
        wb = load_workbook(CE_LOGIC_PATH, read_only=True, data_only=True)
        ws = wb['Supported Product Verticle']
        for r in range(2, ws.max_row + 1):
            name    = ws.cell(r, 1).value
            pv_id   = ws.cell(r, 2).value
            family  = ws.cell(r, 3).value
            if name and str(name).strip() not in ('', 'nan'):
                n = str(name).strip()
                pv_map[n] = {
                    'pv_id':  str(pv_id).strip() if pv_id else '',
                    'family': str(family).strip() if family else n,
                }
                pv_list.append(n)
        wb.close()
    except Exception as e:
        print(f'Warning: could not load CE Unified PV list: {e}')
    return pv_map, pv_list


def _load_ce_uni_template_headers():
    result = {}
    try:
        wb = load_workbook(CE_LOGIC_PATH, read_only=True, data_only=True)
        ws = wb['Templates']
        rows = list(ws.rows)
        i = 0
        while i < len(rows):
            row_vals = [c.value for c in rows[i]]
            if row_vals and row_vals[0] == 'Category *':
                hdr_row = [str(v).strip() if v else '' for v in row_vals]
                while hdr_row and not hdr_row[-1]:
                    hdr_row.pop()
                if i + 1 < len(rows):
                    data_vals = [c.value for c in rows[i + 1]]
                    pv_name = str(data_vals[3]).strip() if data_vals[3] else None
                    if pv_name:
                        static = {}
                        for ci, col in enumerate(hdr_row):
                            if col and ci < len(data_vals) and data_vals[ci] is not None:
                                v = str(data_vals[ci]).strip()
                                if v and v not in ('nan', 'None', 'NaN'):
                                    static[col] = v
                        result[pv_name] = {
                            'hdr_row': hdr_row,
                            'static':  static,
                        }
                i += 3
            else:
                i += 1
        wb.close()
    except Exception as e:
        print(f'Warning: could not load CE Unified template headers: {e}')
    return result


def _load_ce_uni_title_conventions():
    conventions = {}
    try:
        wb = load_workbook(CE_LOGIC_PATH, read_only=True, data_only=True)
        ws = wb['PV Level Title Conventions']
        i = 2
        rows = list(ws.rows)
        while i < len(rows):
            pv_name = rows[i - 1][0].value
            if pv_name and str(pv_name).strip() not in ('', 'nan', 'Product Verticle'):
                if i < len(rows) and str(rows[i][0].value or '').strip() == 'Title Convention':
                    title_conv    = str(rows[i][1].value or '').strip()
                    internal_conv = str(rows[i][2].value or '').strip()
                    conventions[str(pv_name).strip()] = {
                        'title_conv':    title_conv,
                        'internal_conv': internal_conv,
                    }
            i += 2
        wb.close()
    except Exception as e:
        print(f'Warning: could not load CE title conventions: {e}')
    return conventions


def _load_ce_uni_mapping():
    mapping = []
    try:
        wb = load_workbook(CE_LOGIC_PATH, read_only=True, data_only=True)
        ws = wb['ProductVerticle<>Mapping Column']
        for r in range(2, ws.max_row + 1):
            metric  = ws.cell(r, 1).value
            atype   = ws.cell(r, 2).value
            logic   = ws.cell(r, 3).value
            subtypes_raw = ws.cell(r, 4).value
            if not metric: continue
            subtypes = [s.strip() for s in str(subtypes_raw or '').split(',') if s.strip()]
            mapping.append({
                'metric':   str(metric).strip(),
                'attr_type': str(atype or '').strip(),
                'logic':    str(logic or '').strip(),
                'subtypes': subtypes,
            })
        wb.close()
    except Exception as e:
        print(f'Warning: could not load CE mapping: {e}')
    return mapping


# ── Build caches at module load time ───────────────────────────
try:
    CE_UNI_PV_MAP, CE_UNI_PV_LIST = _load_ce_uni_pv_list()
except Exception as e:
    print(f'Warning: CE UNI PV load failed: {e}')
    CE_UNI_PV_MAP, CE_UNI_PV_LIST = {}, []

try:
    CE_UNI_TEMPLATE_HDRS = _load_ce_uni_template_headers()
except Exception as e:
    print(f'Warning: CE UNI template headers load failed: {e}')
    CE_UNI_TEMPLATE_HDRS = {}

try:
    CE_UNI_TITLE_CONV = _load_ce_uni_title_conventions()
except Exception as e:
    print(f'Warning: CE UNI title conventions load failed: {e}')
    CE_UNI_TITLE_CONV = {}

try:
    CE_UNI_MAPPING = _load_ce_uni_mapping()
except Exception as e:
    print(f'Warning: CE UNI mapping load failed: {e}')
    CE_UNI_MAPPING = []


# ── Default config ─────────────────────────────────────────────
CE_UNI_DEFAULT_CONFIG = {
    "brands":             {},
    "biz_cat_id":         "BCAT-139438",
    "biz_cat_name":       "Consumer Electronics",
    "relationship":       "Parent",
    "catalog_status":     "ACTIVE",
    "status_remark":      "Ready to Launch",
    "tax_master_status":  "active",
    "gst_cgst":           50,
    "gst_sgst":           50,
    "gst_igst":           0,
    "country_of_origin":  "India",
    "product_condition":  "Fresh",
    "manufacturing_year": "2026",
    "discovery_cat":      "DISCAT-135528",
}


def get_ce_uni_config_from_disk():
    return _load_config(CE_UNI_CONFIG_PATH, CE_UNI_DEFAULT_CONFIG)


# ── Column hints for the UNIFIED INPUT template ────────────────
CE_UNI_INPUT_COL_HINTS = {
    'product_verticle':  ['Product Verticle'],
    'model_name':        ['Model Name'],
    'sku':               ['Seller SKU ID'],
    'mrp':               ['MRP'],
    'sp':                ['Selling Price'],
    'moq':               ['MOQ'],
    'brand':             ['Brand'],
    'image1':            ['imageURL1'],
    'image2':            ['imageURL2'],
    'image3':            ['imageURL3'],
    'image4':            ['imageURL4'],
    'image5':            ['imageURL5'],
    'image6':            ['imageURL6'],
    'product_desc':      ['Product Description'],
    'colour':            ['Colour'],
    'product_condition': ['Product Condition'],
    'packing_type':      ['Packaging Type'],
    'material':          ['Product Material'],
    'adapter_connector': ['Adapter Connector Type'],
    'country':           ['Country of Origin'],
    'num_ports':         ['Number of Ports'],
    'connector_type':    ['Connector Type'],
    'output_voltage':    ['Output Voltage'],
    'port_type':         ['Port type'],
    'dims':              ['Product Dimension (LXBXH)'],
    'dim_uom':           ['Unit of Measurement'],
    'num_connectors':    ['No of Connetors'],
    'product_name':      ['Product Name'],
    'weight':            ['Product Weight (KG)'],
    'hsn':               ['HSN'],
    'gst':               ['GST'],
    'compatible_model':  ['Compatible Brand + Model Name'],
    'case_cover_type':   ['Case & Cover Type'],
    'case_closure':      ['Case & Cover Closure'],
    'pattern':           ['Pattern'],
    'compatible_brand':  ['Compatible Brand'],
    'coverage':          ['Coverage'],
    'screen_guard_type': ['Screen Guard type'],
    'thickness':         ['Screen Card Thickness'],
    'bluetooth_ver':     ['Bluetooth Version'],
    'connector_type_audio': ['Connector type'],
    'speaker_type':      ['Speaker Type'],
    'wired_or_wireless': ['Wired or Unwired'],
    'mic_type':          ['Mic Type'],
    'battery_capacity':  ['Battery Capacity'],
    'display_size':      ['Display Size'],
    'ram':               ['RAM'],
    'storage':           ['Internal Storage'],
    'os':                ['Operating System'],
    'os_version':        ['Operating System Version'],
    'processor_core':    ['Processor Core'],
    'back_camera':       ['Back Camera'],
    'front_camera':      ['Front Camera'],
    'output_ports_no':   ['No of Output Ports'],
    'output_ports_type': ['Output Ports Type'],
    'card_type':         ['Card type'],
    'speed_class':       ['Speed Class'],
    'storage_capacity':  ['Storage Capacity'],
    'holder_type':       ['Holder Type'],
    'lock_mechanism':    ['Lock Mechanism'],
    'rotation_type':     ['Rotation type'],
    'display_shape':     ['Display Shape'],
}

CE_UNI_BASE_COL_HINTS = {
    'article': ['Model Name'],
    'sku':     ['Seller SKU ID'],
}


# ── Title builders per PV ──────────────────────────────────────

def _ce_uni_safe_join(*parts):
    return ' '.join(p.strip() for p in parts if p and str(p).strip())


def _ce_uni_build_title(pv_name, drow, col_map, brand, condition):
    def _get(key):
        c = col_map.get(key)
        return safe(drow.get(c, '')) if c else ''

    model       = _get('model_name')
    colour      = title_case_color(_get('colour'))
    display     = _get('display_size')
    ram         = _get('ram')
    storage     = _get('storage')
    back_cam    = _get('back_camera')
    battery     = _get('battery_capacity')
    storage_cap = _get('storage_capacity')
    card_type   = _get('card_type')
    speed_class = _get('speed_class')
    output_v    = _get('output_voltage')
    adapter_con = _get('adapter_connector')
    num_conn    = _get('num_connectors')
    conn_type   = _get('connector_type')
    compat_model= _get('compatible_model')
    case_type   = _get('case_cover_type')
    speaker_type= _get('speaker_type')
    rotation    = _get('rotation_type')
    sg_type     = _get('screen_guard_type')
    coverage    = _get('coverage')
    output_ports_no   = _get('output_ports_no')
    output_ports_type = _get('output_ports_type')
    num_ports   = _get('num_ports')
    port_type   = _get('port_type')

    pv_lower = pv_name.lower()

    def _display_str(d):
        if not d: return ''
        d = str(d).strip()
        if '"' not in d and "'" not in d:
            return f'{d}" Display'
        return f'{d} Display'

    def _cam_str(c):
        if not c: return ''
        c = str(c).strip()
        m = re.search(r'(\d+)', c)
        num = m.group(1) if m else c
        return f'{num}MP Camera'

    ram_clean = re.sub(r'\s+', '', str(ram)) if ram else ''
    sto_clean = re.sub(r'\s+', '', str(storage)) if storage else ''
    ram_rom   = f'{ram_clean} + {sto_clean}' if (ram_clean and sto_clean) else (ram_clean or sto_clean)

    if 'mobile adapters' in pv_lower:
        base  = _ce_uni_safe_join(brand, condition, output_v, adapter_con)
        title = f'{base}, {colour}' if colour else base
        int_base  = _ce_uni_safe_join(brand, condition, model, output_v, adapter_con)
        int_title = f'{int_base}, {colour}' if colour else int_base

    elif 'speakers' in pv_lower:
        base  = _ce_uni_safe_join(brand, model, condition, speaker_type)
        title = base
        int_title = f'{_ce_uni_safe_join(brand, model, condition, speaker_type)}, {colour}' if colour else base

    elif 'mobile case' in pv_lower:
        base  = _ce_uni_safe_join(brand, condition, compat_model, _get('material'), case_type)
        title = f'{base}, {colour}' if colour else base
        int_base  = _ce_uni_safe_join(brand, condition, model, compat_model, _get('material'), case_type)
        int_title = f'{int_base}, {colour}' if colour else int_base

    elif 'feature phone' in pv_lower:
        base  = _ce_uni_safe_join(brand, model, _display_str(display), pv_name)
        title = f'{base}, {colour} ({condition})' if colour else f'{base}, ({condition})'
        int_title = title

    elif pv_lower in ('headsets', 'tws ear buds', 'earphones', 'neck bands'):
        label = {'tws ear buds': 'Earbuds', 'earphones': 'Wired Earphones',
                 'neck bands': 'Neckband'}.get(pv_lower, 'Headsets')
        base  = _ce_uni_safe_join(brand, model, condition, label)
        title = base
        int_title = f'{base}, {colour}' if colour else base

    elif 'memory card' in pv_lower:
        base  = _ce_uni_safe_join(brand, storage_cap, card_type, pv_name)
        title = f'{base}, {speed_class}' if speed_class else base
        int_title = title

    elif 'mobile cable' in pv_lower:
        conn_label = _ce_uni_safe_join(num_conn, conn_type) if (num_conn or conn_type) else ''
        base  = _ce_uni_safe_join(brand, condition, conn_label, pv_name)
        title = f'{base}, {colour}' if colour else base
        int_base  = _ce_uni_safe_join(brand, condition, model, conn_label, pv_name)
        int_title = f'{int_base}, {colour}' if colour else int_base

    elif 'mobile holder' in pv_lower:
        base  = _ce_uni_safe_join(brand, condition, pv_name, rotation)
        title = f'{base}, {colour}' if colour else base
        int_base  = _ce_uni_safe_join(brand, condition, model, pv_name, rotation)
        int_title = f'{int_base}, {colour}' if colour else int_base

    elif 'screen guard' in pv_lower:
        coverage_for = f'{sg_type} {coverage} for {compat_model}' if compat_model else f'{sg_type} {coverage}'
        base  = _ce_uni_safe_join(brand, condition, coverage_for.strip())
        title = base
        int_base  = _ce_uni_safe_join(brand, condition, model, sg_type, coverage,
                                      f'for {compat_model}' if compat_model else '')
        int_title = int_base

    elif 'power bank' in pv_lower:
        ports_label = ''
        if output_ports_no or output_ports_type:
            ports_label = _ce_uni_safe_join(output_ports_no, f'Output ({output_ports_type})')
        base  = _ce_uni_safe_join(brand, condition, battery, pv_name, ports_label)
        title = f'{base}, {colour}' if colour else base
        int_base  = _ce_uni_safe_join(brand, condition, model, battery, pv_name, ports_label)
        int_title = f'{int_base}, {colour}' if colour else int_base

    elif 'smartphone' in pv_lower:
        cam_str = _cam_str(back_cam)
        base  = _ce_uni_safe_join(brand, model, cam_str, pv_name)
        suffix_parts = [ram_rom, colour, f'({condition})' if condition else '']
        suffix = ', '.join(p for p in suffix_parts if p)
        title     = f'{base}, {suffix}' if suffix else base
        int_title = title

    elif 'smart watch' in pv_lower:
        base  = _ce_uni_safe_join(brand, model, condition, _display_str(display), pv_name)
        title = f'{base}, {colour}' if colour else base
        int_title = title

    else:
        base  = _ce_uni_safe_join(brand, model, condition, pv_name)
        title = f'{base}, {colour}' if colour else base
        int_title = title

    return title, int_title


def _ce_uni_set_details(pv_name, drow, col_map):
    pv_lower = pv_name.lower()
    if 'mobile case' in pv_lower:
        c = col_map.get('compatible_model')
        return safe(drow.get(c, '')) if c else ''
    if 'smartphone' in pv_lower or 'feature phone' in pv_lower:
        ram_c  = col_map.get('ram')
        sto_c  = col_map.get('storage')
        ram    = safe(drow.get(ram_c, '')) if ram_c else ''
        sto    = safe(drow.get(sto_c, '')) if sto_c else ''
        ram_s  = re.sub(r'\s+', '', ram) if ram else ''
        sto_s  = re.sub(r'\s+', '', sto) if sto else ''
        return f'{ram_s}+{sto_s}' if (ram_s and sto_s) else (ram_s or sto_s or '')
    return '1pc'


def fill_ce_uni_template(ws, headers, rows_df, col_map, pv_name, pv_static,
                         existing_articles, existing_skus):
    tcol        = {h: i + 1 for i, h in enumerate(headers) if h}
    _cfg        = get_ce_uni_config_from_disk()
    brands_dict = normalize_brands(_cfg.get('brands', {}))
    fallback_brand, fallback_id = ('', '')
    if brands_dict:
        fallback_brand, fallback_id = next(iter(brands_dict.items()))

    skipped, filled = [], []

    for _, drow in rows_df.iterrows():
        brand, brand_id = get_brand_info(drow, col_map, brands_dict)
        if not brand and fallback_brand:
            brand    = fallback_brand
            brand_id = fallback_id

        model_name = safe(drow.get(col_map.get('model_name', ''), ''))
        sku_raw    = safe(drow.get(col_map.get('sku', ''), ''))
        article    = model_name if model_name else sku_raw

        if article.upper() in existing_articles or sku_raw.upper() in existing_skus:
            skipped.append({'sku': sku_raw, 'article': article,
                            'reason': 'Already exists in base data'})
            continue

        filled_count = len(filled) + 1
        row_idx      = filled_count + 1

        condition  = safe(drow.get(col_map.get('product_condition', ''), '')) or _cfg.get('product_condition', 'Fresh')
        colour     = title_case_color(safe(drow.get(col_map.get('colour', ''), '')))
        country    = safe(drow.get(col_map.get('country', ''), '')) or _cfg.get('country_of_origin', 'India')
        packing    = safe(drow.get(col_map.get('packing_type', ''), '')) or 'BOX'
        dim_raw    = safe(drow.get(col_map.get('dims', ''), ''))
        dim_uom    = safe(drow.get(col_map.get('dim_uom', ''), '')) or 'cm'
        weight_raw = safe(drow.get(col_map.get('weight', ''), ''))
        product_desc = safe(drow.get(col_map.get('product_desc', ''), ''))

        mrp_raw = drow.get(col_map.get('mrp', ''), '')
        sp_raw  = drow.get(col_map.get('sp', ''), '')
        moq_raw = drow.get(col_map.get('moq', ''), 1)
        hsn_raw = drow.get(col_map.get('hsn', ''), '')
        gst_raw = drow.get(col_map.get('gst', ''), 18)

        try:    mrp = float(mrp_raw) if str(mrp_raw).strip() not in ('', 'nan') else ''
        except: mrp = ''
        try:    sp  = float(sp_raw)  if str(sp_raw).strip()  not in ('', 'nan') else ''
        except: sp  = ''
        try:    moq = int(float(moq_raw))
        except: moq = 1
        try:    hsn = int(float(hsn_raw)) if str(hsn_raw).strip() not in ('', 'nan') else ''
        except: hsn = ''
        try:    gst = int(float(gst_raw))
        except: gst = 18

        weight_clean = ''
        if weight_raw:
            m = re.search(r'([0-9.]+)', str(weight_raw))
            if m: weight_clean = float(m.group(1))

        L, B, H = parse_lbh(dim_raw)

        def _img(key):
            c = col_map.get(key)
            return safe(drow.get(c, '')) if c else ''

        img1 = _img('image1'); img2 = _img('image2'); img3 = _img('image3')
        img4 = _img('image4'); img5 = _img('image5'); img6 = _img('image6')

        def _attr(key):
            c = col_map.get(key)
            return safe(drow.get(c, '')) if c else ''

        title, internal_title = _ce_uni_build_title(pv_name, drow, col_map, brand, condition)

        set_details  = _ce_uni_set_details(pv_name, drow, col_map)
        set_desc     = f'1pc of {pv_name}'

        ram_c  = col_map.get('ram');     ram   = safe(drow.get(ram_c, '')) if ram_c else ''
        sto_c  = col_map.get('storage'); sto   = safe(drow.get(sto_c, '')) if sto_c else ''
        ram_s  = re.sub(r'\s+', '', str(ram)) if ram else ''
        sto_s  = re.sub(r'\s+', '', str(sto)) if sto else ''
        ram_rom = f'{ram_s}+{sto_s}' if (ram_s and sto_s) else (ram_s or sto_s or '')

        row_data = {}

        for col_name, val in pv_static.items():
            row_data[col_name] = val

        row_data['SubType'] = ''

        row_data.update({
            'BusinessCategoryId *':   _cfg.get('biz_cat_id', 'BCAT-139438'),
            'BusinessCategoryName *': _cfg.get('biz_cat_name', 'Consumer Electronics'),
            'ProductCode *':          article,
            'Relationship *':         _cfg.get('relationship', 'Parent'),
            'ParentProductId *':      sku_raw,
            'ChildSKU *':             sku_raw,
            'MRP *':                  mrp,
            'SellingPrice *':         sp,
            'MOQ *':                  moq,
            'title *':                title,
            'internalTitle *':        internal_title,
            'brandId *':              brand_id,
            'brandName *':            brand,
            'imageURL1 *':            img1,
            'imageURL2':              img2,
            'imageURL3':              img3,
            'imageURL4':              img4,
            'imageURL5':              img5,
            'imageURL6':              img6,
            'videoURL1':              '',
            'videoURL2':              '',
            'sizeChartURLImage':      '',
            'catalogStatus *':        _cfg.get('catalog_status', 'ACTIVE'),
            'statusRemark':           _cfg.get('status_remark', 'Ready to Launch'),
            'discoveryCategoryIds':   pv_static.get('discoveryCategoryIds', _cfg.get('discovery_cat', 'DISCAT-135528')),
            'productDescription *':   product_desc,
            'PRODUCT_IDENTIFIER *':   'Set',
            'SET_NAME *':             'Set of 1',
            'SET_COUNT *':            1,
            'PACK_NAME *':            'Pack of 1',
            'PACK_OF *':              1,
            'IS_COMBO *':             'yes',
            'AVAILABLE_SIZES *':      '1',
            'SET_DETAILS *':          set_details,
            'SET_DESCRIPTION *':      set_desc,
            'PRODUCT_COLOR *':        colour,
            'ARTICLE_NUMBER *':       article,
            'MODEL_NAME *':           model_name,
            'PRODUCT_CONDITION *':    condition,
            'UNIT_OF_MEASUREMENT_SINGULAR *':             'Piece',
            'UNIT_OF_MEASUREMENT_PLURAL *':               'Pieces',
            'UNIT_OF_MEASUREMENT_SINGULAR_ABBREVIATION *': 'Pc',
            'UNIT_OF_MEASUREMENT_PLURAL_ABBREVIATION *':  'Pcs',
            'SELLER_SKU_ID *':        sku_raw,
            'PACKAGING_TYPE *':       packing,
            'DESCRIPTION':            '',
            'MATERIAL *':             _attr('material'),
            'ADAPTER_CONNECTOR_TYPE *': _attr('adapter_connector'),
            'CABLE_LENGTH_IN_METER *': '1m',
            'CABLE_MATERIAL *':       _attr('material'),
            'CABLE_TYPE':             '',
            'COUNTRY_OF_ORIGIN *':    country,
            'EAN':                    '',
            'IMPORTED_BY':            '',
            'KEY_FEATURES':           '',
            'MANUFACTURING_YEAR':     '',
            'NO_OF_ADAPTER_PORTS *':  _attr('num_ports'),
            'OUTPUT_CURRENT_OR_VOLTAGE *': _attr('output_voltage'),
            'PORT_TYPE *':            _attr('port_type'),
            'PRODUCT_BREADTH *':      B,
            'PRODUCT_DIMENSION_UOM *': dim_uom,
            'PRODUCT_HEIGHT *':       H,
            'PRODUCT_LENGTH *':       L,
            'PRODUCT_TYPE *':         _attr('product_name'),
            'PRODUCT_WEIGHT_IN_KG *': weight_clean,
            'PRODUCT_MANUFACTURING_CITY':  '',
            'PRODUCT_MANUFACTURING_STATE': '',
            'SUITABLE_FOR':           '',
            'WARRANTY':               '',
            'MANUFACTURER':           '',
            'hsnCode *':              hsn,
            'gstPercentage *':        gst,
            'cgstShare *':            _cfg.get('gst_cgst', 50),
            'sgstShare *':            _cfg.get('gst_sgst', 50),
            'igstShare *':            _cfg.get('gst_igst', 0),
            'cess':                   '',
            'sinTax':                 '',
            'vatPercentage':          '',
            'otherCess':              '',
            'validityPeriodStartDate': '',
            'validityPeriodEndDate':   '',
            'declarationForm':         '',
            'taxMasterStatus':         _cfg.get('tax_master_status', 'active'),
            # ── Audio ──
            'BATTERY_LIFE_FOR_WIRELESS':      '',
            'BLUETOOTH_VERSION_FOR_WIRELESS *': _attr('bluetooth_ver'),
            'CHANNELS':                        '',
            'CHARGING_TIME':                   '',
            'CONNECTOR_TYPE_FOR_WIRED *':      _attr('connector_type_audio'),
            'CONTROL_OPTIONS':                 '',
            'MIC_TYPE':                        '',
            'MOUNTING_OR_PLACEMENT_TYPE':      '',
            'PACKAGE_CONTENTS':                '',
            'PORT_TYPE':                       '',
            'SPEAKER_TYPE *':                  _attr('speaker_type'),
            'WATER_RESISTANCE':                '',
            'WIRED_OR_WIRELESS *':             _attr('wired_or_wireless'),
            # ── Case & Covers / Screen Guards ──
            'COMPATIBLE_BRAND_MODEL *':        _attr('compatible_model'),
            'CASE_COVER_TYPE *':               _attr('case_cover_type'),
            'CLOSURE_TYPE *':                  _attr('case_closure'),
            'DESIGN *':                        _attr('pattern'),
            'THEME':                           '',
            'COMPATIBLE_BRAND *':              _attr('compatible_brand'),
            'COVERAGE *':                      _attr('coverage'),
            'EDGE_TYPE':                       '',
            'SCREEN_GUARD_OR_PROTECTOR_TYPE *': _attr('screen_guard_type'),
            'THICKNESS *':                     _attr('thickness'),
            # ── Mobile Holders ──
            'HOLDER_TYPE *':                   _attr('holder_type'),
            'LOCK_MECHANISM *':                _attr('lock_mechanism'),
            'ROTATION_OR_ADJUSTABILITY *':     _attr('rotation_type'),
            # ── Phones ──
            'BATTERY_CAPACITY_MAH *':          _attr('battery_capacity'),
            'CHARGING_TYPE_SUPPORTED':         '',
            'OPERATING_SYSTEM_OS':             _attr('os'),
            'OS_VERSION':                      _attr('os_version'),
            'DISPLAY_SIZE *':                  _attr('display_size'),
            'DISPLAY_TYPE':                    '',
            'DISPLAY_RESOLUTION':              '',
            'RAM *':                           ram,
            'INTERNAL_STORAGE *':              sto,
            'EXPANDABLE_STORAGE':              '',
            'SIM_TYPE':                        '',
            'BLUETOOTH_VERSION':               '',
            'EXPANDABLE_STORAGE_TYPE':         '',
            'EXPANDABLE_STORAGE_CAPACITY_MAX': '',
            'BATTERY_TYPE':                    '',
            'REMOVABLE_BATTERY':               '',
            'HYBRID_SIM_SLOT':                 '',
            'NETWORK_TYPE_SUPPORTED':          '',
            'AUDIO_JACK':                      '',
            'FM_RADIO':                        '',
            'TORCH_OR_FLASHLIGHT':             '',
            'RAM_ROM *':                       ram_rom,
            # ── Headsets ──
            'ACTIVE_NOISE_CANCELLATION_ANC':   '',
            'ADJUSTABLE_OR_FOLDABLE':          '',
            'CABLE_LENGTH_IN_METER':           '',
            'MIC_TYPE *':                      _attr('mic_type'),
            'WEARING_STYLE':                   '',
            # ── Smartphones extra ──
            'PACKAGING_TYPE':                  packing,
            'OPERATING_SYSTEM_OS *':           _attr('os'),
            'OS_VERSION *':                    _attr('os_version'),
            'PROCESSOR_BRAND_AND_MODEL_NAME':  '',
            'NUMBER_OF_PROCESSOR_CORES *':     _attr('processor_core'),
            'PRIMARY_CAMERA_RESOLUTION *':     _attr('back_camera'),
            'FRONT_CAMERA_RESOLUTION *':       _attr('front_camera'),
            'REAR_FLASH':                      '',
            'SIM_SIZE':                        '',
            'WIFI':                            '',
            'FINGERPRINT_SENSOR':              '',
            'CLOCK_SPEED':                     '',
            'REFRESH_RATE':                    '',
            'TOUCHSCREEN_TYPE':                '',
            'PRIMARY_CAMERA_SETUP':            '',
            'FRONT_FLASH':                     '',
            'VIDEO_RECORDING_RESOLUTION':      '',
            'FAST_CHARGING_WATTAGE':           '',
            'WIRELESS_CHARGING_SUPPORT':       '',
            'GPS_SUPPORT':                     '',
            'NFC_SUPPORT':                     '',
            'INFRARED_IR_BLASTER':             '',
            'FINGERPRINT_SENSOR_POSITION':     '',
            'FACE_UNLOCK':                     '',
            'WATER_RESISTANCE_RATING':         '',
            # ── Memory Cards ──
            'MEMORY_CARD_TYPE *':              _attr('card_type'),
            'SPEED_CLASS *':                   _attr('speed_class'),
            'STORAGE_CAPACITY *':              _attr('storage_capacity'),
            'COMPATIBLE_BRAND':               '',
            'ADAPTER_INCLUDED':               '',
            'CONNECTION_INTERFACE':           '',
            # ── Power Bank ──
            'PACKAGING_CLASSIFICATION':        '',
            'SUB_BRAND':                       '',
            'COLOR *':                         colour,
            'FOOD_NON-FOOD':                   '',
            'NO_OF_OUTPUT_PORTS *':            _attr('output_ports_no'),
            'OUTPUT_PORTS_TYPE *':             _attr('output_ports_type'),
            # ── Smart Watches ──
            'STRAP_COLOR *':                   colour,
            'DISPLAY_SHAPE *':                 _attr('display_shape'),
            'CORE_BRAND *':                    '',
        })

        for col_name, val in row_data.items():
            if col_name in tcol and val is not None and str(val) not in ('None', ''):
                ws.cell(row=row_idx, column=tcol[col_name]).value = val

        if 'SubType' in tcol:
            ws.cell(row=row_idx, column=tcol['SubType']).value = None

        filled.append({'sku': sku_raw, 'article': article})

    return len(filled), skipped


def get_ce_uni_template_for_pv(pv_name):
    tab = CE_PV_TO_TAB.get(pv_name)
    if not tab:
        pv_info = CE_UNI_PV_MAP.get(pv_name, {})
        family  = pv_info.get('family', pv_name)
        tab = CE_PV_TO_TAB.get(family, 'Mobile Accessories')

    try:
        wb_src  = load_workbook(CE_UNIFIED_TEMPLATE_PATH)
        ws_src  = wb_src[tab]
        headers = [ws_src.cell(3, c).value for c in range(1, ws_src.max_column + 1)]
        while headers and headers[-1] is None:
            headers.pop()
    except Exception as e:
        print(f'Warning: Could not load CE Unified template tab {tab}: {e}')
        headers = []

    tpl_info = CE_UNI_TEMPLATE_HDRS.get(pv_name, {})
    out_headers = tpl_info.get('hdr_row', []) if tpl_info else []

    if not out_headers:
        for k, v in CE_UNI_TEMPLATE_HDRS.items():
            out_headers = v.get('hdr_row', [])
            if out_headers:
                break

    wb_new       = Workbook()
    ws_new       = wb_new.active
    ws_new.title = 'CE - PV Template'
    for ci, h in enumerate(out_headers, 1):
        ws_new.cell(1, ci).value = h

    return wb_new, out_headers


# ── In-Memory File Storage ─────────────────────────────────────
FILE_STORE = {}


# ═══════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════

# ── Global error handlers ─────────────────────────────────────
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

@app.route('/logs')
def get_logs():
    return jsonify({'logs': read_logs(500)})

@app.route('/config', methods=['GET'])
def config_get_route():
    return jsonify(get_config())

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

# ── CE Unified Config Routes ───────────────────────────────────
@app.route('/ce_uni_config', methods=['GET'])
def ce_uni_config_get():
    return jsonify(get_ce_uni_config_from_disk())

@app.route('/ce_uni_config', methods=['POST'])
def ce_uni_config_post():
    cfg  = get_ce_uni_config_from_disk()
    data = request.json
    if 'brands' in data:
        data['brands'] = normalize_brands(data['brands'])
    cfg.update(data)
    _save_config(CE_UNI_CONFIG_PATH, cfg)
    write_log('anonymous', 'ce_uni_config_updated', f"brands={cfg.get('brands')}")
    return jsonify({'status': 'ok'})

@app.route('/ce_uni_pvlist')
def get_ce_uni_pvlist():
    return jsonify({'pvs': CE_UNI_PV_LIST, 'pv_map': CE_UNI_PV_MAP})

@app.route('/detect_ce_uni_pvs', methods=['POST'])
def detect_ce_uni_pvs():
    """Auto-detect PVs from Column A ('Product Verticle') of input file."""
    try:
        dump_file = request.files.get('dump')
        if not dump_file:
            return jsonify({'pvs': []})

        xl     = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        col_map  = build_col_map(all_dump, CE_UNI_INPUT_COL_HINTS)
        pv_col   = col_map.get('product_verticle')

        if pv_col and pv_col in all_dump.columns:
            found   = [str(v).strip() for v in all_dump[pv_col].dropna().unique()
                       if str(v).strip() not in ('nan', 'None', '')]
            matched = [v for v in found if v in CE_UNI_PV_MAP]
            return jsonify({'pvs': matched, 'all_found': found})
        return jsonify({'pvs': [], 'all_found': []})
    except Exception as e:
        return jsonify({'pvs': [], 'error': str(e)})

@app.route('/process_ce_uni', methods=['POST'])
def process_ce_uni():
    """
    Consumer Electronics Unified Template processor.
    Accepts: pvs (JSON list), dump file, optional base_data file.
    Returns: ZIP of filled xlsx files, one per PV.
    """
    try:
        pvs_raw = request.form.get('pvs', '')
        try:    pvs = json.loads(pvs_raw)
        except: pvs = [s.strip() for s in pvs_raw.split(',') if s.strip()]

        base_file = request.files.get('base_data')
        dump_file = request.files.get('dump')

        if not pvs:
            return jsonify({'error': 'Please select at least one Product Verticle'}), 400
        if not dump_file:
            return jsonify({'error': 'Input file is required'}), 400

        dump_bytes = dump_file.read()
        xl         = pd.ExcelFile(io.BytesIO(dump_bytes))
        frames     = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

        if all_dump.empty:
            return jsonify({'error': 'Could not read any data from input file'}), 400

        col_map = build_col_map(all_dump, CE_UNI_INPUT_COL_HINTS)
        pv_col  = col_map.get('product_verticle')

        existing_articles, existing_skus = set(), set()
        if base_file:
            bxl = pd.ExcelFile(io.BytesIO(base_file.read()))
            for sname in bxl.sheet_names:
                try:
                    bdf  = bxl.parse(sname)
                    bcol = build_col_map(bdf, CE_UNI_BASE_COL_HINTS)
                    if 'article' in bcol:
                        existing_articles |= set(bdf[bcol['article']].dropna().astype(str).str.strip().str.upper())
                    if 'sku' in bcol:
                        existing_skus |= set(bdf[bcol['sku']].dropna().astype(str).str.strip().str.upper())
                except: pass

        results, all_skipped, grand_filled = [], [], 0
        preview_rows, preview_cols = [], []

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for pv_name in pvs:
                if pv_col and pv_col in all_dump.columns:
                    mask     = all_dump[pv_col].astype(str).str.strip().str.lower() == pv_name.lower()
                    filtered = all_dump[mask].copy()
                    if filtered.empty:
                        mask2    = all_dump[pv_col].astype(str).str.lower().str.contains(
                                        re.escape(pv_name.lower()), na=False)
                        filtered = all_dump[mask2].copy()
                    if filtered.empty:
                        filtered = all_dump.copy()
                else:
                    filtered = all_dump.copy()

                tpl_info  = CE_UNI_TEMPLATE_HDRS.get(pv_name, {})
                pv_static = tpl_info.get('static', {})
                wb, headers = get_ce_uni_template_for_pv(pv_name)
                ws = wb.active

                filled, skipped = fill_ce_uni_template(
                    ws, headers, filtered, col_map, pv_name, pv_static,
                    existing_articles, existing_skus
                )
                all_skipped.extend(skipped)
                grand_filled += filled

                safe_pv = re.sub(r'[^\w\s-]', '', pv_name).replace(' ', '_')
                fname   = f'ce_{safe_pv}.xlsx'
                xls_buf = io.BytesIO()
                wb.save(xls_buf)
                zout.writestr(fname, xls_buf.getvalue())
                results.append({'pv': pv_name, 'filled': filled,
                                 'skipped': len(skipped), 'filename': fname})

                if not preview_cols:
                    pcols = ['title *', 'ChildSKU *', 'ARTICLE_NUMBER *', 'MRP *',
                             'SellingPrice *', 'PRODUCT_COLOR *', 'PRODUCT_CONDITION *',
                             'hsnCode *']
                    preview_cols = [c for c in pcols if c in headers]
                for r in range(2, min(filled + 2, 52)):
                    rdata = {}
                    for c in preview_cols:
                        if c in headers:
                            rdata[c] = ws.cell(r, headers.index(c) + 1).value
                    if any(v for v in rdata.values()):
                        preview_rows.append({**rdata, '_pv': pv_name})

        zip_buf.seek(0)
        if len(pvs) == 1:
            safe_pv  = re.sub(r'[^\w\s-]', '', pvs[0]).replace(' ', '_')
            out_name = f'ce_{safe_pv}.xlsx'
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
        write_log('anonymous', 'ce_uni_catalog_generated',
                  f'pvs={pvs} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status':          'ok',
            'grand_filled':    grand_filled,
            'grand_skipped':   len(all_skipped),
            'results':         results,
            'skipped_details': all_skipped[:50],
            'preview':         preview_rows,
            'preview_cols':    preview_cols,
            'download_token':  file_token,
            'filename':        out_name,
            'is_zip':          len(pvs) > 1,
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


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


@app.route('/process', methods=['POST'])
def process():
    """Footwear catalog processor."""
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


def _generate_blank_template(subtype):
    """Generate a blank Excel template containing only the header row for a specific subtype."""
    if subtype not in SUBTYPE_HEADER_ROW:
        return None, f'SubType "{subtype}" not found in template'

    wb_src = load_workbook(TEMPLATE_PATH)
    ws_src = wb_src['PV Template']
    hdr_row = SUBTYPE_HEADER_ROW.get(subtype, 1)

    headers = []
    for c in range(1, ws_src.max_column + 1):
        val = ws_src.cell(hdr_row, c).value
        if val is not None:
            headers.append(str(val).strip())

    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.title = 'PV Template'

    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h

    for ci, h in enumerate(headers, 1):
        ws_new.column_dimensions[ws_new.cell(1, ci).column_letter].width = max(len(h) + 2, 15)

    buf = io.BytesIO()
    wb_new.save(buf)
    buf.seek(0)
    return buf, None


@app.route('/download_template/<path:category>')
def download_template(category):
    category_lower = category.lower().strip()
    subtype = request.args.get('subtype', '').strip()

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

    if 'electronic' in category_lower or category_lower == 'ce':
        path  = CE_UNIFIED_TEMPLATE_PATH
        fname = 'CE_Unified_Template_v2.xlsx'
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
    cfg        = get_config()
    ce_uni_cfg = get_ce_uni_config_from_disk()
    return jsonify({
        'config_path':              CONFIG_PATH,
        'ce_uni_config_path':       CE_UNI_CONFIG_PATH,
        'config_file_exists':       os.path.exists(CONFIG_PATH),
        'ce_uni_config_file_exists': os.path.exists(CE_UNI_CONFIG_PATH),
        'footwear_brands':          cfg.get('brands', {}),
        'ce_uni_brands':            ce_uni_cfg.get('brands', {}),
        'footwear_config':          cfg,
        'ce_uni_config':            ce_uni_cfg,
    })


if __name__ == '__main__':
    app.run(debug=False, port=5050)
