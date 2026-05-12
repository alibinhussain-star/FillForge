from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd, re, io, os, json, random, string, time, zipfile
from datetime import datetime
from openpyxl import load_workbook

app = Flask(__name__, template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# In-memory store: token -> (bytes, filename, mimetype)
FILE_STORE = {}

# Logs (JSON-line file)
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

# ── Template constants ──────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Logic___Template_File.xlsx')

def _build_header_row_map():
    wb = load_workbook(TEMPLATE_PATH); ws = wb['PV Template']
    hdr_map, static_map = {}, {}
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
                            if col in ('Category *','SubCategory *','CategoryType *','SubType','PVID *','discoveryCategoryIds'):
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
        pv = pd.read_excel(TEMPLATE_PATH, sheet_name='PV List')
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
    from openpyxl import Workbook
    try:
        wb_src = load_workbook(TEMPLATE_PATH)
        ws_src = wb_src['PV Template']
        hdr_row = SUBTYPE_HEADER_ROW.get(subtype, 1)
        headers = [ws_src.cell(hdr_row, c).value for c in range(1, ws_src.max_column + 1)]
        while headers and headers[-1] is None: headers.pop()
    except Exception as e:
        print(f"Warning: Could not load template for {subtype}: {e}")
        headers = []
    wb_new = Workbook()
    ws_new = wb_new.active
    ws_new.title = 'PV Template'
    for ci, h in enumerate(headers, 1):
        ws_new.cell(1, ci).value = h
    return wb_new, headers

# Default config
DEFAULT_CONFIG = {
    "brand_name":"", "brand_id":"", "biz_cat_id":"BCAT-139461", "biz_cat_name":"Footwear",
    "relationship":"Parent", "catalog_status":"ACTIVE", "status_remark":"Ready to Launch",
    "tax_master_status":"active", "gst_cgst":50, "gst_sgst":50, "gst_igst":0,
    "country_of_origin":"India", "product_condition":"Fresh", "manufacturing_year":"2026",
    "discovery_cat":"DISCAT-135542",
}
config = {k: v for k, v in DEFAULT_CONFIG.items()}

# ─── Helper functions (stubs — replace with your originals) ───

DUMP_COL_HINTS = {
    'title': ['title', 'product title', 'product_name', 'name'],
    'sku': ['sku', 'childsku', 'sku id', 'product sku'],
    'article': ['article', 'article number', 'article_no', 'art no'],
    'mrp': ['mrp', 'max retail price', 'marked price'],
    'sp': ['selling price', 'sale price', 'price', 'sp'],
    'color': ['color', 'colour', 'product color'],
    'size': ['size', 'sizes', 'available sizes'],
    'set': ['set', 'set details', 'set_count'],
    'vertical': ['vertical', 'subtype', 'product vertical', 'category'],
    'hsn': ['hsn', 'hsn code', 'hsncode'],
    'upper': ['upper material', 'upper'],
    'sole': ['sole material', 'sole'],
}

BASE_COL_HINTS = {
    'sku': ['sku', 'childsku', 'sku id'],
    'article': ['article', 'article number'],
}

def safe(val):
    if pd.isna(val) or val is None:
        return ''
    return str(val).strip()

def detect_col(df, hints):
    cols = [c.strip() for c in df.columns]
    for hint in hints:
        hint_lower = hint.lower().strip()
        for col in cols:
            if hint_lower in col.lower():
                return col
    return None

def build_col_map(df, hints_dict):
    return {key: detect_col(df, hints) for key, hints in hints_dict.items()}

def title_case_color(color_str):
    if not color_str:
        return ''
    parts = re.split(r'[/&+,]', color_str)
    return '/'.join(p.strip().title() for p in parts if p.strip())

def merge_colors(color_series):
    colors = [safe(c) for c in color_series if safe(c)]
    if not colors:
        return ''
    unique = []
    for c in colors:
        for part in re.split(r'[/&+,]', c):
            p = part.strip().title()
            if p and p not in unique:
                unique.append(p)
    return '/'.join(unique)

def extract_article(title_str):
    if not title_str:
        return ''
    m = re.search(r'\b(Art[.\s]*\d+|ART-\d+|\d{4,})\b', str(title_str), re.I)
    return m.group(1) if m else ''

def expand_size_range(size_str):
    if not size_str:
        return []
    sizes = []
    for part in re.split(r'[,;/]', str(size_str)):
        part = part.strip()
        if not part:
            continue
        m = re.match(r'(\d+(?:\.5)?)\s*-\s*(\d+(?:\.5)?)', part)
        if m:
            start, end = float(m.group(1)), float(m.group(2))
            while start <= end:
                sizes.append(str(int(start)) if start == int(start) else str(start))
                start += 1
        else:
            sizes.append(part)
    return sizes

def build_set_details(set_str, count_str):
    s = safe(set_str)
    c = safe(count_str)
    if s and c:
        return f"{s} ({c} pcs)"
    return s or c or ''

def parse_lbh(dimension_str):
    if not dimension_str:
        return '', '', ''
    m = re.findall(r'(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)', str(dimension_str))
    if m:
        return m[0][0], m[0][1], m[0][2]
    return '', '', ''

def derive_gender(title_str, category_str=''):
    text = f"{title_str} {category_str}".lower()
    if any(w in text for w in ['women', 'woman', 'ladies', 'female', 'girl']):
        return 'Women'
    if any(w in text for w in ['men', 'man', 'male', 'boy', 'gents', 'gent']):
        return 'Men'
    if any(w in text for w in ['kid', 'kids', 'child', 'children', 'baby', 'infant']):
        return 'Kids'
    return 'Unisex'

def make_title(row_dict, subtype):
    parts = []
    brand = safe(row_dict.get('brand', ''))
    if brand:
        parts.append(brand.title())
    gender = safe(row_dict.get('gender', derive_gender(row_dict.get('title', ''))))
    if gender:
        parts.append(gender)
    parts.append(subtype)
    color = safe(row_dict.get('color', ''))
    if color:
        parts.append(title_case_color(color))
    article = safe(row_dict.get('article', extract_article(row_dict.get('title', ''))))
    if article:
        parts.append(f"Art {article}")
    return ' '.join(parts)

def make_internal_title(title):
    return safe(title)[:100]

def make_description(row_dict):
    parts = []
    title = safe(row_dict.get('title', ''))
    if title:
        parts.append(title)
    color = safe(row_dict.get('color', ''))
    if color:
        parts.append(f"Color: {title_case_color(color)}")
    upper = safe(row_dict.get('upper_material', ''))
    if upper:
        parts.append(f"Upper Material: {upper}")
    sole = safe(row_dict.get('sole_material', ''))
    if sole:
        parts.append(f"Sole Material: {sole}")
    return ' | '.join(parts)

def fill_template(dump_df, base_df, headers, wb):
    filled_rows = []
    skipped_rows = []
    base_skus = set()
    if base_df is not None and not base_df.empty:
        base_map = build_col_map(base_df, BASE_COL_HINTS)
        sku_col = base_map.get('sku')
        if sku_col and sku_col in base_df.columns:
            base_skus = set(str(v).strip().lower() for v in base_df[sku_col].dropna())
    col_map = build_col_map(dump_df, DUMP_COL_HINTS)
    ws = wb.active
    row_idx = 2
    for _, row in dump_df.iterrows():
        row_dict = {k: row.get(v) if v and v in dump_df.columns else '' for k, v in col_map.items()}
        sku = safe(row_dict.get('sku', ''))
        article = safe(row_dict.get('article', ''))
        if sku and sku.lower() in base_skus:
            skipped_rows.append({'sku': sku, 'article': article, 'reason': 'Duplicate SKU in base data'})
            continue
        filled = {}
        for h in headers:
            h_lower = str(h).lower().strip() if h else ''
            if h == 'title *':
                filled[h] = make_title(row_dict, 'Footwear')
            elif h == 'ChildSKU *':
                filled[h] = sku or f"SKU-{article or row_idx}"
            elif h == 'ARTICLE_NUMBER *':
                filled[h] = article or extract_article(row_dict.get('title', ''))
            elif h == 'MRP *':
                filled[h] = safe(row_dict.get('mrp', ''))
            elif h == 'SellingPrice *':
                filled[h] = safe(row_dict.get('sp', ''))
            elif h == 'PRODUCT_COLOR *':
                filled[h] = title_case_color(row_dict.get('color', ''))
            elif h == 'AVAILABLE_SIZES *':
                sizes = expand_size_range(row_dict.get('size', ''))
                filled[h] = ', '.join(sizes)
            elif h == 'SET_DETAILS *':
                filled[h] = build_set_details(row_dict.get('set', ''), row_dict.get('set_count', ''))
            elif h == 'UPPER_MATERIAL *':
                filled[h] = safe(row_dict.get('upper', ''))
            elif h == 'SOLE_MATERIAL *':
                filled[h] = safe(row_dict.get('sole', ''))
            elif h == 'hsnCode *':
                filled[h] = safe(row_dict.get('hsn', ''))
            elif h == 'SET_COUNT *':
                filled[h] = safe(row_dict.get('set_count', ''))
            elif h == 'description':
                filled[h] = make_description(row_dict)
            elif h == 'internal_title':
                filled[h] = make_internal_title(filled.get('title *', ''))
            elif h == 'gender':
                filled[h] = derive_gender(row_dict.get('title', ''), '')
            elif h == 'length':
                l, _, _ = parse_lbh(row_dict.get('dimensions', ''))
                filled[h] = l
            elif h == 'breadth':
                _, b, _ = parse_lbh(row_dict.get('dimensions', ''))
                filled[h] = b
            elif h == 'height':
                _, _, h_val = parse_lbh(row_dict.get('dimensions', ''))
                filled[h] = h_val
            else:
                filled[h] = ''
        for ci, h in enumerate(headers, 1):
            ws.cell(row_idx, ci).value = filled.get(h, '')
        filled_rows.append(filled)
        row_idx += 1
    return filled_rows, skipped_rows

# ── Routes ─────────────────────────────────────────────────────

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'pv_list_count': len(PV_LIST)})

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception:
        pass
    try:
        html_path = os.path.join(os.path.dirname(__file__), 'index.html')
        if os.path.exists(html_path):
            with open(html_path, 'r', encoding='utf-8') as f:
                return f.read()
    except Exception:
        pass
    return "<h1>Error</h1><p>Could not load index.html. Make sure it's in the templates/ folder or same directory as app.py.</p>", 500

@app.route('/subtypes')
def get_subtypes():
    return jsonify({'subtypes': PV_LIST})

@app.route('/config', methods=['GET'])
def get_config(): return jsonify(config)

@app.route('/config', methods=['POST'])
def update_config():
    global config
    config.update(request.json)
    write_log('anonymous', 'config_updated')
    return jsonify({'status': 'ok'})

@app.route('/logs')
def get_logs():
    return jsonify({'logs': read_logs(500)})

@app.route('/download_template/<vertical>')
def download_template(vertical):
    if vertical.lower() != 'footwear':
        return jsonify({'error': 'This vertical is under development'}), 403
    write_log('anonymous', 'template_downloaded', vertical)
    return send_file(TEMPLATE_PATH, as_attachment=True,
                     download_name='Footwear_Template.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/detect_verticals', methods=['POST'])
def detect_verticals():
    try:
        dump_file = request.files.get('dump')
        if not dump_file: return jsonify({'verticals': []})
        xl = pd.ExcelFile(io.BytesIO(dump_file.read()))
        frames = []
        for sname in xl.sheet_names:
            try: frames.append(xl.parse(sname))
            except: pass
        all_dump = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        col_map = build_col_map(all_dump, DUMP_COL_HINTS)
        vert_col = col_map.get('vertical')
        if vert_col and vert_col in all_dump.columns:
            found = [str(v).strip() for v in all_dump[vert_col].dropna().unique()
                     if str(v).strip() not in ('nan','None','')]
            matched = [v for v in found if v in SUBTYPE_MAP]
            return jsonify({'verticals': matched, 'all_found': found})
        return jsonify({'verticals': [], 'all_found': []})
    except Exception as e:
        return jsonify({'verticals': [], 'error': str(e)})

@app.route('/process', methods=['POST'])
def process():
    try:
        subtypes = json.loads(request.form.get('subtypes', '[]'))
        dump_file = request.files.get('dump')
        base_file = request.files.get('base_data')

        if not dump_file or not subtypes:
            return jsonify({'error': 'Missing dump file or subtypes'}), 400

        dump_df = pd.read_excel(io.BytesIO(dump_file.read()))
        base_df = None
        if base_file:
            base_df = pd.read_excel(io.BytesIO(base_file.read()))

        results = []
        all_skipped = []
        grand_filled = 0
        grand_skipped = 0
        preview = []
        preview_cols = []

        for subtype in subtypes:
            if subtype not in SUBTYPE_MAP:
                continue

            col_map = build_col_map(dump_df, DUMP_COL_HINTS)
            subtype_col = col_map.get('vertical') or col_map.get('subtype')
            if subtype_col and subtype_col in dump_df.columns:
                subtype_df = dump_df[dump_df[subtype_col].astype(str).str.strip() == subtype].copy()
            else:
                subtype_df = dump_df.copy()

            wb, headers = get_template_wb_for_subtype(subtype)
            filled_rows, skipped_rows = fill_template(subtype_df, base_df, headers, wb)

            # Save to BytesIO instead of temp file
            token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            file_bytes = output.getvalue()

            filled_count = len(filled_rows)
            skipped_count = len(skipped_rows)
            grand_filled += filled_count
            grand_skipped += skipped_count
            all_skipped.extend(skipped_rows)

            results.append({
                'subtype': subtype,
                'filename': f'{subtype.replace(" ", "_")}_Catalog.xlsx',
                'filled': filled_count,
                'skipped': skipped_count,
                'token': token
            })

            # Store in memory
            FILE_STORE[token] = (file_bytes, f'{subtype.replace(" ", "_")}_Catalog.xlsx',
                                 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            if not preview and filled_rows:
                preview_cols = [c for c in headers if c in ('title *', 'ChildSKU *', 'ARTICLE_NUMBER *', 'MRP *', 'SellingPrice *', 'PRODUCT_COLOR *', 'AVAILABLE_SIZES *', 'SET_DETAILS *', 'UPPER_MATERIAL *', 'SOLE_MATERIAL *', 'hsnCode *', 'SET_COUNT *')]
                for row in filled_rows[:50]:
                    preview.append({col: row.get(col, '') for col in preview_cols})

        is_zip = len(results) > 1
        if is_zip:
            zip_output = io.BytesIO()
            with zipfile.ZipFile(zip_output, 'w') as zf:
                for r in results:
                    file_bytes, fname, _ = FILE_STORE.get(r['token'], (b'', '', ''))
                    if file_bytes:
                        zf.writestr(fname, file_bytes)
            zip_output.seek(0)
            zip_bytes = zip_output.getvalue()
            zip_token = ''.join(random.choices(string.ascii_letters + string.digits, k=32))
            FILE_STORE[zip_token] = (zip_bytes, 'Catalog_Files.zip', 'application/zip')
            download_token = zip_token
            download_filename = 'Catalog_Files.zip'
        else:
            download_token = results[0]['token'] if results else ''
            download_filename = results[0]['filename'] if results else 'filled_template.xlsx'

        write_log('anonymous', 'catalog_generated',
                  f'subtypes={subtypes} filled={grand_filled} skipped={len(all_skipped)}')

        return jsonify({
            'status': 'ok',
            'results': results,
            'grand_filled': grand_filled,
            'grand_skipped': grand_skipped,
            'download_token': download_token,
            'filename': download_filename,
            'is_zip': is_zip,
            'preview': preview,
            'preview_cols': preview_cols,
            'skipped_details': all_skipped[:100]
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download/<token>')
def download(token):
    if '..' in token or '/' in token or '\\' in token:
        return 'Invalid', 400

    file_data = FILE_STORE.get(token)
    if not file_data:
        return 'File not found', 404

    file_bytes, default_fname, mtype = file_data
    fname = request.args.get('filename', default_fname)

    write_log('anonymous', 'file_downloaded', fname)
    return send_file(
        io.BytesIO(file_bytes),
        as_attachment=True,
        download_name=fname,
        mimetype=mtype
    )

if __name__ == '__main__':
    app.run(debug=False, port=5050)
