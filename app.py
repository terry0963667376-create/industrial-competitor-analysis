"""Universal industrial datasheet MVP v0.5.1: two-PDF comparison, no paid API/AI."""
import hashlib
import io
import re

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title='工業規格書｜競品比較 MVP v0.5.1', page_icon='📊', layout='wide')
st.title('工業規格書標準化與競品比較｜MVP v0.5.1')
st.caption('雙 PDF 動態擷取 → 人工修正 → 規格配對 → Excel 比較報告｜免費規則式版本，尚未使用 AI')
st.info('僅支援有文字層的 PDF；比較結果為候選資料，須人工確認。請上傳公開或有權使用的文件。')

COLUMNS = ['原始欄位', '原始數值', '標準欄位', '標準數值', '單位', '來源頁碼', '來源原文', '確認狀態']
ALIASES = {
    'model name': '產品型號', 'model number': '產品型號', 'part number': '料號',
    'module type': '產品類型', 'product type': '產品類型', 'form factor': '外型規格',
    'input i/f': '輸入介面', 'input interface': '輸入介面', 'host interface': '輸入介面',
    'output i/f': '輸出介面', 'output interface': '輸出介面',
    'output port': '輸出埠數量', 'output ports': '輸出埠數量', 'number of ports': '輸出埠數量',
    'output connector': '輸出連接器', 'connector': '連接器',
    'max.power consumption': '最大功耗', 'max. power consumption': '最大功耗',
    'maximum power consumption': '最大功耗', 'power consumption (max.)': '最大功耗',
    'dimension (w x l x h/mm)': '尺寸', 'dimensions': '尺寸', 'dimension': '尺寸',
    'operation': '工作溫度', 'operating temp.': '工作溫度',
    'operating temperature': '工作溫度', 'storage': '儲存溫度',
    'storage temperature': '儲存溫度', 'warranty': '保固', 'warranty period': '保固',
    'capacity': '容量', 'interface': '介面', 'sequential read': '連續讀取速度',
    'sequential write': '連續寫入速度', 'mtbf': '平均故障間隔', 'tbw': '耐寫總量',
    'vibration': '耐震', 'shock': '衝擊',
    'capacities': '容量', 'storage capacity': '容量', 'available capacities': '容量',
    'sata interface': '介面', 'interface type': '介面', 'transfer interface': '介面',
    'form factors': '外型規格', 'physical form factor': '外型規格', 'msata': '外型規格',
    'endurance': '耐寫總量', 'tbw (max.)': '耐寫總量', 'total bytes written': '耐寫總量',
    'flash type': '快閃記憶體類型', 'nand type': '快閃記憶體類型',
    'sequential read speed': '連續讀取速度', 'sequential write speed': '連續寫入速度',
}
SECTION_HEADERS = ('specifications', 'order info', 'order info.', 'ordering information',
                   'mechanical drawing', 'highlight features', 'features', 'notes')

def clean(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()

def normalized_key(value):
    return re.sub(r'[^\w\u4e00-\u9fff]+', '', clean(value).casefold())

def standard_name(name):
    key = clean(name).lower().rstrip(':：')
    return ALIASES.get(key, clean(name))

def detect_unit(value, label):
    """Respect case: Gb/s is transfer rate; GB is capacity."""
    v = clean(value)
    checks = [
        (r'(?i)\b(?:mb|mbyte)s?\s*/\s*(?:s|sec)\b', 'MB/s'),
        (r'\b(?:Gb|Gbit)\s*/\s*s\b|(?i:gbps)', 'Gb/s'),
        (r'\b(?:Mb|Mbit)\s*/\s*s\b|(?i:mbps)', 'Mb/s'),
        (r'(?i)\bMHz\b', 'MHz'), (r'(?i)\bGHz\b', 'GHz'),
        (r'°\s*C\b', '°C'), (r'(?i)\b(?:years?|yrs?)\b', '年'),
        (r'(?i)\bmm\b', 'mm'), (r'(?i)\bTBW\b', 'TBW'),
        (r'\b(?:TB|GB|MB)\b', None), (r'(?i)(?<![a-z])\d+(?:\.\d+)?\s*W\b', 'W'),
    ]
    for pattern, unit in checks:
        m = re.search(pattern, v)
        if m:
            return unit or m.group(0)
    return 'mm' if re.search(r'(?i)(?:/mm|\(mm\)|\bmm\b)', label) else ''

def candidate_row(label, value, page, evidence):
    label, value = clean(label).strip(' :：'), clean(value)
    if not label or not value or len(label) > 85 or len(value) > 600:
        return None
    if label.lower() in SECTION_HEADERS or label.lower() in ('p/n', 'description'):
        return None
    # Reject common schematic/OCR fragments; keep unknown legitimate specification labels.
    if re.search(r'(?i)(?:nnc+aa|cache\s*blocks|main storage blocks)', label):
        return None
    if label.lower() in ('dram cache', 'slc blocks') and value.lower() in ('4/8ch', 'main storage blocks'):
        return None
    if label.lower() == 'msata' and re.search(r'(?i)\bmm\b', value):
        label = '尺寸／重量（mSATA 版本）'  # preserve applicability; not form factor
    return {'原始欄位': label, '原始數值': value,
            '標準欄位': standard_name(label), '標準數值': value,
            '單位': detect_unit(value, label), '來源頁碼': f'第 {page} 頁',
            '來源原文': clean(evidence)[:650], '確認狀態': '待人工確認'}

@st.cache_data(show_spinner=False)
def parse_pdf(raw):
    pages = []
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        for number, page in enumerate(pdf.pages, 1):
            pages.append({'number': number, 'layout': page.extract_text(layout=True) or '',
                          'plain': page.extract_text(layout=False) or '',
                          'tables': page.extract_tables() or []})
    return pages

def prose_candidates(page):
    """Conservative generic prose patterns; preserve literal evidence, no invented numbers."""
    text = page['plain']
    n = page['number']
    patterns = [
        ('產品型號', r'(?im)^\s*(SanDisk[®\s]+[A-Z]\d{2,5}\s+SSD)\b'),
        ('產品型號', r'(?im)^\s*(?:model\s+(?:name|number)\s*[:：]?\s*)([A-Z0-9][A-Z0-9-]{3,})\b'),
        ('介面', r'(?i)\b(SATA\s+(?:revision|rev\.?|III|3\.?[0-9]?)\s*[0-9.]*\s*[,;]?\s*(?:6\s*Gb/s)?\s*(?:interface)?)'),
        ('容量', r'(?i)\bcapacities\s+of\s+((?:\d+(?:\.\d+)?\s*(?:GB|TB)\s*[,;]?(?:\s+and\s+)?\s*){2,})'),
        ('容量', r'(?i)\b(?:up to)\s+(\d+(?:\.\d+)?\s*(?:TB|GB))\s+capacities'),
        ('外型規格', r'(?i)(2\.5[”"\s/0-9.]*\s*cased\s*,\s*M\.2\s*2280\s*,?\s*and\s*mSATA)\s+form\s+factors'),
        ('耐寫總量', r'(?i)(tested\s+for\s+\d+\s*TBW\s*\([^)]*\)\s*and\s*\d+\s*TBW\s*\([^)]*\))'),
    ]
    found = []
    for label, pattern in patterns:
        for m in re.finditer(pattern, text):
            value = clean(m.group(1))
            if label == '產品型號' and value.lower().startswith('sandisk'):
                value = re.sub(r'(?i)^sandisk[®\s]+', '', value)
            row = candidate_row(label, value, n, text[max(0, m.start()-75):min(len(text), m.end()+75)])
            if row:
                found.append(row)
    # Capture multi-line descriptions, including the SanDisk X300 form-factor sentence.
    flat = clean(text)
    for match in re.finditer(r'(?i)(2\.5\s*[”"\']?\s*/\s*7\s*mm\s*cased\s*,\s*M\.2\s*2280\s*,?\s*and\s*mSATA)\s+form\s+factors', flat):
        row = candidate_row('外型規格', match.group(1), n, flat[max(0,match.start()-70):match.end()+70])
        if row:
            found.append(row)
    return found


def table_candidates(page):
    """Best-effort two-column table extraction, without pretending every cell is a spec."""
    found = []
    for table in page.get('tables', []):
        for cells in table:
            values = [clean(c) for c in cells]
            nonempty = [(i, v) for i, v in enumerate(values) if v]
            if len(nonempty) != 2:
                continue
            (i, label), (j, value) = nonempty
            if i >= j or not re.search(r'[A-Za-z\u4e00-\u9fff]', label):
                continue
            if len(label) > 55 or len(value) > 250 or label.lower() in ('msata', 'm.2', '2.5"'):
                continue
            if not (re.search(r'\d', value) or len(value) <= 70):
                continue
            row = candidate_row(label, value, page['number'], f'{label}: {value}')
            if row:
                found.append(row)
    return found


def extract_dynamic(pages):
    results, seen = [], set()
    for page in pages:
        number = page['number']
        in_spec = False
        for line in page['layout'].splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            lower = stripped.lower().rstrip(' .')
            if lower in ('specifications', 'technical specifications', 'specification'):
                in_spec = True
                continue
            if lower.startswith(('order info', 'ordering information', 'mechanical drawing', 'contact')):
                in_spec = False
            if not in_spec:
                continue
            cells = [clean(c) for c in re.split(r'\s{2,}', stripped) if clean(c)]
            if len(cells) >= 2:
                row = candidate_row(cells[0], ' '.join(cells[1:]), number, line)
            else:
                m = re.match(r'^\s*([\w][\w /().-]{1,65}?)\s*[:：]\s*(\S.+)$', line)
                row = candidate_row(m.group(1), m.group(2), number, line) if m else None
            if row:
                sig = (number, row['原始欄位'].casefold(), row['原始數值'])
                if sig not in seen:
                    results.append(row)
                    seen.add(sig)
        if not any(r['來源頁碼'] == f'第 {number} 頁' for r in results):
            for line in page['plain'].splitlines():
                m = re.match(r'^\s*([A-Za-z][A-Za-z /().-]{2,65})\s*[:：]\s*(\S.+)$', line)
                if m:
                    row = candidate_row(m.group(1), m.group(2), number, line)
                    if row:
                        sig = (number, row['原始欄位'].casefold(), row['原始數值'])
                        if sig not in seen:
                            results.append(row)
                            seen.add(sig)
        # Run complementary extraction on every page, not only when the line parser fails.
        for row in table_candidates(page) + prose_candidates(page):
            sig = (number, normalized_key(row['標準欄位']), normalized_key(row['標準數值']))
            if not any((r['來源頁碼'] == f'第 {number} 頁' and
                        normalized_key(r['標準欄位']) == sig[1] and
                        normalized_key(r['標準數值']) == sig[2]) for r in results):
                results.append(row)
    # Deduplicate exact semantic/value repeats, including repeated page-2 labels.
    unique, seen_values = [], set()
    for row in results:
        sig = (normalized_key(row['標準欄位']), normalized_key(row['標準數值']))
        if sig in seen_values:
            continue
        seen_values.add(sig)
        unique.append(row)
    return pd.DataFrame(unique, columns=COLUMNS)

def prepare(uploaded, prefix):
    raw = uploaded.getvalue()
    if len(raw) > 15 * 1024 * 1024:
        st.error(f'{prefix} PDF 超過 15 MB。')
        st.stop()
    try:
        pages = parse_pdf(raw)
    except Exception as exc:
        st.error(f'{prefix} PDF 無法解析：{exc}')
        st.stop()
    if not any(p['plain'].strip() for p in pages):
        st.error(f'{prefix} PDF 沒有可擷取文字，暫不支援掃描檔 OCR。')
        st.stop()
    doc_key = hashlib.sha256(raw).hexdigest()
    if st.session_state.get(f'{prefix}_key') != doc_key:
        st.session_state[f'{prefix}_key'] = doc_key
        st.session_state[f'{prefix}_table'] = extract_dynamic(pages)
        st.session_state[f'{prefix}_revision'] = st.session_state.get(f'{prefix}_revision', 0) + 1
    return pages

def interpret_pair(field, a_value, b_value):
    """Conservative interpretation; never infer product superiority from raw strings."""
    if not a_value or not b_value:
        return '資料不足', '其中一方缺少可用資料；不代表產品未提供此規格'
    key = standard_name(field)
    if key == '產品型號':
        return '需人工判讀', '不同型號為正常現象，不代表性能差異'
    if key == '介面':
        a_sata = bool(re.search(r'(?i)\bsata\b', a_value))
        b_sata = bool(re.search(r'(?i)\bsata\b', b_value))
        if a_sata and b_sata:
            a_6 = bool(re.search(r'(?i)6\s*(?:\.?0\s*)?gb\s*/\s*s', a_value))
            b_6 = bool(re.search(r'(?i)6\s*(?:\.?0\s*)?gb\s*/\s*s', b_value))
            if a_6 and b_6:
                return '規格部分一致', '均標示 SATA 6 Gb/s；修訂版本與其他特性仍需確認'
            return '需人工判讀', '均標示 SATA，但速度或版本資訊不足'
    if key == '容量':
        return '需人工判讀', '容量範圍與多容量清單需按具體型號比較'
    if key == '外型規格':
        if re.search(r'2\.5', a_value) and re.search(r'2\.5', b_value):
            return '規格部分一致', '均包含 2.5 吋外型；厚度等細節需確認'
        return '需人工判讀', '需確認外型、尺寸與連接器是否相容'
    if key == '耐寫總量':
        return '需人工判讀', 'TBW 必須確認容量、單位與測試條件；不得直接比較數字'
    if normalized_key(a_value) == normalized_key(b_value):
        return '文字一致', '僅表示文字一致；尚未驗證測試條件'
    return '需人工判讀', '兩邊都有資料，但未確認單位、版本或適用條件'


def match_specs(left, right):
    """Pair on normalized labels; retain unmatched records and explain uncertainty."""
    left_rows = left.fillna('').to_dict('records')
    right_rows = right.fillna('').to_dict('records')
    available = set(range(len(right_rows)))
    output = []
    for a in left_rows:
        key = normalized_key(a.get('標準欄位') or a.get('原始欄位'))
        matching = [j for j in sorted(available) if key and key == normalized_key(
            right_rows[j].get('標準欄位') or right_rows[j].get('原始欄位'))]
        j = matching[0] if matching else None
        b = right_rows[j] if j is not None else {}
        if j is not None:
            available.remove(j)
        a_value, b_value = clean(a.get('標準數值')), clean(b.get('標準數值'))
        field = clean(a.get('標準欄位') or a.get('原始欄位'))
        state, note = interpret_pair(field, a_value, b_value)
        output.append({'比較項目': field, '產品A數值': a_value, '產品B數值': b_value,
                       '比較狀態': state, '產品A原始欄位': clean(a.get('原始欄位')),
                       '產品B原始欄位': clean(b.get('原始欄位')),
                       '產品A來源': clean(a.get('來源頁碼')),
                       '產品B來源': clean(b.get('來源頁碼')),
                       '人工判讀': '', '備註': note})
    for j in sorted(available):
        b = right_rows[j]
        output.append({'比較項目': clean(b.get('標準欄位') or b.get('原始欄位')),
                       '產品A數值': '', '產品B數值': clean(b.get('標準數值')),
                       '比較狀態': '資料不足', '產品A原始欄位': '',
                       '產品B原始欄位': clean(b.get('原始欄位')),
                       '產品A來源': '', '產品B來源': clean(b.get('來源頁碼')),
                       '人工判讀': '', '備註': '產品 A 未擷取到對應規格；不代表未提供'})
    return pd.DataFrame(output, columns=['比較項目', '產品A數值', '產品B數值', '比較狀態',
                                         '產品A原始欄位', '產品B原始欄位',
                                         '產品A來源', '產品B來源', '人工判讀', '備註'])

def canonical_field(field, value=''):
    field = standard_name(field)
    if field == '外型規格' and re.search(r'(?i)\bmm\b', clean(value)) and re.search(r'(?i)\b(?:g|gram)\b', clean(value)):
        return '尺寸／重量（mSATA 版本）'
    return field


def collapse_specs(df):
    """Consolidate equivalent observations for comparison; retain original rows for audit.

    Conflicting values are NOT silently resolved: they remain a review-required row.
    """
    groups = {}
    for row in df.fillna('').to_dict('records'):
        if clean(row.get('確認狀態')) == '不適用':
            continue
        field = canonical_field(clean(row.get('標準欄位') or row.get('原始欄位')), row.get('標準數值'))
        if not field: continue
        groups.setdefault(normalized_key(field), []).append((field,row))
    output=[]
    for items in groups.values():
        field = items[0][0]
        def score(item):
            value=clean(item[1].get('標準數值'))
            # Prefer explicit interface speed over incomplete fragments.
            return (int(bool(re.search(r'(?i)\b(?:gb/s|gbps|mb/s)\b',value))),len(value))
        chosen=max(items,key=score)[1].copy()
        values=[clean(r.get('標準數值')) for _,r in items if clean(r.get('標準數值'))]
        if field == '介面':
            # SATA III, SATA III 6.0 and SATA III 6.0 Gb/s can be same spec.
            distinct = {normalized_key(re.sub(r'(?i)\b(?:6(?:\.0)?\s*gb/s|6\.0)\b','',v)) for v in values}
            same_family = all(re.search(r'(?i)\bsata\b',v) for v in values)
            conflicting = not same_family and len(set(normalized_key(v) for v in values))>1
        else:
            conflicting = len(set(normalized_key(v) for v in values))>1
        chosen['標準欄位']=field
        chosen['來源頁碼']='、'.join(dict.fromkeys(clean(r.get('來源頁碼')) for _,r in items if clean(r.get('來源頁碼'))))
        chosen['來源原文']=' | '.join(dict.fromkeys(clean(r.get('來源原文')) for _,r in items if clean(r.get('來源原文'))))[:650]
        if conflicting:
            chosen['確認狀態']='需修正'
            chosen['標準數值']=' ｜ '.join(dict.fromkeys(values))
        chosen['_合併註記'] = ('同欄位有不同候選值，請人工確認；' if conflicting else '') + f'已整合 {len(items)} 筆來源候選' if len(items)>1 else ('同欄位數值衝突，請人工確認' if conflicting else '')
        output.append(chosen)
    return pd.DataFrame(output,columns=COLUMNS+['_合併註記'])


def review_priority(row):
    field=clean(row.get('比較項目'))
    state=clean(row.get('比較狀態'))
    if field in ('產品型號','料號'):
        return '低', '不同產品識別資訊屬正常情形，不需以此判定性能差異'
    if field in ('耐寫總量','連續讀取速度','連續寫入速度','工作溫度','最大功耗','容量'):
        return '高', '優先核對適用型號、容量、單位與測試條件'
    if state == '人工確認不同':
        return '高', '確認差異是否影響使用需求'
    if state == '資料不足':
        return '中', '核對原始文件是否有未擷取的資料'
    return '中', '核對雙方來源及比較條件'

# 報告分類採通用欄位字典；未知類型保留在「其他規格」，不推測產品優劣。
CATEGORY_MAP = {
    '產品型號': '基本資訊', '產品類型': '基本資訊', '料號': '基本資訊',
    '介面': '連接與相容性', '輸入介面': '連接與相容性', '輸出介面': '連接與相容性',
    '外型規格': '連接與相容性', '連接器': '連接與相容性', '輸出連接器': '連接與相容性',
    '輸出埠數量': '連接與相容性', '容量': '容量與效能', '連續讀取速度': '容量與效能',
    '連續寫入速度': '容量與效能', '耐寫總量': '可靠性與耐用度',
    '平均故障間隔': '可靠性與耐用度', '工作溫度': '環境與電力',
    '儲存溫度': '環境與電力', '最大功耗': '環境與電力',
    '尺寸': '機構與其他', '尺寸／重量（mSATA 版本）': '機構與其他', '保固': '基本資訊', '快閃記憶體類型': '容量與效能',
}


def excel_bytes(left, right, comparison, names):
    """五工作表的可稽核分析報告；結論僅由比較狀態及備註產生。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins
    from datetime import datetime

    wb = Workbook()
    summary = wb.active
    summary.title = '01_分析摘要'
    detail = wb.create_sheet('02_核心規格比較')
    issues = wb.create_sheet('03_差異與待確認')
    raw = wb.create_sheet('04_原始規格資料')
    sources = wb.create_sheet('05_來源與方法')
    navy, teal, pale = '17324D', '087E8B', 'E9F3F5'
    gray, white = '5C6875', 'FFFFFF'
    status_colors = {'資料不足': 'FCE8E6', '需人工判讀': 'FFF2D8',
                     '規格部分一致': 'E2F1F1', '文字一致': 'E7F2E4',
                     '人工確認相同': 'D8F0D9', '人工確認不同': 'FCE8E6'}
    def title(ws, text, width):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
        cell=ws.cell(1,1,text); cell.fill=PatternFill('solid',fgColor=navy)
        cell.font=Font(name='Microsoft JhengHei',size=17,bold=True,color=white)
        cell.alignment=Alignment(vertical='center'); ws.row_dimensions[1].height=36
        for row in ws.iter_rows(min_row=1,max_row=1,max_col=width):
            for c in row: c.fill=PatternFill('solid',fgColor=navy)
    def header(ws,row,labels):
        for i,label in enumerate(labels,1):
            c=ws.cell(row,i,label); c.fill=PatternFill('solid',fgColor=teal)
            c.font=Font(name='Microsoft JhengHei',bold=True,color=white)
            c.alignment=Alignment(wrap_text=True,vertical='center')
        ws.row_dimensions[row].height=27
    def body(ws, row, vals, highlight=None):
        for i,val in enumerate(vals,1):
            c=ws.cell(row,i, '' if val is None else str(val))
            c.font=Font(name='Microsoft JhengHei',size=10,color=navy)
            c.alignment=Alignment(wrap_text=True,vertical='top')
            if row%2==0: c.fill=PatternFill('solid',fgColor='F3F7FA')
            if highlight and i==highlight:
                color=status_colors.get(str(val))
                if color: c.fill=PatternFill('solid',fgColor=color)
        ws.row_dimensions[row].height=36
    def setup(ws, widths, freeze='A4', filter_row=None):
        ws.freeze_panes=freeze
        for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
        ws.sheet_view.showGridLines=False
        if filter_row: ws.auto_filter.ref=f'A{filter_row}:{get_column_letter(len(widths))}{ws.max_row}'
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
        ws.print_options.horizontalCentered=True
        ws.page_margins=PageMargins(left=.25,right=.25,top=.5,bottom=.5,header=.2,footer=.2)
    rows=comparison.fillna('').to_dict('records')
    matched=[r for r in rows if clean(r.get('產品A數值')) and clean(r.get('產品B數值'))]
    missing=[r for r in rows if not clean(r.get('產品A數值')) or not clean(r.get('產品B數值'))]
    review=[r for r in rows if r.get('比較狀態') in ('需人工判讀','人工確認不同') and clean(r.get('比較項目')) not in ('產品型號','料號')]
    partial=[r for r in rows if r.get('比較狀態') in ('規格部分一致','文字一致','人工確認相同')]
    def model(df, fallback):
        found=df[df['標準欄位'].astype(str).str.strip()=='產品型號'] if '標準欄位' in df else pd.DataFrame()
        return clean(found.iloc[0]['標準數值']) if not found.empty else fallback
    name_a,name_b=model(left,names[0]),model(right,names[1])
    title(summary,'工業產品｜競品規格分析摘要',5)
    summary.merge_cells('A2:E2'); summary['A2']='規則式候選分析｜需核對原始 PDF；不構成產品優劣或採購建議'
    summary['A2'].font=Font(name='Microsoft JhengHei',italic=True,color=gray,size=10)
    header(summary,4,['比較產品','產品 A','產品 B','資料範圍','產出日期'])
    body(summary,5,['型號／檔案',name_a,name_b,'文字型 PDF',datetime.now().strftime('%Y-%m-%d')])
    header(summary,7,['比較總列數','雙方有資料','部分一致／文字一致','待人工判讀','單方資料不足'])
    body(summary,8,[len(rows),len(matched),len(partial),len(review),len(missing)])
    summary.merge_cells('A10:E10'); summary['A10']='主要發現（僅依目前擷取及人工修正結果）'
    summary['A10'].font=Font(name='Microsoft JhengHei',bold=True,size=12,color=navy)
    notes=[]
    for r in partial[:4]:
        notes.append(f"【{clean(r.get('比較項目'))}】{clean(r.get('備註')) or '兩側文字相符，仍需核對來源'}")
    if not notes: notes.append('目前沒有足以列為一致的比較項目。')
    notes.append(f'有 {len(review)} 列需人工判讀、{len(missing)} 列單方資料不足；未擷取到不代表原文件未提供。')
    notes.append('TBW、效能、功耗等數值須先確認容量、單位與測試條件，才能比較。')
    for i,t in enumerate(notes,11):
        summary.merge_cells(start_row=i,start_column=1,end_row=i,end_column=5)
        body(summary,i,[f'• {t}']); summary.row_dimensions[i].height=42
    setup(summary,[26,32,32,30,22],freeze='A5')

    title(detail,'核心規格比較｜按主題分類',9)
    detail.merge_cells('A2:I2'); detail['A2']='所有擷取項目均保留；「其他規格」不代表無關，需視產品類別判斷。'
    cols=['規格分類','比較項目','產品 A','產品 B','比較狀態','分析備註','A 來源頁','B 來源頁','人工判讀']
    header(detail,3,cols)
    ordered=sorted(rows,key=lambda r:(list(dict.fromkeys(CATEGORY_MAP.values())).index(CATEGORY_MAP.get(clean(r.get('比較項目')),'機構與其他')) if CATEGORY_MAP.get(clean(r.get('比較項目')),'機構與其他') in CATEGORY_MAP.values() else 99,clean(r.get('比較項目'))))
    for i,r in enumerate(ordered,4):
        body(detail,i,[CATEGORY_MAP.get(clean(r.get('比較項目')),'其他規格'),r.get('比較項目'),r.get('產品A數值'),r.get('產品B數值'),r.get('比較狀態'),r.get('備註'),r.get('產品A來源'),r.get('產品B來源'),r.get('人工判讀')],highlight=5)
    setup(detail,[20,24,34,34,20,62,14,14,28],filter_row=3)

    title(issues,'差異與待確認事項｜人工審查清單',8)
    issues.merge_cells('A2:H2'); issues['A2']='優先檢查單方缺漏、數值條件不明與人工標註；無法確認時不判定優劣。'
    header(issues,3,['優先級','比較項目','產品 A','產品 B','問題類型','建議處理','來源 A','來源 B'])
    flagged=[r for r in rows if r.get('比較狀態') not in ('文字一致','人工確認相同','規格部分一致')]
    flagged.sort(key=lambda r:({'高':0,'中':1,'低':2}[review_priority(r)[0]], clean(r.get('比較項目'))))
    for i,r in enumerate(flagged,4):
        state=clean(r.get('比較狀態'))
        priority, advice=review_priority(r)
        recommendation=advice + '；' + (clean(r.get('備註')) or '核對 PDF 原始資料與適用型號')
        body(issues,i,[priority,r.get('比較項目'),r.get('產品A數值'),r.get('產品B數值'),state,recommendation,r.get('產品A來源'),r.get('產品B來源')],highlight=5)
    setup(issues,[12,25,35,35,22,66,14,14],filter_row=3)

    title(raw,'原始規格與修正紀錄｜保留證據',10)
    raw.merge_cells('A2:J2'); raw['A2']='來源原文與確認狀態供稽核；人工修改後的標準值不等於已經查證。'
    raw_cols=['產品','原始欄位','原始數值','標準欄位','標準數值','單位','來源頁碼','來源原文','確認狀態','檔名']
    header(raw,3,raw_cols)
    i=4
    for label,df,filename in [('A',left,names[0]),('B',right,names[1])]:
        for r in df.fillna('').to_dict('records'):
            body(raw,i,[label,r.get('原始欄位'),r.get('原始數值'),r.get('標準欄位'),r.get('標準數值'),r.get('單位'),r.get('來源頁碼'),r.get('來源原文'),r.get('確認狀態'),filename]); i+=1
    setup(raw,[10,24,34,24,34,12,14,75,18,38],filter_row=3)

    title(sources,'資料來源、方法與限制',3)
    header(sources,3,['項目','內容','說明'])
    entries=[('產品 A 原始檔',names[0],'來源頁碼詳見核心比較及原始規格資料'),
             ('產品 B 原始檔',names[1],'來源頁碼詳見核心比較及原始規格資料'),
             ('擷取方式','pdfplumber 文字、表格與規則式辨識','目前不支援掃描 PDF OCR'),
             ('標準化方式','欄位別名映射與人工修正','未知欄位仍保留原名稱'),
             ('比較方式','同名標準欄位配對與保守規則判讀','不會自動評估產品優劣'),
             ('資料不足','表示至少一方未擷取到可用值','不代表產品或 PDF 不存在該規格'),
             ('文字一致','只表示標準數值文字相同','不代表測試條件一致'),
             ('部分一致','有部分共同特性','不代表完全相容或規格相同'),
             ('耐寫總量 TBW','依容量、測試條件判讀','不得直接比較不同容量的數字'),
             ('人工確認','編輯器可標註已確認','報告不會自行驗證人工修正的正確性'),
             ('產品定位','不同產品類型不可直接判斷優劣','建議選擇相同用途、容量與介面的產品')]
    for i,entry in enumerate(entries,4): body(sources,i,entry)
    setup(sources,[26,68,78],freeze='A4')
    out=io.BytesIO(); wb.save(out); return out.getvalue()

st.subheader('1｜上傳兩份產品規格書')
a_col, b_col = st.columns(2)
with a_col:
    file_a = st.file_uploader('產品 A（例如宜鼎）PDF', type=['pdf'], key='pdf_a')
with b_col:
    file_b = st.file_uploader('產品 B（競品）PDF', type=['pdf'], key='pdf_b')
if not file_a or not file_b:
    st.warning('請分別上傳產品 A、B 的文字型 PDF，才能開始比較。')
    st.stop()

pages_a = prepare(file_a, 'A')
pages_b = prepare(file_b, 'B')
st.success(f'產品 A：{len(pages_a)} 頁、{len(st.session_state.A_table)} 筆候選規格；'
           f'產品 B：{len(pages_b)} 頁、{len(st.session_state.B_table)} 筆候選規格。')
st.caption('建議選擇相同用途的產品。若欄位名稱不同，可先修改「標準欄位」為相同名稱，再重新產生比較表。')

st.subheader('2｜檢視與修正雙方規格')
left_col, right_col = st.columns(2)
with left_col:
    st.markdown('**產品 A 規格**')
    edited_a = st.data_editor(st.session_state.A_table, use_container_width=True,
                              hide_index=True, num_rows='dynamic',
                              disabled=['來源頁碼', '來源原文'],
                              key=f'A_editor_{st.session_state.A_revision}',
                              column_config={'確認狀態': st.column_config.SelectboxColumn(
                                  options=['待人工確認', '已確認', '需修正', '不適用'])})
    with st.expander('產品 A PDF 原始文字'):
        page_a = st.selectbox('A 頁碼', [p['number'] for p in pages_a])
        st.text_area('A 原文', next(p['plain'] for p in pages_a if p['number'] == page_a), height=300)
with right_col:
    st.markdown('**產品 B 規格**')
    edited_b = st.data_editor(st.session_state.B_table, use_container_width=True,
                              hide_index=True, num_rows='dynamic',
                              disabled=['來源頁碼', '來源原文'],
                              key=f'B_editor_{st.session_state.B_revision}',
                              column_config={'確認狀態': st.column_config.SelectboxColumn(
                                  options=['待人工確認', '已確認', '需修正', '不適用'])})
    with st.expander('產品 B PDF 原始文字'):
        page_b = st.selectbox('B 頁碼', [p['number'] for p in pages_b])
        st.text_area('B 原文', next(p['plain'] for p in pages_b if p['number'] == page_b), height=300)

st.subheader('3｜競品規格比較')
comparison = match_specs(collapse_specs(edited_a), collapse_specs(edited_b))
# Merge provenance/ambiguity notes into comparison for audit.
for _idx, _r in comparison.iterrows():
    _field = clean(_r['比較項目'])
    _notes = []
    for _df in (collapse_specs(edited_a), collapse_specs(edited_b)):
        _hits = _df[_df['標準欄位'].astype(str).map(normalized_key) == normalized_key(_field)]
        if not _hits.empty and clean(_hits.iloc[0].get('_合併註記')):
            _notes.append(clean(_hits.iloc[0]['_合併註記']))
    if _notes:
        comparison.at[_idx, '備註'] = clean(_r['備註']) + '；' + '；'.join(_notes)
        if any('不同候選值' in x for x in _notes):
            comparison.at[_idx, '比較狀態'] = '需人工判讀'
matched = (comparison['產品A數值'].astype(str).str.strip().ne('') & comparison['產品B數值'].astype(str).str.strip().ne('')).sum() if not comparison.empty else 0
st.write(f'共 {len(comparison)} 個比較列，配對 {matched} 列；其餘為單方資料，保留供人工確認。')
st.caption('已將同名規格整合成單列，原始紀錄仍完整保留在報告中；若數值衝突，需人工確認。資料不足不代表產品沒有該規格；規格部分一致不代表完全相同。TBW、容量等需確認適用型號及測試條件。')
comparison_edited = st.data_editor(
    comparison, use_container_width=True, hide_index=True, num_rows='dynamic',
    column_config={'比較狀態': st.column_config.SelectboxColumn(
        options=['文字一致', '規格部分一致', '需人工判讀', '資料不足', '人工確認相同', '人工確認不同', '不適用'])},
    key=f'comparison_{st.session_state.A_revision}_{st.session_state.B_revision}')
st.download_button('📥 下載專業競品分析報告 Excel',
                   data=excel_bytes(edited_a, edited_b, comparison_edited, (file_a.name, file_b.name)),
                   file_name='industrial_competitor_report_v051.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
