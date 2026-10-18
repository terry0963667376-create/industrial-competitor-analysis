"""Universal industrial datasheet MVP v0.7: two-PDF comparison, no paid API/AI."""
import hashlib
import io
import re
import json
import os
import urllib.request
import urllib.error

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title='工業規格書｜競品比較 MVP v0.8', page_icon='📊', layout='wide')
st.title('工業規格書標準化與競品比較｜MVP v0.8')
st.caption('雙 PDF 動態擷取 → 人工修正 → 規格配對 → Excel 比較報告｜規則式／Gemini AI 雙模式；AI 結果需人工確認')
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

def gemini_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")

        if not key:
            st.error("Secrets 中找不到 GEMINI_API_KEY")
            return ""

        if not isinstance(key, str):
            st.error("API Key 的資料格式不正確")
            return ""

        if key != key.strip():
            st.warning("API Key 前後存在空白字元")

        st.success("已成功從 Streamlit Secrets 讀取 API Key")

        return key.strip()

    except Exception:
        st.error("無法讀取 Streamlit Secrets")
        return ""


def gemini_connection_test(api_key, model_name='gemini-3.8-flash'):
    """Minimal independent Gemini request. Never include API key in output."""
    if not re.fullmatch(r'[a-zA-Z0-9._-]+', model_name):
        return False, '模型名稱格式不正確。'
    url = ('https://generativelanguage.googleapis.com/v1beta/models/'
           + model_name + ':generateContent')
    payload = {'contents': [{'parts': [{'text': 'Reply with OK only.'}]}],
               'generationConfig': {'maxOutputTokens': 16}}
    request = urllib.request.Request(
        url, data=json.dumps(payload).encode('utf-8'), method='POST',
        headers={'Content-Type': 'application/json', 'x-goog-api-key': api_key})
    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            data = json.load(response)
        parts = data.get('candidates', [{}])[0].get('content', {}).get('parts', [])
        reply = ''.join(p.get('text', '') for p in parts).strip()
        return True, f'Gemini API 連線成功（HTTP 200）。模型回覆：{reply[:80] or "空白回覆"}'
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode('utf-8', errors='replace')
        try:
            error = json.loads(raw).get('error', {})
            reasons = [d.get('reason', '') for d in error.get('details', []) if isinstance(d, dict)]
            reason = next((r for r in reasons if r), '')
            message = str(error.get('message', ''))
        except (ValueError, TypeError):
            reason, message = '', ''
        # Never echo arbitrary server text, which could include request details.
        if reason == 'API_KEY_INVALID' or 'API key not valid' in message:
            advice = 'API_KEY_INVALID：Google 不接受目前金鑰。請檢查 AI Studio 金鑰、專案與金鑰限制。'
        elif exc.code == 403:
            advice = 'HTTP 403：請檢查 API 權限、金鑰應用程式限制及專案設定。'
        elif exc.code == 429:
            advice = 'HTTP 429：請檢查免費額度或呼叫速率限制。'
        elif exc.code == 404:
            advice = 'HTTP 404：模型名稱或 API 版本可能不支援。'
        else:
            advice = f'HTTP {exc.code}：Google 拒絕請求；請查看 Streamlit 日誌或 Google Cloud 設定。'
        return False, advice
    except (urllib.error.URLError, TimeoutError):
        return False, '連線逾時或網路錯誤，請稍後再試。'
    except Exception:
        return False, '測試發生非預期錯誤，請查看 Streamlit 日誌（勿公開金鑰）。'


# v0.8: 固定規格 ID + 適用條件。未知類型保留原欄位，不強制配對。
FIELD_CATALOG = {
    'MODEL': '產品型號', 'PART_NUMBER': '料號', 'PRODUCT_TYPE': '產品類型',
    'INTERFACE': '介面', 'FORM_FACTOR': '外型規格', 'CAPACITY': '容量',
    'SEQ_READ': '連續讀取速度', 'SEQ_WRITE': '連續寫入速度',
    'RANDOM_READ': '隨機讀取速度', 'RANDOM_WRITE': '隨機寫入速度',
    'TBW': '耐寫總量', 'MTBF': '平均故障間隔',
    'OPERATING_TEMP': '工作溫度', 'STORAGE_TEMP': '儲存溫度',
    'WARRANTY': '保固', 'SHOCK': '衝擊', 'VIBRATION': '耐震',
    'MAX_POWER': '最大功耗', 'IDLE_POWER': '閒置功耗',
    'DIMENSIONS': '尺寸', 'NAND_TYPE': '快閃記憶體類型',
    'CONNECTOR': '連接器',
}

# 僅明確同義詞可規則化；不把隨機速度和連續速度混為一談。
FIELD_SYNONYMS = {
    '傳輸介面': '介面', '傳輸接口': '介面', '傳輸界面': '介面',
    '保固年限': '保固', '保固期': '保固', '保固期限': '保固',
    '抗衝擊': '衝擊', '耐衝擊力': '衝擊', '衝擊耐受度': '衝擊',
    '作業溫度': '工作溫度', '操作溫度': '工作溫度',
    '外觀規格': '外型規格', '外觀尺寸': '尺寸',
    '循序讀取速度': '連續讀取速度', '循序寫入速度': '連續寫入速度',
    '順序讀取速度': '連續讀取速度', '順序寫入速度': '連續寫入速度',
    '最大耗電量': '最大功耗', '待機功耗': '閒置功耗',
}


def field_variant(name):
    """保留容量/版本等限定條件；未知限定條件不猜測。"""
    name = clean(name)
    m = re.search(r'[（(]([^()（）]+)[）)]', name)
    return clean(m.group(1)) if m else ''


def normalize_field(name, field_id='', variant=''):
    raw = clean(name)
    if field_id in FIELD_CATALOG:
        base = FIELD_CATALOG[field_id]
    else:
        base = re.sub(r'\s*[（(][^()（）]+[）)]', '', raw).strip()
        base = FIELD_SYNONYMS.get(base, standard_name(base))
    qualifier = clean(variant) or field_variant(raw)
    # 不把通用描述或單位當成版本條件。
    if qualifier and normalized_key(qualifier) not in {'max', 'maximum', '最大', 'min', 'minimum'}:
        return f'{base}（{qualifier}）'
    return base


def pairing_key(field):
    field = clean(field)
    variant = field_variant(field)
    base = re.sub(r'\s*[（(][^()（）]+[）)]', '', field).strip()
    base = FIELD_SYNONYMS.get(base, standard_name(base))
    return normalized_key(base), normalized_key(variant)


def ai_extract(pages, api_key, model_name):
    # Send extracted PDF text to Gemini.
    # Cap text size to keep free-tier requests manageable; surface truncation.
    chunks = []
    remaining = 44000
    for p in pages:
        t = p['plain'].strip()
        if not t or remaining <= 0:
            continue
        take = t[:remaining]
        chunks.append(f"[PAGE {p['number']}]\n{take}")
        remaining -= len(take)
    if not chunks:
        raise ValueError('沒有可供 AI 分析的 PDF 文字。')
    prompt = """你是工業產品規格書的嚴格資料擷取助手。只擷取文件明確出現的規格，不得補寫或推測。
請輸出 JSON 物件，頂層只有 specs 陣列。每筆包含：
label (原始欄位), value (原始數值), standard_field (繁體中文名稱),
field_id (以下固定 ID 之一；不適用請用 UNKNOWN),
variant (適用容量/外型/版本/測試條件；未明示用空字串),
page (整數頁碼), evidence (逐字出現在該頁原文的短句)。
可用 ID：""" + ', '.join(FIELD_CATALOG) + """。
重要規則：
1. 不同容量、外型或版本的數值分別擷取，variant 必須保留適用條件。
2. SEQ_READ 與 SEQ_WRITE 必須區分，MAX_POWER 與 IDLE_POWER 必須區分。
3. 如果文件將讀寫速度合寫為「550 / 510 MB/s」，只有在文件明示順序時才拆成兩筆；否則保留原項目並設 UNKNOWN。
4. 不得混淆 GB 與 Gb/s。不要推論未提供的值；證據與頁碼無法核實時省略。
5. evidence 必須是原文中連續出現的文字，不要輸出 Markdown。
文件如下：\n"""
    payload = {'contents': [{'parts': [{'text': prompt + '\n'.join(chunks)}]}],
               'generationConfig': {'responseMimeType': 'application/json', 'temperature': 0.0}}
    url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model_name + ':generateContent'
    request = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'),
                                     headers={'Content-Type': 'application/json', 'x-goog-api-key': api_key},
                                     method='POST')
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.load(response)
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode('utf-8', errors='replace')[:500]
        raise RuntimeError(f'Gemini HTTP {exc.code}：{detail}') from None
    parts = result.get('candidates', [{}])[0].get('content', {}).get('parts', [])
    content = ''.join(p.get('text', '') for p in parts)
    data = json.loads(content)
    rows, rejected = [], 0
    page_text = {p['number']: p['plain'] for p in pages}
    for spec in data.get('specs', [])[:160]:
        try:
            pg = int(spec['page'])
            label = clean(spec['label'])
            value = clean(spec['value'])
            evidence = clean(spec['evidence'])
            field = clean(spec['standard_field'])
            field_id = clean(spec.get('field_id', 'UNKNOWN')).upper()
            variant = clean(spec.get('variant', ''))
            # Conservative evidence verification: exact evidence appears on the cited page
            # (whitespace normalized to handle PDF extraction line wraps).
            if not label or not value or not field or not evidence or evidence not in clean(page_text.get(pg, '')):
                rejected += 1
                continue
            row = candidate_row(label, value, pg, evidence)
            if row is None:
                rejected += 1
                continue
            row['標準欄位'] = normalize_field(field, field_id, variant)
            row['確認狀態'] = 'AI候選－待人工確認'
            rows.append(row)
        except (KeyError, TypeError, ValueError):
            rejected += 1
    if not rows:
        raise ValueError('AI 沒有回傳可驗證來源的規格，請使用規則式模式或檢查 PDF 文字層。')
    return pd.DataFrame(rows, columns=COLUMNS), rejected


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
        key = pairing_key(a.get('標準欄位') or a.get('原始欄位'))
        matching = [j for j in sorted(available) if key[0] and key == pairing_key(
            right_rows[j].get('標準欄位') or right_rows[j].get('原始欄位'))]
        j = matching[0] if matching else None
        b = right_rows[j] if j is not None else {}
        if j is not None:
            available.remove(j)
        a_value, b_value = clean(a.get('標準數值')), clean(b.get('標準數值'))
        field = clean(a.get('標準欄位') or a.get('原始欄位'))
        state, note = interpret_pair(field, a_value, b_value)
        if j is None and key[1]:
            note = '未找到相同適用條件的對照數值；其他容量/版本可能有資料，請人工核對'
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
                       '人工判讀': '', '備註': '未擷取到相同欄位及適用條件；不代表產品未提供此規格'})
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
        field = normalize_field(clean(row.get('標準欄位') or row.get('原始欄位')))
        field = canonical_field(field, row.get('標準數值'))
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
             ('標準化方式','AI 固定欄位 ID、同義詞映射與人工修正','未知欄位仍保留原名稱'),
             ('比較方式','標準化欄位與適用條件配對、保守規則判讀','不會自動評估產品優劣'),
             ('資料不足','表示至少一方未擷取到可用值','不代表產品或 PDF 不存在該規格'),
             ('文字一致','只表示標準數值文字相同','不代表測試條件一致'),
             ('部分一致','有部分共同特性','不代表完全相容或規格相同'),
             ('耐寫總量 TBW','依容量、測試條件判讀','不得直接比較不同容量的數字'),
             ('人工確認','編輯器可標註已確認','報告不會自行驗證人工修正的正確性'),
             ('產品定位','不同產品類型不可直接判斷優劣','建議選擇相同用途、容量與介面的產品')]
    for i,entry in enumerate(entries,4): body(sources,i,entry)
    setup(sources,[26,68,78],freeze='A4')
    out=io.BytesIO(); wb.save(out); return out.getvalue()

with st.expander('🔍 Gemini API 連線診斷（不需上傳 PDF）', expanded=False):
    st.caption('此測試會傳送一段簡短文字到 Google Gemini，不會傳送 PDF，也不會顯示金鑰。')
    diagnostic_model = st.text_input('測試模型', value='gemini-3.8-flash', key='diagnostic_model')
    if st.button('測試 Gemini API 金鑰', key='test_gemini_key'):
        diagnostic_key = gemini_key()
        if diagnostic_key:
            with st.spinner('正在測試 Gemini API…'):
                ok, result = gemini_connection_test(diagnostic_key, diagnostic_model)
            if ok:
                st.success(result)
            else:
                st.error(result)

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
st.subheader('AI 輔助擷取（選用）')
mode = st.radio('規格擷取方式', ['規則式（免費、資料不離開網站處理環境）', 'Gemini AI（將 PDF 擷取文字傳送至 Google）'], horizontal=True)
if mode.startswith('Gemini'):
    st.warning('AI 模式會將兩份 PDF 擷取出的文字傳送至 Google Gemini API。請勿上傳機密文件；使用前請確認有權將內容提供第三方處理。')
    consent = st.checkbox('我確認兩份文件可傳送至 Google Gemini API 進行分析。')
    model_name = st.text_input('Gemini 模型名稱（可依 AI Studio 可用模型調整）', value='gemini-3.8-flash')
    if st.button('✨ 執行 Gemini AI 擷取', disabled=not consent, type='primary'):
        key = gemini_key()
        if not key:
            st.error('找不到 GEMINI_API_KEY。請至 Streamlit App Settings → Secrets 設定。')
        elif not re.fullmatch(r'[a-zA-Z0-9._-]+', model_name):
            st.error('模型名稱格式不正確。')
        else:
            with st.spinner('AI 正在分析兩份規格書，請稍候…'):
                try:
                    ai_a, reject_a = ai_extract(pages_a, key, model_name)
                    ai_b, reject_b = ai_extract(pages_b, key, model_name)
                    for prefix, df in [('A', ai_a), ('B', ai_b)]:
                        st.session_state[f'{prefix}_table'] = df
                        st.session_state[f'{prefix}_revision'] += 1
                    st.success(f'AI 擷取完成：A {len(ai_a)} 筆、B {len(ai_b)} 筆；因來源無法驗證而排除 A {reject_a}、B {reject_b} 筆。')
                except Exception as exc:
                    st.error(f'AI 擷取失敗：{exc}；原有規格表保留，可使用規則式模式。')
    if st.button('↩ 恢復規則式擷取結果'):
        for prefix, pages in [('A', pages_a), ('B', pages_b)]:
            st.session_state[f'{prefix}_table'] = extract_dynamic(pages)
            st.session_state[f'{prefix}_revision'] += 1
        st.success('已恢復規則式擷取結果。')

st.success(f'產品 A：{len(pages_a)} 頁、{len(st.session_state.A_table)} 筆候選規格；'
           f'產品 B：{len(pages_b)} 頁、{len(st.session_state.B_table)} 筆候選規格。')
st.caption('v0.8：AI 使用固定規格 ID，程式進行同義詞標準化及適用條件配對；不同容量/版本不強行合併。請人工核對。')

st.subheader('2｜檢視與修正雙方規格')
left_col, right_col = st.columns(2)
with left_col:
    st.markdown('**產品 A 規格**')
    edited_a = st.data_editor(st.session_state.A_table, use_container_width=True,
                              hide_index=True, num_rows='dynamic',
                              disabled=['來源頁碼', '來源原文'],
                              key=f'A_editor_{st.session_state.A_revision}',
                              column_config={'確認狀態': st.column_config.SelectboxColumn(
                                  options=['待人工確認', 'AI候選－待人工確認', '已確認', '需修正', '不適用'])})
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
                                  options=['待人工確認', 'AI候選－待人工確認', '已確認', '需修正', '不適用'])})
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
                   file_name='industrial_competitor_report_v08.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
