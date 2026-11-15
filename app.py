import streamlit as st
import html
import json
import os
import re
import pandas as pd
import plotly.express as px
from io import BytesIO
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Page Config
st.set_page_config(
    page_title="Ethiopian Cassation Precedents Intelligence (Vol 1–28 + Collections)",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS: "Ethiopian Cassation Precedents" design system, version 2
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+Ethiopic:wght@500;600;700&family=Noto+Sans+Ethiopic:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap');

    :root {
        --ink: #0b1f33; --ink-muted: #5b6472; --ink-on-dark-muted: #c9d2dc;
        --ground: #f4f5f7; --surface: #ffffff; --surface-sunken: #eef1f4;
        --rule: #dfe3e8; --control-border: #c3cad3;
        --accent: #0b6b4f; --accent-tint: #e6f2ed;
        --plenary-fg: #8a5a00; --plenary-bg: #fbf3db; --plenary-border: #ecd9a6;
        --overruled-fg: #b42318; --overruled-bg: #fdecea; --overruled-border: #f3c2bc;
        --unconfirmed-fg: #9a4a00; --unconfirmed-bg: #fff1e5; --unconfirmed-border: #f2cfae;
        --font-serif: 'Noto Serif Ethiopic', Georgia, serif;
        --font-sans: 'Inter', 'Noto Sans Ethiopic', sans-serif;
    }

    html, body, [class*="css"], .stApp, .stMarkdown, button, input, textarea, select {
        font-family: var(--font-sans);
        color: var(--ink);
    }
    .stApp { background: var(--ground); }
    h1, h2, h3, h4, h5 { font-family: var(--font-serif) !important; color: var(--ink); }

    section[data-testid="stSidebar"] { background: var(--surface); border-right: 1px solid var(--rule); }

    .stTabs [data-baseweb="tab-list"] { gap: 0.2rem; border-bottom: 1px solid var(--rule); width: 100%; }
    .stTabs [data-baseweb="tab"] { padding: 0.45rem 0.65rem !important; font-size: 0.86rem !important; border-radius: 6px 6px 0 0; color: var(--ink-muted); white-space: nowrap; }
    .stTabs [aria-selected="true"] { color: var(--accent) !important; font-weight: 700; }
    .stTabs [data-baseweb="tab-highlight"] { background-color: var(--accent); }

    .stButton > button, .stDownloadButton > button {
        border-radius: 6px; min-height: 44px; font-weight: 600;
        border: 1px solid var(--control-border); background: var(--surface); color: var(--ink);
    }
    .stDownloadButton > button, .stButton > button[kind="primary"] {
        background: var(--accent); border-color: var(--accent); color: #ffffff;
    }
    .stButton > button:hover { border-color: var(--accent); color: var(--accent); }
    .stDownloadButton > button:hover { background: #095a43; color: #ffffff; }

    div[data-testid="stMetric"] {
        background: var(--surface); border: 1px solid var(--rule); border-radius: 8px; padding: 0.9rem 1rem;
    }
    div[data-testid="stMetricLabel"] { color: var(--ink-muted); }

    .main-header {
        background: var(--ink); color: #ffffff;
        padding: 1.25rem 1.5rem; border-radius: 8px; margin-bottom: 1.25rem;
    }
    .main-header h1 { font-family: var(--font-serif) !important; color: #ffffff !important; margin: 0; font-size: 1.6rem; font-weight: 700; }
    .main-header p { margin: 0.35rem 0 0 0; font-size: 0.85rem; color: var(--ink-on-dark-muted); }
    .main-header .eyebrow { font-size: 0.7rem; letter-spacing: 0.08em; color: var(--ink-on-dark-muted); margin: 0 0 0.35rem 0; }

    .overruled-banner, .plenary-banner, .unconfirmed-banner {
        border-radius: 8px; padding: 0.85rem 1rem; margin-bottom: 0.75rem;
        font-size: 0.875rem; line-height: 1.55; font-weight: 500;
    }
    .overruled-banner strong, .plenary-banner strong, .unconfirmed-banner strong { font-weight: 700; }
    .overruled-banner { background: var(--overruled-bg); color: var(--overruled-fg); border: 1px solid var(--overruled-border); }
    .plenary-banner { background: var(--plenary-bg); color: var(--plenary-fg); border: 1px solid var(--plenary-border); }
    .unconfirmed-banner { background: var(--unconfirmed-bg); color: var(--unconfirmed-fg); border: 1px solid var(--unconfirmed-border); }

    blockquote.holding {
        margin: 0 0 1rem 0; padding: 1rem 1.25rem; border: 1px solid var(--rule); border-radius: 8px;
        background: var(--surface); font-family: var(--font-serif); font-size: 1.05rem; line-height: 1.85; color: var(--ink);
    }

    .proc-card {
        background: var(--surface); border: 1px solid var(--rule); border-radius: 8px;
        padding: 1.25rem 1.5rem; margin: 1rem 0;
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;
    }
    .proc-card h3 { margin: 0; font-size: 1.25rem; }
    .proc-card .title { margin: 0.3rem 0 0.5rem 0; font-size: 0.9rem; color: var(--ink-muted); }
    .proc-card .figure { font-size: 1.5rem; font-weight: 700; font-variant-numeric: tabular-nums; }
    .proc-card .caption { margin: 0; font-size: 0.75rem; color: var(--ink-muted); }

    .chip {
        display: inline-block; font-size: 0.75rem; font-weight: 600; padding: 0.15rem 0.5rem;
        border-radius: 4px; background: var(--surface-sunken); color: var(--ink);
    }

    .law-card {
        background: var(--surface); border: 1px solid var(--rule); border-radius: 8px;
        padding: 1rem 1.1rem; margin-bottom: 0.6rem;
    }
    .law-card .head { display: flex; justify-content: space-between; align-items: baseline; flex-wrap: wrap; gap: 0.5rem; }
    .law-card h4 { margin: 0; font-family: var(--font-sans) !important; font-size: 1rem; font-weight: 700; }
    .law-card .quote {
        margin: 0.6rem 0 0 0; padding: 0.6rem 0.8rem; background: var(--ground); border-radius: 6px;
        font-family: var(--font-serif); font-size: 0.85rem; line-height: 1.7; color: var(--ink);
    }
    .law-card .quote strong { display: block; font-family: var(--font-sans); font-size: 0.75rem; color: var(--ink-muted); margin-bottom: 0.2rem; }

    mark { background: #fbf3db; color: inherit; padding: 0 0.1em; border-radius: 2px; }
    .issue-box {
        margin: 0 0 1rem 0; padding: 0.85rem 1rem; border-radius: 8px;
        background: var(--surface-sunken); font-size: 0.95rem; line-height: 1.75; color: var(--ink);
    }
    .filter-bar {
        background: var(--accent-tint); color: var(--accent); border: 1px solid #b9dbcc;
        border-radius: 8px; padding: 0.7rem 1rem; font-weight: 600; font-size: 0.9rem; margin-bottom: 0.5rem;
    }
    [class*="st-key-res_"] button {
        justify-content: flex-start; text-align: left; height: auto; min-height: 44px;
        padding: 0.7rem 0.9rem; font-weight: 400; border-radius: 8px; border: 1px solid var(--rule); background: var(--surface);
    }
    [class*="st-key-res_"] button > div, [class*="st-key-res_"] button div[data-testid="stMarkdownContainer"] { width: 100%; justify-content: flex-start; text-align: left; }
    [class*="st-key-res_"] button p { text-align: left; font-size: 0.86rem; line-height: 1.6; }
    [class*="st-key-res_"] button:hover { border-color: var(--accent); color: var(--ink); }
    .st-key-res_sel button { border: 1.5px solid var(--accent); background: var(--accent-tint); }
    .st-key-mainnav div[role="radiogroup"] { gap: 0.1rem; border-bottom: 1px solid var(--rule); flex-wrap: nowrap; overflow-x: auto; margin-bottom: 0.5rem; }
    .st-key-mainnav div[data-testid="stRadio"], .st-key-mainnav div[data-testid="stRadio"] > div { width: 100%; max-width: 100%; }
    .st-key-mainnav div[role="radiogroup"] label p { white-space: nowrap; }
    .st-key-mainnav div[role="radiogroup"] label {
        padding: 0.5rem 0.5rem; margin: 0; border-radius: 6px 6px 0 0; cursor: pointer;
        border-bottom: 2px solid transparent; background: transparent;
    }
    .st-key-mainnav div[role="radiogroup"] label > div > div:first-child { display: none; }
    .st-key-mainnav div[role="radiogroup"] label p { color: var(--ink-muted); font-size: 0.85rem; }
    .st-key-mainnav div[role="radiogroup"] label:has(input:checked) { border-bottom-color: var(--accent); }
    .st-key-mainnav div[role="radiogroup"] label:has(input:checked) p { color: var(--accent); font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# Ge'ez character normalizer
def normalize_amharic(text):
    if not text: return ""
    s = str(text).lower()
    s = re.sub(r'[ሐኀኃኻ]', 'ሀ', s)
    s = re.sub(r'[ሑኁሗ]', 'ሁ', s)
    s = re.sub(r'[ሒኂ]', 'ሂ', s)
    s = re.sub(r'[ሓኃ]', 'ሃ', s)
    s = re.sub(r'[ሔኄ]', 'ሄ', s)
    s = re.sub(r'[ሕኅ]', 'ህ', s)
    s = re.sub(r'[ሖኆ]', 'ሆ', s)
    s = re.sub(r'ሠ', 'ሰ', s).replace('ሡ', 'ሱ').replace('ሢ', 'ሲ').replace('ሣ', 'ሳ').replace('ሤ', 'ሴ').replace('ሥ', 'ስ').replace('ሦ', 'ሶ')
    s = re.sub(r'ዐ', 'አ', s).replace('ዑ', 'ኡ').replace('ዒ', 'ኢ').replace('ዓ', 'ኣ').replace('ዔ', 'ኤ').replace('ዕ', 'እ').replace('ዖ', 'ኦ')
    s = re.sub(r'ፀ', 'ጸ', s).replace('ፁ', 'ጹ').replace('ፂ', 'ጺ').replace('ፃ', 'ጻ').replace('ፄ', 'ጼ').replace('ፅ', 'ጽ').replace('ፆ', 'ጾ')
    s = s.replace('ዉሌ', 'ውል').replace('ሔግ', 'ህግ').replace('ሥርዒት', 'ሥርዓት')
    return s

# Load Database
@st.cache_data
def load_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "all_cassation_cases.json"),
        os.path.join(os.path.dirname(base_dir), "all_cassation_cases.json"),
        r"C:\Users\mekdem.belayneh\.gemini\antigravity\scratch\ethiopian_laws\cassation\all_cassation_cases.json",
        "all_cassation_cases.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    raise FileNotFoundError("Could not locate all_cassation_cases.json")

cases_data = load_data()

# Load Statutory Citator Graph
@st.cache_data
def load_citator_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "statutory_citator_compact.json"),
        os.path.join(r"C:\Users\mekdem.belayneh\.gemini\antigravity\scratch\ethiopian_laws", "statutory_citator_compact.json"),
        "statutory_citator_compact.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    return {}

# Load Reverse Citator Graph (Subordinate -> Enabling Statute & Article)
@st.cache_data
def load_reverse_citator_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "reverse_citator_compact.json"),
        os.path.join(r"C:\Users\mekdem.belayneh\.gemini\antigravity\scratch\ethiopian_laws\cassation", "reverse_citator_compact.json"),
        "reverse_citator_compact.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    return []

# Load Laws Catalog (6,192 Ministry of Justice Enactments across 15 Subjects)
@st.cache_data
def load_laws_catalog():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "laws_catalog_compact.json"),
        os.path.join(r"C:\Users\mekdem.belayneh\.gemini\antigravity\scratch\ethiopian_laws\cassation", "laws_catalog_compact.json"),
        "laws_catalog_compact.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    return []

# ==========================================
# LINKS BETWEEN DECISIONS AND LAWS
# ==========================================
# Each decision's "provisions" text is read for the laws it cites:
#   - the five codes, by their usual abbreviations (e.g. የፍ/ብ/ሕ/ቁ. 1179)
#   - proclamations by number and year (e.g. አዋጅ ቁጥር 377/96 አንቀጽ 26)
# and an index is built: provision -> the decisions that cite it.
CODE_DEFS = [
    # key, label shown, highest article number, pattern (on normalised text)
    ('civpro', 'ፍ/ብ/ሥ/ሥ/ሕ/ቁ', 483,
     r'ፍ/?ብ?/?ስ/?ስ/?ህ(?:ግ)?/?\s*(?:ቁጥር|ቁ)|ፍት[ሀሃ]\s*ብ[ሄሀሃ]ር\s*ስነ[\s\-]*ስር[አኣ]ት\s*ህግ\s*(?:ቁጥር|ቁ|አንቀጽ)'),
    ('crimpro', 'ወ/መ/ሥ/ሥ/ሕ/ቁ', 224,
     r'ወ/?መ?/?ስ/?ስ/?ህ(?:ግ)?/?\s*(?:ቁጥር|ቁ)'),
    ('civil', 'ፍ/ብ/ሕ/ቁ', 3367,
     r'ፍ/?ብ?/?ህ(?:ግ)?/?\s*(?:ቁጥር|ቁ)|ፍት[ሀሃ]\s*ብ[ሄሀሃ]ር\s*ህግ\s*(?:ቁጥር|ቁ|አንቀጽ)'),
    ('crim', 'ወ/ሕ/ቁ', 865,
     r'ወ/?ህ(?:ግ)?/?\s*(?:ቁጥር|ቁ)|ወንጀል\s*ህግ\s*(?:ቁጥር|ቁ|አንቀጽ)'),
    ('comm', 'ን/ሕ/ቁ', 1175,
     r'ን/?ህ(?:ግ)?/?\s*(?:ቁጥር|ቁ)|ንግድ\s*ህግ\s*(?:ቁጥር|ቁ|አንቀጽ)'),
    ('family', 'የቤተሰብ ሕግ አንቀጽ', 323,
     r'ቤተሰብ\s*ህግ\s*(?:ቁጥር|ቁ|አንቀጽ)'),
    ('const', 'ሕገ መንግሥት አንቀጽ', 106,
     r'ህገ[\s\-]*መንግስት\s*አንቀጽ'),
]
CODE_LABEL = {k: lbl for k, lbl, _, _ in CODE_DEFS}
CODE_MAX = {k: mx for k, _, mx, _ in CODE_DEFS}
_MARKER_RE = re.compile(
    '|'.join(f'(?P<{k}>{pat})' for k, _, _, pat in CODE_DEFS)
    + r'|(?P<proc>አዋጅ\s*(?:ቁጥር|ቁ)?\s*\.?\s*(?P<pno>\d{1,4})\s*/\s*(?P<pyr>\d{2,4}))'
)
_ARTLIST_RE = re.compile(r'^[\s\.\:፡፦]*((?:\d{1,4}(?:\s*[\(/][^\)/\s]{1,5}[\)/])*\s*(?:[፣,፤;\-–—]|እና|እስከ)?\s*)+)')
_PROC_ART_RE = re.compile(r'^[\s፣,፤;]*አንቀ[ጽጸ]\s*')

def _article_numbers(list_text, max_no):
    s = re.sub(r'[\(/][^\)/\s]{1,5}[\)/]', ' ', list_text)   # drop sub-articles: (1), /2/
    nums = []
    for a, b in re.findall(r'(\d{1,4})\s*[\-–—]\s*(\d{1,4})', s):  # ranges such as 284-287
        a, b = int(a), int(b)
        if 0 < b - a <= 20:
            nums.extend(range(a, b + 1))
    nums.extend(int(n) for n in re.findall(r'\d{1,4}', s))
    out = []
    for n in nums:
        if 1 <= n <= max_no and n not in out:
            out.append(n)
    return out

def parse_provisions(text):
    """Return the provision keys cited in a provisions text, in order.
    Keys: 'civil:1179' (code and article), 'proc:377/1996' (proclamation),
    'proc:377/1996:26' (proclamation and article)."""
    norm = normalize_amharic(text)
    marks = list(_MARKER_RE.finditer(norm))
    keys = []
    for i, m in enumerate(marks):
        tail = norm[m.end(): marks[i + 1].start() if i + 1 < len(marks) else len(norm)]
        if m.group('proc'):
            yr = m.group('pyr')
            yr = '19' + yr if len(yr) == 2 else yr
            base = f"proc:{int(m.group('pno'))}/{yr}"
            if base not in keys:
                keys.append(base)
            am = _PROC_ART_RE.match(tail)
            if am:
                lm = _ARTLIST_RE.match(tail[am.end():])
                if lm:
                    for n in _article_numbers(lm.group(1), 2000):
                        k = f"{base}:{n}"
                        if k not in keys:
                            keys.append(k)
        else:
            code = m.lastgroup
            lm = _ARTLIST_RE.match(tail)
            if lm:
                for n in _article_numbers(lm.group(1), CODE_MAX[code]):
                    k = f"{code}:{n}"
                    if k not in keys:
                        keys.append(k)
    return keys

def provision_label(key):
    parts = key.split(':')
    if parts[0] == 'proc':
        return f"አዋጅ {parts[1]}" + (f" አንቀጽ {parts[2]}" if len(parts) > 2 else "")
    return f"{CODE_LABEL[parts[0]]} {parts[1]}"

@st.cache_data
def build_link_index(_cases):
    """provision key -> positions of the decisions citing it; and each decision's keys."""
    index, per_case = {}, []
    for pos, c in enumerate(_cases):
        keys = parse_provisions(c.get('provisions') or '')
        per_case.append(keys)
        for k in keys:
            index.setdefault(k, []).append(pos)
    return index, per_case

@st.cache_data
def build_overrule_links(_cases):
    """For each 7-Judges ruling that changed an earlier interpretation: the file numbers it changed.
    Returns (changed_by: file no. -> positions of the rulings that changed it,
             changes: position of ruling -> list of file numbers it changed)."""
    by_no = {}
    for pos, c in enumerate(_cases):
        if c.get('case_no'):
            by_no.setdefault(str(c['case_no']), []).append(pos)
    changed_by, changes = {}, {}
    for pos, c in enumerate(_cases):
        if c.get('status') != 'PLENARY_OVERRULING':
            continue
        src = str(c.get('overruled_precedents') or '')
        holding = str(c.get('holding') or '')
        if '[የቀድሞ ውሳኔ ተለውጧል]' in holding:
            src += ' ' + holding.split('[አዲስ አስገዳጅ ትርጓሜ]')[0]
        nos = []
        for n in re.findall(r'\d{5,6}', src):
            if n != str(c.get('case_no')) and n not in nos:
                nos.append(n)
        changes[pos] = nos
        for n in nos:
            changed_by.setdefault(n, []).append(pos)
    return changed_by, changes, by_no

def highlight_html(text, terms):
    """Escape text for HTML and mark the searched words, matching the same
    spelling variants the search itself accepts."""
    raw = str(text or '')
    norm = normalize_amharic(raw)
    spans = []
    if terms and len(norm) == len(raw):
        for t in terms:
            start = 0
            while t and (i := norm.find(t, start)) != -1:
                spans.append((i, i + len(t)))
                start = i + len(t)
    if not spans:
        return html.escape(raw).replace('\n', '<br>')
    spans.sort()
    merged = [list(spans[0])]
    for a, b in spans[1:]:
        if a <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    out, last = [], 0
    for a, b in merged:
        out.append(html.escape(raw[last:a]))
        out.append('<mark>' + html.escape(raw[a:b]) + '</mark>')
        last = b
    out.append(html.escape(raw[last:]))
    return ''.join(out).replace('\n', '<br>')

# Session state initialization
if 'selected_for_draft' not in st.session_state:
    st.session_state.selected_for_draft = []
if 'active_case_idx' not in st.session_state:
    st.session_state.active_case_idx = 0
if 'focus_pos' not in st.session_state:
    st.session_state.focus_pos = None      # a decision opened through a link
if 'prov_filter' not in st.session_state:
    st.session_state.prov_filter = None    # "decisions citing this provision" filter

TAB_NAMES = [
    "የሰበር ውሳኔዎች (Precedents)",
    "አቤቱታ አዘጋጅ (Drafter)",
    "የተሻሩ ውሳኔዎች (Overruled)",
    "የሕግ ትንተና (Analytics)",
    "የሥልጣን ትስስር (Citator)",
    "የሕጎች ማውጫ (Library)"
]
if st.session_state.get('nav') not in TAB_NAMES:
    st.session_state.nav = TAB_NAMES[0]

def open_case(pos):
    """Open one decision in the Precedents screen."""
    st.session_state.focus_pos = pos
    st.session_state.nav = TAB_NAMES[0]

def show_citing(key):
    """Show every decision that cites a provision, in the Precedents screen."""
    st.session_state.prov_filter = key
    st.session_state.focus_pos = None
    st.session_state.search_q = ""          # so the list shows every decision citing it
    st.session_state.nav = TAB_NAMES[0]

def clear_citing():
    st.session_state.prov_filter = None

def clear_focus():
    st.session_state.focus_pos = None

# App Header
st.markdown("""
<div class="main-header">
    <p class="eyebrow">የፌዴራል ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት · FEDERAL SUPREME COURT CASSATION DIVISION</p>
    <h1>የኢትዮጵያ ሰበር አስገዳጅ የሕግ ትርጓሜዎች መፈለጊያ እና ረዳት</h1>
    <p>ቅጽ 1 – 28፣ 7 ዳኞች ሰበር እና የ2015–2016 ስብስብ · 3,246 አስገዳጅ ውሳኔዎች</p>
</div>
""", unsafe_allow_html=True)

# Top KPI Metrics
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
total_cases = len(cases_data)
plenary_cases = len([c for c in cases_data if "7 ዳኞች" in str(c.get('bench', '')) or c.get('status') in ['PLENARY', 'PLENARY_OVERRULING']])
overruled_cases = len([c for c in cases_data if c.get('status') == 'OVERRULED'])
verified_cases = len([c for c in cases_data if "Verified" in str(c.get('verification_status', ''))])
selected_count = len(st.session_state.selected_for_draft)

kpi1.metric("📚 ጠቅላላ ውሳኔዎች", f"{total_cases:,}")
kpi2.metric("⭐ 7 ዳኞች ሰበር (ምልዓተ-ጉባዔ)", plenary_cases)
kpi3.metric("⚠️ የተሻሩ የቀድሞ ውሳኔዎች", overruled_cases)
kpi4.metric("✅ በህትመቱ የተረጋገጡ", f"{verified_cases:,}")
kpi5.metric("📝 የተመረጡ ውሳኔዎች", selected_count)

# Sidebar Filters
st.sidebar.header("🔍 የውሳኔዎች ፍተሻ እና አጣሪ")

search_query = st.sidebar.text_input(
    "ሁሉን-አቀፍ ፍለጋ (Universal Search)",
    key="search_q",
    placeholder="መዝገብ ቁጥር፣ ተከራካሪ፣ አንቀጽ ወይም ጭብጥ...",
    help="በመዝገብ ቁጥር (e.g. 18180, 242314), በአንቀጽ (1845, 1000) ወይም በሕግ ቃል ይፈልጉ"
)

# Subjects filter
all_subjects = sorted(list(set(c.get('legal_subject', 'ልዩ ልዩ') for c in cases_data)))
selected_subjects = st.sidebar.multiselect(
    "የሕግ ዘርፍ (Subject)",
    options=all_subjects,
    default=[]
)

# Volume filter
def sort_vol_key(vol):
    m = re.search(r'\d+', vol)
    return int(m.group()) if m else 999

all_volumes = sorted(list(set(c.get('volume', 'ሌሎች') for c in cases_data)), key=sort_vol_key)
selected_volumes = st.sidebar.multiselect(
    "ቅጽ / ምንጭ (Volume)",
    options=all_volumes,
    default=[]
)

# Status filter
status_choice = st.sidebar.radio(
    "የአስገዳጅነት ደረጃ (Status)",
    options=["ሁሉንም ውሳኔዎች (All)", "✅ ጸንቶ የሚሰራ ብቻ (Valid)", "⭐ 7 ዳኞች ሰበር ብቻ (Plenary Bench)", "⚠️ የተሻሩ ውሳኔዎች ብቻ (Overruled Decisions)", "⚖️ የቀድሞ ውሳኔን የሻሩ ውሳኔዎች"]
)

# Verification filter
verify_choice = st.sidebar.selectbox(
    "የህትመት ማረጋገጫ (Verification)",
    options=["ሁሉንም (All)", "✅ በህትመቱ የተረጋገጠ (Verified)", "📘 ከአዲስ ቅጽ የታከለ (Added)", "⚠️ ማጣራት የሚፈልግ (Check Required)"]
)

# Sidebar: Pleading Document Settings
st.sidebar.markdown("---")
st.sidebar.subheader("📝 የፍርድ ቤት አቤቱታ አዘጋጅ (Pleading Settings)")
court_target = st.sidebar.text_input("የፍርድ ቤቱ ስም", value="ለኢ.ፌ.ዴ.ሪ ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት")
pleading_type = st.sidebar.selectbox("የአቤቱታው ዓይነት", ["የሰበር አቤቱታ", "የመከላከያ መልስ", "የይግባኝ አቤቱታ", "የሕግ ማስታወሻ"])
client_app = st.sidebar.text_input("አመልካች (Plaintiff/Petitioner)", value="[የአመልካች ስም]")
client_resp = st.sidebar.text_input("ተጠሪ (Defendant/Respondent)", value="[የተጠሪ ስም]")

# Filter Data
norm_query = normalize_amharic(search_query)
terms = norm_query.split() if norm_query else []

def filter_case(c):
    if selected_subjects and c.get('legal_subject') not in selected_subjects:
        return False
    if selected_volumes and c.get('volume') not in selected_volumes:
        return False
        
    if status_choice == "✅ ጸንቶ የሚሰራ ብቻ (Valid)":
        if c.get('status') != 'VALID':
            return False
    elif status_choice == "⭐ 7 ዳኞች ሰበር ብቻ (Plenary Bench)":
        if "7 ዳኞች" not in str(c.get('bench', '')) and c.get('status') not in ['PLENARY', 'PLENARY_OVERRULING']:
            return False
    elif status_choice == "⚠️ የተሻሩ ውሳኔዎች ብቻ (Overruled Decisions)":
        if c.get('status') != 'OVERRULED':
            return False
    elif status_choice == "⚖️ የቀድሞ ውሳኔን የሻሩ ውሳኔዎች":
        if c.get('status') != 'PLENARY_OVERRULING':
            return False

    if verify_choice == "✅ በህትመቱ የተረጋገጠ (Verified)":
        if "Verified" not in str(c.get('verification_status', '')):
            return False
    elif verify_choice == "📘 ከአዲስ ቅጽ የታከለ (Added)":
        if "Added" not in str(c.get('verification_status', '')):
            return False
    elif verify_choice == "⚠️ ማጣራት የሚፈልግ (Check Required)":
        if "Not confirmed" not in str(c.get('verification_status', '')):
            return False
            
    if terms:
        searchable = normalize_amharic(" ".join([
            str(c.get('case_no', '')),
            str(c.get('applicant', '')),
            str(c.get('respondent', '')),
            str(c.get('legal_subject', '')),
            str(c.get('volume', '')),
            str(c.get('legal_issue', '')),
            str(c.get('holding', '')),
            str(c.get('provisions', '')),
            str(c.get('date', '')),
            str(c.get('notes', '')),
            str(c.get('source', ''))
        ]))
        return all(t in searchable for t in terms)
        
    return True

link_index, case_links = build_link_index(cases_data)
changed_by, changes_made, cases_by_no = build_overrule_links(cases_data)

_citing = set(link_index.get(st.session_state.prov_filter, [])) if st.session_state.prov_filter else None
filtered_pos = [i for i, c in enumerate(cases_data) if (_citing is None or i in _citing) and filter_case(c)]
# a searched number that matches exactly (e.g. Art. 1179, not File 117953) is listed first
_num_terms = [t for t in terms if t.isdigit()]
if _num_terms:
    def _exact_hits(i):
        c = cases_data[i]
        txt = " ".join(str(c.get(f, '')) for f in ('case_no', 'provisions', 'legal_issue', 'holding'))
        return sum(1 for t in _num_terms if re.search(r'(?<!\d)' + t + r'(?!\d)', txt))
    filtered_pos.sort(key=_exact_hits, reverse=True)
filtered_cases = [cases_data[i] for i in filtered_pos]

# Tabs Interface
with st.container(key="mainnav"):
    nav = st.radio("Screen", TAB_NAMES, horizontal=True, key="nav", label_visibility="collapsed")

# ==========================================
# TAB 1: SPLIT SCREEN RESEARCH & INSPECTOR
# ==========================================
if nav == TAB_NAMES[0]:
    if st.session_state.prov_filter:
        _fb1, _fb2 = st.columns([8, 2])
        with _fb1:
            st.markdown(f'<div class="filter-bar">ማጣሪያ (Filter): {html.escape(provision_label(st.session_state.prov_filter))} የሚጠቅሱ ውሳኔዎች (decisions citing this provision)</div>', unsafe_allow_html=True)
        with _fb2:
            st.button("ማጣሪያውን አስወግድ (Clear)", on_click=clear_citing, key="clear_citing", use_container_width=True)
    st.write(f"**የተገኙ ውሳኔዎች:** `{len(filtered_cases):,}` / {total_cases:,}")
    
    col_left, col_right = st.columns([5, 7])
    
    with col_left:
        st.subheader("📋 የውሳኔዎች ዝርዝር")
        
        def _short(text, n=150):
            t = " ".join(str(text or "").split())
            t = t if len(t) <= n else t[:n].rstrip() + "…"
            return re.sub(r'([\\`*_\[\]#$~|<>])', r'\\\1', t)   # keep the text literal inside a button label

        def _status_word(c):
            if c.get('status') == 'OVERRULED':
                return " · የተሻረ"
            if "7 ዳኞች" in str(c.get('bench', '')) or c.get('status') in ['PLENARY', 'PLENARY_OVERRULING']:
                return " · 7 ዳኞች"
            return ""

        def pick_result(k):
            st.session_state.active_case_idx = k
            st.session_state.focus_pos = None

        def turn_page(step):
            st.session_state.res_page = max(0, st.session_state.get('res_page', 0) + step)

        if filtered_cases:
            # back to the first result whenever the search or filters change
            _sig = (search_query, tuple(selected_subjects), tuple(selected_volumes), status_choice, verify_choice, st.session_state.prov_filter)
            if st.session_state.get('res_sig') != _sig:
                st.session_state.res_sig = _sig
                st.session_state.res_page = 0
                st.session_state.active_case_idx = 0
            PAGE = 10
            _pages = max(1, (len(filtered_cases) + PAGE - 1) // PAGE)
            _page = min(st.session_state.get('res_page', 0), _pages - 1)
            st.caption("ውሳኔ ለመክፈት ይጫኑ (Click a decision to open it)")
            for k in range(_page * PAGE, min(len(filtered_cases), (_page + 1) * PAGE)):
                c = filtered_cases[k]
                _no = c.get('case_no') or f"[{c.get('entry_no')}]"
                _label = f"**መ/ቁ {_no}** · {c.get('volume')}{_status_word(c)} — {_short(c.get('legal_issue'))}"
                _is_sel = st.session_state.focus_pos is None and k == st.session_state.active_case_idx
                with st.container(key="res_sel" if _is_sel else f"res_{k}"):
                    st.button(_label, key=f"resbtn_{k}", on_click=pick_result, args=(k,), use_container_width=True)
            _n1, _n2, _n3 = st.columns([3, 4, 3])
            _n1.button("← የቀደመ (Previous)", on_click=turn_page, args=(-1,), disabled=_page == 0, key="res_prev", use_container_width=True)
            _n2.markdown(f"<p style='text-align:center; margin-top:0.5rem; color: var(--ink-muted);'>ገጽ {_page + 1} ከ {_pages}</p>", unsafe_allow_html=True)
            _n3.button("ቀጣይ (Next) →", on_click=turn_page, args=(1,), disabled=_page >= _pages - 1, key="res_next", use_container_width=True)
        else:
            st.warning("የተፈለገውን መስፈርት የሚያሟላ ውሳኔ አልተገኘም።")

    with col_right:
        if st.session_state.focus_pos is not None or filtered_cases:
            if st.session_state.focus_pos is not None:
                active_pos = st.session_state.focus_pos
                st.button("← ወደ ዝርዝሩ ተመለስ (Back to results)", on_click=clear_focus, key="back_to_results")
            else:
                active_pos = filtered_pos[min(st.session_state.active_case_idx, len(filtered_pos) - 1)]
            active_c = cases_data[active_pos]
            
            # Case Header Card
            case_no_disp = active_c.get('case_no') or f"የመዝገብ ቁጥር ያልተረጋገጠ [ተ.ቁ {active_c.get('entry_no')}]"
            st.subheader(f"📖 ሰበር መዝገብ ቁጥር {case_no_disp}")
            
            # Status Alerts
            if active_c.get('status') == 'OVERRULED':
                st.markdown(f"""
                <div class="overruled-banner">
                    <strong>የተሻረ ውሳኔ (OVERRULED PRECEDENT):</strong><br>
                    {active_c.get('status_note')}
                </div>
                """, unsafe_allow_html=True)
            elif active_c.get('status') == 'PLENARY_OVERRULING':
                st.markdown(f"""
                <div class="plenary-banner">
                    <strong>የሰበር ምልዓተ-ጉባዔ ታሪካዊ ውሳኔ (PLENARY OVERRULING):</strong><br>
                    {active_c.get('status_note')}
                </div>
                """, unsafe_allow_html=True)
            elif "7 ዳኞች" in str(active_c.get('bench', '')) or active_c.get('status') == 'PLENARY':
                st.markdown(f"""
                <div class="plenary-banner">
                    <strong>በ 7 ዳኞች ምልዓተ-ጉባዔ የተሰጠ አስገዳጅ ውሳኔ</strong>
                </div>
                """, unsafe_allow_html=True)

            # Links: the ruling that changed this decision / the decisions this ruling changed
            _by = [p for p in changed_by.get(str(active_c.get('case_no')), []) if p != active_pos]
            if _by:
                st.caption("ይህን ትርጉም የለወጠው የ7 ዳኞች ውሳኔ (Changed by):")
                for p in _by:
                    st.button(f"መ/ቁ {cases_data[p].get('case_no')} ክፈት (Open)", on_click=open_case, args=(p,), key=f"chg_by_{p}")
            _made = changes_made.get(active_pos, [])
            if _made:
                st.caption("ይህ ውሳኔ የለወጣቸው የቀድሞ ውሳኔዎች (Earlier decisions it changed):")
                _cols = st.columns(min(len(_made), 4))
                for j, n in enumerate(_made):
                    _targets = cases_by_no.get(n, [])
                    with _cols[j % len(_cols)]:
                        if _targets:
                            st.button(f"መ/ቁ {n} ክፈት (Open)", on_click=open_case, args=(_targets[0],), key=f"chg_made_{n}")
                        else:
                            st.caption(f"መ/ቁ {n} — በመረጃ ቋቱ ውስጥ የለም (not in the database)")

            # Verification Warning
            if "Not confirmed" in str(active_c.get('verification_status', '')):
                st.markdown("""
                <div class="unconfirmed-banner">
                    <strong>በህትመቱ አልተረጋገጠም:</strong> ይህ መዝገብ ቁጥር ወይም ቅጽ በታተመው መጽሐፍ ማውጫ ውስጥ በግልጽ አልተረጋገጠም፤ ከመጥቀስዎ በፊት ህትመቱን ይመልከቱ።
                </div>
                """, unsafe_allow_html=True)

            # Metadata grid
            m1, m2 = st.columns(2)
            with m1:
                st.write(f"**የሕግ ዘርፍ:** {active_c.get('legal_subject')}")
                st.write(f"**ቅጽ / ምንጭ:** {active_c.get('volume')}")
                st.write(f"**ገጽ:** {active_c.get('page_in_vol') or '-'}")
            with m2:
                st.write(f"**የተወሰነበት ቀን:** {active_c.get('date', '-')}")
                st.write(f"**የችሎት አደረጃጀት:** {active_c.get('bench', 'የሰበር መደበኛ ችሎት (5 ዳኞች)')}")
                st.write(f"**የህትመት ማረጋገጫ:** {active_c.get('verification_status', '-')}")

            st.write(f"**አመልካች:** {active_c.get('applicant', '-')}")
            st.write(f"**ተጠሪ:** {active_c.get('respondent', '-')}")

            # Legal Issue
            st.markdown("##### 📌 የተያዘው ጭብጥ (Legal Issue)")
            st.markdown('<div class="issue-box">' + highlight_html(active_c.get('legal_issue', 'የሕግ ትርጓሜ'), terms) + '</div>', unsafe_allow_html=True)

            # Ratio Decidendi
            st.markdown("##### ⚖️ አስገዳጅ የሕግ ትርጓሜ (Authoritative Holding)")
            st.markdown('<blockquote class="holding">' + highlight_html(active_c.get('holding', ''), terms) + '</blockquote>', unsafe_allow_html=True)

            # Statutory Provisions
            if active_c.get('provisions'):
                st.markdown(f"**📜 የተተረጎሙ ድንጋጌዎች:** `{active_c.get('provisions')}`")
                _keys = [k for k in case_links[active_pos] if len(link_index.get(k, [])) > 1][:12]
                if _keys:
                    st.caption("ተመሳሳይ ድንጋጌ የተረጎሙ ሌሎች ውሳኔዎች (Other decisions on the same provision):")
                    _pc = st.columns(2)
                    for j, k in enumerate(_keys):
                        with _pc[j % 2]:
                            st.button(f"{provision_label(k)} ({len(link_index[k])} ውሳኔዎች)", on_click=show_citing, args=(k,), key=f"prov_{active_pos}_{k}", use_container_width=True)

            # Provenance & Correction Notes
            if active_c.get('source') or active_c.get('notes'):
                with st.expander("ℹ️ የመረጃው ምንጭ እና የማረጋገጫ ማስታወሻ (Provenance & Notes)", expanded=False):
                    if active_c.get('source'):
                        st.write(f"**ምንጭ:** {active_c.get('source')}")
                    if active_c.get('notes'):
                        st.write(f"**ማስታወሻ:** {active_c.get('notes')}")

            # Actions
            b1, b2 = st.columns(2)
            with b1:
                page_str = f"፣ ገጽ {active_c.get('page_in_vol')}" if active_c.get('page_in_vol') else ""
                formal_cite = f"በፌዴራል ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት በመዝገብ ቁጥር {active_c.get('case_no')}፣ {active_c.get('volume')}{page_str} ({active_c.get('bench')})፣ አመልካች {active_c.get('applicant')} እና ተጠሪ {active_c.get('respondent')} በተወሰነው አስገዳጅ ውሳኔ መሠረት..."
                st.text_area("📋 የፍርድ ቤት ጥቅስ (Formal Citation):", value=formal_cite, height=75)
            with b2:
                c_no = str(active_c.get('case_no') or active_c.get('entry_no'))
                if c_no in [str(c.get('case_no') or c.get('entry_no')) for c in st.session_state.selected_for_draft]:
                    if st.button("➖ ከአቤቱታ ረቂቅ አስወግድ (Remove from Draft)"):
                        st.session_state.selected_for_draft = [c for c in st.session_state.selected_for_draft if str(c.get('case_no') or c.get('entry_no')) != c_no]
                        st.rerun()
                else:
                    if st.button("➕ ወደ አቤቱታ ረቂቅ ጨምር (Add to Draft)"):
                        st.session_state.selected_for_draft.append(active_c)
                        st.rerun()

# ==========================================
# TAB 2: AUTOMATED WORD PLEADING DRAFTER
# ==========================================
if nav == TAB_NAMES[1]:
    st.subheader("📝 የፍርድ ቤት አቤቱታ እና ማጣቀሻ አዘጋጅ (Microsoft Word Drafter)")
    st.write("በምርምር ክፍል የመረጧቸውን አስገዳጅ ውሳኔዎች በማካተት የተሟላ **የፍርድ ቤት አቤቱታ የሕግ ማጣቀሻ ሰነድ (.docx)** በቀጥታ ማመንጨት ይችላሉ።")

    if not st.session_state.selected_for_draft:
        st.info("💡 በአሁኑ ሰዓት የተመረጠ ውሳኔ የለም። በትር 1 (Research) ውስጥ የሚፈልጓቸውን ውሳኔዎች መርጠው **'ወደ አቤቱታ ረቂቅ ጨምር'** የሚለውን ይጫኑ።")
    else:
        st.write(f"**የተመረጡ አስገዳጅ ውሳኔዎች ({len(st.session_state.selected_for_draft)}):**")
        
        for i, sc in enumerate(st.session_state.selected_for_draft):
            c_col1, c_col2 = st.columns([9, 1])
            c_col1.write(f"**{i+1}. መ/ቁ {sc.get('case_no') or f'[Entry {sc.get('entry_no')}]'}** ({sc.get('volume')}) — {str(sc.get('legal_issue'))[:80]}...")
            if c_col2.button("❌", key=f"del_{i}"):
                st.session_state.selected_for_draft.pop(i)
                st.rerun()

        def generate_docx():
            doc = docx.Document()
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Court Heading
            h1 = doc.add_paragraph()
            h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_court = h1.add_run(court_target)
            run_court.font.name = "Nyala"
            run_court.font.size = Pt(14)
            run_court.font.bold = True

            h2 = doc.add_paragraph()
            h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = h2.add_run("አዲስ አበባ\n")
            r2.font.name = "Nyala"
            r2.font.size = Pt(12)

            p_parties = doc.add_paragraph()
            r_app = p_parties.add_run(f"አመልካች፡- {client_app}\n")
            r_app.font.name = "Nyala"
            r_app.font.size = Pt(11)
            r_app.font.bold = True
            
            r_resp = p_parties.add_run(f"ተጠሪ፡- {client_resp}\n")
            r_resp.font.name = "Nyala"
            r_resp.font.size = Pt(11)
            r_resp.font.bold = True

            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_title = p_title.add_run(f"ጉዳዩ፡- {pleading_type} አቅርቦ ስለመከራከር")
            r_title.font.name = "Nyala"
            r_title.font.size = Pt(12)
            r_title.font.bold = True
            r_title.underline = True

            doc.add_paragraph()

            p_table_title = doc.add_paragraph()
            r_tt = p_table_title.add_run("በአዋጅ ቁጥር 454/97 አንቀጽ 2(4) እና በአዋጅ ቁጥር 1234/2013 አንቀጽ 10(2) መሠረት ለዚህ ክርክር አስገዳጅነት ያላቸው የሰበር ውሳኔዎች ዝርዝር፡-")
            r_tt.font.name = "Nyala"
            r_tt.font.size = Pt(11)
            r_tt.font.bold = True

            table = doc.add_table(rows=1, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False

            headers = ["ተ.ቁ", "የሰ/መ/ቁ እና ቅጽ", "ተከራካሪ ወገኖች", "የተያዘው ጭብጥ", "አስገዳጅ የሕግ ትርጓሜ (Ratio Decidendi)"]
            col_widths = [Inches(0.5), Inches(1.1), Inches(1.3), Inches(1.6), Inches(2.5)]

            hdr_cells = table.rows[0].cells
            for idx, text in enumerate(headers):
                hdr_cells[idx].text = text
                hdr_cells[idx].width = col_widths[idx]
                for p in hdr_cells[idx].paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = "Nyala"
                        r.font.size = Pt(10)
                        r.font.bold = True

            for idx, sc in enumerate(st.session_state.selected_for_draft):
                row_cells = table.add_row().cells
                row_cells[0].text = str(idx + 1)
                page_p = f"፣ ገጽ {sc.get('page_in_vol')}" if sc.get('page_in_vol') else ""
                row_cells[1].text = f"መ/ቁ {sc.get('case_no') or '-'}\n({sc.get('volume')}{page_p})"
                row_cells[2].text = f"{sc.get('applicant')}\nእና\n{sc.get('respondent')}"
                row_cells[3].text = str(sc.get('legal_issue'))
                
                holding_str = str(sc.get('holding'))
                if sc.get('status') == 'OVERRULED':
                    holding_str = f"[⚠️ ይህ ውሳኔ በ 7 ዳኞች ሰበር ተሽሯል!]\n" + holding_str
                row_cells[4].text = holding_str

                for c_idx in range(5):
                    row_cells[c_idx].width = col_widths[c_idx]
                    for p in row_cells[c_idx].paragraphs:
                        for r in p.runs:
                            r.font.name = "Nyala"
                            r.font.size = Pt(9.5)

            doc.add_paragraph("\n")

            p_concl = doc.add_paragraph()
            r_c = p_concl.add_run("ስለሆነም ክቡር ፍርድ ቤቱ ከላይ የተጠቀሱትን አስገዳጅ የሰበር ውሳኔዎች መሠረት በማድረግ ፍትሐዊ ውሳኔ ይሰጠን ዘንድ በአክብሮት እንጠይቃለን፡፡\n\nአመልካች፡- ______________________\nፊርማ፡- ______________________\nቀን፡- ______________________")
            r_c.font.name = "Nyala"
            r_c.font.size = Pt(11)

            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)
            return bio

        docx_bio = generate_docx()
        st.download_button(
            label="📥 ይፋዊ የፍርድ ቤት አቤቱታ ሰነድ አውርድ (.docx)",
            data=docx_bio,
            file_name=f"Ethiopian_Cassation_Brief_{pleading_type}_{client_app[:15]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ==========================================
# TAB 3: SHEPARD'S OVERRULE MAP
# ==========================================
if nav == TAB_NAMES[2]:
    st.subheader("⚖️ የተሻሩ ውሳኔዎች እና የ 7 ዳኞች አስገዳጅ ውሳኔዎች ማውጫ (Shepard's Citator)")
    st.write("በአዋጅ ቁጥር 1234/2013 አንቀጽ 10(4) እና 26(1) መሠረት በሰባት (7) ዳኞች የተሰየመው የሰበር ሰሚ ችሎት የቀድሞ ባለ 5 ዳኞች ውሳኔዎችን የመሻር እና የመለወጥ ስልጣን ተሰጥቶታል፡፡")

    ov_cases = [c for c in cases_data if c.get('status') == 'OVERRULED']
    st.markdown(f"#### ⚠️ የተሻሩ/የተለወጡ ውሳኔዎች ዝርዝር ({len(ov_cases)})")
    
    df_ov = pd.DataFrame([
        {
            "ተ.ቁ": c.get('entry_no'),
            "የተሻረው መ/ቁ": c.get('case_no'),
            "ቅጽ": c.get('volume'),
            "የሕግ ዘርፍ": c.get('legal_subject'),
            "የተሻረበት ጉዳይ": c.get('legal_issue'),
            "ማብራሪያ": c.get('status_note')
        }
        for c in ov_cases
    ])
    st.dataframe(df_ov, use_container_width=True)

    st.markdown("---")
    st.markdown("#### ⭐ የቀድሞ ውሳኔዎችን የሻሩ እና ያረጋገጡ የ 7 ዳኞች ታሪካዊ ውሳኔዎች")
    
    plen_cases = [c for c in cases_data if c.get('status') in ['PLENARY_OVERRULING', 'PLENARY']]
    df_plen = pd.DataFrame([
        {
            "ተ.ቁ": c.get('entry_no'),
            "የ 7 ዳኞች መ/ቁ": c.get('case_no'),
            "ቀን": c.get('date'),
            "ቅጽ": c.get('volume'),
            "የሕግ ዘርፍ": c.get('legal_subject'),
            "አስገዳጅ መርህ / ጭብጥ": c.get('legal_issue'),
            "የተሻሩ የቀድሞ መዛግብት": c.get('overruled_precedents') or "-"
        }
        for c in plen_cases
    ])
    st.dataframe(df_plen, use_container_width=True)

    st.info("💡 **ልዩ የሕግ ማስታወሻ:** በሰባት ዳኞች ሰበር መ/ቁ 237423 (የካቲት 29/2016 ዓ.ም) በተሰጠ ውሳኔ መ/ቁ 48850፣ 58008፣ 74237፣ 94913፣ 102640 እና 218957 የቀረበውን የመሻር ጥያቄ ውድቅ በማድረግ ውሳኔዎቹ ጸንተው እንዲቀጥሉ ወስኗል።")

# ==========================================
# TAB 4: 20-YEAR JUDICIAL ANALYTICS
# ==========================================
if nav == TAB_NAMES[3]:
    st.subheader("📊 የ 20 ዓመታት የሰበር አስገዳጅ ውሳኔዎች ትንተና (1997 - 2017 ዓ.ም)")
    
    # 1. Subject distribution
    subj_counts = pd.Series([c.get('legal_subject') for c in cases_data]).value_counts().reset_index()
    subj_counts.columns = ['የሕግ ዘርፍ', 'የውሳኔዎች ብዛት']
    
    fig_subj = px.bar(
        subj_counts,
        x='የሕግ ዘርፍ',
        y='የውሳኔዎች ብዛት',
        color='የውሳኔዎች ብዛት',
        color_continuous_scale=['#e6f2ed', '#0b6b4f'],
        title="አስገዳጅ የሰበር ውሳኔዎች በሕግ ዘርፍ (Precedents by Subject - 3,246 Cases)"
    )
    st.plotly_chart(fig_subj, use_container_width=True)
    
    # 2. Volumes Timeline
    vol_counts = pd.Series([c.get('volume') for c in cases_data]).value_counts().reset_index()
    vol_counts.columns = ['ቅጽ', 'የውሳኔዎች ብዛት']
    
    fig_vol = px.bar(
        vol_counts.head(28),
        x='ቅጽ',
        y='የውሳኔዎች ብዛት',
        title="በየቅጹ የታተሙ አስገዳጅ ውሳኔዎች ብዛት (Volumes 1 – 28 & Collections)"
    )
    st.plotly_chart(fig_vol, use_container_width=True)


# ==========================================
# TAB 5: STATUTORY CITATOR (FORWARD & REVERSE)
# ==========================================
if nav == TAB_NAMES[4]:
    st.subheader("የኢትዮጵያ ሕጎች የሥልጣን ትስስር (Statutory Citator)")
    st.write("በዋና አዋጆች እና በአስፈፃሚ መመሪያዎች/ደንቦች መካከል የተዘረጋ ባለሁለት-አቅጣጫ የሥልጣን ትስስር ማውጫ")

    citator_mode = st.radio(
        "የፍተሻ አቅጣጫ ይምረጡ (Select Direction):",
        [
            "ከአዋጅ ወደ አስፈፃሚ መመሪያዎች/ደንቦች (Statute to Subordinate Laws)",
            "ከመመሪያ/ደንብ ወደ ሥልጣን ሰጪ አዋጅ (Subordinate Law to Enabling Authority)"
        ],
        horizontal=True
    )

    citator_data = load_citator_data()

    if "Subordinate Laws" in citator_mode:
        if not citator_data:
            st.warning("የሕጎች ትስስር መረጃ ቋት አልተጫነም (statutory_citator_compact.json not found).")
        else:
            # Search controls
            c_p1, c_p2, c_p3 = st.columns([4, 3, 3])
            
            # Prepare list of proclamations
            proc_options = sorted(list(citator_data.keys()), key=lambda x: int(x) if x.isdigit() else 9999)
            
            default_idx = 0
            if "1363" in proc_options:
                default_idx = proc_options.index("1363")
            elif "1333" in proc_options:
                default_idx = proc_options.index("1333")

            with c_p1:
                selected_proc_no = st.selectbox(
                    "አዋጅ ይምረጡ (Select Proclamation No.):",
                    options=proc_options,
                    index=default_idx,
                    format_func=lambda x: f"አዋጅ ቁጥር {x} — {citator_data[x]['title'][:50]}... ({citator_data[x]['total']} የበታች ሕጎች)"
                )
                
            proc_entry = citator_data[selected_proc_no]
            available_articles = ["ሁሉንም አንቀጾች (All Articles)"] + sorted(list(proc_entry["articles"].keys()))

            with c_p2:
                selected_article = st.selectbox(
                    "የተጠቀሰውን አንቀጽ ይምረጡ (Select Article Cited):",
                    options=available_articles
                )

            with c_p3:
                law_type_filter = st.selectbox(
                    "የበታች ሕግ ዓይነት (Law Type):",
                    ["ሁሉንም (All)", "መመሪያ ብቻ (Directives Only)", "ደንብ ብቻ (Regulations Only)"]
                )

            # Proclamation Header Card
            st.markdown(f"""
            <div class="proc-card">
                <div>
                    <h3>አዋጅ ቁጥር {selected_proc_no}</h3>
                    <p class="title">{proc_entry['title']}</p>
                    <span class="chip">ዘርፍ፡ {proc_entry.get('subject', 'ጠቅላላ')}</span>
                </div>
                <div style="text-align: right;">
                    <div class="figure">{proc_entry['total']}</div>
                    <p class="caption">የተገኙ የበታች ሕጎች</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if proc_entry.get("link"):
                st.markdown(f"[የአዋጁን ሙሉ PDF በ Google Drive ይመልከቱ]({proc_entry['link']})")

            # Collect matching subordinate laws (Strict Deduplication & Article Aggregation)
            if selected_article == "ሁሉንም አንቀጾች (All Articles)":
                law_agg = {}
                for art, items in proc_entry["articles"].items():
                    for it in items:
                        t = it['title']
                        art_lbl = f"አንቀጽ {art}"
                        if t not in law_agg:
                            law_agg[t] = {'item': it, 'articles': [art_lbl]}
                        else:
                            if art_lbl not in law_agg[t]['articles']:
                                law_agg[t]['articles'].append(art_lbl)
                for it in proc_entry.get("general", []):
                    t = it['title']
                    if t not in law_agg:
                        law_agg[t] = {'item': it, 'articles': ["ጠቅላላ የሥልጣን ውክልና"]}

                matching_subordinates = [( "፣ ".join(v['articles']), v['item']) for v in law_agg.values()]
            else:
                art_clean = selected_article.replace("አንቀጽ ", "").strip()
                seen_sub_titles = set()
                matching_subordinates = []
                for it in proc_entry["articles"].get(art_clean, []):
                    if it['title'] not in seen_sub_titles:
                        seen_sub_titles.add(it['title'])
                        matching_subordinates.append((f"አንቀጽ {art_clean}", it))

            # Filter by Law Type
            if law_type_filter == "መመሪያ ብቻ (Directives Only)":
                matching_subordinates = [x for x in matching_subordinates if "Directive" in x[1].get("type", "")]
            elif law_type_filter == "ደንብ ብቻ (Regulations Only)":
                matching_subordinates = [x for x in matching_subordinates if "Regulation" in x[1].get("type", "")]

            st.markdown(f"##### የተገኙ አስፈፃሚ ሕጎች (`{len(matching_subordinates)}`)፡")

            if not matching_subordinates:
                st.info(f"በአዋጅ ቁጥር {selected_proc_no} {selected_article} መሠረት የወጣ መመሪያ ወይም ደንብ አልተገኘም። (ይህ አንቀጽ እስካሁን በአዲስ መመሪያ ያልተደነገገ ወይም በቀድሞው መመሪያ የመሸጋገሪያ ድንጋጌ የሚገዛ ሊሆን ይችላል)።")
            else:
                for idx, (art_lbl, sub) in enumerate(matching_subordinates, 1):
                    with st.container():
                        st.markdown(f"""
                        <div class="law-card">
                            <div class="head">
                                <h4>{idx}. [{sub.get('type')}] {sub.get('title')}</h4>
                                <span class="chip">{art_lbl}</span>
                            </div>
                            <p class="quote">
                                <strong>የሥልጣን መነሻ ዐረፍተ-ነገር (Quoted Authority Clause)</strong>
                                "{sub.get('snippet', '-')}"
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                        if sub.get("link"):
                            st.markdown(f"[የመመሪያውን PDF በ Drive ክፈት]({sub['link']})")
                        st.write("")

    else:
        # REVERSE CITATOR: ENABLING AUTHORITY
        reverse_data = load_reverse_citator_data()
        if not reverse_data:
            st.warning("የተገላቢጦሽ የሥልጣን ትስስር መረጃ አልተገኘም (reverse_citator_compact.json not found).")
        else:
            st.write("**የሥልጣን ውክልና መነሻ (Enabling Authority):** መመሪያዎች እና ደንቦች የተመሰረቱባቸው ዋና አዋጆች እና የተወሰኑ አንቀጾች ማውጫ")

            r_col1, r_col2 = st.columns([3, 7])
            with r_col1:
                rev_type_filter = st.selectbox("የሕግ ዓይነት (Filter Type):", ["ሁሉንም (All)", "መመሪያዎች ብቻ (Directives)", "ደንቦች ብቻ (Regulations)"])
                rev_keyword = st.text_input("በቁጥር ወይም ቃል ፈልግ (Search Subordinate Law):", placeholder="ለምሳሌ፡ 1094, 1060, ነዳጅ, ግዥ, ንግድ...")

            # Filter reverse data
            filtered_reverse = reverse_data
            if rev_type_filter == "መመሪያዎች ብቻ (Directives)":
                filtered_reverse = [x for x in filtered_reverse if x.get('category') == 'Directive']
            elif rev_type_filter == "ደንቦች ብቻ (Regulations)":
                filtered_reverse = [x for x in filtered_reverse if x.get('category') == 'Regulation']
            
            if rev_keyword:
                kw_norm = normalize_amharic(rev_keyword)
                filtered_reverse = [
                    x for x in filtered_reverse
                    if kw_norm in normalize_amharic(x.get('title', '')) or kw_norm in str(x.get('law_number', ''))
                ]

            with r_col2:
                if not filtered_reverse:
                    st.warning("ምንም መመሪያ ወይም ደንብ አልተገኘም።")
                    selected_sub_idx = None
                else:
                    selected_sub_idx = st.selectbox(
                        f"መመሪያ/ደንብ ይምረጡ ({len(filtered_reverse):,} የተገኙ):",
                        options=range(len(filtered_reverse)),
                        format_func=lambda i: f"[{filtered_reverse[i].get('category', 'Directive')[:3]}] {filtered_reverse[i].get('title', '')[:70]}..."
                    )

            if selected_sub_idx is not None and filtered_reverse:
                active_sub = filtered_reverse[selected_sub_idx]
                
                # Header card for selected directive
                st.markdown(f"""
                <div class="proc-card">
                    <div>
                        <h3>{active_sub.get('title')}</h3>
                        <p class="title">ዓይነት፡ {active_sub.get('category')} | ቁጥር፡ {active_sub.get('law_number', '-')} | ዓመት፡ {active_sub.get('year', '-')}</p>
                        <span class="chip">ዘርፍ፡ {active_sub.get('subject', 'ጠቅላላ')}</span>
                    </div>
                    <div style="text-align: right;">
                        <div class="figure">{len(active_sub.get('authorities', []))}</div>
                        <p class="caption">የተጠቀሱ ዋና አዋጆች</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if active_sub.get("link"):
                    st.markdown(f"[የዚህን የበታች ሕግ ሙሉ PDF በ Drive ክፈት]({active_sub['link']})")

                st.markdown(f"##### ይህ መመሪያ ሥልጣን ያገኘባቸው ዋና አዋጆች እና አንቀጾች ({len(active_sub.get('authorities', []))})፡")

                for a_idx, auth in enumerate(active_sub.get('authorities', []), 1):
                    with st.container():
                        is_fed = auth.get('is_federal', True)
                        art_display = auth.get('article', 'ጠቅላላ የሥልጣን ውክልና')
                        jurisdiction_badge = "የፌዴራል አዋጅ (Federal Statute)" if is_fed else "የከተማ/የክልል ሕግ (Regional/City Law)"
                        
                        st.markdown(f"""
                        <div class="law-card">
                            <div class="head">
                                <h4>{a_idx}. {auth.get('proc_title')}</h4>
                                <div>
                                    <span class="chip">{jurisdiction_badge}</span>
                                    <span class="chip">{art_display}</span>
                                </div>
                            </div>
                            <p class="quote">
                                <strong>በመመሪያው መግቢያ/ድንጋጌ የተጠቀሰ ሥልጣን ሰጪ ዐረፍተ-ነገር (Enabling Authority Clause):</strong>
                                "{auth.get('snippet', '-')}"
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                        if is_fed and auth.get("proc_link"):
                            st.markdown(f"[የዋናውን አዋጅ PDF በ Drive ይመልከቱ]({auth['proc_link']})")
                        st.write("")


# ==========================================
# TAB 6: LAWS LIBRARY BY SUBJECT (3,696 DISTINCT ENACTMENTS)
# ==========================================
if nav == TAB_NAMES[5]:
    st.subheader("የኢትዮጵያ ሕጎች ቤተ-መጻሕፍት በዘርፍ (Ministry of Justice Laws Library)")
    st.write("በፍትህ ሚኒስቴር የታተሙ 3,696 የተጣሩ አዋጆች፣ ደንቦች፣ መመሪያዎች እና ስምምነቶች በ15 የሕግ ዘርፎች የተመደቡበት ማውጫ")

    laws_catalog = load_laws_catalog()

    if not laws_catalog:
        st.warning("የሕጎች ማውጫ መረጃ አልተገኘም (laws_catalog_compact.json not found).")
    else:
        # Top KPI stats for distinct laws
        tot_laws = len(laws_catalog)
        procs_cnt = len([x for x in laws_catalog if "አዋጅ" in x.get('norm', '') or "Proclamation" in x.get('norm', '')])
        regs_cnt = len([x for x in laws_catalog if "ደንብ" in x.get('norm', '') or "Regulation" in x.get('norm', '')])
        dirs_cnt = len([x for x in laws_catalog if "መመሪያ" in x.get('norm', '') or "Directive" in x.get('norm', '')])
        treaty_cnt = len([x for x in laws_catalog if "ስምምነት" in x.get('norm', '') or "Treaty" in x.get('norm', '')])
        gen_cnt = len([x for x in laws_catalog if "ጠቅላላ ሰነድ" in x.get('norm', '') or "General Document" in x.get('norm', '')])

        l_kpi1, l_kpi2, l_kpi3, l_kpi4, l_kpi5 = st.columns(5)
        l_kpi1.metric("ጠቅላላ ሕጎች (Distinct Laws)", f"{tot_laws:,}")
        l_kpi2.metric("አዋጆች (Proclamations)", f"{procs_cnt:,}")
        l_kpi3.metric("ደንቦች (Regulations)", f"{regs_cnt:,}")
        l_kpi4.metric("መመሪያዎች (Directives)", f"{dirs_cnt:,}")
        l_kpi5.metric("ስምምነቶች እና ሰነዶች", f"{treaty_cnt + gen_cnt:,}")

        st.markdown("---")

        # Filters
        l_f1, l_f2, l_f3 = st.columns([4, 3, 3])

        # Dynamic subjects with live counts from deduplicated catalog
        subject_counts = {}
        subject_am_names = {}
        for l in laws_catalog:
            s_en = l.get('subject_en', 'Other')
            s_am = l.get('subject_am', '')
            subject_counts[s_en] = subject_counts.get(s_en, 0) + 1
            if s_en not in subject_am_names and s_am:
                subject_am_names[s_en] = s_am

        # Build dynamic subject options sorted by count
        dynamic_subject_options = ["ሁሉም የሕግ ዘርፎች (All Subjects)"]
        for s_en, cnt in sorted(subject_counts.items(), key=lambda x: x[1], reverse=True):
            s_am = subject_am_names.get(s_en, s_en)
            dynamic_subject_options.append(f"{s_am} ({cnt:,}) • {s_en}")

        with l_f1:
            sel_subject_label = st.selectbox("የሕግ ዘርፍ ይምረጡ (Select Subject):", dynamic_subject_options)

        with l_f2:
            sel_norm = st.selectbox(
                "የሕግ ደረጃ (Norm Level):",
                [
                    "ሁሉም የሕግ ደረጃዎች (All)",
                    "አዋጅ (Proclamation)",
                    "ደንብ (Regulation)",
                    "መመሪያ (Directive)",
                    "ስምምነት (Treaty)",
                    "ጠቅላላ ሰነድ (General Document)"
                ]
            )

        with l_f3:
            law_search_kw = st.text_input("የሕግ ቁጥር ወይም ቃል ፍለጋ:", placeholder="ምሳሌ፡ 1333, 1094, ታክስ, መሬት, ግዥ...")

        # Filter application
        filtered_catalog = laws_catalog
        if "All Subjects" not in sel_subject_label:
            en_part = sel_subject_label.split(" • ")[-1].strip()
            filtered_catalog = [x for x in filtered_catalog if x.get('subject_en') == en_part]

        if "አዋጅ" in sel_norm:
            filtered_catalog = [x for x in filtered_catalog if "አዋጅ" in x.get('norm', '') or "Proclamation" in x.get('norm', '')]
        elif "ደንብ" in sel_norm:
            filtered_catalog = [x for x in filtered_catalog if "ደንብ" in x.get('norm', '') or "Regulation" in x.get('norm', '')]
        elif "መመሪያ" in sel_norm:
            filtered_catalog = [x for x in filtered_catalog if "መመሪያ" in x.get('norm', '') or "Directive" in x.get('norm', '')]
        elif "ስምምነት" in sel_norm:
            filtered_catalog = [x for x in filtered_catalog if "ስምምነት" in x.get('norm', '') or "Treaty" in x.get('norm', '')]
        elif "ጠቅላላ ሰነድ" in sel_norm:
            filtered_catalog = [x for x in filtered_catalog if "ጠቅላላ ሰነድ" in x.get('norm', '') or "General Document" in x.get('norm', '')]

        if law_search_kw:
            kw_n = normalize_amharic(law_search_kw)
            filtered_catalog = [
                x for x in filtered_catalog
                if kw_n in normalize_amharic(x.get('title_am', '')) 
                or kw_n in normalize_amharic(x.get('title_en', '')) 
                or kw_n in str(x.get('no', ''))
                or kw_n in str(x.get('year', ''))
                or kw_n in str(x.get('year_gc', ''))
            ]

        st.markdown(f"##### የተገኙ ሕጎች ዝርዝር (`{len(filtered_catalog):,}`)፡")

        if not filtered_catalog:
            st.info("በተመረጠው መስፈርት የተገኘ ሕግ የለም።")
        else:
            # Pagination for fast rendering
            page_size = 20
            total_pages = max(1, (len(filtered_catalog) + page_size - 1) // page_size)
            
            p_col1, p_col2 = st.columns([2, 8])
            with p_col1:
                cur_page = st.number_input("ገጽ (Page):", min_value=1, max_value=total_pages, value=1, step=1)
            with p_col2:
                st.write(f"ገጽ {cur_page} ከ {total_pages} (በአጠቃላይ {len(filtered_catalog):,} ሕጎች)")

            start_idx = (cur_page - 1) * page_size
            end_idx = min(start_idx + page_size, len(filtered_catalog))

            for idx, law in enumerate(filtered_catalog[start_idx:end_idx], start=start_idx + 1):
                with st.container():
                    ec_val = law.get('year', '-')
                    gc_val = law.get('year_gc', '-')
                    yr_display = f"{ec_val} ዓ.ም. ({gc_val} G.C.)"
                        
                    st.markdown(f"""
                    <div class="law-card">
                        <div class="head">
                            <h4>{idx}. [{law.get('norm')}] {law.get('title_am')}</h4>
                            <div>
                                <span class="chip">ቁጥር፡ {law.get('no', '-')}</span>
                                <span class="chip">{yr_display}</span>
                                <span class="chip">{law.get('subject_am', '')}</span>
                            </div>
                        </div>
                        <p style="margin: 0.4rem 0 0.2rem 0; font-size: 0.85rem; color: var(--ink-muted);">
                            {law.get('title_en', '')}
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    if law.get('link'):
                        st.markdown(f"[የሕጉን ሙሉ PDF በ Google Drive ክፈት ({law.get('size_mb', 0)} MB)]({law['link']})")
                    # decisions that cite this proclamation (matched on number and year)
                    _lk = None
                    _num = str(law.get('no', '')).split('/')[0]
                    if "አዋጅ" in str(law.get('norm', '')) and _num.isdigit():
                        for _y in (law.get('year'), law.get('year_gc')):
                            if _y and f"proc:{int(_num)}/{_y}" in link_index:
                                _lk = f"proc:{int(_num)}/{_y}"
                                break
                    if _lk:
                        _hits = link_index[_lk]
                        with st.expander(f"{len(_hits)} የሰበር ውሳኔዎች ይህን አዋጅ ይጠቅሳሉ ({len(_hits)} decisions cite this proclamation)"):
                            st.button("ሁሉንም አሳይ (Show all in Precedents)", on_click=show_citing, args=(_lk,), key=f"lib_all_{law.get('id')}_{idx}")
                            for _p in _hits[:5]:
                                _c = cases_data[_p]
                                _l1, _l2 = st.columns([8, 2])
                                _l1.write(f"**መ/ቁ {_c.get('case_no') or '-'}** ({_c.get('volume')}) — {' '.join(str(_c.get('legal_issue') or '').split())[:140]}")
                                _l2.button("ክፈት (Open)", on_click=open_case, args=(_p,), key=f"lib_open_{law.get('id')}_{idx}_{_p}")
                    st.write("")
