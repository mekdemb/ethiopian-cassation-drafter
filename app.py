import streamlit as st
import html
import json
import os
import re
import datetime
import urllib.request
import urllib.error
from urllib.parse import quote, urlsplit
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

    .block-container { padding-top: 2.75rem; }
    .main-header.compact {
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;
        gap: 0.6rem 1.5rem; padding: 0.85rem 1.25rem; margin-bottom: 0.75rem;
    }
    .main-header.compact h1 { font-size: 1.2rem; }
    .main-header.compact .eyebrow { margin: 0 0 0.2rem 0; }
    .main-header .stats { display: flex; flex-wrap: wrap; gap: 0.35rem 1.1rem; font-size: 0.8rem; color: var(--ink-on-dark-muted); }
    .main-header .stats b { color: #ffffff; font-size: 0.95rem; font-variant-numeric: tabular-nums; margin-right: 0.25rem; }
    .st-key-global_search input { font-size: 1rem; padding: 0.7rem 0.9rem; }
    .start-panel { background: var(--surface); border: 1px solid var(--rule); border-radius: 8px; padding: 1.5rem 1.75rem 0.75rem; margin-top: 0.5rem; }
    .start-panel h3 { margin: 0 0 0.4rem 0; font-size: 1.35rem; }
    .start-panel p { margin: 0 0 0.25rem 0; color: var(--ink-muted); font-size: 0.9rem; }
    .section-label { font-size: 0.75rem; font-weight: 700; letter-spacing: 0.06em; color: var(--ink-muted); margin: 1rem 0 0.4rem 0; }
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

# Load Volume Drive Map (linking Volumes to direct Google Drive PDFs)
@st.cache_data
def load_volume_drive_map():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "volume_drive_map.json"),
        os.path.join(os.path.dirname(base_dir), "volume_drive_map.json"),
        r"C:\Users\mekdem.belayneh\.gemini\antigravity\scratch\ethiopian_laws\cassation\volume_drive_map.json",
        "volume_drive_map.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    return {}

volume_drive_map = load_volume_drive_map()

# Load Drive Dockets Index (20,461 raw decision files across 15,940 dockets in CASSF)
@st.cache_data
def load_drive_dockets():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "drive_dockets_compact.json"),
        os.path.join(os.path.dirname(base_dir), "drive_dockets_compact.json"),
        r"C:\My Files\cassation\cloud_deploy\drive_dockets_compact.json",
        r"C:\My Files\cassation\drive_dockets_compact.json",
        "drive_dockets_compact.json"
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return json.load(f)
    return {}

drive_dockets = load_drive_dockets()

LEGAL_BILINGUAL_DICT = {
    # Tax & Finance
    'tax': ['ታክስ', 'ግብር', 'ገቢ ግብር'],
    'taxes': ['ታክስ', 'ግብር'],
    'taxation': ['ታክስ', 'ግብር'],
    'withholding': ['ተቀናሽ', 'በምንጭ ተቀናሽ', 'ቅናሽ'],
    'withhold': ['ተቀናሽ', 'ማስቀረት'],
    'vat': ['ቫት', 'የተጨማሪ እሴት ታክስ'],
    'tot': ['ተርን ኦቨር', 'የተርን ኦቨር ታክስ'],
    'customs': ['ጉምሩክ', 'ቀረጥ'],
    'duty': ['ቀረጥ'],
    # Tenancy & Property
    'tenant': ['ተከራይ', 'ኪራይ'],
    'tenants': ['ተከራይ', 'ኪራይ'],
    'landlord': ['አከራይ', 'ኪራይ'],
    'landlords': ['አከራይ', 'ኪራይ'],
    'rent': ['ኪራይ'],
    'rental': ['ኪራይ'],
    'lease': ['ኪራይ', 'ሊዝ'],
    'eviction': ['ማስለቀቅ', 'መልቀቅ'],
    'evict': ['ማስለቀቅ'],
    'house': ['ቤት'],
    'building': ['ህንፃ', 'ቤት'],
    'property': ['ንብረት', 'የማይንቀሳቀስ ንብረት'],
    'land': ['መሬት', 'ይዞታ'],
    'ownership': ['ባለቤትነት'],
    'possession': ['ይዞታ'],
    # Labor & Employment
    'labor': ['አሰሪና ሠራተኛ', 'ሠራተኛ'],
    'employment': ['አሰሪና ሠራተኛ', 'ቅጥር'],
    'employee': ['ሠራተኛ'],
    'employees': ['ሠራተኛ'],
    'employer': ['አሰሪ'],
    'employers': ['አሰሪ'],
    'severance': ['ስንብት', 'የስንብት ክፍያ'],
    'dismissal': ['ስንብት', 'ከስራ ማሰናበት'],
    'probation': ['የሙከራ ጊዜ'],
    'termination': ['የውል መቋረጥ', 'ስንብት'],
    'wages': ['ደመወዝ'],
    'salary': ['ደመወዝ'],
    # Contract Law
    'contract': ['ውል'],
    'contracts': ['ውል'],
    'agreement': ['ውል', 'ስምምነት'],
    'breach': ['ውል ማፍረስ', 'የውል መቋረጥ'],
    'damages': ['ካሳ', 'ጉዳት'],
    'compensation': ['ካሳ'],
    'obligation': ['ግዴታ'],
    'performance': ['ግዴታን መወጣት', 'ውል'],
    # Limitation / Prescription
    'limitation': ['ይርጋ', 'የይርጋ ጊዜ'],
    'prescription': ['ይርጋ'],
    'period': ['ይርጋ', 'ጊዜ'],
    # Family & Succession
    'marriage': ['ጋብቻ'],
    'divorce': ['ፍቺ'],
    'spouse': ['ባልና ሚስት', 'የትዳር አጋር'],
    'alimony': ['ቀለብ'],
    'maintenance': ['ቀለብ', 'አስተዳደግ'],
    'custody': ['አስተዳደግ', 'የልጆች አስተዳደግ'],
    'inheritance': ['ውርስ'],
    'succession': ['ውርስ', 'የውርስ ሀብት'],
    'will': ['ኑዛዜ'],
    'testament': ['ኑዛዜ'],
    'heir': ['ወራሽ'],
    'heirs': ['ወራሽ'],
    # Banking, Debt & Commerce
    'loan': ['ብድር', 'ዕዳ'],
    'debt': ['ዕዳ', 'ባለዕዳ'],
    'debtor': ['ባለዕዳ'],
    'creditor': ['ባለገንዘብ'],
    'mortgage': ['መያዣ', 'የማይንቀሳቀስ ንብረት መያዣ'],
    'pledge': ['መያዣ'],
    'guarantor': ['ዋስ', 'ዋስትና'],
    'guarantee': ['ዋስትና'],
    'check': ['ቼክ'],
    'cheque': ['ቼክ'],
    'share': ['አክሲዮን'],
    'shares': ['አክሲዮን'],
    'shareholder': ['ባለአክሲዮን'],
    'company': ['የንግድ ማህበር', 'ማህበር'],
    'bankruptcy': ['ኪሳራ'],
    # Agency & Representation
    'agency': ['ውክልና'],
    'agent': ['ወኪል'],
    'principal': ['ውክልና ሰጪ', 'ወካይ'],
    'power of attorney': ['ውክልና', 'የውክልና ስልጣን'],
    # Procedural & Judicial
    'appeal': ['ይግባኝ'],
    'cassation': ['ሰበር'],
    'jurisdiction': ['የዳኝነት ስልጣን', 'የስረ-ነገር ስልጣን'],
    'execution': ['ፍርድ አፈፃፀም'],
    'attachment': ['ዕግድ', 'ማስያዝ'],
    'injunction': ['ዕግድ'],
    'plenary': ['ምልዓተ-ጉባዔ', '7 ዳኞች'],
    'overruled': ['ተሽሯል', 'ሽሯል', 'የተሻረ'],
    # Torts & Criminal
    'tort': ['የፍትሐብሔር ኃላፊነት', 'ካሳ'],
    'injury': ['ጉዳት', 'አደጋ'],
    'negligence': ['ጥፋት', 'ቸልተኝነት'],
    'crime': ['ወንጀል'],
    'criminal': ['ወንጀል'],
    'penal': ['ወንጀል'],
    'bail': ['ዋስትና'],
    'theft': ['ስርቆት'],
    'corruption': ['ሙስና', 'የሙስና ወንጀል'],
    'fraud': ['ማታለል']
}

SEARCH_STOP_WORDS = set([
    'what', 'is', 'the', 'law', 'and', 'related', 'decisions', 'decision', 'regards', 'regard', 'to', 'his', 'her',
    'a', 'an', 'for', 'of', 'on', 'at', 'by', 'with', 'about', 'between', 'from', 'rule', 'rules', 'case', 'cases',
    'court', 'legal', 'in', 'relation', 'does', 'do', 'under', 'ethiopian', 'ethiopia', 'who', 'how', 'why', 'when',
    'where', 'which', 'can', 'should', 'would', 'could', 'give', 'me', 'tell', 'explain', 'show', 'please', 'all',
    'any', 'some', 'my', 'their', 'its', 'as', 'or', 'if', 'into', 'out', 'up', 'down', 'then', 'so', 'there',
    'ምን', 'ነው', 'ሕግ', 'ስለ', 'ላይ', 'ጋር', 'የተመለከተ', 'እና', 'ወይም', 'እንደ', 'ውስጥ', 'ይህ', 'ያለ', 'ያለው'
])

def retrieve_relevant_cases(query, cases, max_results=5):
    """Bilingual, semantic-boosted retrieval over 3,246 binding precedents."""
    if not query:
        return []

    norm_q = normalize_amharic(query.lower())
    raw_words = [w for w in re.findall(r'[\w\d/]+', norm_q) if len(w) > 1]
    if not raw_words:
        return []

    meaningful = [w for w in raw_words if w not in SEARCH_STOP_WORDS]
    if not meaningful:
        meaningful = raw_words

    search_terms = set(meaningful)
    for w in meaningful:
        if w in LEGAL_BILINGUAL_DICT:
            for syn in LEGAL_BILINGUAL_DICT[w]:
                search_terms.add(normalize_amharic(syn))

    digits = [w for w in raw_words if w.isdigit()]

    scored = []
    for i, c in enumerate(cases):
        score = 0
        c_no = str(c.get('case_no') or '')
        issue = normalize_amharic(str(c.get('legal_issue') or ''))
        holding = normalize_amharic(str(c.get('holding') or ''))
        subject = normalize_amharic(str(c.get('legal_subject') or ''))
        provisions = normalize_amharic(str(c.get('provisions') or ''))
        parties = normalize_amharic(str(c.get('applicant') or '') + ' ' + str(c.get('respondent') or ''))
        full_text = f"{c_no} {issue} {holding} {subject} {provisions} {parties}"

        # Exact docket match
        for d in digits:
            if d == c_no:
                score += 350
            elif d in c_no:
                score += 120
            elif d in provisions:
                score += 40

        # Term matching across fields
        matched_terms = set()
        for t in search_terms:
            if t in full_text:
                matched_terms.add(t)
                if t in issue:
                    score += 30
                if t in holding:
                    score += 25
                if t in provisions:
                    score += 20
                if t in subject:
                    score += 15
                if t in parties:
                    score += 5

        # Multi-concept synergy bonus (rewards cases matching multiple distinct facets of the query)
        if len(matched_terms) > 1:
            score += (len(matched_terms) ** 2) * 15

        if score > 0:
            if c.get('status') in ['PLENARY', 'PLENARY_OVERRULING'] or '7 ዳኞች' in str(c.get('bench', '')):
                score += 25
            scored.append((score, i, c))

    scored.sort(key=lambda x: x[0], reverse=True)
    return [item[2] for item in scored[:max_results]]

def call_gemini_counsel(question, retrieved_cases, api_key):
    """Query Gemini with strict grounding on Ethiopian Supreme Court precedents."""
    if not retrieved_cases:
        return "No directly matching binding Cassation precedents were found in the database for your query. Please rephrase with specific legal terms, article numbers, or docket numbers."

    context_blocks = []
    for idx, c in enumerate(retrieved_cases, 1):
        status_alert = ""
        if c.get('status') == 'OVERRULED':
            status_alert = f"⚠️ [OVERRULED PRECEDENT: This decision was explicitly overruled by the 7-Judges Plenary Bench! Note: {c.get('status_note', '')}]"
        elif c.get('status') == 'PLENARY_OVERRULING':
            status_alert = f"⭐ [HISTORIC 7-JUDGES PLENARY PRECEDENT: Overrules prior contradictory rulings! Note: {c.get('status_note', '')}]"

        # Trim holding and issue to avoid exceeding token limits while preserving essence
        holding_val = str(c.get('holding', '-')).strip()
        if len(holding_val) > 1200:
            holding_val = holding_val[:1200] + "..."
        issue_val = str(c.get('legal_issue', '-')).strip()
        if len(issue_val) > 600:
            issue_val = issue_val[:600] + "..."

        block = f"""[ሰበር መዝገብ {idx}]
መዝገብ ቁጥር (Docket No): {c.get('case_no') or 'ያልተረጋገጠ'}
ቅጽ / ምንጭ (Volume): {c.get('volume', '-')} (ገጽ: {c.get('page_in_vol', '-')})
የችሎት አደረጃጀት (Bench): {c.get('bench', '5 ዳኞች')}
የተወሰነበት ቀን (Date): {c.get('date', '-')}
አመልካች (Applicant): {c.get('applicant', '-')}
ተጠሪ (Respondent): {c.get('respondent', '-')}
የሕግ ዘርፍ (Field): {c.get('legal_subject', '-')}
የተተረጎሙ ድንጋጌዎች (Provisions): {c.get('provisions', '-')}
የአስገዳጅነት ደረጃ (Status): {c.get('status', 'VALID')} {status_alert}
የተያዘው ጭብጥ (Issue): {issue_val}
አስገዳጅ የሕግ ትርጓሜ (Ratio Decidendi): {holding_val}
"""
        context_blocks.append(block)

    context_str = "\n".join(context_blocks)

    system_prompt = (
        "You are the Senior Judicial Legal Counsel for the Federal Supreme Court of Ethiopia Cassation Division (የኢፌዴሪ ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት).\n"
        "Your task is to provide an authoritative, comprehensive, source-grounded legal analysis to the user's question, based on the provided binding precedents (under Proclamation No. 454/1997 & 1234/2021) and statutory provisions.\n\n"
        "CORE JUDICIAL RULES:\n"
        "1. STRICT CITATION & SOURCING: For every legal principle or rule you state, cite the authentic precedent: ሰበር መ/ቁ [Docket Number], ቅጽ [Volume], አመልካች [Applicant] v. ተጠሪ [Respondent], and ቀን [Date].\n"
        "2. LANGUAGE MATCHING: If the user asks in English, write your complete response, headings, reasoning, and practical takeaways in fluent, professional English (keeping authentic docket numbers and party names). If the user asks in Amharic, respond fully in professional Amharic.\n"
        "3. HIGHLY STRUCTURED BENCH MEMORANDUM FORMAT:\n"
        "   - **Direct Answer / Executive Summary**: Provide a clear, definitive statement resolving the user's exact legal question.\n"
        "   - **Statutory Framework**: State the applicable statutory provisions (e.g. Civil Code, Commercial Code, Tax Proclamations, Labor Proclamations).\n"
        "   - **Binding Cassation Precedent Analysis**: Analyze the cited precedent(s) in detail, explaining how the Cassation Division interpreted the law and resolved the issue (Ratio Decidendi).\n"
        "   - **Practical Judicial Takeaway & Guidance**: Provide actionable guidance for litigants, counsel, tenants, or businesses.\n"
        "4. SHEPARD'S CITATOR WARNING: If any cited case is marked OVERRULED, prominently highlight it with '⚠️ OVERRULED PRECEDENT' and explain the prevailing rule established by the 7-Judges Plenary Bench.\n"
        "5. TONE: Objective, scholarly, authoritative, and judicial."
    )

    user_prompt = f"""User's Legal Inquiry:
{question}

Authentic Binding Cassation Precedents Context:
{context_str}

Please provide a thorough, authoritative judicial analysis answering the inquiry based directly on these binding Ethiopian Supreme Court Cassation precedents and governing laws.
"""

    # Dynamically discover supported models for this user's API key
    discovered_models = []
    bad_tokens = ["tts", "audio", "speech", "image", "imagen", "embed", "realtime", "robotics"]
    try:
        list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        list_req = urllib.request.Request(list_url, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(list_req, timeout=10) as l_resp:
            l_data = json.loads(l_resp.read().decode('utf-8'))
            for m in l_data.get('models', []):
                m_name = m.get('name', '').replace('models/', '')
                if 'generateContent' in m.get('supportedGenerationMethods', []):
                    # Strictly exclude TTS, audio, image, embedding models
                    if not any(b in m_name.lower() for b in bad_tokens):
                        discovered_models.append(m_name)
    except Exception:
        pass

    # Preferred models in order of priority (most stable production text models first)
    candidate_order = [
        "gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-1.5-flash-001", "gemini-1.5-flash-002",
        "gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-2.0-flash-exp",
        "gemini-1.5-flash-8b",
        "gemini-1.5-pro", "gemini-1.5-pro-latest", "gemini-1.5-pro-001", "gemini-1.5-pro-002",
        "gemini-pro"
    ]

    models_to_try = []
    for c in candidate_order:
        if c in discovered_models and c not in models_to_try:
            models_to_try.append(c)
    for d in discovered_models:
        if "gemini" in d.lower() and d not in models_to_try:
            if not any(b in d.lower() for b in bad_tokens):
                models_to_try.append(d)
    if not models_to_try:
        models_to_try = candidate_order

    last_err = None

    for model in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        payload = {
            "system_instruction": {
                "parts": [{"text": system_prompt}]
            },
            "contents": [
                {"role": "user", "parts": [{"text": user_prompt}]}
            ],
            "generationConfig": {
                "temperature": 0.2,
                "maxOutputTokens": 3000
            }
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode('utf-8'),
            headers={"Content-Type": "application/json"}
        )

        try:
            with urllib.request.urlopen(req, timeout=35) as resp:
                res_data = json.loads(resp.read().decode('utf-8'))
                candidates = res_data.get('candidates', [])
                if candidates and 'content' in candidates[0]:
                    parts = candidates[0]['content'].get('parts', [])
                    if parts:
                        text = parts[0].get('text', '')
                        footer = f"\n\n---\n*⚡ Legal analysis powered by Google `{model}` • Grounded on authentic Federal Supreme Court precedents.*"
                        return text + footer
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8', errors='ignore')
            last_err = f"HTTP {e.code} ({model}): {err_body}"
            # If API key is invalid/unauthorized, stop immediately
            if e.code in [400, 403] and ("API_KEY_INVALID" in err_body or "PERMISSION_DENIED" in err_body):
                return f"❌ Invalid or unauthorized API Key ({e.code}). Please check your Gemini API key in the configuration above."
            # On 503 (high demand), 429 (rate/quota limit), 404 (model not found), or 500/502/504 server errors:
            # seamlessly failover to the next candidate model in the list!
            continue
        except Exception as e:
            last_err = str(e)
            continue

    return f"❌ Unable to generate AI analysis (All models busy or unavailable). Last error: {last_err}"

def generate_notebooklm_dossier(question, answer, cases):
    """Generate a structured research dossier text file for NotebookLM."""
    lines = []
    lines.append("=" * 70)
    lines.append("ETHIOPIAN CASSATION PRECEDENTS RESEARCH DOSSIER FOR NOTEBOOKLM")
    lines.append(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("Source: Federal Supreme Court Cassation Division (Vols 1-28 + Plenary)")
    lines.append("=" * 70)
    lines.append(f"\nLEGAL INQUIRY:\n{question}\n")
    lines.append(f"\nSYNTHESIS & ANALYSIS:\n{answer}\n")
    lines.append("=" * 70)
    lines.append(f"AUTHENTIC CITED PRECEDENTS ({len(cases)} Cases)\n")

    for i, c in enumerate(cases, 1):
        lines.append(f"\n[CASE #{i}]")
        lines.append(f"Docket No: {c.get('case_no') or '-'}")
        lines.append(f"Volume: {c.get('volume', '-')} (Page: {c.get('page_in_vol', '-')})")
        lines.append(f"Date: {c.get('date', '-')}")
        lines.append(f"Bench: {c.get('bench', '-')}")
        lines.append(f"Parties: {c.get('applicant', '-')} vs {c.get('respondent', '-')}")
        lines.append(f"Subject: {c.get('legal_subject', '-')}")
        lines.append(f"Provisions: {c.get('provisions', '-')}")
        lines.append(f"Status: {c.get('status', 'VALID')}")
        if c.get('status_note'):
            lines.append(f"Status Note: {c.get('status_note')}")
        lines.append(f"Legal Issue:\n{c.get('legal_issue', '-')}")
        lines.append(f"Ratio Decidendi:\n{c.get('holding', '-')}")
        lines.append("-" * 50)

    return "\n".join(lines)

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

TAB_PRECEDENTS = "የሰበር ውሳኔዎች (Precedents)"
TAB_AI = "🤖 የሕግ ረዳት (AI Counsel)"
TAB_DRAFTER = "አቤቱታ አዘጋጅ (Drafter)"
TAB_OVERRULED = "የተሻሩ ውሳኔዎች (Overruled)"
TAB_ANALYTICS = "የሕግ ትንተና (Analytics)"
TAB_CITATOR = "የሥልጣን ትስስር (Citator)"
TAB_LIBRARY = "የሕጎች ማውጫ (Library)"

TAB_NAMES = [
    TAB_PRECEDENTS,
    TAB_AI,
    TAB_DRAFTER,
    TAB_OVERRULED,
    TAB_ANALYTICS,
    TAB_CITATOR,
    TAB_LIBRARY
]
if st.session_state.get('nav') not in TAB_NAMES:
    st.session_state.nav = TAB_PRECEDENTS

def open_case(pos):
    """Open one decision in the Precedents screen."""
    st.session_state.focus_pos = pos
    st.session_state.browse_all = True     # "Back to results" then shows the list, not the start screen
    st.session_state.nav = TAB_PRECEDENTS

def show_citing(key):
    """Show every decision that cites a provision, in the Precedents screen."""
    st.session_state.prov_filter = key
    st.session_state.focus_pos = None
    st.session_state.search_q = ""          # so the list shows every decision citing it
    st.session_state.nav = TAB_PRECEDENTS

def clear_citing():
    st.session_state.prov_filter = None

def clear_focus():
    st.session_state.focus_pos = None

STATUS_OPTS = ["ሁሉንም ውሳኔዎች (All)", "ጸንቶ የሚሰራ ብቻ (Valid)", "7 ዳኞች ሰበር ብቻ (Plenary Bench)", "የተሻሩ ውሳኔዎች ብቻ (Overruled Decisions)", "የቀድሞ ውሳኔን የሻሩ ውሳኔዎች"]
VERIFY_OPTS = ["ሁሉንም (All)", "በህትመቱ የተረጋገጠ (Verified)", "ከአዲስ ቅጽ የታከለ (Added)", "ማጣራት የሚፈልግ (Check Required)"]
PLEADING_TYPES = ["የሰበር አቤቱታ", "የመከላከያ መልስ", "የይግባኝ አቤቱታ", "የሕግ ማስታወሻ"]

# Settings that must be remembered while their controls are off screen
_DEFAULTS = {
    'search_q': "", 'lib_kw': "",
    'f_subjects': [], 'f_volumes': [], 'f_status': STATUS_OPTS[0], 'f_verify': VERIFY_OPTS[0],
    'p_court': "ለኢ.ፌ.ዴ.ሪ ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት", 'p_type': PLEADING_TYPES[0],
    'p_app': "[የአመልካች ስም]", 'p_resp': "[የተጠሪ ስም]",
    'browse_all': False,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v
    else:
        st.session_state[_k] = st.session_state[_k]   # keeps the value when its control is not shown

# ---- Details to update by hand when they change ----
APP_URL = "https://ethiopian-cassation-drafter.streamlit.app/"   # address used in "link to this decision"
DATA_UPDATED = "15 November 2026"                                 # date the decisions database was last changed
REPORT_EMAIL = "mekdemb@gmail.com"                                # where "report an error" emails go

def move_draft(i, step):
    """Drafter: move a selected precedent up or down."""
    lst = st.session_state.selected_for_draft
    j = i + step
    if 0 <= i < len(lst) and 0 <= j < len(lst):
        lst[i], lst[j] = lst[j], lst[i]

def remove_draft(i):
    lst = st.session_state.selected_for_draft
    if 0 <= i < len(lst):
        lst.pop(i)

def load_draft():
    """Drafter: restore a selection saved earlier with 'Save draft'."""
    f = st.session_state.get('draft_upload')
    if f is None:
        return
    try:
        data = json.loads(f.getvalue().decode('utf-8'))
        by_entry = {str(c.get('entry_no')): c for c in cases_data}
        picked = [by_entry[str(e)] for e in data.get('entries', []) if str(e) in by_entry]
        st.session_state.selected_for_draft = picked
        if data.get('court'):
            st.session_state.p_court = str(data['court'])
        if data.get('type') in PLEADING_TYPES:
            st.session_state.p_type = data['type']
        if data.get('applicant'):
            st.session_state.p_app = str(data['applicant'])
        if data.get('respondent'):
            st.session_state.p_resp = str(data['respondent'])
        st.session_state.draft_msg = ('ok', len(picked))
    except Exception:
        st.session_state.draft_msg = ('error', 0)

if not st.session_state.get('link_done'):
    st.session_state.link_done = True
    _qc, _qe = st.query_params.get('case'), st.query_params.get('e')
    if _qc or _qe:
        _hit = None
        if _qe:
            _hit = next((i for i, c in enumerate(cases_data) if str(c.get('entry_no')) == str(_qe)), None)
        if _hit is None and _qc:
            _hit = next((i for i, c in enumerate(cases_data) if str(c.get('case_no')) == str(_qc)), None)
        if _hit is not None:
            st.session_state.focus_pos = _hit
            st.session_state.browse_all = True
            st.session_state.nav = TAB_PRECEDENTS

shown_case = None   # the decision on screen, if any (keeps the address bar in step)

def run_example(q):
    """Start screen: run an example search."""
    st.session_state.search_q = q
    st.session_state.focus_pos = None
    st.session_state.prov_filter = None

def browse_status(option):
    st.session_state.f_status = option

def browse_everything():
    st.session_state.browse_all = True

def go_decisions():
    st.session_state.focus_pos = None
    st.session_state.nav = TAB_PRECEDENTS

def go_laws():
    """Carry the search words over to the Laws Library."""
    st.session_state.lib_kw = st.session_state.search_q
    st.session_state.nav = TAB_LIBRARY

@st.cache_data
def law_search_texts(_catalog):
    return [normalize_amharic(" ".join(str(x.get(f, '')) for f in ('title_am', 'title_en', 'no', 'year', 'year_gc'))) for x in _catalog]

# Counts shown in the header
total_cases = len(cases_data)
# 7-Judges decisions, each counted once (three are printed in two volumes)
plenary_cases = len({str(c.get('case_no') or c.get('entry_no')) for c in cases_data if "7 ዳኞች" in str(c.get('bench', '')) or c.get('status') in ['PLENARY', 'PLENARY_OVERRULING']})
overruled_cases = len([c for c in cases_data if c.get('status') == 'OVERRULED'])
verified_cases = len([c for c in cases_data if "Verified" in str(c.get('verification_status', ''))])
selected_count = len(st.session_state.selected_for_draft)

# App Header (compact, with the counts inside it)
st.markdown(f"""
<div class="main-header compact">
    <div>
        <p class="eyebrow">የፌዴራል ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት · FEDERAL SUPREME COURT CASSATION DIVISION</p>
        <h1>የኢትዮጵያ ሰበር አስገዳጅ የሕግ ትርጓሜዎች መፈለጊያ እና ረዳት</h1>
    </div>
    <div class="stats">
        <span><b>{total_cases:,}</b>ጠቅላላ ውሳኔዎች</span>
        <span><b>{plenary_cases}</b>7 ዳኞች ሰበር</span>
        <span><b>{overruled_cases}</b>የተሻሩ</span>
        <span><b>{verified_cases:,}</b>በህትመቱ የተረጋገጡ</span>
        <span><b>{selected_count}</b>የተመረጡ (Draft)</span>
    </div>
</div>
""", unsafe_allow_html=True)

# One search box for decisions and laws
with st.container(key="global_search"):
    search_query = st.text_input(
        "ሁሉን-አቀፍ ፍለጋ (Universal Search)",
        key="search_q",
        placeholder="ውሳኔዎችን እና ሕጎችን ይፈልጉ (Search decisions and laws): መዝገብ ቁጥር፣ ተከራካሪ፣ አንቀጽ ወይም ቃል...",
        help="በመዝገብ ቁጥር (e.g. 18180, 242314), በአንቀጽ (1845, 1000) ወይም በሕግ ቃል ይፈልጉ",
        label_visibility="collapsed",
    )

# Filters for decisions: shown in the sidebar only on the Precedents screen
all_subjects = sorted(list(set(c.get('legal_subject', 'ልዩ ልዩ') for c in cases_data)))

def sort_vol_key(vol):
    m = re.search(r'\d+', vol)
    return int(m.group()) if m else 999

all_volumes = sorted(list(set(c.get('volume', 'ሌሎች') for c in cases_data)), key=sort_vol_key)

if st.session_state.nav == TAB_PRECEDENTS:
    st.sidebar.header("የውሳኔዎች አጣሪ (Filters)")
    selected_subjects = st.sidebar.multiselect("የሕግ ዘርፍ (Subject)", options=all_subjects, key="f_subjects")
    selected_volumes = st.sidebar.multiselect("ቅጽ / ምንጭ (Volume)", options=all_volumes, key="f_volumes")
    status_choice = st.sidebar.radio("የአስገዳጅነት ደረጃ (Status)", options=STATUS_OPTS, key="f_status")
    verify_choice = st.sidebar.selectbox("የህትመት ማረጋገጫ (Verification)", options=VERIFY_OPTS, key="f_verify")
else:
    selected_subjects = st.session_state.f_subjects
    selected_volumes = st.session_state.f_volumes
    status_choice = st.session_state.f_status
    verify_choice = st.session_state.f_verify

# Filter Data
norm_query = normalize_amharic(search_query)
terms = norm_query.split() if norm_query else []

def filter_case(c):
    if selected_subjects and c.get('legal_subject') not in selected_subjects:
        return False
    if selected_volumes and c.get('volume') not in selected_volumes:
        return False
        
    if status_choice == "ጸንቶ የሚሰራ ብቻ (Valid)":
        if c.get('status') != 'VALID':
            return False
    elif status_choice == "7 ዳኞች ሰበር ብቻ (Plenary Bench)":
        if "7 ዳኞች" not in str(c.get('bench', '')) and c.get('status') not in ['PLENARY', 'PLENARY_OVERRULING']:
            return False
    elif status_choice == "የተሻሩ ውሳኔዎች ብቻ (Overruled Decisions)":
        if c.get('status') != 'OVERRULED':
            return False
    elif status_choice == "የቀድሞ ውሳኔን የሻሩ ውሳኔዎች":
        if c.get('status') != 'PLENARY_OVERRULING':
            return False

    if verify_choice == "በህትመቱ የተረጋገጠ (Verified)":
        if "Verified" not in str(c.get('verification_status', '')):
            return False
    elif verify_choice == "ከአዲስ ቅጽ የታከለ (Added)":
        if "Added" not in str(c.get('verification_status', '')):
            return False
    elif verify_choice == "ማጣራት የሚፈልግ (Check Required)":
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

# Where the search found something: decisions and laws
laws_catalog_all = load_laws_catalog()
if terms and laws_catalog_all:
    _law_texts = law_search_texts(laws_catalog_all)
    law_hits = sum(1 for t in _law_texts if all(w in t for w in terms))
else:
    law_hits = 0
if terms:
    _s1, _s2, _s3 = st.columns([3, 3, 6])
    _s1.button(f"ውሳኔዎች (Decisions): {len(filtered_cases):,}", on_click=go_decisions, key="go_decisions", use_container_width=True, disabled=not filtered_cases)
    _s2.button(f"ሕጎች (Laws): {law_hits:,}", on_click=go_laws, key="go_laws", use_container_width=True, disabled=law_hits == 0)

# A first visit with nothing searched or filtered shows a start screen
show_start = (
    not terms and not selected_subjects and not selected_volumes
    and status_choice == STATUS_OPTS[0] and verify_choice == VERIFY_OPTS[0]
    and not st.session_state.prov_filter and st.session_state.focus_pos is None
    and not st.session_state.browse_all
)

# Tabs Interface
with st.container(key="mainnav"):
    nav = st.radio("Screen", TAB_NAMES, horizontal=True, key="nav", label_visibility="collapsed")

# ==========================================
# TAB 1: SPLIT SCREEN RESEARCH & INSPECTOR
# ==========================================
if nav == TAB_PRECEDENTS and show_start:
    st.markdown("""
    <div class="start-panel">
        <h3>ምን መፈለግ ይፈልጋሉ? (What are you looking for?)</h3>
        <p>በመዝገብ ቁጥር፣ በተከራካሪ ስም፣ በአንቀጽ ወይም በቃል ከላይ ባለው ሳጥን ይፈልጉ።</p>
        <p>Type a file number, a party, an article or a word in the search box above. It searches decisions and laws together.</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<p class="section-label">ለምሳሌ · FOR EXAMPLE</p>', unsafe_allow_html=True)
    _examples = [
        ("ይርጋ (Period of limitation)", "ይርጋ"),
        ("ፍ/ብ/ሕ/ቁ 1845", "1845"),
        ("መ/ቁ 242314", "242314"),
        ("ኑዛዜ (Wills)", "ኑዛዜ"),
        ("አዋጅ 377/96 (Labour)", "377/96"),
    ]
    _ec = st.columns(len(_examples))
    for _j, (_lbl, _q) in enumerate(_examples):
        _ec[_j].button(_lbl, on_click=run_example, args=(_q,), key=f"example_{_j}", use_container_width=True)
    st.markdown('<p class="section-label">ወይም ይመልከቱ · OR BROWSE</p>', unsafe_allow_html=True)
    _bc = st.columns(3)
    _bc[0].button("የ7 ዳኞች ሰበር ውሳኔዎች (7-Judges rulings)", on_click=browse_status, args=(STATUS_OPTS[2],), key="browse_plenary", use_container_width=True)
    _bc[1].button("የተሻሩ ውሳኔዎች (Overruled decisions)", on_click=browse_status, args=(STATUS_OPTS[3],), key="browse_overruled", use_container_width=True)
    _bc[2].button("ሁሉንም ውሳኔዎች አሳይ (Show all decisions)", on_click=browse_everything, key="browse_all_btn", use_container_width=True)

if nav == TAB_PRECEDENTS and not show_start:
    if st.session_state.prov_filter:
        _fb1, _fb2 = st.columns([8, 2])
        with _fb1:
            st.markdown(f'<div class="filter-bar">ማጣሪያ (Filter): {html.escape(provision_label(st.session_state.prov_filter))} የሚጠቅሱ ውሳኔዎች (decisions citing this provision)</div>', unsafe_allow_html=True)
        with _fb2:
            st.button("ማጣሪያውን አስወግድ (Clear)", on_click=clear_citing, key="clear_citing", use_container_width=True)
    st.write(f"**የተገኙ ውሳኔዎች:** `{len(filtered_cases):,}` / {total_cases:,}")
    
    col_left, col_right = st.columns([5, 7])
    
    with col_left:
        st.subheader("የውሳኔዎች ዝርዝር")
        
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
            st.subheader(f"ሰበር መዝገብ ቁጥር {case_no_disp}")
            
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

            # Raw Google Drive Court Decisions from CASSF
            c_docket = str(active_c.get('case_no', '')).strip()
            raw_drive_matches = drive_dockets.get(c_docket) or []
            if raw_drive_matches:
                st.markdown("##### 📁 የዚህ ውሳኔ ሰነዶች በ Google Drive (Raw Files in CASSF)")
                d_cols = st.columns(min(len(raw_drive_matches), 3))
                for idx_rf, rf in enumerate(raw_drive_matches):
                    with d_cols[idx_rf % 3]:
                        st.markdown(f"[📄 **{rf['n']}** (ፎልደር: {rf['f']})](https://drive.google.com/file/d/{rf['id']}/view)")

            # Legal Issue
            st.markdown("##### የተያዘው ጭብጥ (Legal Issue)")
            st.markdown('<div class="issue-box">' + highlight_html(active_c.get('legal_issue', 'የሕግ ትርጓሜ'), terms) + '</div>', unsafe_allow_html=True)

            # Ratio Decidendi
            st.markdown("##### አስገዳጅ የሕግ ትርጓሜ (Authoritative Holding)")
            st.markdown('<blockquote class="holding">' + highlight_html(active_c.get('holding', ''), terms) + '</blockquote>', unsafe_allow_html=True)

            # Statutory Provisions
            if active_c.get('provisions'):
                st.markdown(f"**የተተረጎሙ ድንጋጌዎች:** `{active_c.get('provisions')}`")
                _keys = [k for k in case_links[active_pos] if len(link_index.get(k, [])) > 1][:12]
                if _keys:
                    st.caption("ተመሳሳይ ድንጋጌ የተረጎሙ ሌሎች ውሳኔዎች (Other decisions on the same provision):")
                    _pc = st.columns(2)
                    for j, k in enumerate(_keys):
                        with _pc[j % 2]:
                            st.button(f"{provision_label(k)} ({len(link_index[k])} ውሳኔዎች)", on_click=show_citing, args=(k,), key=f"prov_{active_pos}_{k}", use_container_width=True)

            # Provenance & Correction Notes
            if active_c.get('source') or active_c.get('notes'):
                with st.expander("የመረጃው ምንጭ እና የማረጋገጫ ማስታወሻ (Provenance & Notes)", expanded=False):
                    if active_c.get('source'):
                        st.write(f"**ምንጭ:** {active_c.get('source')}")
                    if active_c.get('notes'):
                        st.write(f"**ማስታወሻ:** {active_c.get('notes')}")

            # Actions
            b1, b2 = st.columns(2)
            with b1:
                page_str = f"፣ ገጽ {active_c.get('page_in_vol')}" if active_c.get('page_in_vol') else ""
                formal_cite = f"በፌዴራል ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት በመዝገብ ቁጥር {active_c.get('case_no')}፣ {active_c.get('volume')}{page_str} ({active_c.get('bench')})፣ አመልካች {active_c.get('applicant')} እና ተጠሪ {active_c.get('respondent')} በተወሰነው አስገዳጅ ውሳኔ መሠረት..."
                st.text_area("የፍርድ ቤት ጥቅስ (Formal Citation):", value=formal_cite, height=75)
            with b2:
                c_no = str(active_c.get('case_no') or active_c.get('entry_no'))
                if c_no in [str(c.get('case_no') or c.get('entry_no')) for c in st.session_state.selected_for_draft]:
                    if st.button("− ከአቤቱታ ረቂቅ አስወግድ (Remove from Draft)"):
                        st.session_state.selected_for_draft = [c for c in st.session_state.selected_for_draft if str(c.get('case_no') or c.get('entry_no')) != c_no]
                        st.rerun()
                else:
                    if st.button("+ ወደ አቤቱታ ረቂቅ ጨምር (Add to Draft)"):
                        st.session_state.selected_for_draft.append(active_c)
                        st.rerun()

            # Link to this decision, data date, report an error
            shown_case = active_c
            try:
                _parts = urlsplit(st.context.url)
                if not _parts.netloc:
                    _base = APP_URL
                elif _parts.netloc.endswith(".streamlit.app"):
                    _base = f"https://{_parts.netloc}/"
                else:
                    _base = f"{_parts.scheme}://{_parts.netloc}{_parts.path or '/'}"
            except Exception:
                _base = APP_URL
            _cn, _en = active_c.get('case_no'), active_c.get('entry_no')
            _link = f"{_base}?case={quote(str(_cn))}&e={_en}" if _cn else f"{_base}?e={_en}"
            st.caption("የዚህ ውሳኔ ማስፈንጠሪያ (Link to this decision) — ለመቅዳት በቀኝ ያለውን ምልክት ይጫኑ (use the copy button on the right):")
            st.code(_link, language=None)
            _subject = quote(f"Cassation database: error in File {_cn or '-'} (entry {_en})")
            _body = quote(f"File no.: {_cn or '-'}\nEntry no.: {_en}\nVolume: {active_c.get('volume')}\n\nWhat is wrong:\n")
            st.caption(f"የመረጃ ቋቱ የመጨረሻ ማሻሻያ (Database last updated): {DATA_UPDATED} · [ስህተት ያሳውቁ (Report an error)](mailto:{REPORT_EMAIL}?subject={_subject}&body={_body})")

# ==========================================
# TAB: AI LEGAL COUNSEL (የሕግ ረዳት)
# ==========================================
if nav == TAB_AI:
    st.subheader("🤖 የሰበር ሕግ ረዳት (AI Legal Counsel & Research Workspace)")
    st.caption("በኢትዮጵያ ፌዴራል ጠቅላይ ፍርድ ቤት ሰበር አስገዳጅ ውሳኔዎች (ቅጽ 1 - 28 እና 7 ዳኞች ምልዓተ-ጉባዔ) ላይ ብቻ የተመሰረተ ምንጭ-ተኮር (Source-Grounded) የሕግ ጥያቄና መልስ ረዳት")

    # API Key Configuration
    api_key = (
        st.session_state.get("gemini_api_key")
        or os.environ.get("GEMINI_API_KEY")
        or (st.secrets.get("GEMINI_API_KEY") if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets else None)
    )

    with st.expander("🔑 የ Gemini API ቁልፍ ማስተካከያ (API Key Configuration)", expanded=not bool(api_key)):
        st.markdown("""
        ይህ የ AI ረዳት የሚሰራው በ **Google Gemini API** አማካኝነት ነው። 
        [Google AI Studio (aistudio.google.com)](https://aistudio.google.com/) በመግባት በነጻ የራስዎን API Key ማግኘት ይችላሉ።
        """)
        entered_key = st.text_input(
            "Gemini API Key:",
            value=api_key or "",
            type="password",
            placeholder="AIzaSy...",
            help="በ Google AI Studio የተፈጠረውን የነጻ API ቁልፍ ያስገቡ"
        )
        if entered_key:
            st.session_state.gemini_api_key = entered_key.strip()
            api_key = entered_key.strip()
            st.success("✅ የ API ቁልፍ ተቀምጧል! ጥያቄዎን ማቅረብ ይችላሉ።")

    # Example query buttons
    st.markdown('<p class="section-label">የተመረጡ የሕግ ጥያቄዎች ምሳሌ (CLICK AN EXAMPLE QUERY)</p>', unsafe_allow_html=True)
    ai_examples = [
        ("ይርጋ በ7 ዳኞች ሰበር", "የይርጋ ጊዜ መቋረጥን አስመልክቶ በ7 ዳኞች ሰበር የተሰጠው አዲስ አስገዳጅ አቋም ምንድን ነው?"),
        ("ያልተመዘገበ የቤት ሽያጭ", "ያልተመዘገበ ወይም በሰነዶች ማረጋገጫ ያልጸደቀ የቤት ሽያጭ ውል የፈረሱ ውጤት ምንድን ነው?"),
        ("የሰራተኛ ስንብት ካሳ", "የሰራተኛ የስራ ውል ያለበቂ ምክንያት ሲቋረጥ የሚከፈሉ የካሳ እና የአገልግሎት ክፍያዎች ምን ይመስላሉ?"),
        ("የኑዛዜ ውል ተቀባይነት", "የኑዛዜ ውል አቀራረፅና ተቀባይነት በፍ/ብ/ሕ/ቁ 881 እና ተከታዮቹ መሠረት የሰበር ውሳኔዎች ምን ይላሉ?")
    ]
    ai_cols = st.columns(len(ai_examples))
    chosen_prompt = None
    for _idx, (_short, _full) in enumerate(ai_examples):
        if ai_cols[_idx].button(_short, key=f"ai_ex_{_idx}", use_container_width=True):
            chosen_prompt = _full

    # Chat history state
    if "ai_chat_history" not in st.session_state:
        st.session_state.ai_chat_history = []

    if st.session_state.ai_chat_history:
        _c_col1, _c_col2 = st.columns([8, 2])
        if _c_col2.button("🗑️ Clear Chat", key="clear_ai_chat", help="Clear previous conversation history"):
            st.session_state.ai_chat_history = []
            st.rerun()

    # Display chat messages
    for msg_idx, msg in enumerate(st.session_state.ai_chat_history):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

            if msg["role"] == "assistant" and msg.get("cases"):
                with st.expander(f"📚 የተጠቀሱ አስገዳጅ ውሳኔዎች ({len(msg['cases'])} Precedents Cited & Grounded)", expanded=False):
                    for c_idx, sc in enumerate(msg["cases"], 1):
                        _cno = sc.get('case_no') or f"[ተ.ቁ {sc.get('entry_no')}]"
                        _vol = sc.get('volume', '-')
                        _pvol = f"፣ ገጽ {sc.get('page_in_vol')}" if sc.get('page_in_vol') else ""
                        _status_badge = "⚠️ የተሻረ (OVERRULED)" if sc.get('status') == 'OVERRULED' else ("⭐ 7 ዳኞች ሰበር" if ("7 ዳኞች" in str(sc.get('bench', '')) or sc.get('status') in ['PLENARY', 'PLENARY_OVERRULING']) else "✅ ጸንቶ የሚሰራ")

                        st.markdown(f"""
                        <div class="law-card">
                            <div class="head">
                                <h4>{c_idx}. መዝገብ ቁጥር {_cno} ({_vol}{_pvol})</h4>
                                <span class="chip">{_status_badge}</span>
                            </div>
                            <p style="margin: 0.25rem 0; font-size: 0.85rem; color: var(--ink-muted);">
                                <strong>አመልካች:</strong> {sc.get('applicant', '-')} | <strong>ተጠሪ:</strong> {sc.get('respondent', '-')} | <strong>ቀን:</strong> {sc.get('date', '-')}
                            </p>
                            <div class="quote">
                                <strong>አስገዳጅ የሕግ ትርጓሜ (Holding):</strong> {str(sc.get('holding'))[:280]}...
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        btn_c1, btn_c2, btn_c3 = st.columns([4, 4, 4])
                        # Action 1: Add to Draft
                        is_drafted = any(str(d.get('case_no') or d.get('entry_no')) == str(sc.get('case_no') or sc.get('entry_no')) for d in st.session_state.selected_for_draft)
                        if is_drafted:
                            btn_c1.info("✓ በአቤቱታ ረቂቅ ውስጥ ተካቷል")
                        else:
                            if btn_c1.button("➕ ወደ አቤቱታ ረቂቅ ጨምር", key=f"ai_add_{msg_idx}_{c_idx}"):
                                st.session_state.selected_for_draft.append(sc)
                                st.rerun()

                        # Action 2: Open in Precedents
                        pos = next((i for i, x in enumerate(cases_data) if str(x.get('entry_no')) == str(sc.get('entry_no'))), None)
                        if pos is not None:
                            btn_c2.button("📖 ውሳኔውን ክፈት", on_click=open_case, args=(pos,), key=f"ai_open_{msg_idx}_{c_idx}")

                        # Action 3: Direct Google Drive Link
                        _sc_docket = str(sc.get('case_no', '')).strip()
                        _raw_files = drive_dockets.get(_sc_docket) or []
                        if _raw_files:
                            rf = _raw_files[0]
                            btn_c3.markdown(f"[📄 በ Drive ክፈት ({rf['n']})](https://drive.google.com/file/d/{rf['id']}/view)")
                        else:
                            _vol_match = re.search(r'\d+', str(sc.get('volume', '')))
                            _vol_key = _vol_match.group() if _vol_match else None
                            _drive_file = volume_drive_map.get(_vol_key) if _vol_key else None
                            if _drive_file and _drive_file.get("id"):
                                btn_c3.markdown(f"[📄 በ Drive ክፈት (ቅጽ {_vol_key})](https://drive.google.com/file/d/{_drive_file['id']}/view)")
                            elif sc.get("link"):
                                btn_c3.markdown(f"[📄 በ Drive ክፈት]({sc['link']})")

                    # Export for NotebookLM Button (Option 3 Bridge)
                    dossier_text = generate_notebooklm_dossier(msg.get("question", ""), msg["content"], msg["cases"])
                    st.download_button(
                        label="📥 የዚህን ጥናት ማጠቃለያ ለ NotebookLM አውርድ (.txt)",
                        data=dossier_text,
                        file_name=f"Cassation_Research_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain",
                        key=f"dossier_dl_{msg_idx}"
                    )

    # Chat Input Handling
    user_input = st.chat_input("የሕግ ጥያቄዎን እዚህ ያስገቡ... (Ask your legal question in Amharic or English)") or chosen_prompt
    if user_input:
        if not api_key:
            st.error("⚠️ እባክዎ ከላይ የ Gemini API ቁልፍዎን ያስገቡ። (Please enter your Gemini API key above to proceed).")
        else:
            top_cases = retrieve_relevant_cases(user_input, cases_data, max_results=5)

            st.session_state.ai_chat_history.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.markdown(user_input)

            with st.chat_message("assistant"):
                with st.spinner("በ 3,246 አስገዳጅ የሰበር ውሳኔዎች እና የሕጎች ማውጫ ውስጥ በመፈለግ ላይ..."):
                    ai_response = call_gemini_counsel(user_input, top_cases, api_key)
                    st.markdown(ai_response)

            st.session_state.ai_chat_history.append({
                "role": "assistant",
                "content": ai_response,
                "cases": top_cases,
                "question": user_input
            })
            st.rerun()

# ==========================================
# TAB: AUTOMATED WORD PLEADING DRAFTER
# ==========================================
if nav == TAB_DRAFTER:
    st.subheader("የፍርድ ቤት አቤቱታ እና ማጣቀሻ አዘጋጅ (Microsoft Word Drafter)")
    st.write("በምርምር ክፍል የመረጧቸውን አስገዳጅ ውሳኔዎች በማካተት የተሟላ **የፍርድ ቤት አቤቱታ የሕግ ማጣቀሻ ሰነድ (.docx)** በቀጥታ ማመንጨት ይችላሉ።")

    st.markdown('<p class="section-label">የአቤቱታው ዝርዝር · PLEADING SETTINGS</p>', unsafe_allow_html=True)
    _p1, _p2 = st.columns([6, 4])
    court_target = _p1.text_input("የፍርድ ቤቱ ስም", key="p_court")
    pleading_type = _p2.selectbox("የአቤቱታው ዓይነት", PLEADING_TYPES, key="p_type")
    _p3, _p4 = st.columns(2)
    client_app = _p3.text_input("አመልካች (Plaintiff/Petitioner)", key="p_app")
    client_resp = _p4.text_input("ተጠሪ (Defendant/Respondent)", key="p_resp")

    st.markdown('<p class="section-label">ረቂቁን ማስቀመጥ · SAVE OR LOAD A DRAFT</p>', unsafe_allow_html=True)
    _d1, _d2 = st.columns([4, 6])
    with _d1:
        _saved = json.dumps({
            "app": "ethiopian-cassation-drafter",
            "saved": datetime.date.today().isoformat(),
            "entries": [c.get('entry_no') for c in st.session_state.selected_for_draft],
            "court": court_target, "type": pleading_type,
            "applicant": client_app, "respondent": client_resp,
        }, ensure_ascii=False, indent=2)
        st.download_button(
            "ረቂቁን አስቀምጥ (Save draft)", data=_saved.encode('utf-8'),
            file_name=f"cassation_draft_{datetime.date.today().isoformat()}.json", mime="application/json",
            disabled=not st.session_state.selected_for_draft, use_container_width=True, key="draft_save",
        )
        st.caption("የተመረጡ ውሳኔዎችን እና ዝርዝሩን በፋይል ያስቀምጣል (Saves the selected decisions and the settings as a file).")
    with _d2:
        st.file_uploader("የተቀመጠ ረቂቅ ክፈት (Load a saved draft)", type=["json"], key="draft_upload", on_change=load_draft)
    _msg = st.session_state.pop('draft_msg', None)
    if _msg and _msg[0] == 'ok':
        st.success(f"{_msg[1]} ውሳኔዎች ተጭነዋል ({_msg[1]} decisions loaded).")
    elif _msg:
        st.error("ፋይሉ ሊነበብ አልቻለም (The file could not be read). Use a file saved with 'Save draft'.")
    st.markdown("---")

    if not st.session_state.selected_for_draft:
        st.info("በአሁኑ ሰዓት የተመረጠ ውሳኔ የለም። በትር 1 (Research) ውስጥ የሚፈልጓቸውን ውሳኔዎች መርጠው **'ወደ አቤቱታ ረቂቅ ጨምር'** የሚለውን ይጫኑ።")
    else:
        st.write(f"**የተመረጡ አስገዳጅ ውሳኔዎች ({len(st.session_state.selected_for_draft)}):**")
        
        _n_sel = len(st.session_state.selected_for_draft)
        for i, sc in enumerate(st.session_state.selected_for_draft):
            c_col1, c_up, c_down, c_del = st.columns([17, 1, 1, 1])
            c_col1.write(f"**{i+1}. መ/ቁ {sc.get('case_no') or f'[Entry {sc.get('entry_no')}]'}** ({sc.get('volume')}) — {str(sc.get('legal_issue'))[:110]}")
            c_up.button("↑", key=f"up_{i}", on_click=move_draft, args=(i, -1), disabled=i == 0, help="ወደ ላይ (Move up)")
            c_down.button("↓", key=f"down_{i}", on_click=move_draft, args=(i, 1), disabled=i == _n_sel - 1, help="ወደ ታች (Move down)")
            c_del.button("×", key=f"del_{i}", on_click=remove_draft, args=(i,), help="አስወግድ (Remove)")

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
                    holding_str = f"[ይህ ውሳኔ በ 7 ዳኞች ሰበር ተሽሯል!]\n" + holding_str
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
            label="ይፋዊ የፍርድ ቤት አቤቱታ ሰነድ አውርድ (.docx)",
            data=docx_bio,
            file_name=f"Ethiopian_Cassation_Brief_{pleading_type}_{client_app[:15]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ==========================================
# TAB: SHEPARD'S OVERRULE MAP
# ==========================================
if nav == TAB_OVERRULED:
    st.subheader("የተሻሩ ውሳኔዎች እና የ 7 ዳኞች አስገዳጅ ውሳኔዎች ማውጫ (Shepard's Citator)")
    st.write("በአዋጅ ቁጥር 1234/2013 አንቀጽ 10(4) እና 26(1) መሠረት በሰባት (7) ዳኞች የተሰየመው የሰበር ሰሚ ችሎት የቀድሞ ባለ 5 ዳኞች ውሳኔዎችን የመሻር እና የመለወጥ ስልጣን ተሰጥቶታል፡፡")

    ov_cases = [c for c in cases_data if c.get('status') == 'OVERRULED']
    st.markdown(f"#### የተሻሩ/የተለወጡ ውሳኔዎች ዝርዝር ({len(ov_cases)})")
    
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
    st.markdown("#### የቀድሞ ውሳኔዎችን የሻሩ እና ያረጋገጡ የ 7 ዳኞች ታሪካዊ ውሳኔዎች")
    
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

    st.info("**ልዩ የሕግ ማስታወሻ:** በሰባት ዳኞች ሰበር መ/ቁ 237423 (የካቲት 29/2016 ዓ.ም) በተሰጠ ውሳኔ መ/ቁ 48850፣ 58008፣ 74237፣ 94913፣ 102640 እና 218957 የቀረበውን የመሻር ጥያቄ ውድቅ በማድረግ ውሳኔዎቹ ጸንተው እንዲቀጥሉ ወስኗል።")

# ==========================================
# TAB: 20-YEAR JUDICIAL ANALYTICS
# ==========================================
if nav == TAB_ANALYTICS:
    st.subheader("የ 20 ዓመታት የሰበር አስገዳጅ ውሳኔዎች ትንተና (1997 - 2017 ዓ.ም)")
    
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
# TAB: STATUTORY CITATOR (FORWARD & REVERSE)
# ==========================================
if nav == TAB_CITATOR:
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
# TAB: LAWS LIBRARY BY SUBJECT (3,696 DISTINCT ENACTMENTS)
# ==========================================
if nav == TAB_LIBRARY:
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
            law_search_kw = st.text_input("የሕግ ቁጥር ወይም ቃል ፍለጋ:", key="lib_kw", placeholder="ምሳሌ፡ 1333, 1094, ታክስ, መሬት, ግዥ...")

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
            _words = normalize_amharic(law_search_kw).split()
            def _law_text(x):
                return normalize_amharic(" ".join(str(x.get(f, '')) for f in ('title_am', 'title_en', 'no', 'year', 'year_gc')))
            filtered_catalog = [x for x in filtered_catalog if all(w in _law_text(x) for w in _words)]

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

# Keep the address bar in step with the decision on screen, so the page address can be shared
_want = {}
if shown_case is not None:
    if shown_case.get('case_no'):
        _want['case'] = str(shown_case.get('case_no'))
    _want['e'] = str(shown_case.get('entry_no'))
if dict(st.query_params) != _want:
    st.query_params.from_dict(_want)
