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

# Session state initialization
if 'selected_for_draft' not in st.session_state:
    st.session_state.selected_for_draft = []
if 'active_case_idx' not in st.session_state:
    st.session_state.active_case_idx = 0

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

filtered_cases = [c for c in cases_data if filter_case(c)]

# Tabs Interface
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "የሰበር ውሳኔዎች (Precedents)",
    "አቤቱታ አዘጋጅ (Drafter)",
    "የተሻሩ ውሳኔዎች (Overruled)",
    "የሕግ ትንተና (Analytics)",
    "የሥልጣን ትስስር (Citator)",
    "የሕጎች ማውጫ (Library)"
])

# ==========================================
# TAB 1: SPLIT SCREEN RESEARCH & INSPECTOR
# ==========================================
with tab1:
    st.write(f"**የተገኙ ውሳኔዎች:** `{len(filtered_cases):,}` / {total_cases:,}")
    
    col_left, col_right = st.columns([5, 7])
    
    with col_left:
        st.subheader("📋 የውሳኔዎች ዝርዝር")
        
        df_preview = pd.DataFrame([
            {
                "መ/ቁ": c.get('case_no') or f"[Entry {c.get('entry_no')}]",
                "ቅጽ": c.get('volume'),
                "ዘርፍ": c.get('legal_subject', '')[:16],
                "ማረጋገጫ": "✓ የተረጋገጠ" if "Verified" in str(c.get('verification_status', '')) else ("⚠️ ማጣራት" if "Not confirmed" in str(c.get('verification_status', '')) else "+ የታከለ"),
                "ደረጃ": "⚠️ የተሻረ" if c.get('status') == 'OVERRULED' else ("⭐ 7 ዳኞች" if ("7 ዳኞች" in str(c.get('bench', '')) or c.get('status') in ['PLENARY', 'PLENARY_OVERRULING']) else "መደበኛ")
            }
            for c in filtered_cases
        ])
        
        if not df_preview.empty:
            selected_row = st.selectbox(
                "ለመመርመር ውሳኔ ይምረጡ:",
                range(len(filtered_cases)),
                format_func=lambda i: f"መ/ቁ {filtered_cases[i].get('case_no') or f'[Entry {filtered_cases[i].get('entry_no')}]'} | {filtered_cases[i].get('volume')} | {filtered_cases[i].get('legal_subject')[:18]}"
            )
            st.session_state.active_case_idx = selected_row
            st.dataframe(df_preview, use_container_width=True, height=450)
        else:
            st.warning("የተፈለገውን መስፈርት የሚያሟላ ውሳኔ አልተገኘም።")

    with col_right:
        if filtered_cases:
            active_c = filtered_cases[st.session_state.active_case_idx]
            
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
            st.info(active_c.get('legal_issue', 'የሕግ ትርጓሜ'))

            # Ratio Decidendi
            st.markdown("##### ⚖️ አስገዳጅ የሕግ ትርጓሜ (Authoritative Holding)")
            st.markdown('<blockquote class="holding">' + html.escape(str(active_c.get('holding', ''))).replace('\n', '<br>') + '</blockquote>', unsafe_allow_html=True)

            # Statutory Provisions
            if active_c.get('provisions'):
                st.markdown(f"**📜 የተተረጎሙ ድንጋጌዎች:** `{active_c.get('provisions')}`")

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
with tab2:
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
with tab3:
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
with tab4:
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
with tab5:
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
with tab6:
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
                    "ጠቅላላ ሰነድ (General Documents)"
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
                    st.write("")
