import streamlit as st
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

# Custom CSS for Ethiopic font & clean legal theme
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Ethiopic:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Noto Sans Ethiopic', 'Inter', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 1rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .overruled-banner {
        background-color: #fee2e2;
        border-left: 5px solid #ef4444;
        color: #991b1b;
        padding: 0.85rem;
        border-radius: 0.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .plenary-banner {
        background-color: #fef3c7;
        border-left: 5px solid #d97706;
        color: #78350f;
        padding: 0.85rem;
        border-radius: 0.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .unconfirmed-banner {
        background-color: #fefce8;
        border-left: 5px solid #eab308;
        color: #854d0e;
        padding: 0.85rem;
        border-radius: 0.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
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

# Session state initialization
if 'selected_for_draft' not in st.session_state:
    st.session_state.selected_for_draft = []
if 'active_case_idx' not in st.session_state:
    st.session_state.active_case_idx = 0

# App Header
st.markdown("""
<div class="main-header">
    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 1rem;">
        <div>
            <h1 style="margin: 0; font-size: 1.6rem; font-weight: 800; color: white;">
                ⚖️ የኢትዮጵያ ሰበር አስገዳጅ የሕግ ትርጓሜዎች መፈለጊያ እና ረዳት (Desktop App v2)
            </h1>
            <p style="margin: 0.3rem 0 0 0; font-size: 0.85rem; color: #94a3b8;">
                የፌዴራል ጠቅላይ ፍርድ ቤት ሰበር ሰሚ ችሎት ውሳኔዎች • ቅጽ 1 – 28፣ 7 ዳኞች ሰበር እና የ2015–2016 ስብስብ (3,246 አስገዳጅ ውሳኔዎች)
            </p>
        </div>
        <div style="text-align: right;">
            <span style="background: rgba(16, 185, 129, 0.2); color: #34d399; padding: 0.3rem 0.8rem; border-radius: 9999px; font-size: 0.75rem; font-weight: 700; border: 1px solid rgba(16, 185, 129, 0.4);">
                🟢 Local Intelligence Engine Active (v2)
            </span>
        </div>
    </div>
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
tab1, tab2, tab3, tab4 = st.tabs([
    "🔍 የሰበር ውሳኔዎች ምርምር (Research & Split Inspector)",
    "📝 የፍርድ ቤት አቤቱታ ረዳት (Word Pleading Drafter)",
    "⚖️ የተሻሩ ውሳኔዎች ማውጫ (Shepard's Overrule Map)",
    "📊 የ 20 ዓመታት የሕግ ትንተና (Analytics & Trends)"
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
                    ⚠️ <strong>የተሻረ ውሳኔ (OVERRULED PRECEDENT):</strong><br>
                    {active_c.get('status_note')}
                </div>
                """, unsafe_allow_html=True)
            elif active_c.get('status') == 'PLENARY_OVERRULING':
                st.markdown(f"""
                <div class="plenary-banner">
                    ⭐ <strong>የሰበር ምልዓተ-ጉባዔ ታሪካዊ ውሳኔ (PLENARY OVERRULING):</strong><br>
                    {active_c.get('status_note')}
                </div>
                """, unsafe_allow_html=True)
            elif "7 ዳኞች" in str(active_c.get('bench', '')) or active_c.get('status') == 'PLENARY':
                st.markdown(f"""
                <div class="plenary-banner">
                    ⭐ <strong>በ 7 ዳኞች ምልዓተ-ጉባዔ የተሰጠ አስገዳጅ ውሳኔ</strong>
                </div>
                """, unsafe_allow_html=True)

            # Verification Warning
            if "Not confirmed" in str(active_c.get('verification_status', '')):
                st.markdown("""
                <div class="unconfirmed-banner">
                    ⚠️ <strong>በህትመቱ አልተረጋገጠም:</strong> ይህ መዝገብ ቁጥር ወይም ቅጽ በታተመው መጽሐፍ ማውጫ ውስጥ በግልጽ አልተረጋገጠም፤ ከመጥቀስዎ በፊት ህትመቱን ይመልከቱ።
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
            st.markdown(f"```text\n{active_c.get('holding', '')}\n```")

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
        color_continuous_scale='Blues',
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
