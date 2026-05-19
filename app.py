import streamlit as st
import pandas as pd
import re
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time
import os

# --- 1. INITIAL CONFIG ---
st.set_page_config(
    page_title="NexMap Extractor Pro", 
    page_icon="📊", 
    initial_sidebar_state="expanded"
)

# --- 2. CUSTOM CSS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        font-size: 14px; 
    }
    
    .main { background-color: #f4f6f9; }
    
    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #ffffff 0%, #fcfcfc 100%);
        padding: 15px 20px; 
        border-radius: 10px; 
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04);
        border-left: 4px solid #EE2D24;
        border-top: 1px solid #f0f0f0;
        border-right: 1px solid #f0f0f0;
        border-bottom: 1px solid #f0f0f0;
        transition: transform 0.2s ease-in-out;
    }
    
    div[data-testid="stMetricValue"] > div {
        font-size: 1.8rem !important; 
        font-weight: 600;
    }
    
    div[data-testid="stMetricLabel"] {
        font-weight: 500;
        color: #666 !important;
        font-size: 0.9rem !important;
    }

    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-size: 14px !important; 
        font-weight: 600;
        background: linear-gradient(135deg, #EE2D24 0%, #cc1e16 100%) !important;
        color: white !important;
        border: none !important;
        height: 3em; 
        box-shadow: 0 4px 6px rgba(238, 45, 36, 0.2);
    }
    
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 12px rgba(238, 45, 36, 0.3);
    }
    
    .header-container {
        padding-bottom: 5px;
        border-bottom: 1px solid #eaeaea;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. HEADER LAYOUT ---
st.markdown('<div class="header-container">', unsafe_allow_html=True)
header_col1, header_col2 = st.columns([1, 6]) 
with header_col1:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/6/6f/Telkom_Indonesia_logo.svg/1200px-Telkom_Indonesia_logo.svg.png", width=90) 
with header_col2:
    st.markdown("<h2 style='color: #EE2D24; margin-bottom: -10px; font-weight: 800;'>NEXMAP MO EXTRACTOR <span style='color: #ccc; font-size: 16px; font-weight: 400;'>| Pro Edition</span></h2>", unsafe_allow_html=True) 
    st.markdown("<p style='color: #555; font-size: 14px;'>Automatic Field Data Sync — <b>PT Telkom Infrastruktur Indonesia</b></p>", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# --- 4. FUNGSI REGEX ---
def extract_track_id(text):
    if pd.isna(text): return ""
    match = re.search(r'-((?:MO|AO|DO|PDA)[ki][^_]+)_', str(text)) 
    if match: return match.group(1)
    match_fallback = re.search(r'((?:MO|AO|DO|PDA)[ki][a-zA-Z0-9]+)', str(text))
    return match_fallback.group(1) if match_fallback else ""

def extract_homepass_id(text):
    if pd.isna(text): return ""
    text_str = str(text)
    match = re.search(r'[-_ ]?(MO|AO|DO|PDA)[ki]', text_str)
    if match: return text_str[:match.start()].strip()
    return ""

# --- 5. GOOGLE SHEETS CONNECTION (WITH DEBUG) ---
@st.cache_resource(ttl=1800)
def get_gsheets_connection():
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        
        # URUTAN 1: Cek File Lokal
        if os.path.exists('credentials.json'):
            st.sidebar.info("📁 Mode: credentials.json lokal")
            credentials = Credentials.from_service_account_file('credentials.json', scopes=scopes)
            
        # URUTAN 2: Cek Secrets (Streamlit Cloud)
        elif "gcp_service_account" in st.secrets:
            st.sidebar.info("🔑 Mode: Streamlit Secrets")
            creds_info = dict(st.secrets["gcp_service_account"])
            creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
            credentials = Credentials.from_service_account_info(creds_info, scopes=scopes)
            
        # URUTAN 3: Tidak ditemukan — tampilkan semua key yang ada
        else:
            st.error("❌ Secrets tidak ditemukan!")
            st.sidebar.error("Keys tersedia: " + str(list(st.secrets.keys())))
            return None
            
        gc = gspread.authorize(credentials)
        spreadsheet_id = "14dcsgUepQqbfSr4y8h4Si1niy-VC2XbmRI87oZ9CmUI"
        doc = gc.open_by_key(spreadsheet_id)
        doc.sheet1.get('A1')
        return doc
        
    except Exception as e:
        st.error(f"❌ Connection Error: {e}")
        st.exception(e)  # Tampilkan full traceback
        return None

# --- 6. SIDEBAR ---
with st.sidebar:
    st.subheader("⚙️ System Control")
    st.info("Otomatis mendistribusikan data NexMap ke tab **MO, AO, DO, dan PDA**.")
    
    if st.button("🧪 TEST CONNECTION", use_container_width=True):
        get_gsheets_connection.clear()  # Force clear cache sebelum test
        doc = get_gsheets_connection()
        if doc: st.success("✅ Connected!")
        else: st.error("❌ Failed!")
    
    st.divider()
    if st.button("🔄 Clear Session Cache", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

# --- 7. CORE LOGIC ---
st.markdown("#### 📂 Upload Panel")
uploaded_file = st.file_uploader("Drop file Excel NexMap", type=["xlsx", "xls"], label_visibility="collapsed")

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📊 Analysis Overview")
        
        m1, m2, m3 = st.columns(3)
        with m1: st.metric("Total Data Excel", f"{len(df_raw)} Baris")
        with m2: st.metric("Status File", "Valid Excel", delta="Ready")
        with m3: st.metric("Waktu Sistem", datetime.now().strftime("%H:%M WIB"))

        st.markdown("<hr style='border:1px dashed #ddd; margin: 15px 0;'>", unsafe_allow_html=True)

        doc = get_gsheets_connection()
        if not doc: st.stop()
        
        tab_mapping = {
            'MO': 'DATA MO',
            'AO': 'DATA AO',
            'DO': 'DATA DO',
            'PDA': 'DATA PDA'
        }

        if st.button(f"🚀 JALANKAN DISTRIBUSI KE 4 TAB (MO, AO, DO, PDA)", use_container_width=True):
            with st.spinner('Memproses dan mendistribusikan data...'):
                
                df_extracted = pd.DataFrame()
                df_extracted['HOMEPAS ID'] = df_raw['SC Order No/Track ID/CSRM No'].apply(extract_homepass_id)
                df_extracted['ORDER DATE'] = df_raw['Date Created'].astype(str)
                df_extracted['WONUM'] = df_raw['Workorder'].astype(str).str.strip()
                df_extracted['TRACK ID'] = df_raw['SC Order No/Track ID/CSRM No'].apply(extract_track_id)
                df_extracted['ND INTERNET'] = df_raw['Service No.'].fillna("")
                df_extracted['NAMA PELANGGAN'] = df_raw['Customer Name'].fillna("")
                df_extracted['ALAMAT'] = df_raw['Address'].fillna("")
                df_extracted['STO'] = df_raw['Workzone'].fillna("")
                df_extracted['NO HP'] = df_raw['Contact Number'].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
                df_extracted = df_extracted[df_extracted['TRACK ID'] != ""]
                df_extracted['KCONTACT'] = (
                    df_extracted['HOMEPAS ID'] + ";" + 
                    df_extracted['TRACK ID'] + ";" + 
                    df_extracted['NAMA PELANGGAN'] + ";" + 
                    df_extracted['NO HP']
                )

                total_inserted = 0
                
                for prefix, sheet_name in tab_mapping.items():
                    df_tab = df_extracted[df_extracted['TRACK ID'].str.startswith(prefix)].copy()
                    
                    if len(df_tab) == 0:
                        continue
                    
                    try:
                        sheet = doc.worksheet(sheet_name)
                    except Exception as e:
                        st.error(f"❌ Tab '{sheet_name}' tidak ditemukan.")
                        continue
                        
                    raw_headers = sheet.row_values(1)
                    sheet_headers = [h.strip() for h in raw_headers if h.strip() != ""]
                    all_values = sheet.get_all_values()
                    last_filled_row = 0
                    
                    for idx, row in enumerate(all_values):
                        if len(row) > 1:
                            last_filled_row = idx + 1
                    
                    wonum_idx = 1
                    if "WONUM" in sheet_headers:
                        wonum_idx = sheet_headers.index("WONUM")
                        
                    existing_wonums = []
                    if last_filled_row > 1:
                        existing_wonums = [row[wonum_idx].strip() for row in all_values[1:last_filled_row] if len(row) > wonum_idx]
                        
                    df_new = df_tab[~df_tab['WONUM'].isin(existing_wonums)].copy()

                    if len(df_new) > 0:
                        for col in sheet_headers:
                            if col not in df_new.columns: 
                                df_new[col] = ""
                        
                        df_final = df_new[sheet_headers]
                        data_to_upload = df_final.fillna("").astype(str).values.tolist()
                        start_row = last_filled_row + 1
                        range_name = f"A{start_row}"
                        
                        try:
                            sheet.update(range_name=range_name, values=data_to_upload)
                        except TypeError:
                            sheet.update(data_to_upload, range_name)
                        
                        total_inserted += len(df_new)
                        st.success(f"✅ Masuk: **{len(df_new)} data** ke tab **{sheet_name}**.")
                        
                        with st.expander(f"Pratinjau Data {sheet_name}:"):
                            st.dataframe(df_final.head(3), use_container_width=True)
                    else:
                        st.info(f"ℹ️ Tidak ada data baru untuk tab **{sheet_name}**.")

                if total_inserted > 0:
                    st.balloons()
                    st.success(f"🎉 **SELESAI! Total {total_inserted} data baru berhasil didistribusikan.**")
                else:
                    st.warning("Semua data sudah ada di dalam sistem (Duplikat).")

    except Exception as e:
        st.error(f"❌ Kesalahan Sistem: {e}")
        st.exception(e)

# --- 8. FOOTER ---
st.divider()
st.markdown(f"<center style='color: #999; font-size: 12px;'>© {datetime.now().year} PT Telkom Infrastruktur Indonesia | Developed by Gramaldy</center>", unsafe_allow_html=True)