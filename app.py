import streamlit as st
import pandas as pd
import re
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import time
import os

# --- 1. INITIAL CONFIG (Premium UI) ---
st.set_page_config(
    page_title="NexMap MO Extractor", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. CUSTOM CSS (Modern UI) ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    div[data-testid="stMetric"] {
        background-color: white;
        padding: 15px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-left: 5px solid #EE2D24;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: bold;
        background-color: #EE2D24 !important;
        color: white !important;
        border: none !important;
        height: 3em;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. HEADER LAYOUT ---
header_col1, header_col2 = st.columns([1, 6])
with header_col1:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/6/6f/Telkom_Indonesia_logo.svg/1200px-Telkom_Indonesia_logo.svg.png", width=100)
with header_col2:
    st.markdown("<h1 style='color: #EE2D24; margin-bottom: 0;'>NEXMAP MO EXTRACTOR</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #666; font-size: 1.1em;'>Automatic Field Data Sync — PT Telkom Infrastruktur Indonesia</p>", unsafe_allow_html=True)

# --- 4. FUNGSI REGEX (SAMA PERSIS) ---
def extract_track_id(text):
    if pd.isna(text): return ""
    match = re.search(r'-(MO[ki][^_]+)_', str(text)) 
    if match: return match.group(1)
    match_fallback = re.search(r'(MO[ki][a-zA-Z0-9]+)', str(text))
    return match_fallback.group(1) if match_fallback else ""

def extract_homepass_id(text):
    if pd.isna(text): return ""
    text_str = str(text)
    match = re.search(r'[-_ ]?MO[ki]', text_str)
    if match: return text_str[:match.start()].strip()
    return ""

# --- 5. FIXED GOOGLE SHEETS CONNECTION ---
@st.cache_resource(ttl=1800)
def get_gsheets_connection():
    """Fixed connection dengan proper error handling"""
    try:
        # CEK FILE credentials.json
        if not os.path.exists('credentials.json'):
            st.error("❌ **credentials.json TIDAK DITEMUKAN**")
            st.info("📥 Download dari Google Cloud Console → Service Accounts → JSON Key")
            return None
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = Credentials.from_service_account_file('credentials.json', scopes=scopes)
        gc = gspread.authorize(credentials)
        
        spreadsheet_id = "14dcsgUepQqbfSr4y8h4Si1niy-VC2XbmRI87oZ9CmUI"
        doc = gc.open_by_key(spreadsheet_id)
        
        # Test connection
        doc.sheet1.get('A1')
        return doc
        
    except Exception as e:
        if "invalid_grant" in str(e) or "Invalid JWT" in str(e):
            st.error(f"❌ **AUTH ERROR**: {e}")
            st.info("""
            🔧 **FIX SEKARANG:**
            1. Download ulang `credentials.json` 
            2. Share Google Sheets ke email service account (Editor)
            3. Enable Google Sheets API & Drive API
            """)
        else:
            st.error(f"❌ Connection Error: {e}")
        return None

# --- 6. SIDEBAR (DITAMBAH TEST BUTTON) ---
with st.sidebar:
    st.subheader("⚙️ System Control")
    st.info("Sistem ini akan otomatis mengirim hasil ekstraksi NexMap ke tab **DATA**.")
    
    # TEST CONNECTION BUTTON
    if st.button("🧪 TEST CONNECTION", use_container_width=True):
        doc = get_gsheets_connection()
        if doc:
            st.success("✅ Google Sheets Connected!")
        else:
            st.error("❌ Connection Failed!")
    
    st.divider()
    if st.button("Clear Session Cache", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

# --- 7. CORE LOGIC (LOGIKA SAMA PERSIS) ---
uploaded_file = st.file_uploader("📂 Pilih file Excel NexMap", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # Metrics (SAMA PERSIS)
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Total Data Excel", f"{len(df_raw)} Baris")
        with m2:
            st.metric("Status File", "Valid", delta="Ready")
        with m3:
            st.metric("Waktu Sistem", datetime.now().strftime("%H:%M"))

        # FIXED GSHEETS SETUP
        doc = get_gsheets_connection()
        if not doc:
            st.stop()
        
        spreadsheet_id = "14dcsgUepQqbfSr4y8h4Si1niy-VC2XbmRI87oZ9CmUI"
        target_sheet_name = "DATA"
        sheet = doc.worksheet(target_sheet_name)

        if st.button(f"🚀 SINKRONISASI KE TAB {target_sheet_name}", use_container_width=True):
            with st.spinner('Menganalisis baris kosong dan membersihkan duplikat...'):
                
                # --- A. PEMBERSIHAN HEADER (SAMA PERSIS) ---
                raw_headers = sheet.row_values(1)
                sheet_headers = [h.strip() for h in raw_headers if h.strip() != ""]
                
                # --- B. DETEKSI BARIS KOSONG ASLI (SAMA PERSIS) ---
                all_values = sheet.get_all_values()
                last_filled_row = 0
                for idx, row in enumerate(all_values):
                    if len(row) > 1 and row[1].strip() != "":
                        last_filled_row = idx + 1

                # --- C. MAPPING DATA (SAMA PERSIS) ---
                df_extracted = pd.DataFrame()
                df_extracted['HOMEPAS ID'] = df_raw['SC Order No/Track ID/CSRM No'].apply(extract_homepass_id)
                df_extracted['ORDER DATE'] = df_raw['Date Created'].astype(str)
                df_extracted['WONUM'] = df_raw['Workorder'].astype(str).str.strip()
                df_extracted['TRACK ID'] = df_raw['SC Order No/Track ID/CSRM No'].apply(extract_track_id)
                df_extracted['ND INTERNET'] = df_raw['Service No.'].fillna("")
                df_extracted['NAMA PELANGGAN'] = df_raw['Customer Name'].fillna("")
                df_extracted['ALAMAT'] = df_raw['Address'].fillna("")
                df_extracted['STO'] = df_raw['Workzone'].fillna("")
                df_extracted['NO HP'] = df_raw['Contact Number'].fillna("")

                # Filtering MOk/MOi (SAMA PERSIS)
                df_extracted = df_extracted[df_extracted['TRACK ID'] != ""]
                
                # Cek Duplikat WONUM (SAMA PERSIS)
                existing_wonums = []
                if last_filled_row > 1:
                    existing_wonums = [row[1].strip() for row in all_values[1:last_filled_row] if len(row) > 1]
                
                df_new = df_extracted[~df_extracted['WONUM'].isin(existing_wonums)].copy()

                if len(df_new) == 0:
                    st.warning(f"🚫 Tidak ada data baru. Semua data sudah ada di tab {target_sheet_name}.")
                else:
                    # Sesuaikan urutan kolom dengan GSheets (SAMA PERSIS)
                    for col in sheet_headers:
                        if col not in df_new.columns: 
                            df_new[col] = ""
                    
                    df_final = df_new[sheet_headers]
                    data_to_upload = df_final.fillna("").values.tolist()

                    # Penentuan Range (SAMA PERSIS)
                    start_row = last_filled_row + 1
                    range_name = f"A{start_row}"
                    
                    # Update data
                    sheet.update(range_name=range_name, values=data_to_upload)
                    
                    st.balloons()
                    st.success(f"✅ Berhasil! {len(df_new)} data masuk ke tab {target_sheet_name} mulai baris {start_row}.")
                    st.dataframe(df_final.head(), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Kesalahan: {e}")
        st.info("Tips: Jika data masih melompat, hapus baris kosong (2-2869) di GSheets secara manual dengan klik kanan > Delete Rows.")

# --- 8. FOOTER ---
st.divider()
st.caption(f"© {datetime.now().year} Developed by Gramaldy | NexMap MO Extractor v2.2 (Fixed)")