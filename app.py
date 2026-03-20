import pytz
from google.genai import types
import io
import streamlit as st
import pandas as pd
import json
from google import genai
from supabase import create_client, Client

# --- 1. CONFIGURATION ---
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    GENAI_API_KEY = st.secrets["GENAI_API_KEY"]
except:
    # สำหรับใช้รันในเครื่องตัวเอง (Local) ถ้ายังไม่ได้ตั้งค่า Secrets
    st.error("❌ ไม่พบ API Keys ในระบบ Secrets กรุณาตั้งค่าที่ Settings > Secrets")
    st.stop()

# Initialize Clients
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
genai_client = genai.Client(api_key=GENAI_API_KEY)
MODEL_ID = "models/gemini-3.1-flash-lite-preview"

st.set_page_config(page_title="DHL Booking Cloud Extractor", layout="wide")

# --- 2. CSS CUSTOM THEME ---
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #FFCC00 0%, #FFD700 50%, #ba9500 100%); }
    .block-container { background-color: white; padding: 40px; border-radius: 25px; box-shadow: 0 15px 35px rgba(0,0,0,0.3); border: 6px solid #D40511; margin-top: 20px; margin-bottom: 20px; }
    .stDataFrame, div[data-testid="stTable"], .pinned-row-container { background-color: #f0f2f6 !important; border-radius: 10px; padding: 10px; box-shadow: inset 2px 2px 5px rgba(0,0,0,0.05); }
    .stButton>button { background-color: #D40511; color: white; border-radius: 10px; border: none; box-shadow: 0 4px #990000; transition: 0.2s; width: 100%; }
    .stButton>button:hover { background-color: #ff0000; transform: translateY(-2px); box-shadow: 0 6px #990000; }
</style>
""", unsafe_allow_html=True)

# --- UI TITLE ---
container_html = """
    <div style="display: flex; align-items: center; margin-bottom: 20px; padding: 15px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 10px solid #FFCC00;">
        <div style="width: 140px; height: 80px; background-color: #FFCC00; border: 3px solid #333; border-radius: 6px; display: flex; justify-content: center; align-items: center; margin-right: 25px; position: relative; box-shadow: 4px 4px 0px #ba9500; flex-shrink: 0;">
            <span style="color: #D40511; font-family: 'Arial Black', sans-serif; font-size: 32px; font-weight: 900; letter-spacing: -2px; z-index: 2;">DHL</span>
        </div>
        <div style="flex-grow: 1;">
            <h1 style="margin: 0; color: #333; font-size: 30px; line-height: 1.2;">DSC: CTC FG Export</h1>
            <p style="margin: 0; color: #666; font-size: 18px;">Booking Cloud Extractor (Vision Engine)</p>
        </div>
    </div>
"""

# --- FUNCTIONS ---

def extract_info_from_pdf_vision(file_bytes):
    """ส่ง PDF ทั้งก้อนให้ AI (Vision) อ่านเหมือนหน้าแชท"""
    prompt = """
    You are an expert DHL Logistics Analyst specializing in Export Operations.
    Extract shipping information from this PDF into JSON format.

    Required fields: booking_no, fcl_or_lcl, by_air_or_sea, country, port_of_destination, liner_name, vessel_name,
    no_container, no_pallet, liner_cutoff, vgm_cutoff, si_cutoff, return_date_1st, cy_date, cy_at, etd, eta,
    return_place, container_type, paperless_code.
    
    Rules: 
    - cy_date: means 1st date that i can receive containers. Look for 'Empty Pick up date' or date to pick up empty container.
    - return_date_1st: refer to '1st return date' or 'Turn-In Date'.
    - country: 
  1. First, try to extract the final destination country from 'CONSIGNEE', 'NOTIFY', or 'PL. OF DELIVERY'. 
  2. If these fields are missing, OR if the address found belongs to the origin country (e.g., Thailand), you MUST infer the destination country based on the 'Port of Discharge' (e.g., if Port of Discharge is "FOS SUR MER", output "France"). 
  3. STRICTLY DO NOT output the origin/shipper's country (e.g., "Thailand").
    - Dates: dd/mm/yyyy. Cut-offs: dd/mm/yyyy hh:mm.
    - fcl_or_lcl: "FCL" or "LCL". by_air_or_sea: "Air" or "Sea".
    - ETD must be less than ETA.
    - If any field is not found, use null. Return ONLY JSON.
    - For LCL shipment, you may take time to process because the format is informal.
    Sometimes, they use day instead of date so you may find out the ref date in the file first then calculate again.
    - liner_cutoff: Look for 'Liner Cut-off', 'Gate Closing', 'Closing Date', or 'Last Load'.
    - If a cut-off is mentioned as a day of the week (e.g., "THU"), calculate the actual date based on the document date or ETD. 
    - Example: If the document date is 17/02/2026 (Tue) and Last Load is "THU", the liner_cutoff should be 19/02/2026.
    - The container type/s should be filled only "40HC" or "20GP"
    - si_cutoff: Look for 'SI Cut-off', 'Document Close Date', 'Doc Cut-off', or 'Shipping Particular Cut-off'.
    - If a cut-off is mentioned as a day of the week, calculate the date relative to the 'Run Date' or 'Date of Issue' in the document.
    - paperless_code: FIRST, actively search for the exact 4-digit number written explicitly next to "PAPERLESS CODE"
      in the document (e.g., 2836). ONLY if it is completely missing from the document, then you may use fallback logic based on the terminal. 
    - vessel_name: Use the first vessel/voyage mentioned (e.g., DP WORLD JEDDAH). 
      If there is a connecting voyage, you may include both as "First Vessel / Connecting Vessel".
    - Validation Rule: ETD must always be an earlier date than ETA.
    - booking_no: The document contains multiple reference numbers (e.g., 'Carrier Ref' and 'Booking Ref'). You MUST extract ONLY the 'Carrier Ref' (the one issued by the shipping line, usually numeric). DO NOT extract the 'Booking Ref'.
    """
    try:
        # 2. ปรับโครงสร้างการส่งข้อมูลใหม่ให้ถูก format
        response = genai_client.models.generate_content(
            model=MODEL_ID,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_text(text=prompt),
                        types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
                    ]
                )
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0,
                seed=999
            )
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"AI Extraction Error: {e}")
        return None

def save_to_cloud(data_list):
    try:
        df_temp = pd.DataFrame(data_list)
        df_temp = df_temp.replace({pd.NA: None, float('nan'): None})
        df_temp = df_temp.where(pd.notnull(df_temp), None)
        
        # กรองเอาเฉพาะที่มีเลข Booking
        df_clean = df_temp.dropna(subset=['booking_no'])
        df_unique = df_clean.drop_duplicates(subset=['booking_no'], keep='last')
        
        supabase.table("bookings").upsert(df_unique.to_dict(orient='records')).execute()
        supabase.table("booking_revisions").insert(df_temp.to_dict(orient='records')).execute()
        return True
    except Exception as e:
        st.error(f"Database Error: {e}")
        return False

def format_thai_timezone(df, column_name):
    if column_name in df.columns:
        try:
            df[column_name] = pd.to_datetime(df[column_name])
            df[column_name] = df[column_name].dt.tz_convert('Asia/Bangkok')
            df[column_name] = df[column_name].dt.strftime('%d/%m/%Y %H:%M')
        except: pass
    return df

def get_excel_download_link(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- 3. UI: UPLOAD ---
st.markdown(container_html, unsafe_allow_html=True)

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

uploaded_files = st.file_uploader("ลากไฟล์ PDF มาวางที่นี่ (วางได้มากกว่า 1 ไฟล์)", type="pdf", accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}")

if uploaded_files:
    all_extracted_data = []
    progress_bar = st.progress(0)
    for i, file in enumerate(uploaded_files):
        with st.spinner(f"กำลังสแกน: {file.name}"):
            data = extract_info_from_pdf_vision(file.read())
            if data:
                if isinstance(data, dict): data = [data]
                for item in data:
                    item['source_file'] = file.name
                    all_extracted_data.append(item)
        progress_bar.progress((i + 1) / len(uploaded_files))

    if all_extracted_data:
        if save_to_cloud(all_extracted_data):
            st.success(f"🎉 บันทึก {len(all_extracted_data)} รายการเรียบร้อย")
            st.session_state['uploader_key'] += 1
            st.rerun()

# --- 4. UI: LIVE VIEW ---
st.divider()
st.subheader("📊 รายการ Booking ทั้งหมด (Live View)")
try:
    res = supabase.table("bookings").select("*").order("updated_at", desc=True).execute()
    if res.data:
        df_live = pd.DataFrame(res.data)
        df_live = format_thai_timezone(df_live, 'updated_at')
        
        my_columns = [
            "booking_no", "fcl_or_lcl", "by_air_or_sea", "country", "port_of_destination", 
            "liner_name", "vessel_name", "no_container", "container_type", "no_pallet",
            "etd", "eta", "liner_cutoff", "vgm_cutoff", "si_cutoff",
            "cy_date", "cy_at", "return_date_1st", "return_place", "paperless_code", "updated_at"
        ]
        
        existing_cols = [c for c in my_columns if c in df_live.columns]
        df_sorted = df_live[existing_cols]
        
        search = st.text_input("🔍 ค้นหา...")
        if search:
            mask = df_sorted.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            df_to_show = df_sorted[mask]
        else:
            df_to_show = df_sorted

        if not df_to_show.empty:
            df_to_show.index = range(1, len(df_to_show) + 1)
            st.dataframe(df_to_show, width='stretch')
            st.download_button("📥 Export Excel", get_excel_download_link(df_to_show), "Current_Bookings.xlsx")
    else:
        st.info("📌 ยังไม่มีข้อมูล")
except Exception as e:
    st.error(f"Load Error: {e}")

# --- 5. UI: HISTORY SECTION ---
st.divider()
st.subheader("📜 ประวัติการบันทึกย้อนหลัง (Revision Logs)")
if st.button("🔍 โหลดประวัติทั้งหมด"):
    st.session_state['show_history'] = True

if st.session_state.get('show_history'):
    try:
        # ดึงข้อมูลจากตาราง revision
        rev_res = supabase.table("booking_revisions").select("*").order("created_at", desc=True).execute()
        
        if rev_res.data:
            df_rev = pd.DataFrame(rev_res.data)
            df_rev = format_thai_timezone(df_rev, 'created_at')
            
            # จัดเรียงคอลัมน์ให้เหมือนหน้า Live View ตามตัวแปร my_columns ที่คุณตั้งไว้
            history_cols = [c for c in my_columns if c in df_rev.columns] + ["created_at"]
            df_rev_sorted = df_rev[history_cols]
            
            # --- เพิ่มช่องค้นหาสำหรับประวัติย้อนหลัง ---
            search_hist = st.text_input("🔍 ค้นหาในประวัติ...", key="search_history") 
            
            if search_hist:
                # สร้าง Filter กรองข้อมูลจากทุกคอลัมน์
                mask_hist = df_rev_sorted.astype(str).apply(lambda x: x.str.contains(search_hist, case=False)).any(axis=1)
                df_rev_to_show = df_rev_sorted[mask_hist]
            else:
                df_rev_to_show = df_rev_sorted
            # ---------------------------------------

            if not df_rev_to_show.empty:
                df_rev_to_show.index = range(1, len(df_rev_to_show) + 1)
                st.dataframe(df_rev_to_show, use_container_width=True)
                st.download_button("📥 Export History", get_excel_download_link(df_rev_to_show), "History.xlsx")
            else:
                st.warning("🔎 ไม่พบข้อมูลที่ตรงกับการค้นหาในประวัติ")
                
    except Exception as e:
        st.error(f"History Error: {e}")
