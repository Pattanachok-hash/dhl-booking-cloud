import pytz
import io
import streamlit as st
import pandas as pd
import fitz
import json
from google import genai
from supabase import create_client, Client

# --- 1. CONFIGURATION (Security Update) ---
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
MODEL_ID = "gemini-3.1-flash-lite-preview"

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

# --- UI TITLE (DHL Container Style) ---
container_html = """
    <div style="display: flex; align-items: center; margin-bottom: 20px; padding: 15px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 10px solid #FFCC00;">
        <div style="width: 140px; height: 80px; background-color: #FFCC00; border: 3px solid #333; border-radius: 6px; display: flex; justify-content: center; align-items: center; margin-right: 25px; position: relative; box-shadow: 4px 4px 0px #ba9500; flex-shrink: 0;">
            <span style="color: #D40511; font-family: 'Arial Black', sans-serif; font-size: 32px; font-weight: 900; letter-spacing: -2px; z-index: 2;">DHL</span>
            <div style="position: absolute; width: 100%; height: 100%; display: flex; justify-content: space-evenly; z-index: 1;">
                <div style="width: 1px; height: 100%; background: rgba(0,0,0,0.1);"></div>
                <div style="width: 1px; height: 100%; background: rgba(0,0,0,0.1);"></div>
                <div style="width: 1px; height: 100%; background: rgba(0,0,0,0.1);"></div>
            </div>
        </div>
        <div style="flex-grow: 1;">
            <h1 style="margin: 0; color: #333; font-size: 30px; line-height: 1.2;">DSC: CTC FG Export</h1>
            <p style="margin: 0; color: #666; font-size: 18px;">Booking Cloud Extractor</p>
        </div>
    </div>
"""
st.markdown(container_html, unsafe_allow_html=True)
st.info("💡 ลากไฟล์ PDF วางเพื่ออัปเดตข้อมูลล่าสุดลงระบบ Cloud และบันทึกประวัติการแก้ไข (Full Revision)")

# --- FUNCTIONS ---
def extract_info_from_pdf(text):
    # นำ Prompt Logic ชุดใหญ่ที่คุณเคยเขียนไว้กลับมาครบทุกข้อ
    prompt = f"""
    Extract shipping information from this text into JSON. 
    Required fields: booking_no, fcl_or_lcl, by_air_or_sea, country, port_of_destination, liner_name, vessel_name,
    no_container, no_pallet, liner_cutoff, vgm_cutoff, si_cutoff, return_date_1st, cy_date, cy_at, etd, eta,
    return_place, container_type, paperless_code.
    
    Rules: 
    - cy_date means 1st date that can receive the containers; sometimes, it is issued in CY AT
    - cy_at means the place that can receive the container
    - return_place or RTN means the place that can receive the container
    - country refer to the country name in the port of destination
    - return_date_1st: refer to '1st return date' or 'Turn-In Date' that can return the container
    - CY cut-off date or gate closing date or closing date mean the last date that can return the container. This is very
    important information please beware.
    - VGM cut-off date:  This is very important information please beware.
    - SI cut-off date:  This is very important information please beware.
    - No_container column: please use whole number
    - The container type/s should be filled only "40HC" or "20GP"
    - If any field is not found, use null.
    - Return ONLY the JSON object.
    - For the Liner_cutoff, VGM_cutoff, SI_cutoff column, please use format: dd/mm/yyyy hh:mm
    - For the ETD, 1st return date, CY date column, please use format: dd/mm/yyyy
    - If the shipment type is FCL the unit is container but if the shipment type is LCL the unit is pallet
    - Paperless code is the 4-digit number
    - by_air_or_sea should be filled only "Air" or "Sea"
    - fcl_or_lcl should be filled only "FCL" or "LCL"
    - Please beware ETD is ETD not ETA
    - For LCL shipment, you may take time to process because the format is informal.
    Sometimes, they use day instead of date so you may find out the ref date in the file first then calculate again.
    
    Text content:
    {text}
    """
    try:
        response = genai_client.models.generate_content(
            model=MODEL_ID, contents=prompt, config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"AI Extraction Error: {e}")
        return None

def save_to_cloud(data_list):
    try:
        df_temp = pd.DataFrame(data_list)
        # แก้ไขค่า NaN ให้เป็น None เพื่อความถูกต้องของ Database (ตามโค้ดเดิมของคุณ)
        df_temp = df_temp.replace({pd.NA: None, float('nan'): None})
        df_temp = df_temp.where(pd.notnull(df_temp), None)
        
        # จัดการข้อมูลซ้ำสำหรับตารางหลัก (bookings) ตามโค้ดเดิม
        df_unique = df_temp.drop_duplicates(subset=['booking_no'], keep='last')
        clean_data_list = df_unique.to_dict(orient='records')
        supabase.table("bookings").upsert(clean_data_list).execute()
        
        # เพิ่มลงตารางประวัติ (History)
        full_data_list = df_temp.to_dict(orient='records')
        supabase.table("booking_revisions").insert(full_data_list).execute()
        return True
    except Exception as e:
        st.error(f"Database Error: {e}")
        return False
    
def get_excel_download_link(df, filename="export.xlsx"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# --- ฟังก์ชันช่วยแปลงเวลา (ใส่ไว้ในส่วน FUNCTIONS) ---
def format_thai_timezone(df, column_name):
    """แปลงเวลาจากฐานข้อมูลเป็น Thai Timezone และ Format dd/mm/yyyy hh:mm"""
    if column_name in df.columns:
        try:
            # 1. แปลงเป็น datetime object (กรณีมาจาก Supabase มักจะเป็น ISO format)
            df[column_name] = pd.to_datetime(df[column_name])
            
            # 2. กำหนด Timezone เป็น UTC (เพราะ Supabase เก็บเป็น UTC) แล้วแปลงเป็น Asia/Bangkok
            df[column_name] = df[column_name].dt.tz_convert('Asia/Bangkok')
            
            # 3. จัด Format เป็น dd/mm/yyyy hh:mm
            df[column_name] = df[column_name].dt.strftime('%d/%m/%Y %H:%M')
        except Exception as e:
            st.error(f"Error formatting time: {e}")
    return df

# --- 3. UI: UPLOAD SECTION (Auto-Clear Enabled) ---
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

uploaded_files = st.file_uploader(
    "ลากไฟล์ Booking PDF มาวางที่นี่ (รองรับหลายไฟล์)", 
    type="pdf", 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

if uploaded_files:
    all_extracted_data = []
    progress_bar = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        with st.spinner(f"กำลังประมวลผลไฟล์: {file.name}"):
            try:
                doc = fitz.open(stream=file.read(), filetype="pdf")
                full_text = "".join([page.get_text() for page in doc])
                if full_text.strip():
                    data = extract_info_from_pdf(full_text)
                    if data:
                        if isinstance(data, dict): data = [data]
                        if isinstance(data, list):
                            for item in data:
                                item['source_file'] = file.name
                                all_extracted_data.append(item)
                    else:
                        st.warning(f"⚠️ ไม่สามารถสกัดข้อมูลจากไฟล์ {file.name} ได้")
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดกับไฟล์ {file.name}: {e}")
        progress_bar.progress((i + 1) / len(uploaded_files))

    if all_extracted_data:
        if save_to_cloud(all_extracted_data):
            st.success(f"🎉 สำเร็จ! บันทึกข้อมูล {len(all_extracted_data)} รายการเรียบร้อย")
            st.session_state['uploader_key'] += 1
            st.rerun()

# --- 4. UI: LIVE VIEW & EXPORT (ฉบับสมบูรณ์: เรียงคอลัมน์ + แจ้งเตือน) ---
st.divider()
st.subheader("📊 รายการ Booking ทั้งหมดในระบบ (Live View)")

try:
    # 1. ดึงข้อมูลจาก Supabase
    res = supabase.table("bookings").select("*").order("updated_at", desc=True).execute()
    
    # เช็คว่าใน Database มีข้อมูลไหม
    if res.data and len(res.data) > 0:
        df_live = pd.DataFrame(res.data)
        df_live = format_thai_timezone(df_live, 'updated_at')
        
        # 2. กำหนดลำดับคอลัมน์ที่คุณต้องการ (จัดเรียงใหม่ให้หายงง)
        my_columns = [
            "booking_no", "fcl_or_lcl", "by_air_or_sea", "country", 
            "port_of_destination", "liner_name", "vessel_name",
            "no_container", "container_type", "no_pallet",
            "etd", "eta", "liner_cutoff", "vgm_cutoff", "si_cutoff",
            "cy_date", "cy_at", "return_date_1st", "return_place", 
            "paperless_code", "source_file", "updated_at"
        ]

        # เลือกเฉพาะคอลัมน์ที่มีอยู่จริงใน DB มาโชว์
        existing_columns = [c for c in my_columns if c in df_live.columns]
        df_sorted = df_live[existing_columns].copy()

        # 3. ช่องค้นหา (Search Box)
        search_query = st.text_input("🔍 ค้นหาเลข Booking, ชื่อเรือ หรือชื่อ Liner", placeholder="พิมพ์เพื่อค้นหา...")
        
        if search_query:
            mask = df_sorted.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            df_to_show = df_sorted[mask].copy()
        else:
            df_to_show = df_sorted.copy()

        # ปรับเลขลำดับหน้าตารางให้เริ่มที่ 1
        df_to_show.index = range(1, len(df_to_show) + 1)

        # 4. แสดงผลตาราง (ถ้ามีข้อมูลจากการค้นหา)
        if not df_to_show.empty:
            st.dataframe(df_to_show, use_container_width=True)
            
            # ปุ่ม Export Excel (เรียงลำดับตามที่เราจัดไว้)
            excel_data = get_excel_download_link(df_to_show, "Current_Bookings.xlsx")
            st.download_button(
                label="📥 Export Current View to Excel",
                data=excel_data,
                file_name="Current_Bookings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # กรณี Search แล้วไม่เจอ
            st.warning(f"❌ ไม่พบข้อมูลที่ตรงกับ '{search_query}'")
            
    else:
        # กรณีตารางใน Database ยังว่างเปล่า
        st.info("📌 ยังไม่มีข้อมูลในระบบ ลองอัปโหลดไฟล์ PDF ด้านบนเพื่อเริ่มใช้งาน")

except Exception as e:
    # กรณีเกิด Error จากการโหลด (เช่น ลืมเพิ่มคอลัมน์ใน Supabase)
    st.error(f"⚠️ Load Error: {e}")

# --- 5. UI: HISTORY SECTION & EXPORT (ฉบับจัดเรียงคอลัมน์ให้หายงง) ---
st.divider()
st.subheader("📜 ประวัติการบันทึกข้อมูลย้อนหลัง (Full Revision Logs)")

# 1. ปุ่มเรียกดูประวัติ
if st.button("🔍 โหลดประวัติการสแกนทั้งหมด"):
    st.session_state['show_history'] = True

if st.session_state.get('show_history'):
    try:
        # 2. ดึงข้อมูลประวัติจาก Supabase
        rev_res = supabase.table("booking_revisions").select("*").order("created_at", desc=True).execute()
        
        if rev_res.data and len(rev_res.data) > 0:
            df_rev = pd.DataFrame(rev_res.data)
            df_rev = format_thai_timezone(df_rev, 'created_at')
            
            # 3. กำหนดลำดับคอลัมน์ให้เหมือนกับหน้า Live View (เพื่อความต่อเนื่อง)
            history_columns = [
                "booking_no", "fcl_or_lcl", "by_air_or_sea", "country", 
                "port_of_destination", "liner_name", "vessel_name",
                "no_container", "container_type", "no_pallet",
                "etd", "eta", "liner_cutoff", "vgm_cutoff", "si_cutoff",
                "cy_date", "cy_at", "return_date_1st", "return_place", 
                "paperless_code", "source_file", "created_at" # ใช้ created_at เพื่อดูเวลาที่บันทึกจริง
            ]
            
            # เลือกเฉพาะคอลัมน์ที่มีอยู่จริง
            existing_rev_cols = [c for c in history_columns if c in df_rev.columns]
            df_rev_sorted = df_rev[existing_rev_cols].copy()

            # 4. ช่องค้นหาเฉพาะในส่วนประวัติ
            search_history = st.text_input("🔎 ค้นหาเลข Booking ในประวัติ (เพื่อดู Revision ย้อนหลัง)", key="history_search")
            
            if search_history:
                mask_rev = df_rev_sorted.astype(str).apply(lambda x: x.str.contains(search_history, case=False, na=False)).any(axis=1)
                df_rev_show = df_rev_sorted[mask_rev].copy()
            else:
                df_rev_show = df_rev_sorted.copy()

            # ปรับลำดับเลข Index ให้เริ่มที่ 1
            df_rev_show.index = range(1, len(df_rev_show) + 1)

            # 5. แสดงผลตารางประวัติ
            if not df_rev_show.empty:
                st.dataframe(df_rev_show, use_container_width=True)
                
                # ปุ่ม Export สำหรับ History (จะเรียงลำดับตามที่เราจัดไว้ด้วย)
                excel_rev_data = get_excel_download_link(df_rev_show, "Booking_History_Full.xlsx")
                st.download_button(
                    label="📥 Export History to Excel",
                    data=excel_rev_data,
                    file_name="Booking_History_Full.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.caption(f"พบข้อมูลประวัติทั้งหมด {len(df_rev_show)} รายการ")
            else:
                st.warning(f"❌ ไม่พบเลข Booking '{search_history}' ในประวัติ")
        else:
            st.warning("📌 ยังไม่มีข้อมูลในประวัติการบันทึก")
            
    except Exception as e:
        st.error(f"⚠️ History Load Error: {e}")
