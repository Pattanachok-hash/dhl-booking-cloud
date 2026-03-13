import streamlit as st
import pandas as pd
import fitz
import json
from google import genai
from supabase import create_client, Client

# --- 1. CONFIGURATION (Security Update) ---
# วิธีนี้จะช่วยให้แอปดึงรหัสจากระบบความปลอดภัยของ Streamlit ได้
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
MODEL_ID = "gemini-3.1-pro-preview"

st.set_page_config(page_title="DHL Booking Cloud Extractor", layout="wide")

# --- 2. แก้ไขส่วน CSS (ปรับปรุงเพื่อให้พื้นหลังตารางเป็นเทาและดู 3 มิติ) ---
st.markdown("""
<style>
    /* 1. พื้นหลังหลักสีเหลืองไล่ระดับ */
    .stApp {
        background: linear-gradient(135deg, #FFCC00 0%, #FFD700 50%, #ba9500 100%);
    }

    /* 2. กรอบ 3 มิติสีแดงครอบส่วนแอป */
    .block-container {
        background-color: white;
        padding: 40px;
        border-radius: 25px;
        box-shadow: 0 15px 35px rgba(0,0,0,0.3);
        border: 6px solid #D40511; 
        margin-top: 20px;
        margin-bottom: 20px;
    }

    /* 3. ปรับพื้นหลังตารางและตัวครอบให้เป็นเทาอ่อน */
    /* เราต้องสั่งเจาะจงไปที่ส่วนหัวและส่วนเนื้อหาของ Dataframe */
    .stDataFrame, div[data-testid="stTable"], .pinned-row-container {
        background-color: #f0f2f6 !important;
        border-radius: 10px;
        padding: 10px;
        box-shadow: inset 2px 2px 5px rgba(0,0,0,0.05); /* เงาด้านในให้ดูยุบลงไป */
    }

    /* ปรับสีปุ่มให้เป็นสีแดง DHL */
    .stButton>button {
        background-color: #D40511;
        color: white;
        border-radius: 10px;
        border: none;
        box-shadow: 0 4px #990000;
        transition: 0.2s;
    }
    .stButton>button:hover {
        background-color: #ff0000;
        transform: translateY(-2px);
        box-shadow: 0 6px #990000;
    }
</style>
""", unsafe_allow_html=True)

# --- แก้ไขส่วน Title (ลบของเก่าแล้ววางชุดนี้แทน) ---
container_html = """
    <div style="display: flex; align-items: center; margin-bottom: 20px; padding: 15px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 10px solid #FFCC00;">
        <div style="
            width: 140px; 
            height: 80px; 
            background-color: #FFCC00; 
            border: 3px solid #333; 
            border-radius: 6px; 
            display: flex; 
            justify-content: center; 
            align-items: center;
            margin-right: 25px;
            position: relative;
            box-shadow: 4px 4px 0px #ba9500;
            flex-shrink: 0;
        ">
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

# --- 2. FUNCTIONS ---
def extract_info_from_pdf(text):
    """ส่งข้อความให้ AI สกัดข้อมูลเป็น JSON ตามชื่อคอลัมน์ใน SQL"""
    prompt = f"""
    Extract shipping information from this text into JSON. 
    Required fields: booking_no, liner_name, vessel_name, etd, liner_cutoff, 
    vgm_cutoff, si_cutoff, return_date_1st, cy_date, no_container, no_pallet, return_place, container_type, paperless_code, fcl_or_lcl, by_air_or_sea.
    
    Rules: 
    - CY_date means 1st date that can receive the containers
    - return_date_1st: refer to '1st date to return' or 'Turn-In Date' that can return the container
    - CY cut-off date or gate closing date mean the last date that can return the container
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
    Text content:
    {text}
    """
    try:
        response = genai_client.models.generate_content(
            model=MODEL_ID, 
            contents=prompt,
            config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"AI Extraction Error: {e}")
        return None

def save_to_cloud(data_list):
    """บันทึกข้อมูลลง 2 ตาราง: ล่าสุด (Upsert) และ ประวัติ (Insert)"""
    try:
        # 1. แปลง data_list เป็น DataFrame เพื่อจัดการค่า NaN
        df_temp = pd.DataFrame(data_list)
        
        # เปลี่ยนค่า NaN (ที่ทำให้เกิด Error) ให้เป็น None เพื่อให้ Database ยอมรับ
        df_temp = df_temp.replace({pd.NA: None, float('nan'): None})
        df_temp = df_temp.where(pd.notnull(df_temp), None)

        # จัดการข้อมูลซ้ำสำหรับตารางหลัก (bookings)
        df_unique = df_temp.drop_duplicates(subset=['booking_no'], keep='last')
        clean_data_list = df_unique.to_dict(orient='records')

        # อัปเดตตารางหลัก
        supabase.table("bookings").upsert(clean_data_list).execute()
        
        # 2. เพิ่มลงตารางประวัติ (ใช้ข้อมูลที่แก้ค่า NaN แล้ว)
        full_data_list = df_temp.to_dict(orient='records')
        supabase.table("booking_revisions").insert(full_data_list).execute()
        
        return True
    except Exception as e:
        st.error(f"Database Error: {e}")
        return False

# --- 3. UI: UPLOAD SECTION ---
uploaded_files = st.file_uploader("ลากไฟล์ Booking PDF มาวางที่นี่ (รองรับหลายไฟล์)", type="pdf", accept_multiple_files=True)

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
                        # ตรวจสอบรูปแบบข้อมูล: ถ้ามาเป็น Dict ให้ห่อด้วย List
                        if isinstance(data, dict):
                            data = [data]
                        
                        # วนลูปเก็บข้อมูล (รองรับกรณี AI ส่งมาเป็น List หลายรายการในไฟล์เดียว)
                        if isinstance(data, list):
                            for item in data:
                                item['source_file'] = file.name
                                all_extracted_data.append(item)
                    else:
                        st.warning(f"⚠️ ไม่สามารถสกัดข้อมูลจากไฟล์ {file.name} ได้ (AI อ่านไม่สำเร็จ)")
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดกับไฟล์ {file.name}: {e}")
        
        progress_bar.progress((i + 1) / len(uploaded_files))

    # บันทึกข้อมูลทั้งหมดที่สะสมมาได้ลง Cloud
    if all_extracted_data:
        if save_to_cloud(all_extracted_data):
            st.success(f"🎉 สำเร็จ! บันทึกข้อมูล {len(all_extracted_data)} รายการ และเก็บประวัติเรียบร้อย")
            

# --- 4. UI: LIVE VIEW (แสดงผลทันทีและระบบค้นหา) ---
st.divider()
st.subheader("📊 รายการ Booking ทั้งหมดในระบบ (Live View)")

try:
    # 1. ดึงข้อมูลจาก Supabase
    res = supabase.table("bookings").select("*").order("updated_at", desc=True).execute()
    
    if res.data and len(res.data) > 0:
        # 2. แปลงเป็น DataFrame
        df_live = pd.DataFrame(res.data)
        
        # 3. ช่องค้นหา (Search Box)
        search_query = st.text_input("🔍 ค้นหาเลข Booking, ชื่อเรือ หรือชื่อ Liner", placeholder="พิมพ์เพื่อค้นหา...")
        
        # 4. ตรรกะการค้นหา
        if search_query:
            # ค้นหาทุกคอลัมน์ว่ามีคำที่พิมพ์ไหม (ไม่สนตัวพิมพ์เล็ก-ใหญ่)
            mask = df_live.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            df_to_show = df_live[mask].copy()
        else:
            df_to_show = df_live.copy()

        # 5. ปรับเลข Index ให้เริ่มที่ 1
        df_to_show.index = range(1, len(df_to_show) + 1)

        # 6. แสดงผลตาราง
        if not df_to_show.empty:
            st.dataframe(df_to_show, use_container_width=True)
        else:
            st.warning(f"❌ ไม่พบข้อมูลที่ตรงกับ '{search_query}'")
            
    else:
        st.info("📌 ยังไม่มีข้อมูลในระบบ ลองอัปโหลดไฟล์ PDF ด้านบนเพื่อเริ่มใช้งาน")

except Exception as e:
    # ถ้าเกิด Error จะแสดงข้อความเตือนแทนกรอบสีแดงยาวๆ
    st.error(f"⚠️ เกิดข้อผิดพลาดในการโหลดข้อมูล: {e}")

# --- 5. UI: HISTORY SECTION (เพิ่มระบบค้นหาประวัติ) ---
st.divider()
st.subheader("📜 ประวัติการบันทึกข้อมูลย้อนหลัง (Full Revision Logs)")

# 1. ปุ่มเรียกดูประวัติ (เพื่อให้โหลดข้อมูลเฉพาะตอนต้องการดูจริง ๆ)
if st.button("🔍 โหลดประวัติการสแกนทั้งหมด"):
    st.session_state['show_history'] = True

if st.session_state.get('show_history'):
    try:
        # 2. ดึงข้อมูลประวัติจาก Supabase
        rev_res = supabase.table("booking_revisions").select("*").order("created_at", desc=True).execute()
        
        if rev_res.data and len(rev_res.data) > 0:
            df_rev = pd.DataFrame(rev_res.data)
            
            # 3. ช่องค้นหาเฉพาะในส่วนประวัติ
            search_history = st.text_input("🔎 ค้นหาเลข Booking ในประวัติ (เพื่อดู Revision ย้อนหลัง)", key="history_search")
            
            # 4. ตรรกะการกรองข้อมูลประวัติ
            if search_history:
                mask_rev = df_rev.astype(str).apply(lambda x: x.str.contains(search_history, case=False, na=False)).any(axis=1)
                df_rev_show = df_rev[mask_rev].copy()
            else:
                df_rev_show = df_rev.copy()

            # 5. ปรับเลข Index ให้เริ่มที่ 1
            df_rev_show.index = range(1, len(df_rev_show) + 1)

            # 6. แสดงผลตารางประวัติ
            if not df_rev_show.empty:
                st.dataframe(df_rev_show, use_container_width=True)
                
                # แสดงจำนวนรายการที่พบ
                st.caption(f"พบทั้งหมด {len(df_rev_show)} รายการในประวัติ")
            else:
                st.warning(f"❌ ไม่พบเลข Booking '{search_history}' ในประวัติ")
        else:
            st.warning("ยังไม่มีข้อมูลในประวัติ")
            
    except Exception as e:

        st.error(f"⚠️ เกิดข้อผิดพลาดในการโหลดประวัติ: {e}")

