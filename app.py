import io
import json
import base64
import pytz
import streamlit as st
import pandas as pd
from google import genai
from google.genai import types
from supabase import create_client, Client
from datetime import datetime

# ─────────────────────────────────────────
# 1. CONFIGURATION
# ─────────────────────────────────────────
try:
    SUPABASE_URL  = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY  = st.secrets["SUPABASE_KEY"]
    GEMINI_KEY    = st.secrets["GENAI_API_KEY"]
    TBL_BOOKINGS  = "bookings"
    TBL_REVISIONS = "booking_revisions"
except Exception:
    # สำหรับใช้รันในเครื่องตัวเอง (Local) ถ้ายังไม่ได้ตั้งค่า Secrets
    st.error("❌ ไม่พบ API Keys ในระบบ Secrets กรุณาตั้งค่าที่ Settings > Secrets")
    st.stop()
    TBL_BOOKINGS  = "test_bookings"
    TBL_REVISIONS = "test_booking_revisions"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
genai_client     = genai.Client(api_key=GEMINI_KEY)
GEMINI_MODEL     = "models/gemini-3.1-flash-lite-preview"

# ─────────────────────────────────────────
# 2. PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="DHL Booking Cloud Extractor",
    page_icon="🚚",
    layout="wide",
)
 
# ─────────────────────────────────────────
# 3. CSS THEME  (DHL — Golden gradient style)
# ─────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #FFCC00 0%, #FFD700 50%, #ba9500 100%); }
    .block-container { background-color: white; padding: 40px; border-radius: 25px; box-shadow: 0 15px 35px rgba(0,0,0,0.3); border: 6px solid #D40511; margin-top: 20px; margin-bottom: 20px; }
    .stDataFrame, div[data-testid="stTable"] { background-color: #f0f2f6 !important; border-radius: 10px; padding: 10px; box-shadow: inset 2px 2px 5px rgba(0,0,0,0.05); }
    .stButton>button { background-color: #D40511; color: white; border-radius: 10px; border: none; box-shadow: 0 4px #990000; transition: 0.2s; width: 100%; }
    .stButton>button:hover { background-color: #ff0000; transform: translateY(-2px); box-shadow: 0 6px #990000; }
    [data-testid="stDownloadButton"] > button { background-color: #D40511; color: white; border-radius: 10px; border: none; box-shadow: 0 4px #990000; transition: 0.2s; width: auto; }
    [data-testid="stDownloadButton"] > button:hover { background-color: #ff0000; transform: translateY(-2px); box-shadow: 0 6px #990000; }
    [data-testid="metric-container"] { background: #fff9e6; border: 2px solid #FFCC00; border-radius: 12px; padding: 16px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #D40511; font-size: 28px; font-weight: 700; }
    [data-testid="stFileUploader"] { border: 2px dashed #FFCC00 !important; border-radius: 12px !important; padding: 8px !important; background-color: #fffef5 !important; }
    [data-testid="stFileUploader"]:hover { border-color: #D4A900 !important; background-color: #fffbe6 !important; }
    [data-testid="stTextInput"] > div { border: 2px solid #FFCC00 !important; border-radius: 10px !important; background-color: #fffef5 !important; }
    [data-testid="stTextInput"] > div:focus-within { border-color: #D4A900 !important; box-shadow: 0 0 0 3px rgba(255,204,0,0.2) !important; }
    [data-testid="stSelectbox"] > div > div { border: 2px solid #FFCC00 !important; border-radius: 10px !important; background-color: #fffef5 !important; }
    [data-testid="stSelectbox"] > div > div:focus-within { border-color: #D4A900 !important; box-shadow: 0 0 0 3px rgba(255,204,0,0.2) !important; }
</style>
""", unsafe_allow_html=True)
 
# ─────────────────────────────────────────
# 4. HEADER
# ─────────────────────────────────────────
_truck_svg = """<svg viewBox="0 0 230 92" xmlns="http://www.w3.org/2000/svg" width="230" height="92">
  <ellipse cx="115" cy="90" rx="108" ry="4" fill="rgba(0,0,0,0.07)"/>
  <rect x="78" y="71" width="146" height="6" rx="1" fill="#444" stroke="#222" stroke-width="1"/>
  <rect x="74" y="6" width="150" height="65" rx="3" fill="#FFCC00" stroke="#222" stroke-width="2.5"/>
  <rect x="74" y="6" width="150" height="7" rx="2" fill="#D4A900" stroke="#222" stroke-width="1.5"/>
  <rect x="74" y="64" width="150" height="7" rx="1" fill="#D4A900" stroke="#222" stroke-width="1.5"/>
  <line x1="112" y1="13" x2="112" y2="64" stroke="#D4A900" stroke-width="2"/>
  <line x1="150" y1="13" x2="150" y2="64" stroke="#D4A900" stroke-width="2"/>
  <line x1="188" y1="13" x2="188" y2="64" stroke="#D4A900" stroke-width="2"/>
  <rect x="218" y="13" width="6" height="51" rx="1" fill="#D4A900" stroke="#222" stroke-width="1"/>
  <text x="149" y="51" font-family="Arial Black,Arial,sans-serif" font-size="27" font-weight="900" fill="#D40511" text-anchor="middle" letter-spacing="-1">DHL</text>
  <rect x="4" y="10" width="66" height="63" rx="5" fill="#D40511" stroke="#222" stroke-width="2.5"/>
  <rect x="4" y="10" width="13" height="63" rx="4" fill="#B8030F" stroke="#222" stroke-width="1"/>
  <rect x="18" y="15" width="36" height="27" rx="3" fill="#AED6F1" stroke="#333" stroke-width="1.5"/>
  <line x1="22" y1="18" x2="27" y2="37" stroke="white" stroke-width="2" opacity="0.4"/>
  <rect x="54" y="15" width="12" height="18" rx="2" fill="#AED6F1" stroke="#333" stroke-width="1"/>
  <line x1="52" y1="43" x2="52" y2="71" stroke="#aa0000" stroke-width="1.5"/>
  <rect x="54" y="53" width="10" height="3" rx="1.5" fill="#FFCC00" stroke="#D4A900" stroke-width="1"/>
  <rect x="5" y="18" width="10" height="6" rx="2" fill="#FFF9C4" stroke="#333" stroke-width="1"/>
  <rect x="5" y="50" width="10" height="5" rx="2" fill="#FF8A65" stroke="#333" stroke-width="1"/>
  <rect x="5" y="26" width="10" height="22" rx="1" fill="#222" stroke="#111" stroke-width="1"/>
  <line x1="5" y1="31" x2="15" y2="31" stroke="#555" stroke-width="1"/>
  <line x1="5" y1="36" x2="15" y2="36" stroke="#555" stroke-width="1"/>
  <line x1="5" y1="41" x2="15" y2="41" stroke="#555" stroke-width="1"/>
  <rect x="4" y="67" width="66" height="6" rx="2" fill="#999" stroke="#333" stroke-width="1.5"/>
  <rect x="-1" y="18" width="8" height="6" rx="1" fill="#777" stroke="#444" stroke-width="1"/>
  <line x1="4" y1="21" x2="7" y2="21" stroke="#555" stroke-width="1.5"/>
  <rect x="60" y="0" width="5" height="16" rx="2" fill="#888" stroke="#555" stroke-width="1"/>
  <circle cx="63" cy="-1" r="3" fill="#bbb" opacity="0.4"/>
  <circle cx="65" cy="-5" r="2" fill="#bbb" opacity="0.25"/>
  <rect x="62" y="67" width="16" height="5" rx="1" fill="#666" stroke="#333" stroke-width="1"/>
  <circle cx="19" cy="80" r="11" fill="#1a1a1a" stroke="#555" stroke-width="2"/>
  <circle cx="19" cy="80" r="5.5" fill="#666"/>
  <circle cx="19" cy="80" r="2" fill="#333"/>
  <circle cx="52" cy="80" r="11" fill="#1a1a1a" stroke="#555" stroke-width="2"/>
  <circle cx="52" cy="80" r="5.5" fill="#666"/>
  <circle cx="52" cy="80" r="2" fill="#333"/>
  <circle cx="67" cy="80" r="11" fill="#1a1a1a" stroke="#555" stroke-width="2"/>
  <circle cx="67" cy="80" r="5.5" fill="#666"/>
  <circle cx="67" cy="80" r="2" fill="#333"/>
  <circle cx="168" cy="80" r="11" fill="#1a1a1a" stroke="#555" stroke-width="2"/>
  <circle cx="168" cy="80" r="5.5" fill="#666"/>
  <circle cx="168" cy="80" r="2" fill="#333"/>
  <circle cx="207" cy="80" r="11" fill="#1a1a1a" stroke="#555" stroke-width="2"/>
  <circle cx="207" cy="80" r="5.5" fill="#666"/>
  <circle cx="207" cy="80" r="2" fill="#333"/>
</svg>"""

_truck_b64 = base64.b64encode(_truck_svg.encode()).decode()

st.markdown(f"""
    <div style="display: flex; align-items: center; margin-bottom: 20px; padding: 15px; background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 10px solid #FFCC00;">
        <div style="margin-right: 20px; flex-shrink: 0;">
            <img src="data:image/svg+xml;base64,{_truck_b64}" width="230" height="92"/>
        </div>
        <div style="flex-grow: 1;">
            <h1 style="margin: 0; color: #333; font-size: 30px; line-height: 1.2;">DSC: CTC FG Export</h1>
            <p style="margin: 0; color: #666; font-size: 18px;">Booking Cloud Extractor (Vision Engine)</p>
        </div>
    </div>
""", unsafe_allow_html=True)
 
# ─────────────────────────────────────────
# 5. HELPERS
# ─────────────────────────────────────────
COLUMNS_ORDER = [
    "booking_no", "loading_at", "fcl_or_lcl", "by_air_or_sea",
    "country", "port_of_destination", "liner_name", "vessel_name",
    "no_container", "container_type", "no_pallet",
    "etd", "eta", "liner_cutoff", "vgm_cutoff", "si_cutoff",
    "cy_date", "cy_at", "return_date_1st", "return_place",
    "paperless_code", "updated_at",
]
 
PROMPT_DATES = """You are a DHL Logistics Analyst. Extract ONLY date/time fields from this booking PDF.
 
Return ONLY a JSON object (no markdown, no explanation):
{
  "liner_cutoff":    "dd/mm/yyyy hh:mm or null",
  "vgm_cutoff":      "dd/mm/yyyy hh:mm or null",
  "si_cutoff":       "dd/mm/yyyy hh:mm or null",
  "return_date_1st": "dd/mm/yyyy or null",
  "cy_date":         "dd/mm/yyyy or null",
  "etd":             "dd/mm/yyyy or null",
  "eta":             "dd/mm/yyyy or null"
}
 
Rules:
- Dates: dd/mm/yyyy. Cut-offs include hh:mm.
- cy_date: Empty Pick-up date / date to collect empty container.
- return_date_1st: 1st Return Date / Turn-In Date.
- liner_cutoff: Gate Closing / Closing Date / CY Cut-off / Last Load.
- si_cutoff: SI Cut-off / Doc Cut-off / Shipping Particular Cut-off.
- If a cut-off shows only a weekday (e.g. "THU"), calculate the actual date from the document date or ETD.
- ETD must be earlier than ETA.
- null if not found."""
 
PROMPT_GENERAL = """You are a DHL Logistics Analyst. Extract ONLY general shipping info from this booking PDF.
 
Return ONLY a JSON object (no markdown, no explanation):
{
  "booking_no":          "Carrier Ref or Booking No.",
  "fcl_or_lcl":          "FCL or LCL",
  "by_air_or_sea":       "Air or Sea",
  "country":             "destination country (NOT Thailand)",
  "port_of_destination": "port name",
  "liner_name":          "shipping line",
  "vessel_name":         "vessel/voyage (include connecting if any)",
  "no_container":        number or null,
  "container_type":      "40HC or 20GP or null",
  "no_pallet":           number or null,
  "cy_at":               "empty pick-up depot",
  "return_place":        "laden return location",
  "paperless_code":      "4-digit code or null"
}
 
Rules:
- booking_no: prefer Carrier Ref; fallback to Booking No. / Booking Ref.
- country: infer from Port of Discharge if consignee is Thai.
- cy_at: depot for picking up empty container.
- return_place: Laden Return / Return to location.
- paperless_code: exact 4-digit number next to "PAPERLESS CODE"; fallback by terminal only if missing.
- null if not found."""
 
 
def extract_from_pdf(file_bytes: bytes) -> list[dict]:
    """ส่ง PDF ให้ Gemini อ่าน 2 รอบ (dates + general) แล้ว merge"""
    ai_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.0,
        seed=42,
    )
 
    def call(prompt: str) -> dict:
        res = genai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part.from_text(text=prompt),
                        types.Part.from_bytes(
                            data=file_bytes,
                            mime_type="application/pdf",
                        ),
                    ],
                )
            ],
            config=ai_config,
        )
        raw = res.text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw.strip())
 
    dates   = call(PROMPT_DATES)
    general = call(PROMPT_GENERAL)
 
    dates_list   = dates   if isinstance(dates,   list) else [dates]
    general_list = general if isinstance(general, list) else [general]
 
    merged = []
    for i, g in enumerate(general_list):
        d = dates_list[i] if i < len(dates_list) else {}
        merged.append({**g, **d})
    return merged
 
 
def save_to_supabase(data_list: list[dict]) -> bool:
    """Upsert to bookings (by booking_no) and insert to revisions."""
    try:
        df = pd.DataFrame(data_list).replace({pd.NA: None, float("nan"): None})
        df = df.where(pd.notnull(df), None)
        df_clean = df.dropna(subset=["booking_no"]).drop_duplicates(
            subset=["booking_no"], keep="last"
        )
        if df_clean.empty:
            st.warning("⚠️ ไม่พบ Booking No. ในเอกสาร — ไม่ได้บันทึกลงฐานข้อมูล")
            return False
        supabase.table(TBL_BOOKINGS).upsert(df_clean.to_dict(orient="records")).execute()
        supabase.table(TBL_REVISIONS).insert(df.to_dict(orient="records")).execute()
        return True
    except Exception as e:
        st.error(f"❌ Database Error: {e}")
        return False
 
 
def bkk_time(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Convert a UTC timestamp column to Bangkok time string."""
    if col in df.columns:
        try:
            df[col] = (
                pd.to_datetime(df[col])
                .dt.tz_convert("Asia/Bangkok")
                .dt.strftime("%d/%m/%Y %H:%M")
            )
        except Exception:
            pass
    return df
 
 
def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Bookings")
        wb  = writer.book
        ws  = writer.sheets["Bookings"]
 
        # Formats
        hdr_fmt = wb.add_format({
            "bold": True, "font_name": "Arial",
            "bg_color": "#FFCC00", "font_color": "#D40511",
            "border": 1, "align": "center", "valign": "vcenter",
        })
        cell_fmt = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "border": 1, "valign": "vcenter",
        })
        alt_fmt = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "bg_color": "#FFF9E6", "border": 1, "valign": "vcenter",
        })
 
        # Header row
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, hdr_fmt)
            ws.set_column(col_idx, col_idx, max(len(str(col_name)) + 4, 14))
 
        # Data rows
        for row_idx in range(len(df)):
            fmt = alt_fmt if row_idx % 2 else cell_fmt
            for col_idx in range(len(df.columns)):
                ws.write(row_idx + 1, col_idx, df.iloc[row_idx, col_idx], fmt)
 
        ws.set_row(0, 22)
        ws.freeze_panes(1, 0)
 
    return buf.getvalue()
 
 
def render_table(df: pd.DataFrame, table_id: str = "main") -> None:
    """แสดงตารางแบบ HTML มี badge สี ควบคุมได้เต็มที่"""
 
    # CSS สำหรับตาราง (ใส่ครั้งแรกครั้งเดียว)
    st.markdown("""
    <style>
    .dhl-table-wrap {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 500px;
        border: 1px solid #e8e8e8;
        border-radius: 12px;
        margin-bottom: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    .dhl-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
        font-family: 'Inter', sans-serif;
        background: #ffffff;
    }
    .dhl-table thead tr {
        background: #fafafa;
        border-bottom: 1px solid #e8e8e8;
        position: sticky;
        top: 0;
        z-index: 1;
    }
    .dhl-table thead th {
        padding: 11px 14px;
        text-align: left;
        font-size: 10px;
        font-weight: 600;
        color: #aaa;
        letter-spacing: 1px;
        text-transform: uppercase;
        white-space: nowrap;
    }
    .dhl-table tbody tr {
        border-bottom: 1px solid #f0f0f0;
        transition: background 0.1s;
    }
    .dhl-table tbody tr:hover { background: #fdf5f5; }
    .dhl-table tbody td {
        padding: 9px 14px;
        white-space: nowrap;
        color: #444;
    }
    .dhl-num    { color: #ccc !important; }
    .dhl-bkno   { color: #D40511 !important; font-weight: 600; }
    .dhl-code   { color: #D40511 !important; font-weight: 700; letter-spacing: 1px; }
    .dhl-none   { color: #ddd !important; font-style: italic; }
    .dhl-badge  {
        display: inline-block;
        padding: 2px 9px;
        border-radius: 99px;
        font-size: 10px;
        font-weight: 600;
    }
    .b-fcl   { background: #eff6ff; color: #2563eb; }
    .b-lcl   { background: #faf5ff; color: #7c3aed; }
    .b-icd   { background: #eff6ff; color: #3b82f6; }
    .b-alpha { background: #f0fdf4; color: #16a34a; }
    .b-type  { background: #fffbeb; color: #d97706; border: 1px solid #fde68a; }
    .b-sea   { background: #f0f9ff; color: #0284c7; }
    .b-air   { background: #fff7ed; color: #ea580c; }
    .dhl-table-footer {
        padding: 9px 14px;
        border-top: 1px solid #f0f0f0;
        font-size: 11px;
        color: #bbb;
        background: #fafafa;
        border-radius: 0 0 12px 12px;
    }
    </style>
    """, unsafe_allow_html=True)
 
    def val(row, key):
        v = row.get(key, None)
        if v is None or str(v).strip() in ("", "None", "nan", "NaN"):
            return None
        return str(v).strip()
 
    def badge(text, cls):
        return f'<span class="dhl-badge {cls}">{text}</span>'
 
    rows_html = ""
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        row = row.to_dict()
 
        # booking_no
        bkno  = val(row, "booking_no")
        bkno_html = f'<td class="dhl-bkno">{bkno}</td>' if bkno else '<td class="dhl-none">—</td>'
 
        # loading_at badge
        wh    = val(row, "loading_at") or ""
        wh_cls = "b-icd" if "ICD" in wh else "b-alpha"
        wh_html = badge(wh, wh_cls) if wh else "—"
 
        # FCL/LCL badge
        fcl   = val(row, "fcl_or_lcl") or ""
        fcl_cls = "b-fcl" if fcl == "FCL" else "b-lcl"
        fcl_html = badge(fcl, fcl_cls) if fcl else "—"
 
        # Sea/Air badge
        mode  = val(row, "by_air_or_sea") or ""
        mode_cls = "b-sea" if mode == "Sea" else "b-air"
        mode_html = badge(mode, mode_cls) if mode else "—"
 
        # container_type badge
        ctype = val(row, "container_type") or ""
        ctype_html = badge(ctype, "b-type") if ctype else '<span class="dhl-none">—</span>'
 
        # paperless_code
        code  = val(row, "paperless_code")
        code_html = f'<span class="dhl-code">{code}</span>' if code else '<span class="dhl-none">—</span>'
 
        def cell(key):
            v = val(row, key)
            return f'<td>{v}</td>' if v else '<td class="dhl-none">—</td>'
 
        rows_html += f"""
        <tr>
            <td class="dhl-num">{i}</td>
            {bkno_html}
            <td>{wh_html}</td>
            <td>{fcl_html}</td>
            <td>{mode_html}</td>
            {cell("country")}
            {cell("port_of_destination")}
            {cell("liner_name")}
            {cell("vessel_name")}
            {cell("no_container")}
            <td>{ctype_html}</td>
            {cell("no_pallet")}
            {cell("etd")}
            {cell("eta")}
            {cell("liner_cutoff")}
            {cell("vgm_cutoff")}
            {cell("si_cutoff")}
            {cell("cy_date")}
            {cell("cy_at")}
            {cell("return_date_1st")}
            {cell("return_place")}
            <td>{code_html}</td>
            {cell("updated_at")}
        </tr>"""
 
    headers = [
        "#", "Booking No.", "Loading At", "FCL/LCL", "Mode",
        "Country", "Port of Dest.", "Liner", "Vessel",
        "Ctrs", "Type", "Pallets",
        "ETD", "ETA", "Liner Cutoff", "VGM Cutoff", "SI Cutoff",
        "CY Date", "CY At", "1st Return", "Return Place",
        "Code", "Updated",
    ]
    thead = "".join(f"<th>{h}</th>" for h in headers)
 
    st.markdown(f"""
    <div class="dhl-table-wrap">
        <table class="dhl-table" id="dhl-{table_id}">
            <thead><tr>{thead}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
        <div class="dhl-table-footer">แสดง {len(df)} รายการ</div>
    </div>
    """, unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────
# 6. SIDEBAR NAVIGATION
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🚚 DHL Logistics Menu")
    page = st.radio(
        "เลือกเมนู",
        ["📤 Upload & Extract", "📄 Generate SI (Draft)"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "<span style='font-size:11px;color:#555;'>Powered by Ship Co. (CTC Site)<br>"
        "© DHL Supply Chain Thailand</span>",
        unsafe_allow_html=True,
    )
 
# ─────────────────────────────────────────
# 7. PAGE: UPLOAD & EXTRACT
# ─────────────────────────────────────────
if page == "📤 Upload & Extract":
 
    # ── Upload zone ──────────────────────
    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0
 
    col1, col2 = st.columns(2)

    with col1:
        st.info("🏢 **ICD Warehouse**")
        files_icd = st.file_uploader(
            "โยนไฟล์สำหรับ ICD ที่นี่",
            type="pdf",
            accept_multiple_files=True,
            key=f"icd_{st.session_state['uploader_key']}",
        )

    with col2:
        st.success("🏗️ **ALPHA Warehouse**")
        files_alpha = st.file_uploader(
            "โยนไฟล์สำหรับ ALPHA ที่นี่",
            type="pdf",
            accept_multiple_files=True,
            key=f"alpha_{st.session_state['uploader_key']}",
        )

    # ── Auto-process when files are uploaded ─────────────────
    list_icd   = files_icd   or []
    list_alpha = files_alpha or []
    total      = len(list_icd) + len(list_alpha)

    if total > 0:
        all_data      = []
        progress_bar  = st.progress(0)
        processed     = 0

        for f in list_icd:
            with st.spinner(f"กำลังสแกน (ICD): {f.name}"):
                items = extract_from_pdf(f.read())
                if items:
                    for item in items:
                        item["source_file"] = f.name
                        item["loading_at"]  = "ICD"
                    all_data.extend(items)
            processed += 1
            progress_bar.progress(processed / total)

        for f in list_alpha:
            with st.spinner(f"กำลังสแกน (ALPHA): {f.name}"):
                items = extract_from_pdf(f.read())
                if items:
                    for item in items:
                        item["source_file"] = f.name
                        item["loading_at"]  = "ALPHA"
                    all_data.extend(items)
            processed += 1
            progress_bar.progress(processed / total)

        if all_data:
            if save_to_supabase(all_data):
                st.success(
                    f"🎉 บันทึกข้อมูล {len(all_data)} รายการ "
                    f"จากทั้งหมด {total} ไฟล์ เรียบร้อยแล้ว"
                )
                st.session_state["uploader_key"] += 1
                st.rerun()
 
    # ── Live View ────────────────────────
    st.divider()
 
    st.subheader("📊 รายการ Booking ทั้งหมด (Live View)")
 
    try:
        res = (
            supabase.table(TBL_BOOKINGS)
            .select("*")
            .order("updated_at", desc=True)
            .execute()
        )
 
        if res.data:
            df_live = pd.DataFrame(res.data)
            df_live = bkk_time(df_live, "updated_at")
            existing = [c for c in COLUMNS_ORDER if c in df_live.columns]
            df_show  = df_live[existing].copy()
 
            # Metrics row
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📦 Total Bookings", len(df_show))
            m2.metric("🏢 ICD",   int(df_show["loading_at"].str.contains("ICD",   na=False).sum()) if "loading_at" in df_show else 0)
            m3.metric("🏗️ ALPHA", int(df_show["loading_at"].str.contains("ALPHA", na=False).sum()) if "loading_at" in df_show else 0)
            m4.metric("🌊 FCL",   int(df_show["fcl_or_lcl"].str.contains("FCL",   na=False).sum()) if "fcl_or_lcl"  in df_show else 0)
 
            st.markdown("<br>", unsafe_allow_html=True)
 
            search = st.text_input("🔍 ค้นหา...", placeholder="Booking No., Port, Vessel, Country...")
 
            if search:
                mask = df_show.astype(str).apply(
                    lambda x: x.str.contains(search, case=False, na=False)
                ).any(axis=1)
                df_show = df_show[mask]
 
            df_show.index = range(1, len(df_show) + 1)
            render_table(df_show, table_id="live")
 
            col_dl1, col_dl2 = st.columns([1, 5])
            with col_dl1:
                st.download_button(
                    "📥 Export Excel",
                    data=to_excel(df_show),
                    file_name="DHL_Bookings.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            # ── Edit Booking ──────────────────────────────────────
            st.divider()
            st.subheader("✏️ แก้ไขข้อมูล Booking")

            all_bk_nos = [r.get("booking_no") for r in res.data if r.get("booking_no")]
            edit_bk = st.selectbox("เลือก Booking ที่ต้องการแก้ไข",
                                   ["-- เลือก --"] + all_bk_nos, key="edit_bk_select")

            if edit_bk != "-- เลือก --":
                row_data = next((r for r in res.data if r.get("booking_no") == edit_bk), {})

                with st.form("edit_booking_form"):
                    st.markdown(f"**Booking No.: {edit_bk}**")
                    ec1, ec2, ec3 = st.columns(3)

                    loading_at     = ec1.selectbox("Loading At",    ["ICD", "ALPHA"],
                                                    index=["ICD","ALPHA"].index(row_data.get("loading_at","ICD"))
                                                    if row_data.get("loading_at") in ["ICD","ALPHA"] else 0)
                    fcl_or_lcl     = ec2.selectbox("FCL/LCL",       ["FCL","LCL"],
                                                    index=["FCL","LCL"].index(row_data.get("fcl_or_lcl","FCL"))
                                                    if row_data.get("fcl_or_lcl") in ["FCL","LCL"] else 0)
                    by_air_or_sea  = ec3.selectbox("Mode",          ["Sea","Air"],
                                                    index=["Sea","Air"].index(row_data.get("by_air_or_sea","Sea"))
                                                    if row_data.get("by_air_or_sea") in ["Sea","Air"] else 0)

                    ec4, ec5, ec6 = st.columns(3)
                    country        = ec4.text_input("Country",           value=row_data.get("country") or "")
                    port_of_dest   = ec5.text_input("Port of Dest.",     value=row_data.get("port_of_destination") or "")
                    liner_name     = ec6.text_input("Liner",             value=row_data.get("liner_name") or "")

                    ec7, ec8 = st.columns([2, 1])
                    vessel_name    = ec7.text_input("Vessel / Voyage",   value=row_data.get("vessel_name") or "")
                    paperless_code = ec8.text_input("Paperless Code",    value=row_data.get("paperless_code") or "")

                    ec9, ec10, ec11 = st.columns(3)
                    no_container   = ec9.number_input("No. Container",   min_value=0,
                                                       value=int(row_data.get("no_container") or 0))
                    container_type = ec10.text_input("Container Type",   value=row_data.get("container_type") or "")
                    no_pallet      = ec11.number_input("No. Pallet",     min_value=0,
                                                        value=int(row_data.get("no_pallet") or 0))

                    ec12, ec13 = st.columns(2)
                    cy_at          = ec12.text_input("CY At",            value=row_data.get("cy_at") or "")
                    return_place   = ec13.text_input("Return Place",     value=row_data.get("return_place") or "")

                    st.markdown("**วันที่สำคัญ**")
                    ed1, ed2, ed3 = st.columns(3)
                    etd            = ed1.text_input("ETD (dd/mm/yyyy)",  value=row_data.get("etd") or "")
                    eta            = ed2.text_input("ETA (dd/mm/yyyy)",  value=row_data.get("eta") or "")
                    cy_date        = ed3.text_input("CY Date",           value=row_data.get("cy_date") or "")

                    ed4, ed5, ed6 = st.columns(3)
                    liner_cutoff   = ed4.text_input("Liner Cutoff",      value=row_data.get("liner_cutoff") or "")
                    vgm_cutoff     = ed5.text_input("VGM Cutoff",        value=row_data.get("vgm_cutoff") or "")
                    si_cutoff      = ed6.text_input("SI Cutoff",         value=row_data.get("si_cutoff") or "")

                    ed7, ed8 = st.columns(2)
                    return_date    = ed7.text_input("1st Return Date",   value=row_data.get("return_date_1st") or "")

                    submitted = st.form_submit_button("💾 บันทึกการแก้ไข", use_container_width=True)

                if submitted:
                    update_payload = {
                        "booking_no":        edit_bk,
                        "loading_at":        loading_at,
                        "fcl_or_lcl":        fcl_or_lcl,
                        "by_air_or_sea":     by_air_or_sea,
                        "country":           country or None,
                        "port_of_destination": port_of_dest or None,
                        "liner_name":        liner_name or None,
                        "vessel_name":       vessel_name or None,
                        "paperless_code":    paperless_code or None,
                        "no_container":      no_container or None,
                        "container_type":    container_type or None,
                        "no_pallet":         no_pallet or None,
                        "cy_at":             cy_at or None,
                        "return_place":      return_place or None,
                        "etd":               etd or None,
                        "eta":               eta or None,
                        "cy_date":           cy_date or None,
                        "liner_cutoff":      liner_cutoff or None,
                        "vgm_cutoff":        vgm_cutoff or None,
                        "si_cutoff":         si_cutoff or None,
                        "return_date_1st":   return_date or None,
                    }
                    try:
                        supabase.table(TBL_BOOKINGS).upsert(update_payload).execute()
                        st.success(f"✅ บันทึก {edit_bk} เรียบร้อยแล้ว")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Save Error: {e}")

        else:
            st.info("📌 ยังไม่มีข้อมูล — อัปโหลด PDF เพื่อเริ่มต้น")
 
    except Exception as e:
        st.error(f"Load Error: {e}")
 
    # ── Revision History ─────────────────
    st.divider()
    st.subheader("📜 ประวัติการบันทึกย้อนหลัง (Revision Logs)")
 
    if st.button("🔍 โหลดประวัติทั้งหมด"):
        st.session_state["show_history"] = True
 
    if st.session_state.get("show_history"):
        try:
            rev = (
                supabase.table(TBL_REVISIONS)
                .select("*")
                .order("created_at", desc=True)
                .execute()
            )
            if rev.data:
                df_rev = pd.DataFrame(rev.data)
                df_rev = bkk_time(df_rev, "created_at")
                hist_cols = [c for c in COLUMNS_ORDER if c in df_rev.columns] + (
                    ["created_at"] if "created_at" in df_rev.columns else []
                )
                df_rev = df_rev[hist_cols]
 
                s_hist = st.text_input("🔍 ค้นหาในประวัติ...", key="hist_search")
                if s_hist:
                    mask = df_rev.astype(str).apply(
                        lambda x: x.str.contains(s_hist, case=False, na=False)
                    ).any(axis=1)
                    df_rev = df_rev[mask]
 
                df_rev.index = range(1, len(df_rev) + 1)
                render_table(df_rev, table_id="history")
                st.download_button(
                    "📥 Export History",
                    data=to_excel(df_rev),
                    file_name="DHL_Bookings_History.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("ยังไม่มีประวัติ")
        except Exception as e:
            st.error(f"History Error: {e}")
 
# ─────────────────────────────────────────
# 8. PAGE: GENERATE SI
# ─────────────────────────────────────────
elif page == "📄 Generate SI (Draft)":
 
    # ── imports เพิ่มเติมสำหรับ Auto SI ──────────────────────
    import copy
    from openpyxl import load_workbook
 
    PROMPT_INVOICE = """You are a DHL logistics expert. Extract shipping info from this Invoice & Packing List PDF.
Return ONLY a JSON array — one object per INVOICE found. No markdown, no explanation.
[{
  "invoice_no": "e.g. 1075863",
  "description": "SHORT description of goods",
  "shipping_mark": "ALL lines under CASE MARK / SHIPPING MARK section joined with \\n, e.g. 'PO#5400025474\\nHS CODE : 841590'. Copy every line exactly as printed. Do NOT add HS CODE if it is not written in that section.",
  "cartons": "1,389 CARTONS or 40 PP.PALLETS — full package count with unit from TOTAL row",
  "quantity_str": "1,389 SETS or 98,470 PCS — quantity with unit",
  "net_weight_kgs": 33053.00,
  "gross_weight_kgs": 37221.00,
  "measurement_cbm": 282.895,
  "hs_code": "8415.10 or null",
  "consignee_name": "ACCOUNTEE name",
  "consignee_address": "full address, use \\n for line breaks",
  "ship_to_name": "SHIP TO name",
  "ship_to_address": "full address, use \\n for line breaks",
  "vessel_feeder": "feeder vessel + voyage e.g. X-PRESS ANGLESEY V.26002W",
  "vessel_mother": "mother vessel + voyage e.g. ONE HAMMERSMITH V.088W or null",
  "port_of_loading": "FROM port e.g. LAEM CHABANG, THAILAND",
  "port_of_discharge": "TO port e.g. LE HAVRE, FRANCE",
  "transhipment_port": "VIA port or null",
  "etd": "SAILING ON/OR ABOUT date dd/mm/yyyy",
  "carrier": "CARRIER field e.g. EXPEDITORS/ONE — look for 'CARRIER:' label near vessel/voyage info"
}]
Rules: Extract TOTAL row from Packing List. gross_weight/cbm from PL TOTAL.
For 'cartons': always include the unit word (CARTONS, PP.PALLETS, CTNS, etc.) not just the number.
null if not found."""
 
    def _extract_invoices(file_bytes):
        ai_cfg = types.GenerateContentConfig(
            response_mime_type="application/json", temperature=0.0, seed=42
        )
        res = genai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[types.Content(role="user", parts=[
                types.Part.from_text(text=PROMPT_INVOICE),
                types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
            ])],
            config=ai_cfg,
        )
        raw = res.text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"): raw = raw[4:]
        data = json.loads(raw.strip())
        return data if isinstance(data, list) else [data]
 
    def _fill_si(template_bytes, booking, invoices, containers, extra):
        from openpyxl.cell.cell import MergedCell
        from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
        import copy
 
        wb = load_workbook(io.BytesIO(template_bytes))
        ws = wb["Shipping Particular"]
        s  = lambda v: str(v).strip() if v else ""
        first = invoices[0] if invoices else {}
 
        def safe_write(addr, value):
            cell = ws[addr]
            if not isinstance(cell, MergedCell) and (cell.value is None or str(cell.value).strip() == ""):
                cell.value = value
 
        def safe_write_rc(row, col, value):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell, MergedCell) and (cell.value is None or str(cell.value).strip() == ""):
                cell.value = value
 
        def copy_row_style(src_row, dst_row):
            """copy style ทุก cell จาก src_row ไป dst_row"""
            for c in range(1, 12):
                src = ws.cell(row=src_row, column=c)
                dst = ws.cell(row=dst_row, column=c)
                if isinstance(src, MergedCell) or isinstance(dst, MergedCell):
                    continue
                dst.font      = copy.copy(src.font)
                dst.border    = copy.copy(src.border)
                dst.alignment = copy.copy(src.alignment)
                dst.fill      = copy.copy(src.fill)
                dst.number_format = src.number_format
            ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
 
        # ── Header ─────────────────────────────────────────────
        safe_write("J7", s(booking.get("booking_no")))
        if extra.get("revised"):
            safe_write("I9", "REVISED")
 
        # ── Consignee rows 15-18 ────────────────────────────────
        con_name   = s(first.get("consignee_name"))
        con_addr   = s(first.get("consignee_address"))
        addr_lines = [l.strip() for l in con_addr.replace("\\n","\n").split("\n") if l.strip()]
        a15_cell = ws["A15"]
        if not isinstance(a15_cell, MergedCell) and (a15_cell.value is None or str(a15_cell.value).strip() == ""):
            safe_write("A15", con_name)
            for i, line in enumerate(addr_lines[:5]):
                safe_write(f"A{16+i}", line)
 
        # ── Notify rows 22-25 ───────────────────────────────────
        notify_name = s(first.get("ship_to_name")) or con_name
        notify_addr = s(first.get("ship_to_address")) or con_addr
        nlines = [l.strip() for l in notify_addr.replace("\\n","\n").split("\n") if l.strip()]
        a22_cell = ws["A22"]
        if not isinstance(a22_cell, MergedCell) and (a22_cell.value is None or str(a22_cell.value).strip() == ""):
            safe_write("A22", notify_name)
            for i, line in enumerate(nlines[:4]):
                safe_write(f"A{23+i}", line)
 
        # ── Vessel / Port (จาก Invoice SAP) ────────────────────
        safe_write("A28", s(first.get("vessel_feeder")))
        safe_write("D28", s(first.get("port_of_loading")) or "LAEM CHABANG, THAILAND")
        safe_write("I28", s(first.get("port_of_loading")) or "LAEM CHABANG, THAILAND")
 
        etd_str = s(first.get("etd"))
        try:
            etd_dt = datetime.strptime(etd_str, "%d/%m/%Y")
            cell_a30 = ws["A30"]
            if not isinstance(cell_a30, MergedCell):
                cell_a30.value = etd_dt
                cell_a30.number_format = "DD-MMM-YY"
        except Exception:
            safe_write("A30", etd_str)
 
        safe_write("D30", s(first.get("port_of_discharge")))
        safe_write("I30", s(first.get("port_of_discharge")))  # Place of Delivery = same
 
        # A32 = Place of Issue (BANGKOK, THAILAND — static)
        safe_write("A32", "BANGKOK, THAILAND")
        # D32 = Transhipment port
        trans = s(first.get("transhipment_port"))
        safe_write("D32", trans + ",SINGAPORE" if trans and "SINGAPORE" in trans.upper() and "," not in trans else trans)
        safe_write("I32", s(first.get("vessel_mother")))
 
        # ── I13 = Carrier (จาก Invoice SAP) ────────────────────
        carrier = s(first.get("carrier"))
        safe_write("I13", carrier)
 
        # ── Container count label ───────────────────────────────
        safe_write("B35", f"{len(containers)}X40' HC" if containers else "")
 
        # ── I36/J36 = KGS. / CBM (หน่วยใต้ตัวเลข total) ───────
        safe_write("I36", "KGS.")
        safe_write("J36", "CBM")
 
        # ── Cargo block ─────────────────────────────────────────
        # Template layout:
        #   row 37: B=CARTONS (total qty label — มีแค่อันเดียว), formula total
        #   row 38: A=mark1, B=CARTONS(หน่วย—มีแค่ invoice แรก), C=qty, D=CARTONS, E=(qty_str)
        #   row 39: A=mark2, C=description
        #   row 40: A=mark3, C=INVOICE NO.
        #   row 41: C=G.W., D=gw, E=KGS, F=M3, G=cbm, H=CBM
        #   row 42: (blank spacer)
        #   row 43: invoice 2 qty row (ไม่มี B=CARTONS แล้ว)
        #   ...
 
        marks = list(dict.fromkeys(
            s(inv.get("shipping_mark")) for inv in invoices if inv.get("shipping_mark")
        ))
 
        MARK_START = 38
        ROWS_PER   = 5
        MAX_INV    = 10
 
        # ── helper: สีพื้นหลัง ──────────────────────────────────
        HIGHLIGHT_FILL = PatternFill(fill_type="solid", fgColor="BDD7EE")
 
        # ── คำนวณ row ก่อนเพื่อให้รู้ bl_row ──────────────────────
        cargo_end = MARK_START + len(invoices) * ROWS_PER
        hs_row    = cargo_end + 2
        fr_row    = hs_row + 2
        bl_row    = fr_row + 1
 
        # ════════════════════════════════════════════════════════
        # STEP 1: ล้าง border ทั้งหมด row 34 → bl_row+20
        # ════════════════════════════════════════════════════════
        for r in range(34, bl_row + 20):
            for c in range(1, 12):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell, MergedCell): continue
                cell.border = Border()
 
        # ════════════════════════════════════════════════════════
        # STEP 2: เคลียร์ค่า (value) ในพื้นที่ cargo
        # ════════════════════════════════════════════════════════
        clear_end = MARK_START + MAX_INV * ROWS_PER + 2
        for r in range(MARK_START, clear_end):
            for c in range(1, 12):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell, MergedCell): continue
                if isinstance(cell.value, str) and cell.value.startswith("="): continue
                cell.value = None
 
        # ════════════════════════════════════════════════════════
        # STEP 3: วาด border ใหม่ทีละ column (row34 → bl_row)
        # ════════════════════════════════════════════════════════
        thin = Side(border_style="thin")
 
        # col A (1): left+right ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=1)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                left=thin, right=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # col B (2): left ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=2)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                left=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # col C (3): left ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=3)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                left=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # col D-G (4-7): ไม่มี left/right (พื้นที่ description), bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            for c in range(4, 8):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell, MergedCell): continue
                cell.border = Border(
                    bottom=thin if r == bl_row else Side(border_style=None)
                )
 
        # col H (8): right ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=8)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                right=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # col I (9): right ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=9)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                right=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # col J (10): left+right ทุก row, bottom เฉพาะ bl_row
        for r in range(34, bl_row + 1):
            cell = ws.cell(row=r, column=10)
            if isinstance(cell, MergedCell): continue
            cell.border = Border(
                left=thin, right=thin,
                bottom=thin if r == bl_row else Side(border_style=None)
            )
 
        # ════════════════════════════════════════════════════════
        # STEP 4: เพิ่ม top border row 34 (ใต้ header row 33)
        # ════════════════════════════════════════════════════════
        for c in range(1, 11):
            cell = ws.cell(row=34, column=c)
            if isinstance(cell, MergedCell): continue
            old = cell.border
            cell.border = Border(top=thin, left=old.left, right=old.right, bottom=old.bottom)
 
        # Shipping Marks ลง col A — แยกแต่ละบรรทัดลงคนละ row
        mark_row = MARK_START
        for mark in marks[:MAX_INV]:
            for line in str(mark).split("\n"):
                line = line.strip()
                if line:
                    safe_write_rc(mark_row, 1, line)
                    mark_row += 1
 
        gw_cells, cbm_cells, ctn_cells = [], [], []
 
        for idx, inv in enumerate(invoices):
            base = MARK_START + idx * ROWS_PER
            for offset in range(ROWS_PER):
                copy_row_style(38 + offset, base + offset)
 
            # col B = "CARTONS" เฉพาะ invoice แรกเท่านั้น
            if idx == 0:
                safe_write_rc(base, 2, "CARTONS")
 
            # row+0: qty — ใส่สีเฉพาะ C (จำนวน + unit เช่น "1,389 CARTONS" หรือ "40 PP.PALLETS")
            safe_write_rc(base, 3, s(inv.get("cartons")) or 0)
            safe_write_rc(base, 5, f"({s(inv.get('quantity_str'))})")
            cell_c = ws.cell(row=base, column=3)
            if not isinstance(cell_c, MergedCell):
                cell_c.fill = copy.copy(HIGHLIGHT_FILL)
 
            # row+1: description
            safe_write_rc(base+1, 3, s(inv.get("description")))
 
            # row+2: invoice no.
            safe_write_rc(base+2, 3, f"INVOICE NO. {s(inv.get('invoice_no'))}")
 
            # row+3: G.W. — ใส่สีเฉพาะ D (ตัวเลข GW) และ G (ตัวเลข CBM)
            safe_write_rc(base+3, 3, "G.W.  ")
            safe_write_rc(base+3, 4, inv.get("gross_weight_kgs") or 0)
            safe_write_rc(base+3, 5, "KGS")
            safe_write_rc(base+3, 6, "M3")
            safe_write_rc(base+3, 7, inv.get("measurement_cbm") or 0)
            safe_write_rc(base+3, 8, "CBM")
            for col in [4, 7]:   # D=ตัวเลข GW, G=ตัวเลข CBM
                cell = ws.cell(row=base+3, column=col)
                if not isinstance(cell, MergedCell):
                    cell.fill = copy.copy(HIGHLIGHT_FILL)
 
            gw_cells.append(f"D{base+3}")
            cbm_cells.append(f"G{base+3}")
            ctn_cells.append(f"C{base}")
 
        if gw_cells:
            safe_write("I35", "=" + "+".join(gw_cells))
            safe_write("J35", "=" + "+".join(cbm_cells))
        if ctn_cells:
            # B37 = reference C ของ invoice แรก (cartons เป็น string เช่น "1,389 CARTONS")
            safe_write("B37", f"={ctn_cells[0]}")
 
        # ── HS Code, Freight, BL ─────────────────────────────────
        # bl_row / hs_row / fr_row คำนวณและวาด border ไว้แล้วใน STEP 1-4
 
        for r in [hs_row, fr_row, bl_row]:
            try:
                ws.merge_cells(f"C{r}:H{r}")
            except Exception:
                pass
            # ล้าง right border ที่ merge_cells สร้างให้อัตโนมัติ
            for c in range(3, 9):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell, MergedCell): continue
                old = cell.border
                cell.border = Border(
                    top=old.top, left=old.left,
                    right=Side(border_style=None),
                    bottom=old.bottom,
                )
 
        hs = s(extra.get("hs_code_all"))
        if hs:
            safe_write_rc(hs_row, 3, f"HS CODE : {hs}")
            cell = ws.cell(row=hs_row, column=3)
            if not isinstance(cell, MergedCell):
                cell.alignment = Alignment(horizontal="center")
 
        safe_write_rc(fr_row, 3, s(extra.get("freight_terms")) or "FREIGHT COLLECT")
        safe_write_rc(bl_row, 3, s(extra.get("bl_type")) or "Sea Waybill")
        for r in [fr_row, bl_row]:
            cell = ws.cell(row=r, column=3)
            if not isinstance(cell, MergedCell):
                cell.alignment = Alignment(horizontal="center")
 
        # ── Container table (optional — แสดงเฉพาะเมื่อ user กรอก cont_no) ──
        filled_containers = [c for c in containers if s(c.get("cont_no"))]
 
        if filled_containers:
            CONT_START = bl_row + 2
            ws.row_dimensions[CONT_START].height = ws.row_dimensions[59].height
            for c, hdr in enumerate(
                ["CONT. NO.","SEAL NO.","QTY","TYPE OF PACKAGE","GW.","M3","SIZE CONT","TARE WEIGHT","DT","VGM","HS CODE"],
                start=1
            ):
                safe_write_rc(CONT_START, c, hdr)
 
            for idx, cont in enumerate(filled_containers):
                r = CONT_START + 1 + idx
                ws.row_dimensions[r].height = ws.row_dimensions[60].height
                safe_write_rc(r, 1,  s(cont.get("cont_no")))
                safe_write_rc(r, 2,  s(cont.get("seal_no")))
                safe_write_rc(r, 3,  cont.get("cartons"))
                safe_write_rc(r, 4,  "CARTONS")
                safe_write_rc(r, 5,  cont.get("gw"))
                safe_write_rc(r, 6,  cont.get("cbm"))
                safe_write_rc(r, 7,  s(cont.get("size")) or "40 ' HQ")
                safe_write_rc(r, 8,  cont.get("tare"))
                safe_write_rc(r, 9,  cont.get("dt") or 10)
                safe_write_rc(r, 10, f"=E{r}+H{r}+I{r}")
                safe_write_rc(r, 11, s(cont.get("hs_code")))
 
            last_cont  = CONT_START + len(filled_containers)
            total_cont = last_cont + 1
            copy_row_style(65, total_cont)
            safe_write_rc(total_cont, 3, f"=SUM(C{CONT_START+1}:C{last_cont})")
            safe_write_rc(total_cont, 5, f"=SUM(E{CONT_START+1}:E{last_cont})")
            safe_write_rc(total_cont, 6, f"=SUM(F{CONT_START+1}:F{last_cont})")
 
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
 
    # ── UI ──────────────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex;align-items:center;gap:16px;padding:14px 20px;
                background:#ffffff;border:1px solid #e8e8e8;border-left:4px solid #D40511;
                border-radius:12px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
        <div style="background:#FFCC00;padding:8px 16px;border-radius:6px;">
            <span style="color:#D40511;font-family:'Arial Black',sans-serif;font-size:22px;font-weight:900;">DHL</span>
        </div>
        <div>
            <div style="color:#111;font-weight:700;font-size:17px;">Auto SI Generator</div>
            <div style="color:#999;font-size:12px;">เลือก Booking → อัปโหลด Invoice → Generate SI.xlsx</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    # Step 1: Booking
    st.markdown("**Step 1 · เลือก Booking**")
    booking_map = {}
    try:
        res2 = supabase.table(TBL_BOOKINGS).select("*").order("updated_at", desc=True).execute()
        if res2.data:
            for row in res2.data:
                if row.get("booking_no"):
                    booking_map[row["booking_no"]] = row
    except Exception as e:
        st.error(f"Load bookings error: {e}")
 
    bk_options  = ["-- เลือก Booking No. --"] + list(booking_map.keys())
    selected_bk = st.selectbox("Booking No.", bk_options, label_visibility="collapsed")
    bk          = booking_map.get(selected_bk, {})
 
    if bk:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Booking No.", bk.get("booking_no","—"))
        m2.metric("Container", f"{bk.get('no_container','—')}×{bk.get('container_type','—')}")
        m3.metric("Paperless", bk.get("paperless_code","—"))
        m4.metric("SI Cutoff", bk.get("si_cutoff","—"))
 
    st.divider()
 
    # Step 2: Invoice PDFs
    st.markdown("**Step 2 · อัปโหลด Invoice PDF** (โยนพร้อมกันหลายใบได้)")
    inv_files = st.file_uploader(
        "Invoice PDFs", type="pdf", accept_multiple_files=True, key="si_inv_up"
    )
 
    st.divider()
 
    # Step 3: SI Template
    st.markdown("**Step 3 · อัปโหลด SI Template (.xlsx)**")
    tmpl_file = st.file_uploader("SI Template", type="xlsx", key="si_tmpl_up")
 
    st.divider()
 
    # Step 4: ข้อมูลเสริม
    st.markdown("**Step 4 · ข้อมูลเสริม**")
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        freight_terms = st.selectbox("Freight Terms", ["FREIGHT COLLECT","FREIGHT PREPAID"])
        bl_type       = st.selectbox("BL Type", ["Sea Waybill","Original B/L","Telex Release"])
        revised       = st.checkbox("REVISED", value=False)
    with col_e2:
        hs_code_input = st.text_input("HS Code (คั่นด้วย , )", placeholder="8415.10, 3926.90")
 
    # Container table
    st.markdown("**ข้อมูล Container**")
    n_cont = st.number_input("จำนวน Container", min_value=1, max_value=20,
                              value=int(bk.get("no_container") or 1))
 
    cont_header = st.columns([2,2,1.2,1.2,1.2,1.2,1.2,1])
    for h, col in zip(["CONT. NO.","SEAL NO.","CARTONS","G.W.(KGS)","CBM","TARE(KGS)","SIZE","DT"], cont_header):
        col.markdown(f"<div style='font-size:11px;font-weight:600;color:#999;'>{h}</div>", unsafe_allow_html=True)
 
    container_rows = []
    for i in range(int(n_cont)):
        cols = st.columns([2,2,1.2,1.2,1.2,1.2,1.2,1])
        container_rows.append({
            "cont_no": cols[0].text_input("", key=f"cno_{i}",  placeholder=f"CONT {i+1}", label_visibility="collapsed"),
            "seal_no": cols[1].text_input("", key=f"sno_{i}",  placeholder="SEAL",        label_visibility="collapsed"),
            "cartons": cols[2].number_input("", key=f"ctn_{i}", min_value=0, value=0,      label_visibility="collapsed"),
            "gw":      cols[3].number_input("", key=f"cgw_{i}", min_value=0.0, value=0.0, label_visibility="collapsed", format="%.2f"),
            "cbm":     cols[4].number_input("", key=f"ccb_{i}", min_value=0.0, value=0.0, label_visibility="collapsed", format="%.3f"),
            "tare":    cols[5].number_input("", key=f"ctr_{i}", min_value=0, value=3900,   label_visibility="collapsed"),
            "size":    cols[6].selectbox("",   key=f"csz_{i}", options=["40 ' HQ","40 ' GP","20 ' GP"], label_visibility="collapsed"),
            "dt":      cols[7].number_input("", key=f"cdt_{i}", min_value=0, value=10,    label_visibility="collapsed"),
            "hs_code": hs_code_input,
        })
 
    st.divider()
 
    can_gen = selected_bk != "-- เลือก Booking No. --" and inv_files and tmpl_file
    if not can_gen:
        st.info("⬆️ กรุณาเลือก Booking + อัปโหลด Invoice PDF + SI Template ก่อน")
 
    if can_gen and st.button("🚀 GENERATE SI.xlsx", use_container_width=True):
        all_invoices = []
        prog = st.progress(0)
        for idx, f in enumerate(inv_files):
            with st.spinner(f"กำลังอ่าน: {f.name}"):
                try:
                    extracted = _extract_invoices(f.read())
                    all_invoices.extend(extracted)
                    st.success(f"✅ {f.name} → {len(extracted)} invoice(s)")
                except Exception as e:
                    st.error(f"❌ {f.name}: {e}")
            prog.progress((idx+1)/len(inv_files))
 
        if not all_invoices:
            st.error("ไม่พบข้อมูล Invoice")
        else:
            # Preview
            st.markdown("---")
            st.markdown("**📋 Preview — ตรวจสอบก่อน Generate**")
            st.dataframe(pd.DataFrame([{
                "Invoice No.": inv.get("invoice_no"),
                "Description": inv.get("description"),
                "Cartons":     inv.get("cartons"),
                "QTY":         inv.get("quantity_str"),
                "GW (KGS)":    inv.get("gross_weight_kgs"),
                "CBM":         inv.get("measurement_cbm"),
                "Vessel":      inv.get("vessel_feeder"),
                "Port Disc.":  inv.get("port_of_discharge"),
                "ETD":         inv.get("etd"),
            } for inv in all_invoices]), use_container_width=True)
 
            # Generate
            with st.spinner("กำลังสร้างไฟล์ SI..."):
                try:
                    si_bytes = _fill_si(
                        template_bytes = tmpl_file.read(),
                        booking        = bk,
                        invoices       = all_invoices,
                        containers     = container_rows,
                        extra          = {
                            "hs_code_all":   hs_code_input,
                            "freight_terms": freight_terms,
                            "bl_type":       bl_type,
                            "revised":       revised,
                        },
                    )
                    inv_nos  = "_".join(inv.get("invoice_no","") for inv in all_invoices)
                    filename = f"SI_{bk.get('booking_no','BK')}_INV_{inv_nos}.xlsx"
                    st.success("🎉 สร้างไฟล์ SI สำเร็จ!")
                    st.download_button(
                        "📥 ดาวน์โหลด SI.xlsx",
                        data=si_bytes, file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"❌ Generate Error: {e}")
                    import traceback; st.code(traceback.format_exc())
