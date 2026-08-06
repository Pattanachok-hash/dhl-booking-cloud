import io
import re
import json
import base64
import pytz
import streamlit as st
import pandas as pd
import os
from google import genai
from google.genai import types
from supabase import create_client, Client
from datetime import datetime

def get_secret(key):
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key)

# ─────────────────────────────────────────
# 1. CONFIGURATION
# ─────────────────────────────────────────
SUPABASE_URL  = get_secret("SUPABASE_URL")
SUPABASE_KEY  = get_secret("SUPABASE_KEY")
GEMINI_KEY    = get_secret("GENAI_API_KEY")

if SUPABASE_URL and SUPABASE_KEY and GEMINI_KEY:
    TBL_BOOKINGS           = "bookings"
    TBL_REVISIONS          = "booking_revisions"
    TBL_LOCAL_CHARGES      = "local_charges"
    TBL_LOCAL_CHARGES_V2   = "local_charges_v2"
    TBL_LOCAL_CHARGE_ITEMS = "local_charge_items"
else:
    # สำหรับใช้รันในเครื่องตัวเอง (Local) ถ้ายังไม่ได้ตั้งค่า Secrets
    st.error("❌ ไม่พบ API Keys ในระบบ Secrets กรุณาตั้งค่าที่ Settings > Secrets")
    st.stop()
    TBL_BOOKINGS           = "test_bookings"
    TBL_REVISIONS          = "test_booking_revisions"
    TBL_LOCAL_CHARGES      = "test_local_charges"
    TBL_LOCAL_CHARGES_V2   = "test_local_charges_v2"
    TBL_LOCAL_CHARGE_ITEMS = "test_local_charge_items"

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
genai_client     = genai.Client(api_key=GEMINI_KEY)
GEMINI_MODEL_BOOKING = "models/gemini-2.5-flash"
GEMINI_MODEL         = "models/gemini-2.5-flash"

# ─────────────────────────────────────────
# 2. PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="Booking Cloud Extractor",
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
    [data-testid="stTextArea"] > div { border: 2px solid #FFCC00 !important; border-radius: 10px !important; background-color: #fffef5 !important; }
    [data-testid="stTextArea"] > div:focus-within { border-color: #D4A900 !important; box-shadow: 0 0 0 3px rgba(255,204,0,0.2) !important; }
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
            <h1 style="margin: 0; color: #333; font-size: 30px; line-height: 1.2;">CTC FG Export</h1>
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
 
PROMPT_BOOKING = """You are a DHL Logistics Analyst. Extract ALL shipping info from this booking PDF.

Return ONLY a JSON object (no markdown, no explanation):
{
  "booking_no":          "Carrier Ref or Booking No.",
  "fcl_or_lcl":          "FCL or LCL",
  "by_air_or_sea":       "Air or Sea",
  "country":             "destination country (NOT Thailand)",
  "port_of_destination": "final destination of the shipment",
  "liner_name":          "shipping line",
  "vessel_name":         "vessel/voyage (include connecting if any)",
  "no_container":        number or null,
  "container_type":      "ALL container counts+types combined e.g. '1X40HC+1X20GP' or '2X40HC' — null if LCL or no container",
  "no_pallet":           number or null,
  "cy_at":               "empty pick-up depot",
  "return_place":        "laden return location",
  "paperless_code":      "4-digit code from PAPERLESS CODE or PORT CODE label e.g. 2836, or null",
  "liner_cutoff":        "dd/mm/yyyy hh:mm or null",
  "vgm_cutoff":          "dd/mm/yyyy hh:mm or null",
  "si_cutoff":           "dd/mm/yyyy hh:mm or null",
  "return_date_1st":     "dd/mm/yyyy or null",
  "cy_date":             "dd/mm/yyyy or null",
  "etd":                 "dd/mm/yyyy or null",
  "eta":                 "dd/mm/yyyy or null"
}

Rules:
- booking_no: extract using this priority order:
  1. Value next to labels "Carrier Booking No.", "Carrier Booking Reference", "Carrier Ref" — use that directly.
  2. Value next to label "Booking No." or "Booking No" — BUT only if the value looks like an actual booking number (alphanumeric code). If the value is descriptive text (e.g. "LOAD ON SHIPPER NAME", "SEE ATTACHED", or any phrase that is clearly not a code), skip it.
  3. Value next to "Ref No", "Ref No.", "Reference No." — use as fallback if steps 1 and 2 yield nothing.
  4. Return null if nothing found.
  Do NOT use B/L No., forwarder ref (e.g. FLXCB-), CONSOL, or tracking number.
  CRITICAL: Read EVERY character/digit of the booking number EXACTLY as printed — do NOT skip, merge, or drop repeated digits. If you see "80451557" do NOT shorten it to "8045157". Count the digits twice before returning to verify the length matches what's printed.
- port_of_destination: final destination of the shipment — use the "To:" field in the booking header, or the last stop in the Intended Transport Plan. Do NOT use intermediate sea ports or terminal names (e.g. "Guadalajara Castilla La Mancha, Spain" not "APM Terminal Valencia").
- country: use consignee's country if clearly stated in address; otherwise infer from port_of_destination.
- cy_at: depot for picking up empty container. For Maersk bookings: use the location name of the "Empty Container Depot" row from the Load Itinerary table (Page 2).
- return_place: Laden Return / Return to location. For Maersk bookings: use the location name of the "Return Equip Delivery Terminal" row from the Load Itinerary table (Page 2). For LCL shipments, use the value from "Stuffing at" or "Loading at" field instead.
- paperless_code: exact 4-digit number. For Maersk bookings: first find the "Return Equip Delivery Terminal" location from the Load Itinerary table (Page 2), then look up the matching 4-digit code from the "Paperless Code" line on Page 1 — e.g. if return terminal is "Lat Krabang" find the code next to "LKB" (e.g. "B1/LKB/TICT 2811" → 2811); if "Sahathai" or "SHCT" find the code next to "SHCT" (e.g. "SHCT 0520" → 0520); if "TICT" find the code next to "TICT". For other carriers: exact 4-digit number from labels "PAPERLESS CODE", "PORT CODE", or inside parentheses like "(KERRY : 2816)" — extract only the number. If the PAPERLESS CODE section lists multiple codes by terminal (e.g. Yang Ming format with lines like "JTC : EX. LCB KERRY TERMINAL = 2816"), look at the "Turn-In At" field to identify which terminal this booking uses, then find the matching 4-digit code from that terminal's entry in the PAPERLESS CODE section.
- container_type: combine ALL container counts and types e.g. "1X40HC+1X20GP" or "2X40HC". For Maersk bookings: convert the Equipment table format — "40 DRY 9 6" = 40HC, "40 DRY" = 40GP, "20 DRY" = 20GP. Map "40H"/"40HQ" → 40HC. null if LCL or no container.
- Counting rule (applies to ALL formats): include a container size ONLY if a real number is printed as its quantity. If no number is shown for a size, that size = 0 and MUST be omitted — NEVER default an unquantified size to 1.
- [EXPEDITORS] VOLUME-row caution (Expeditors/forwarder forms like "__ X 20'   __ X 40'   __ X 40H   __ X 45"): here "X" is the multiply sign, NOT a checkbox tick — even though the SAME Expeditors form uses "X" as a tick elsewhere (e.g. "COLLECT X", "BKK X"). The quantity is the number in the blank to the LEFT of each "X"; a bare "X <size>" with no number = 0 for that size. Example: "X 20'   X 40'   3 X 40H   X 45" → container_type "3X40HC", no_container 3.
- no_container: sum of the printed per-size quantities only (e.g. 3X40HC → 3).
- Dates: dd/mm/yyyy. Cut-offs include hh:mm.
- Time format: output times as 24-hour HH:MM.
  - Convert ONLY a time written in 12-hour form (hour 1-12) that has an AM/PM marker clearly belonging to THAT time. The marker may sit right after the time or later on the SAME line (e.g. "CUT OFF TIME: 18/06/26 BEFORE 05.00 PM") — associate each AM/PM with the NEAREST time; if a line holds several times, match each marker to its own time and never borrow a marker from a different field.
  - Conversion: 1-11 PM → hour + 12 (05.00 PM → 17:00); 12 PM / noon → 12:00; 12 AM / midnight → 00:00; 1-11 AM → same hour (05.00 AM → 05:00).
  - Never add 12 to an hour already ≥ 13 — a time like "17:00" or "17.00 PM" stays 17:00.
  - If a time has NO AM/PM marker, leave it EXACTLY as printed; do NOT guess (e.g. "11:59", "22:00" stay unchanged).
- cy_date: Date the empty container is available for pick-up. For Maersk bookings: look at the Load Itinerary table (Page 2) → find the "Empty Container Depot" row → read its "Release Date" column. Convert YYYY-MM-DD to dd/mm/yyyy. For other carriers: use the empty pick-up date field.
- return_date_1st: 1st Return Date / Turn-In Date / Gate-In Date. For Maersk bookings: calculate ETD minus 5 days and use that date. For MSC bookings: use the "First Receiving" date from the DRY row in the GATE-IN AT TERMINAL/DEPOT table.
- liner_cutoff: Gate Closing / Closing Date / CY Cut-off / Last Load. For Maersk bookings: look at the Load Itinerary table (Page 2) → find the "Return Equip Delivery Terminal" row → read the location name → match EXACTLY to the cut-off on Page 1 using this mapping: "Lat Krabang" → "Cut-Off (DRY and REEF) Lat Krabang"; "TICT" → "Cut-Off TICT"; "Sahathai"/"SHCT" → "Cut-Off SHCT (Sahathai)"; "Laem Chabang" → "Cut-Off (DRY) Laem Chabang". The location name must match exactly — "Lat Krabang" is NOT the same as "TICT". Example: Return terminal = "Lat Krabang" → correct answer is "Cut-Off (DRY and REEF) Lat Krabang" = 05/04/2026 22:00, NOT Cut-Off TICT 06/04/2026 10:00. Ignore DG and REEF-only cutoffs.
- si_cutoff: SI Cut-off / Doc Cut-off / Shipping Particular Cut-off. For Maersk bookings: use "SI (Transshipment & Intra-Asia)" or "SI (Direct)" line from Page 1.
- vgm_cutoff: VGM line.
- If a cut-off shows only a weekday (e.g. "THU"), calculate the actual date from the document date or ETD.
- If ETA is given as a range (e.g. "19/May/2026 - 22/May/2026"), use the first date.
- ETD must be earlier than ETA.
- For MSC bookings: the "EST. TIME OF ARRIVAL/DEPARTURE" field shows two dates — use the SECOND date as ETD (the first date is vessel arrival at POL, the second is departure from POL).
- null if not found."""


def extract_from_pdf(file_bytes: bytes) -> list[dict]:
    """ส่ง PDF ให้ Gemini อ่านครั้งเดียว (รวม dates + general)"""
    ai_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.0,
        seed=42,
    )
    res = genai_client.models.generate_content(
        model=GEMINI_MODEL_BOOKING,
        contents=[
            types.Content(
                role="user",
                parts=[
                    types.Part.from_text(text=PROMPT_BOOKING),
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
    result = json.loads(raw.strip())
    records = result if isinstance(result, list) else [result]

    # Verify booking_no against PDF text layer (catches AI dropping/merging digits)
    try:
        from pypdf import PdfReader
        import io as _io, re as _re
        _text = "\n".join((p.extract_text() or "") for p in PdfReader(_io.BytesIO(file_bytes)).pages)
        if _text:
            for _rec in records:
                _ai_bk = (_rec.get("booking_no") or "").strip()
                if _ai_bk and _ai_bk not in _text:
                    _matches = _re.findall(rf"\w*{_re.escape(_ai_bk)}\w*", _text)
                    if _matches:
                        _rec["booking_no"] = max(_matches, key=len)
    except Exception:
        pass

    return records


# ─────────────────────────────────────────
# 5c. AIR AWB — PROMPT & FUNCTIONS
# ─────────────────────────────────────────

PROMPT_DETECT_TYPE = """Look at this shipping document and determine if it is:
- "air"  : Air Waybill (AWB / HAWB) — documents for air freight shipment
- "sea"  : Sea Booking Confirmation or Bill of Lading — documents for ocean freight

Reply with ONLY one word: air  or  sea"""


PROMPT_BOOKING_AIR = """You are a DHL Logistics Analyst. Extract shipping info from this Air Waybill (AWB/HAWB).

Return ONLY a JSON object (no markdown, no explanation):
{
  "booking_no":          "11-digit AWB number (digits only)",
  "fcl_or_lcl":          null,
  "by_air_or_sea":       "Air",
  "country":             "destination country",
  "port_of_destination": "destination airport or city",
  "liner_name":          "freight forwarder or issuing agent name",
  "vessel_name":         "flight number",
  "no_container":        null,
  "container_type":      null,
  "no_pallet":           null,
  "cy_at":               null,
  "return_place":        "Don Mueang or Suvarnabhumi",
  "paperless_code":      null,
  "liner_cutoff":        "dd/mm/yyyy or null",
  "vgm_cutoff":          null,
  "si_cutoff":           null,
  "return_date_1st":     null,
  "cy_date":             null,
  "etd":                 "dd/mm/yyyy or null",
  "eta":                 "dd/mm/yyyy or null"
}

Rules:
- booking_no: ALWAYS use the number at the TOP-LEFT corner of the document only.
  It contains letters mixed in (e.g. "940DMK15635815") — strip ALL letters, concatenate digits only → "94015635815".
  Result must be exactly 11 digits. Do NOT use HAWB No., AWB No., or any other reference number.
- liner_name: freight forwarder or issuing agent (e.g. "HELLMANN WORLDWIDE LOGISTICS CO LTD") — look for "Issued by", "Issuing Carrier's Agent", or company letterhead top-right. NOT the airline.
- vessel_name: flight number (e.g. "XJ230/03") — from "Requested Flight/Date" or "Flight/Date" field.
- etd: flight departure date — from "Requested Flight/Date" or "Flight/Date" field.
  The date is the portion after "/" in the flight number (e.g. "XJ230/03" → day 03 of the document's month/year).
  Do NOT use "Executed on (date)" — that is the AWB issue date, not the flight date. Format dd/mm/yyyy.
- liner_cutoff: ETD minus 1 day. Format dd/mm/yyyy.
- return_place: Airport of Departure — map to ONLY one of these two values:
  "Don Mueang" → if departure airport is Don Mueang / DMK / THDMK
  "Suvarnabhumi" → if departure airport is Suvarnabhumi / BKK / THBKK
- eta: arrival date if shown, otherwise null.
- country: consignee's country. If not stated, infer from Airport of Destination.
- port_of_destination: Airport of Destination city (e.g. "DELHI").
- All null fields: leave as null."""


def detect_doc_type_by_ai(file_bytes: bytes) -> str:
    res = genai_client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            types.Content(
                role="user",
                parts=[
                    types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
                    types.Part.from_text(text=PROMPT_DETECT_TYPE),
                ],
            )
        ],
    )
    return "air" if "air" in res.text.strip().lower() else "sea"


def extract_air_awb(file_bytes: bytes) -> list[dict]:
    """Extract Air Waybill — liner_cutoff คำนวณจาก ETD-1 วัน (Python override)"""
    from datetime import timedelta
    ai_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.0,
        seed=42,
    )
    res = genai_client.models.generate_content(
        model=GEMINI_MODEL_BOOKING,
        contents=[
            types.Content(
                role="user",
                parts=[
                    types.Part.from_text(text=PROMPT_BOOKING_AIR),
                    types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
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
    records = json.loads(raw.strip())
    if not isinstance(records, list):
        records = [records]

    for rec in records:
        etd_str = rec.get("etd")
        if etd_str:
            try:
                cutoff_dt = datetime.strptime(etd_str, "%d/%m/%Y") - timedelta(days=1)
                rec["liner_cutoff"] = cutoff_dt.strftime("%d/%m/%Y")
            except Exception:
                pass
    return records


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
    with pd.ExcelWriter(buf, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as writer:
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
                v = df.iloc[row_idx, col_idx]
                if pd.isna(v):
                    ws.write_blank(row_idx + 1, col_idx, None, fmt)
                else:
                    ws.write(row_idx + 1, col_idx, v, fmt)
 
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
        if isinstance(v, float) and v == int(v):
            return str(int(v))
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
# 5b. LOCAL CHARGES — PROMPT & FUNCTIONS
# ─────────────────────────────────────────
PROMPT_LOCAL_CHARGES = """You are a DHL Logistics Analyst. Extract local charge invoice data (in Thai Baht) from this PDF.

Return ONLY a JSON object (no markdown, no explanation):
{
  "agent_invoice_no": <string or null>,
  "pay_to":       <string or null>,
  "tax_name":     <string or null>,
  "tax_id":       <string or null>,
  "delivery_port": <string or null>,
  "etd":          <string DD/MM/YYYY or null>,
  "bl_no":        <string or null>,
  "due_date":     <string DD/MM/YYYY or null>,
  "vat_applicable": <true or false>,
  "items": [
    {
      "description": <string>,
      "category":    <string>,
      "wht_pct":     <0, 1, or 3>,
      "rate":        <number or null>,
      "qty":         <number or null>,
      "total":       <number>
    }
  ]
}

Rules:
- agent_invoice_no: invoice number issued by the freight forwarder/agent — look for:
  1. Labels such as "Invoice No.", "Invoice Number", "INV No.", "Invoice #" near the top of the document
  2. If no label found, look for a prominent alphanumeric code in the document title or heading (e.g. "INVOICE BKK003521Z" → extract "BKK003521Z")
- pay_to: name of the freight forwarder or agent who issued this invoice — look for the company logo, letterhead, or "From" company at the top-right or bottom of the document.
- tax_name: full name AND full address of the freight forwarder who issued this invoice, combined into one string. The issuer is the company whose logo/letterhead appears on the document — look for their address in the footer or "Service provider" section. For invoices with a Thai agent (e.g. "C/O" or "as agent for"), use the Thai local entity's name and address. NOT the "Invoice To" / "Customer" section at the top.
- tax_id: tax identification number of the issuing freight forwarder — search in the footer, bottom of page, or near the issuer's company name/address. It may appear as "Tax ID", "TAX ID", "เลขประจำตัวผู้เสียภาษี", or an unlabeled number near the issuer's details. Do NOT use the tax ID from "Invoice To" / "Customer" / "Billed To" section at the top (that is the recipient's tax ID).
- delivery_port: port of delivery or destination port — format as "Port Name, Country" e.g. "Mombasa, Kenya"
- etd: estimated time of departure in DD/MM/YYYY format
- due_date: payment due date — look for labels "Due Date", "Payment Due", "Due", "วันครบกำหนดชำระ". Format as DD/MM/YYYY. null if not found.
- bl_no: Bill of Lading number — extract using this priority:
  1. Kuehne+Nagel: use "KN TRACKING NO."
  2. Others: "House Bill of Lading", "House B/L", "HAWB", or "HAWB No." first
  3. Fallback: "B/L No.", "Bill of Lading", "OBL NO.", or "SHIPMENT" field
  4. Never use "Master Bill of Lading", "MB/L", or "MAWB"
- vat_applicable: true if invoice mentions "7% VAT", "VAT 7%", "7.00% PURSUANT TO SECTION 80 (2) OF TRC" or has a VAT line item. false otherwise.
  Exception: if the issuer is GEODIS and the VAT column explicitly shows "0%" or "0% = 0" (i.e. zero-rated), set vat_applicable = false regardless of whether a VAT column is present.
- items: list of ALL charge line items found in the invoice (exclude VAT and WHT rows — those are calculated by the system).
  CRITICAL row-handling rules:
  1. ONE row in the invoice table = ONE item in the output. Do NOT split a single row into multiple items even if it contains a sub-description, alias, or alternate name on a second line (e.g. row "Documentation fee origin" with sub-text "SURRENDER FEE / TELEX-RELEASE FEE" and amount 1,500.00 → ONE item only, with the amount 1,500.00 — NOT two items).
  2. The "Amount THB" of each item must come from the SAME row as its Charge Code/description. Do NOT mix amounts across rows. Re-verify row alignment before returning (e.g. if "B/L fee" is on row 4 with amount 1,500.00, do NOT pull the amount 450.00 from row 5).
  3. After listing all items, the sum of their totals MUST equal the invoice "Net Total" / "Sub-total" / "Total before VAT" shown in the invoice. If the sum does not match, re-read the table — likely a row was duplicated or an amount was mis-aligned.
  - description: exact charge name as shown in invoice
  - category: classify this charge into one of these fixed values:
      "thc_40hc"        → Terminal Handling Charge for 40HC container
      "thc_40dv"        → Terminal Handling Charge for 40DV/40GP container
      "thc_20gp"        → Terminal Handling Charge for 20GP container
      "export_handling" → Export Handling / Handling Origin
      "seal"            → Seal Fee
      "bl_fee"          → B/L Fee / Bill of Lading Fee
      "surrender_fee"   → Surrender Fee / Telex Release
      "vgm_fee"         → VGM Fee / VGM Submission / VGM Coordination
      "doc_amendment"   → Documentation Amendment / Doc Amendment
      "detention"       → Detention
      "demurrage"       → Demurrage
      "container_repair"→ Container Repair
      "edi_fee"         → EDI Fee / EDI Transmission
      "late_gate"       → Late Gate / Late Gate Service
      "environmental_fee" → Environmental Fee / Green Fee
      "storage"           → Storage / Storage Fee / Container Storage
      "freight_charge"    → Freight / Freight Charge / Ocean Freight / Air Freight / Sea Freight
      "other"           → anything that does not match the above
  - wht_pct: WHT rate for this item. Determine using this priority:
    1. If the charge has a "(W/H 1%)" or "(W/H 3%)" label next to it → use that rate.
    2. If there is a "WHT IN THB" column with entries like "1%=93.8" or "3%=30" next to the charge:
       - Read the digit BEFORE the % sign as the wht_pct (e.g. "1%=93.8" → wht_pct=1, "3%=30" → wht_pct=3)
       - The number AFTER "=" is the pre-calculated WHT amount — do NOT use it as the item total
       - Item total must come from the CHARGES IN THB column
       - IMPORTANT: If this column exists in the invoice, apply it to ALL items and do NOT use rule 5 (Expeditors) at all — even for SEAL, VGM, HANDLING
    2b. If the table has a "W/T%" column with per-row values like "01" or "03" (e.g. Maritime Alliance format):
       - Read the value for each row as wht_pct ("01" → 1, "03" → 3)
       - This takes priority over rules 3–6. Do NOT use global remarks or issuer-based rules.
    3. If there is no per-item label but the document has a global remark applying WHT to all charges (e.g. "Please deduct 3% withholding tax from total Service Charge", "หัก ณ ที่จ่าย 3%") → apply that rate to ALL items.
    4. If the invoice contains "PLEASE PAY WITHOUT DEDUCTION" or "NO DEDUCTION" → wht_pct = 0 for ALL items. Do NOT apply rule 5 (Expeditors).
    5. If the issuer is Expeditors and no WHT is stated in the invoice, apply based on description keywords:
       - WHT 1%: description contains "THC" (any container type), "B/L", "BL FEE", "BILL OF LADING", or "SURRENDER"
       - WHT 3%: description contains "SEAL", "HANDLING", or "VGM"
       - WHT 0%: all other charges
    5b. If the issuer is CEVA and no per-item WHT is stated → wht_pct = 3 for ALL items.
    5c. If the issuer is DSV and no per-item WHT is stated, apply based on shipment_type context provided:
       - Ocean Export:
         WHT 3%: description is specifically "Export Handling", "Handling Fee", or "Handling Charge" (standalone handling service — NOT Terminal Handling Charge / THC)
         WHT 1%: ALL other charges including THC (Terminal Handling Charge), B/L Fee, Surrender Fee, SEAL, VGM, Environmental Fee, Security Fee, EDI Fee, Late Gate, Detention, Demurrage, and any other non-handling charge
       - Air Export / Air Import / Ocean Import:
         WHT 0%: description contains "FREIGHT" or "OCEAN FREIGHT" or "AIR FREIGHT"
         WHT 1%: description contains "TRANSPORT" or "TRUCKING" or "DELIVERY"
         WHT 3%: ALL other charges (Export Handling, Handling Fee, SEAL, VGM, ENVIRONMENTAL, SECURITY, EDI, etc.)
    5d. If the issuer is GEODIS and no per-item WHT is stated:
       - WHT 0%: description contains "Late Pickup B/L" or "Late Pick-up B/L"
       - WHT 3%: any item whose description does NOT contain "Late Pickup B/L" or "Late Pick-up B/L"
    5e. If the issuer is DACHSER and no per-item WHT is stated → wht_pct = 3 for ALL items.
    5f. If the issuer is LOGWIN and no per-item WHT is stated, apply based on description keywords. IMPORTANT: check the WHT 3% list BEFORE the WHT 1% list (so "BL AMEND FEE" is matched as 3% before the generic "BL" 1% rule):
       - WHT 3%: description contains "ENS", "AFR", "AMS", "ACI", "EDI", "VGM", "HANDLING", "FUMIGATE", "BL AMEND", "LIABILITY", "COMPLIANCE", "DATA TRANSFER", "ICS2", "EMISSION DETERMINATION", "LOGWIN SERVICE FEE", or any oversea/overseas charge
       - WHT 1%: description contains "FREIGHT", "GRI", "LSS", "BUNKER", "DGR", "PEAK SEASON", "SEAL", "THC", "CFS", "DEM", "DET", "BL SUR", "TEX", "DISRUPT", "FUEL", or "BL"
       - WHT 0%: all other charges
    6. Default: 0
  - rate: unit rate in THB. If the invoice has a RATE column in foreign currency with an EXCH RATE column, convert: rate = RATE × EXCH RATE. If no rate is shown (flat fee), set rate = total.
    Exception: for DSV or SEKO THC items where qty is derived from the container count (see qty rule below), if no per-unit rate column is present set rate = round(total / qty, 2).
  - qty: number of units. If the invoice has an explicit QTY column, always use that value — even if the rate is in a foreign currency.
    Special rule for DSV and SEKO — THC items only: if the issuer is DSV or SEKO, the item category is thc_40hc, thc_40dv, or thc_20gp, and there is no explicit QTY column:
      1. Find the CONTAINERS section and count containers by type (formats vary: "TCNU5588680 - 40HC" or "BEAU5332480 (40HC - Forty foot high cube)"):
         thc_40hc → count containers labeled 40HC, 45HC, or 40HQ
         thc_40dv → count containers labeled 40GP, 40DV, or 40ST
         thc_20gp → count containers labeled 20GP, 20DV, or 20ST
      2. Single container type (all containers same type, one THC line): set qty = total container count for that type.
      3. Mixed container types (e.g. both 40HC and 20GP, two THC lines with no type label): try both possible assignments and choose the one where the per-unit rate of the 40HC line (total ÷ count_40hc) is greater than the per-unit rate of the 20GP line (total ÷ count_20gp) — this is always true because 40HC THC rate is always higher than 20GP THC rate. Set qty accordingly for each line.
      4. If no CONTAINERS section is found, fall back to qty = 1.
    Default (all other cases): if no quantity is shown (flat fee), set qty = 1.
  - total: pre-VAT total amount for this line item. If the invoice shows separate columns such as "Subject to VAT Amount", "NON-VAT Amount", and "Total Amount" (VAT-inclusive), use "Subject to VAT Amount" or "NON-VAT Amount" — NOT the "Total Amount" column. Never include VAT in this value.
- Use numeric values only (no currency symbols, no commas). null if not found.
"""


def extract_local_charges(file_bytes: bytes, shipment_type: str = "") -> dict:
    ai_config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.0,
        seed=42,
        thinking_config=types.ThinkingConfig(thinking_budget=0),
    )
    prompt = PROMPT_LOCAL_CHARGES
    if shipment_type:
        prompt += f"\n\nContext: shipment_type = \"{shipment_type}\" (use this for DSV WHT rule 5b)"
    res = genai_client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            types.Content(
                role="user",
                parts=[
                    types.Part.from_text(text=prompt),
                    types.Part.from_bytes(data=file_bytes, mime_type="application/pdf"),
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
    result = json.loads(raw.strip())

    # Handle case where AI returns a list (multiple invoices in one PDF)
    if isinstance(result, list):
        result = result[0] if result else {}
        result["_multi_invoice"] = True

    items = result.get("items") or []

    # For DSV and SEKO: set seal qty = sum of all THC quantities
    pay_to = (result.get("pay_to") or "").upper()
    if "DSV" in pay_to or "SEKO" in pay_to:
        thc_categories = {"thc_40hc", "thc_40dv", "thc_20gp"}
        total_thc_qty = sum(int(it.get("qty") or 1) for it in items if it.get("category") in thc_categories)
        if total_thc_qty > 0:
            for it in items:
                if it.get("category") == "seal":
                    it["qty"] = total_thc_qty
                    if it.get("total") and total_thc_qty > 0:
                        it["rate"] = round(float(it["total"]) / total_thc_qty, 2)

    # Calculate VAT 7% in Python
    if result.get("vat_applicable"):
        subtotal = sum(float(it.get("total") or 0) for it in items)
        if subtotal > 0:
            result["vat_7"] = round(subtotal * 0.07, 2)

    # Calculate WHT 1% and 3% from per-item wht_pct
    wht1_sum = sum(float(it.get("total") or 0) for it in items if int(it.get("wht_pct") or 0) == 1)
    wht3_sum = sum(float(it.get("total") or 0) for it in items if int(it.get("wht_pct") or 0) == 3)
    result["wht_1"] = round(wht1_sum * 0.01, 2) if wht1_sum > 0 else None
    result["wht_3"] = round(wht3_sum * 0.03, 2) if wht3_sum > 0 else None

    return result


def save_local_charge_v2(header: dict, items: list, pdf_bytes: bytes = None, filename: str = None) -> bool:
    try:
        # อัพโหลด invoice PDF ไปยัง Supabase Storage ก่อน
        if pdf_bytes and filename:
            import uuid as _uuid
            import re as _re
            safe_name = _re.sub(r"[^A-Za-z0-9._-]+", "_", filename).strip("_")
            path = f"{_uuid.uuid4()}_{safe_name}"
            supabase.storage.from_("local-charge-invoices").upload(
                path, pdf_bytes, {"content-type": "application/pdf"}
            )
            header["invoice_pdf_path"] = path
        res = supabase.table(TBL_LOCAL_CHARGES_V2).insert(header).execute()
        lc_id = res.data[0]["id"]
        for item in items:
            item["local_charge_id"] = lc_id
        supabase.table(TBL_LOCAL_CHARGE_ITEMS).insert(items).execute()
        return True
    except Exception as e:
        st.error(f"❌ Supabase Error: {e}")
        return False


# ─────────────────────────────────────────
# 5c. EXPORT SUMMARY — PDF GENERATOR
# ─────────────────────────────────────────
def generate_expense_pdf(records: list[dict], prepared_by: str = "", prepared_by_phone: str = "") -> bytes:
    """Generate expense summary PDF. records = list of {header, items}"""
    from fpdf import FPDF

    from pathlib import Path
    import tempfile, urllib.request
    _win = Path("C:/Windows/Fonts")
    if _win.exists() and (_win / "tahoma.ttf").exists():
        FONT_PATH    = str(_win / "tahoma.ttf")
        FONT_PATH_BD = str(_win / "tahomabd.ttf")
    else:
        _tmp = Path(tempfile.gettempdir())
        FONT_PATH    = str(_tmp / "Sarabun-Regular.ttf")
        FONT_PATH_BD = str(_tmp / "Sarabun-Bold.ttf")
        if not Path(FONT_PATH).exists():
            urllib.request.urlretrieve("https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf", FONT_PATH)
        if not Path(FONT_PATH_BD).exists():
            urllib.request.urlretrieve("https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Bold.ttf", FONT_PATH_BD)
    LOGO_PATH    = str(Path(__file__).parent / "Logo.png")

    class PDF(FPDF):
        def header(self):
            # Logo top-left — h=14 วางที่ y=10
            self.image(LOGO_PATH, x=10, y=10, h=14)
            # Company info เริ่มที่ x=55 ให้พ้น logo
            self.set_xy(55, 10)
            self.set_font("Tahoma", "B", 9)
            self.cell(0, 5, "DHL SUPPLY CHAIN (THAILAND)", new_x="LMARGIN", new_y="NEXT")
            self.set_x(55)
            self.set_font("Tahoma", "", 7.5)
            self.cell(0, 4.5, "NO. 9 G TOWER GRAND RAMA 9 (NORTH WING), 26TH FLOOR AND 27TH FLOOR,", new_x="LMARGIN", new_y="NEXT")
            self.set_x(55)
            self.cell(0, 4.5, "RAMA IX RD, HUAI KHWANG, BANGKOK 10310  TEL. (02) 779 9800", new_x="LMARGIN", new_y="NEXT")
            self.ln(5)
            if self.page_no() > 1:
                self.set_font("Tahoma", "B", 13)
                self.cell(0, 8, "EXPENSE DETAIL", new_x="LMARGIN", new_y="NEXT", align="C")
                self.ln(2)

    pdf = PDF()
    pdf.add_font("Tahoma",  "", FONT_PATH)
    pdf.add_font("Tahoma",  "B", FONT_PATH_BD)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_margins(10, 15, 10)

    # column widths: 28+72+25+7+25+23 = 180mm
    CW = {"inv": 28, "name": 72, "rate": 25, "x": 7, "qty": 25, "amt": 23}
    COL_W  = 35
    LABEL_W = CW["inv"] + CW["name"] + CW["rate"] + CW["x"] + CW["qty"]

    def info_row(label, value, multiline=False):
        pdf.set_font("Tahoma", "B", 9)
        pdf.cell(COL_W, 6, label)
        pdf.set_font("Tahoma", "", 9)
        if multiline:
            pdf.multi_cell(0, 6, str(value or ""), new_x="LMARGIN", new_y="NEXT")
        else:
            pdf.cell(0, 6, str(value or ""), new_x="LMARGIN", new_y="NEXT")

    # Summary row colors matching sample PDF
    COLOR = {
        "amount": (255, 213, 128),   # amber
        "vat":    (226, 239, 218),   # light green
        "total":  (226, 239, 218),   # light green
        "wht":    (255, 255, 153),   # light yellow
        "net":    (189, 215, 238),   # light blue
    }

    def sum_row(label, value, color_key="amount"):
        r, g, b = COLOR[color_key]
        bold = color_key == "net"
        pdf.set_font("Tahoma", "B" if bold else "", 8)
        pdf.set_fill_color(r, g, b)
        pdf.cell(LABEL_W, 6, label, border=1, fill=True, align="R")
        pdf.cell(CW["amt"], 6, f"{value:,.2f}", border=1, fill=True, align="R", new_x="LMARGIN", new_y="NEXT")

    # ── Cover Page ──────────────────────────────────────────
    # total width = 40+28+24+32+20+20+21 = 185mm (fits A4 190mm printable)
    CW_COV = {"part": 40, "inv": 28, "country": 24, "payto": 32, "due": 20, "remark": 20, "amt": 21}
    COV_LINE_H = 5  # height per line in cover table

    def _count_lines(pdf_obj, text, col_w):
        """Count lines needed for text in a given column width."""
        import math
        if not text:
            return 1
        effective_w = col_w - 2
        words = str(text).split()
        lines, line_w = 1, 0.0
        for word in words:
            ww = pdf_obj.get_string_width(word + " ")
            if ww > effective_w:
                # Word wider than column (e.g. long invoice no. with no spaces)
                if line_w > 0:
                    lines += 1
                    line_w = 0
                word_lines = math.ceil(ww / effective_w)
                lines += word_lines - 1
                line_w = ww - (word_lines - 1) * effective_w
            elif line_w + ww > effective_w and line_w > 0:
                lines += 1
                line_w = ww
            else:
                line_w += ww
        return lines

    def _draw_cover_row(pdf_obj, cw, vals, aligns, fill=False, bold=False, fill_color=None):
        """Draw one row with auto row-height and word-wrap for part/payto/remark."""
        WRAP_KEYS = {"part", "inv", "payto", "remark"}
        if fill_color:
            pdf_obj.set_fill_color(*fill_color)
        pdf_obj.set_font("Tahoma", "B" if bold else "", 7.5)

        # Calculate row height
        max_lines = 1
        for key in WRAP_KEYS:
            if key in cw:
                max_lines = max(max_lines, _count_lines(pdf_obj, vals.get(key, ""), cw[key]))
        row_h = max_lines * COV_LINE_H

        x0, y0 = pdf_obj.get_x(), pdf_obj.get_y()
        x = x0
        for key, w in cw.items():
            text = str(vals.get(key) or "")
            align = aligns.get(key, "L")
            pdf_obj.set_xy(x, y0)
            if key in WRAP_KEYS:
                # วาด background fill ก่อน (ถ้ามี)
                if fill and fill_color:
                    pdf_obj.set_fill_color(*fill_color)
                    pdf_obj.rect(x, y0, w, row_h, style="F")
                # วาด border รอบ cell ด้วย rect
                pdf_obj.rect(x, y0, w, row_h)
                # วาด text ด้วย multi_cell โดยไม่มี border (center แนวตั้งเหมือน cell())
                n_lines = _count_lines(pdf_obj, text, w)
                v_offset = max(0, (row_h - n_lines * COV_LINE_H) / 2)
                pdf_obj.set_xy(x + 1, y0 + v_offset)
                pdf_obj.multi_cell(w - 2, COV_LINE_H, text, border=0, align=align,
                                   fill=False, new_x="RIGHT", new_y="TOP")
            else:
                pdf_obj.cell(w, row_h, text, border=1, align=align, fill=fill)
            x += w
        pdf_obj.set_xy(x0, y0 + row_h)

    COV_ALIGNS = {"part": "L", "inv": "C", "country": "C", "payto": "L", "due": "C", "remark": "L", "amt": "R"}

    pdf.add_page()
    pdf.set_font("Tahoma", "B", 14)
    pdf.cell(0, 8, "ใบแจ้งค่าใช้จ่ายส่งออก", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Tahoma", "B", 12)
    pdf.cell(0, 7, "EXPORT EXPENSE DETAIL", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(6)

    today_cover = datetime.now(pytz.timezone("Asia/Bangkok")).strftime("%d/%m/%Y")
    pdf.set_font("Tahoma", "", 9)
    pdf.cell(0, 6, f"Date : {today_cover}", new_x="LMARGIN", new_y="NEXT", align="R")
    pdf.ln(4)

    # Cover table header
    pdf.set_fill_color(220, 220, 220)
    pdf.set_font("Tahoma", "B", 7.5)
    hdr_vals = {"part": "รายการ / PARTICULARS", "inv": "INVOICE NO.", "country": "Country",
                "payto": "Pay To", "due": "Due Date", "remark": "Remark", "amt": "Amount"}
    _draw_cover_row(pdf, CW_COV, hdr_vals, {k: "C" for k in CW_COV}, fill=True, bold=True, fill_color=(220, 220, 220))

    cover_total = 0.0
    for rec in records:
        hdr  = rec["header"]
        bk   = rec.get("bk") or {}
        its  = rec["items"]
        subtotal = sum(float(it.get("total") or 0) for it in its)
        vat_7 = float(hdr.get("vat_7") or 0)
        wht_1 = float(hdr.get("wht_1") or 0)
        wht_3 = float(hdr.get("wht_3") or 0)
        net   = round(subtotal + vat_7 - wht_1 - wht_3, 2)
        cover_total += net

        cats = sorted(set(it.get("category") or "other" for it in its))
        default_part = "+ ".join(c.upper().replace("_", "") for c in cats)
        cat_label = rec.get("cover_part") or default_part

        row_vals = {
            "part":    cat_label,
            "inv":     rec.get("cover_inv")     or hdr.get("ctc_invoice_no") or "-",
            "country": rec.get("cover_country") or bk.get("country") or "-",
            "payto":   rec.get("cover_payto")   or hdr.get("pay_to") or "-",
            "due":     rec.get("cover_due")     or hdr.get("due_date") or "-",
            "remark":  rec.get("cover_remark")  or hdr.get("remark") or "",
            "amt":     f"{net:,.2f}",
        }
        _draw_cover_row(pdf, CW_COV, row_vals, COV_ALIGNS)

    # Total row
    total_label_w = sum(CW_COV[k] for k in ["part", "inv", "country", "payto", "due", "remark"])
    pdf.set_font("Tahoma", "B", 8)
    pdf.set_fill_color(189, 215, 238)
    pdf.cell(total_label_w, COV_LINE_H, "Total", border=1, fill=True, align="R")
    pdf.cell(CW_COV["amt"], COV_LINE_H, f"{cover_total:,.2f}", border=1, fill=True, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(6)
    pdf.set_font("Tahoma", "", 9)
    pdf.ln(10)
    name_line = f"Prepared by : {prepared_by}" if prepared_by else "Prepared by ............................................"
    pdf.cell(80, 6, name_line)
    pdf.cell(80, 6, "Approved by ............................................", new_x="LMARGIN", new_y="NEXT")
    if prepared_by_phone:
        pdf.ln(2)
        pdf.set_font("Tahoma", "", 8)
        pdf.cell(80, 5, f"Tel. : {prepared_by_phone}")

    # ── Detail pages ────────────────────────────────────────
    for rec in records:
        hdr   = rec["header"]
        items = rec["items"]

        # ปรับ column widths ตามจำนวน invoice no.
        # ถ้า > 10 ตัว → ขยาย inv column เพื่อให้ wrap 2 ตัว/บรรทัด, ลด description ลง
        _inv_no_tmp = hdr.get("ctc_invoice_no") or ""
        _inv_count = _inv_no_tmp.count("+") + 1 if _inv_no_tmp else 1
        if _inv_count > 10:
            CW["inv"]  = 50
            CW["name"] = 50
        else:
            CW["inv"]  = 28
            CW["name"] = 72
        LABEL_W = CW["inv"] + CW["name"] + CW["rate"] + CW["x"] + CW["qty"]

        pdf.add_page()

        # ── Header info ──
        today_str = datetime.now(pytz.timezone("Asia/Bangkok")).strftime("%d/%m/%Y")

        # DATE: วางที่ขวาบนแบบ absolute (y=10 ระดับเดียวกับ logo/company info)
        y_after_header = pdf.get_y()
        pdf.set_xy(142, 10)
        pdf.set_font("Tahoma", "B", 9)
        pdf.cell(28, 5, "DATE :", align="R")
        pdf.set_font("Tahoma", "", 9)
        pdf.cell(30, 5, today_str, align="R")
        pdf.set_xy(10, y_after_header)

        # SHIPPER row (ไม่มี DATE แล้ว)
        pdf.set_font("Tahoma", "B", 9)
        pdf.cell(COL_W, 6, "SHIPPER :")
        pdf.set_font("Tahoma", "", 9)
        pdf.cell(0, 6, "CARRIER AIR CONDITIONING (THAILAND) CO.,LTD.", new_x="LMARGIN", new_y="NEXT")
        info_row("PAY TO :", hdr.get("pay_to") or "")
        info_row("TAX NAME :", hdr.get("tax_name") or "", multiline=True)
        pdf.ln(2)
        info_row("TAX ID NO. :", hdr.get("tax_id") or "")
        info_row("DELIVERY PORT :", hdr.get("delivery_port") or "")
        info_row("ETD :", hdr.get("etd") or "")
        info_row("BL NO :", hdr.get("bl_no") or "")
        info_row("BOOKING NO. :", hdr.get("booking_no") or "")
        pdf.ln(3)

        # ── Table header ──
        pdf.set_fill_color(220, 220, 220)
        pdf.set_font("Tahoma", "B", 8)
        pdf.cell(CW["inv"],  7, "INVOICE NO.", border=1, fill=True, align="C")
        pdf.cell(CW["name"], 7, "DESCRIPTION", border=1, fill=True, align="C")
        pdf.cell(CW["rate"], 7, "RATE",        border=1, fill=True, align="C")
        pdf.cell(CW["x"],    7, "x",           border=1, fill=True, align="C")
        pdf.cell(CW["qty"],  7, "QUANTITY",    border=1, fill=True, align="C")
        pdf.cell(CW["amt"],  7, "AMOUNT",      border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")

        # ── Table rows — invoice_no merged cell spanning all items ──
        invoice_no = hdr.get("ctc_invoice_no") or ""
        subtotal = 0.0
        pdf.set_font("Tahoma", "", 8)
        _CAT_ORDER = [
            "thc_40hc","thc_40dv","thc_20gp","export_handling","seal","bl_fee",
            "surrender_fee","vgm_fee","doc_amendment","detention","demurrage",
            "container_repair","edi_fee","late_gate","environmental_fee",
            "storage","freight_charge","other",
        ]
        items = sorted(items, key=lambda x: _CAT_ORDER.index(x.get("category") or "other")
                       if (x.get("category") or "other") in _CAT_ORDER else len(_CAT_ORDER))

        ROW_H   = 6
        x0, y0  = pdf.get_x(), pdf.get_y()

        inv_display = invoice_no or ""
        pdf.set_font("Tahoma", "", 8)
        inv_lines = _count_lines(pdf, inv_display, CW["inv"]) if inv_display else 1

        # พื้นที่เหลือบนหน้า เผื่อ summary rows (~50mm)
        page_break_y = pdf.h - pdf.b_margin
        available_h  = max(50, page_break_y - y0 - 50)

        # คำนวณ row heights — กันทั้ง inv และ items ให้ fit ใน available_h
        n_items = len(items) or 1
        max_lines = max(inv_lines, n_items)

        # font 8pt → minimum line height ≈ 3.5mm (cap + descender) เพื่อกัน text overflow
        MIN_INV_LINE_H = 3.5
        if max_lines * ROW_H > available_h:
            scaled_row_h = max(MIN_INV_LINE_H, available_h / max_lines)
        else:
            scaled_row_h = ROW_H

        if inv_lines > n_items:
            inv_row_h  = scaled_row_h
            item_row_h = (inv_lines * scaled_row_h) / n_items
        else:
            inv_row_h  = scaled_row_h
            item_row_h = scaled_row_h
        total_h     = n_items * item_row_h
        inv_h_total = inv_lines * inv_row_h

        # ปิด auto_page_break ชั่วคราว — ใช้ try/finally ป้องกันลืม restore
        pdf.set_auto_page_break(auto=False)
        try:
            pdf.rect(x0, y0, CW["inv"], total_h)
            v_offset = max(0, (total_h - inv_h_total) / 2)
            pdf.set_xy(x0 + 1, y0 + v_offset)
            pdf.multi_cell(CW["inv"] - 2, inv_row_h, inv_display, border=0, align="C",
                           new_x="RIGHT", new_y="TOP")

            for i, it in enumerate(items):
                desc      = it.get("description") or ""
                rate      = float(it.get("rate") or 0)
                qty       = float(it.get("qty") or 0)
                total     = float(it.get("total") or 0)
                subtotal += total

                rate_str  = f"{rate:,.2f}" if rate else ""
                qty_str   = f"{qty:,.3f}" if qty else ""
                total_str = f"{total:,.2f}" if total else "-"

                pdf.set_xy(x0 + CW["inv"], y0 + i * item_row_h)
                pdf.cell(CW["name"], item_row_h, desc[:45],  border=1)
                pdf.cell(CW["rate"], item_row_h, rate_str,   border=1, align="R")
                pdf.cell(CW["x"],    item_row_h, "x",        border=1, align="C")
                pdf.cell(CW["qty"],  item_row_h, qty_str,    border=1, align="R")
                pdf.cell(CW["amt"],  item_row_h, total_str,  border=1, align="R")
        finally:
            pdf.set_auto_page_break(auto=True, margin=15)

        pdf.set_xy(pdf.l_margin, y0 + total_h)

        # ── Summary rows ──
        vat_7     = float(hdr.get("vat_7")  or 0)
        wht_1     = float(hdr.get("wht_1")  or 0)
        wht_3     = float(hdr.get("wht_3")  or 0)
        total_net = subtotal + vat_7 - wht_1 - wht_3

        sum_row("Amount",              subtotal,          color_key="amount")
        sum_row("บวก vat 7%",          vat_7,             color_key="vat")
        sum_row("Total",               subtotal + vat_7,  color_key="total")
        sum_row("หัก ภาษี ณ ที่จ่าย 1%", wht_1,          color_key="wht")
        sum_row("หัก ภาษี ณ ที่จ่าย 3%", wht_3,          color_key="wht")
        sum_row("ยอดจ่ายจริง",         total_net,         color_key="net")

    summary_bytes = bytes(pdf.output())

    # ── Merge invoice PDFs ก่อน summary ของแต่ละ invoice ──
    try:
        from pypdf import PdfWriter, PdfReader
        import io as _io

        writer = PdfWriter()

        # page 0 = cover page, detail pages เริ่มที่ index 1
        sum_reader = PdfReader(_io.BytesIO(summary_bytes))
        writer.add_page(sum_reader.pages[0])  # cover page

        for i, rec in enumerate(records):
            # detail page อยู่ที่ index i+1 (เพราะ page 0 = cover)
            writer.add_page(sum_reader.pages[i + 1])
            # แทรก invoice PDF ต่อท้าย (ถ้ามี)
            inv_path = rec["header"].get("invoice_pdf_path")
            if inv_path:
                try:
                    inv_bytes = supabase.storage.from_("local-charge-invoices").download(inv_path)
                    inv_reader = PdfReader(_io.BytesIO(inv_bytes))
                    for page in inv_reader.pages:
                        writer.add_page(page)
                except Exception:
                    pass  # ถ้าดาวน์โหลดไม่ได้ ข้ามไป

        out = _io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception:
        # fallback: return summary only
        return summary_bytes


# ─────────────────────────────────────────
# 5d. RUN NO. ITEMS — PDF NUMBER OVERLAY
# ─────────────────────────────────────────
def number_items_in_pdf(pdf_bytes: bytes) -> bytes:
    """Add sequential numbers (1,2,3...) to item lines in Invoice/Packing List PDF.
    Invoice items numbered 1-N, PL items also 1-N (matching invoice).
    Handles page-break case where P/O is on page N and item code is on page N+1."""
    import fitz  # PyMuPDF
    import re as _re
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Classify each page as invoice or pl (inherit from previous if ambiguous)
    page_sections = []
    current = "invoice"
    for _pg in doc:
        _txt = _pg.get_text()
        if "PACKING LIST" in _txt:
            current = "pl"
        elif "*** INVOICE ***" in _txt:
            current = "invoice"
        page_sections.append(current)

    # Collect all lines across all pages in document order (page → y → x)
    ordered_lines = []
    for _pn, _page in enumerate(doc):
        _pdict = _page.get_text("dict")
        _plines = []
        for _block in _pdict.get("blocks", []):
            for _line in _block.get("lines", []):
                _spans = _line.get("spans", [])
                if not _spans:
                    continue
                _ltext = "".join(s["text"] for s in _spans).strip()
                if not _ltext:
                    continue
                _plines.append({"page": _pn, "text": _ltext, "bbox": _line["bbox"]})
        _plines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
        ordered_lines.extend(_plines)

    _ITEM_RE = _re.compile(r"^[A-Z][A-Z0-9]*-")
    counters = {"invoice": 0, "pl": 0}

    for _i, _ln in enumerate(ordered_lines):
        if "P/O NO" not in _ln["text"]:
            continue
        # Scan forward up to 100 lines to find the item code line
        # (large window handles cross-page case: P/O at end of page, item code on next page after footer+header)
        # Safe because we break at next P/O — won't false-match items belonging to a later P/O.
        _found = None
        for _j in range(_i + 1, min(_i + 101, len(ordered_lines))):
            _cand = ordered_lines[_j]
            _ctext = _cand["text"]
            if not _ctext:
                continue
            if "P/O NO" in _ctext:
                break  # next P/O — abort, item missing
            if _ITEM_RE.match(_ctext):
                _found = _cand
                break
        if _found is None:
            continue
        _section = page_sections[_found["page"]]
        counters[_section] += 1
        _n = counters[_section]
        _x0, _y0, _x1, _y1 = _found["bbox"]
        _mx = max(15.0, _x0 - 35.0)
        doc[_found["page"]].insert_text(
            (_mx, _y1 - 1.0),
            f"{_n}.",
            fontsize=9,
            color=(0.83, 0.02, 0.07),
        )

    _out = doc.tobytes()
    doc.close()
    return _out


# ─────────────────────────────────────────
# 6. SIDEBAR NAVIGATION
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🚚 เมนู")
    page = st.radio(
        "เลือกเมนู",
        ["📤 Upload & Extract", "📄 Generate SI (Draft)", "💰 Local Charges", "📊 Export Summary", "📝 Run No. Items"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "<span style='font-size:11px;color:#555;'>Powered by Ship Co. (CTC Site)</span>",
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
    from concurrent.futures import ThreadPoolExecutor, as_completed

    list_icd   = files_icd   or []
    list_alpha = files_alpha or []
    total      = len(list_icd) + len(list_alpha)

    if total > 0:
        all_data     = []
        progress_bar = st.progress(0)
        status_text  = st.empty()
        processed    = 0

        # อ่าน bytes ทั้งหมดใน main thread ก่อน (UploadedFile ไม่ thread-safe)
        tasks = (
            [(f.name, f.read(), "ICD")   for f in list_icd] +
            [(f.name, f.read(), "ALPHA") for f in list_alpha]
        )

        def process_file(name, file_bytes, warehouse):
            doc_type = detect_doc_type_by_ai(file_bytes)
            items = extract_air_awb(file_bytes) if doc_type == "air" else extract_from_pdf(file_bytes)
            if items:
                for item in items:
                    item["source_file"] = name
                    item["loading_at"]  = warehouse
            return items or []

        with ThreadPoolExecutor(max_workers=min(total, 5)) as executor:
            futures = {
                executor.submit(process_file, name, fb, wh): name
                for name, fb, wh in tasks
            }
            for future in as_completed(futures):
                name = futures[future]
                try:
                    items = future.result()
                    all_data.extend(items)
                except Exception as e:
                    st.error(f"❌ {name}: {e}")
                processed += 1
                progress_bar.progress(processed / total)
                status_text.caption(f"สแกนแล้ว {processed}/{total} ไฟล์ — {name}")

        status_text.empty()

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
                        "container_type":      container_type or None,
                        "no_pallet":           no_pallet or None,
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

            # ── Delete Booking ─────────────────────────────────────
            st.divider()
            st.subheader("🗑️ ลบ Booking")

            del_bk = st.selectbox("เลือก Booking ที่ต้องการลบ",
                                  ["-- เลือก --"] + all_bk_nos, key="del_bk_select")

            if del_bk != "-- เลือก --":
                st.warning(f"⚠️ จะลบ **{del_bk}** ออกจากฐานข้อมูลถาวร")
                confirm = st.checkbox("ยืนยันการลบ", key="del_confirm")
                if confirm:
                    if st.button("🗑️ ลบเลย", type="primary", key="del_btn"):
                        try:
                            supabase.table(TBL_BOOKINGS).delete().eq("booking_no", del_bk).execute()
                            st.success(f"✅ ลบ {del_bk} เรียบร้อยแล้ว")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Delete Error: {e}")

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
        safe_write("B35", s(booking.get("container_type")) or "")
 
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
        MAX_INV    = 100
 
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
        for i, mark in enumerate(marks[:MAX_INV]):
            if i > 0:
                mark_row += 1  # blank row between invoices
            for line in str(mark).split("\n"):
                line = line.strip()
                if line:
                    safe_write_rc(mark_row, 1, line)
                    mark_row += 1
 
        def _parse_pkg(val):
            m = re.match(r'^([\d,]+)\s+(.*)', str(val or '').strip())
            if m:
                try:
                    num  = int(m.group(1).replace(',', ''))
                    unit = m.group(2).strip()
                    if unit.upper().startswith("PP."):
                        unit = unit[3:]
                    return num, unit
                except Exception:
                    pass
            return 0, "CARTONS"

        first_pkg_unit = _parse_pkg(invoices[0].get("cartons"))[1] if invoices else "CARTONS"

        gw_cells, cbm_cells, ctn_cells = [], [], []
 
        for idx, inv in enumerate(invoices):
            base = MARK_START + idx * ROWS_PER
            for offset in range(ROWS_PER):
                copy_row_style(38 + offset, base + offset)
 
            # col B = "CARTONS" เฉพาะ invoice แรกเท่านั้น
            if idx == 0:
                safe_write_rc(base, 2, first_pkg_unit)
 
            # row+0: qty — ใส่สีเฉพาะ C (จำนวน + unit เช่น "1,389 CARTONS" หรือ "40 PP.PALLETS")
            qty_num, pkg_unit = _parse_pkg(inv.get("cartons"))
            safe_write_rc(base, 3, qty_num or 0)
            safe_write_rc(base, 4, pkg_unit)
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
            # B37 = SUM ตัวเลข C ของทุก invoice
            safe_write("B37", "=" + "+".join(ctn_cells))
 
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
        bl_type       = st.selectbox("BL Type", ["Sea Waybill","Original B/L","Surrender B/L","Telex Release"])
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

# ─────────────────────────────────────────
# 9. PAGE: LOCAL CHARGES
# ─────────────────────────────────────────
if page == "💰 Local Charges":

    st.subheader("💰 Local Charges")

    st.markdown("""
    <style>
    button[data-testid="stNumberInputStepUp"],
    button[data-testid="stNumberInputStepDown"] { display: none; }
    </style>
    """, unsafe_allow_html=True)

    # ── Booking No. dropdown ──────────────
    try:
        bk_res = supabase.table(TBL_BOOKINGS).select("booking_no").order("updated_at", desc=True).execute()
        bk_options = [r["booking_no"] for r in bk_res.data if r.get("booking_no")]
    except Exception:
        bk_options = []

    selected_booking_no = st.selectbox(
        "Booking No. *",
        options=["— เลือก Booking —"] + bk_options,
    )
    booking_selected = selected_booking_no != "— เลือก Booking —"
    if not booking_selected:
        st.warning("⚠️ กรุณาเลือก Booking No. ก่อนอัปโหลดไฟล์")

    # ── Fetch shipment type for DSV WHT rule ──
    shipment_type = ""
    if booking_selected:
        try:
            bk_detail = supabase.table(TBL_BOOKINGS).select("by_air_or_sea, fcl_or_lcl").eq("booking_no", selected_booking_no).limit(1).execute()
            if bk_detail.data:
                b = bk_detail.data[0]
                air_sea = (b.get("by_air_or_sea") or "").strip().lower()
                fcl_lcl = (b.get("fcl_or_lcl") or "").strip().lower()
                if air_sea == "air":
                    shipment_type = "Air Export"
                elif fcl_lcl == "lcl":
                    shipment_type = "Ocean Export"
                else:
                    shipment_type = "Ocean Export"
        except Exception:
            pass

    # ── Upload zone ──────────────────────
    if "lc_uploader_key" not in st.session_state:
        st.session_state["lc_uploader_key"] = 0

    lc_file = st.file_uploader(
        "โยนไฟล์ Local Charge PDF ที่นี่",
        type="pdf",
        accept_multiple_files=False,
        key=f"lc_{st.session_state['lc_uploader_key']}",
        disabled=not booking_selected,
    )

    if lc_file and booking_selected:
        cache_key = f"lc_data_{lc_file.name}_{lc_file.size}"
        if cache_key not in st.session_state:
            with st.spinner(f"กำลังสแกน: {lc_file.name}"):
                try:
                    st.session_state[cache_key] = extract_local_charges(lc_file.read(), shipment_type=shipment_type)
                except Exception as e:
                    st.error(f"❌ AI Error: {e}")
                    st.session_state[cache_key] = None

        data = st.session_state[cache_key]

        if data:
            if data.get("_multi_invoice"):
                st.warning("⚠️ ตรวจพบหลาย invoice ในไฟล์เดียวกัน — แสดงเฉพาะ invoice แรก กรุณาอัพโหลดทีละ invoice")
            st.success("✅ Extract สำเร็จ — ตรวจสอบข้อมูลก่อนบันทึก")

            # ── Header fields ─────────────────────────────────────
            st.markdown("**ข้อมูลทั่วไป**")
            r1c1, r1c2, r1c3 = st.columns(3)
            hdr_agent_invoice_no = r1c1.text_input("Agent Invoice No.", value=str(data.get("agent_invoice_no") or ""), key="lc_agent_invoice_no")
            hdr_pay_to           = r1c2.text_input("Pay To",            value=str(data.get("pay_to")        or ""), key="lc_pay_to")
            hdr_tax_id           = r1c3.text_input("Tax ID No.",        value=str(data.get("tax_id")        or ""), key="lc_tax_id")

            r2c1, r2c2, r2c3 = st.columns(3)
            hdr_tax_name      = r2c1.text_area("Tax Name & Address", value=str(data.get("tax_name")      or ""), height=80, key="lc_tax_name")
            hdr_delivery_port = r2c2.text_input("Delivery Port",     value=str(data.get("delivery_port") or ""), key="lc_delivery_port")
            hdr_etd           = r2c3.text_input("ETD (DD/MM/YYYY)",  value=str(data.get("etd")           or ""), key="lc_etd")

            r3c1, r3c2, r3c3 = st.columns(3)
            hdr_bl_no         = r3c1.text_input("B/L No.",         value=str(data.get("bl_no") or ""), key="lc_bl_no")
            hdr_ctc_invoice_no = r3c2.text_input("CTC Invoice No.", value="", key="lc_ctc_invoice_no")
            hdr_remark        = r3c3.text_input("Remark",           value="", key="lc_remark")

            r4c1, r4c2, r4c3 = st.columns(3)
            hdr_due_date      = r4c1.text_input("Due Date (DD/MM/YYYY)", value=str(data.get("due_date") or ""), key="lc_due_date")

            # ── Dynamic items ─────────────────────────────────────
            st.markdown("**ค่าใช้จ่าย (บาท)**")

            CATEGORY_ORDER = [
                "thc_40hc", "thc_40dv", "thc_20gp",
                "export_handling", "seal", "bl_fee", "surrender_fee", "vgm_fee",
                "doc_amendment", "detention", "demurrage", "container_repair",
                "edi_fee", "late_gate", "environmental_fee", "storage", "freight_charge", "other",
            ]
            CATEGORY_LABEL = {
                "thc_40hc":          "THC (40HC)",
                "thc_40dv":          "THC (40DV)",
                "thc_20gp":          "THC (20GP)",
                "export_handling":   "Export Handling",
                "seal":              "Seal",
                "bl_fee":            "B/L Fee",
                "surrender_fee":     "Surrender Fee",
                "vgm_fee":           "VGM Coordination Fee",
                "doc_amendment":     "Documentation Amendment Charge",
                "detention":         "Detention",
                "demurrage":         "Demurrage",
                "container_repair":  "Container Repair",
                "edi_fee":           "EDI Fee",
                "late_gate":         "Late Gate Service",
                "environmental_fee": "Environmental Fee",
                "storage":           "Storage",
                "freight_charge":    "Freight Charge",
                "other":             None,  # keep original description
            }

            if "lc_items" not in st.session_state or st.session_state.get("lc_items_source") != cache_key:
                raw_items = [
                    {"description": it.get("description", ""), "category": it.get("category") or "other",
                     "wht_pct": int(it.get("wht_pct") or 0),
                     "rate": float(it.get("rate") or 0), "qty": float(it.get("qty") or 0),
                     "total": float(it.get("total") or 0)}
                    for it in (data.get("items") or [])
                ]
                # Replace description with standard label (keep original for "other")
                for it in raw_items:
                    cat = it["category"]
                    label = CATEGORY_LABEL.get(cat)
                    if label:
                        it["description"] = label
                # Sort by category order
                raw_items.sort(key=lambda x: CATEGORY_ORDER.index(x["category"]) if x["category"] in CATEGORY_ORDER else 99)
                st.session_state["lc_items"] = raw_items
                st.session_state["lc_items_source"] = cache_key

            hc1, hc2, hc3, hc4, hc5, hc6 = st.columns([3, 1, 2, 2, 2, 1])
            hc1.markdown("**รายการ**"); hc2.markdown("**WHT %**")
            hc3.markdown("**Rate**"); hc4.markdown("**No.Unit**")
            hc5.markdown("**Total**"); hc6.markdown("")

            items_to_delete = []
            for idx, item in enumerate(st.session_state["lc_items"]):
                c1, c2, c3, c4, c5, c6 = st.columns([3, 1, 2, 2, 2, 1])
                item["description"] = c1.text_input("_", value=item["description"], label_visibility="collapsed", key=f"lc_desc_{idx}")
                item["wht_pct"]     = int(c2.number_input("_", value=int(item["wht_pct"]), min_value=0, max_value=3, step=1, label_visibility="collapsed", key=f"lc_wht_{idx}"))
                item["rate"]        = c3.number_input("_", value=item["rate"], min_value=0.0, step=0.01, format="%.2f", label_visibility="collapsed", key=f"lc_rate_{idx}")
                item["qty"]         = c4.number_input("_", value=item["qty"],  min_value=0.0, step=0.01, format="%.2f", label_visibility="collapsed", key=f"lc_qty_{idx}")
                item["total"]       = c5.number_input("_", value=item["total"], min_value=0.0, step=0.01, format="%.2f", label_visibility="collapsed", key=f"lc_total_{idx}")
                if c6.button("🗑️", key=f"lc_del_{idx}"):
                    items_to_delete.append(idx)

            for idx in reversed(items_to_delete):
                st.session_state["lc_items"].pop(idx)
                st.rerun()

            if st.button("➕ เพิ่มรายการ", key="lc_add"):
                st.session_state["lc_items"].append({"description": "", "category": "other", "wht_pct": 0, "rate": 0.0, "qty": 0.0, "total": 0.0})
                st.rerun()

            # ── Live summary ──────────────────────────────────────
            current_items = st.session_state["lc_items"]
            charges_subtotal = sum(float(it.get("total") or 0) for it in current_items)
            wht1_sum = sum(float(it.get("total") or 0) for it in current_items if int(it.get("wht_pct") or 0) == 1)
            wht3_sum = sum(float(it.get("total") or 0) for it in current_items if int(it.get("wht_pct") or 0) == 3)
            calc_wht1 = round(wht1_sum * 0.01, 2)
            calc_wht3 = round(wht3_sum * 0.03, 2)

            st.divider()
            sub1, _, _, _, sub5, _ = st.columns([3, 1, 2, 2, 2, 1])
            sub1.markdown("<div style='padding-top:8px'>**รวมค่าใช้จ่าย**</div>", unsafe_allow_html=True)
            sub5.markdown(f"<div style='padding-top:8px; text-align:right'><b>{charges_subtotal:,.2f}</b></div>", unsafe_allow_html=True)

            st.markdown("**สรุป**")
            sc1, _, _, _, sc5, _ = st.columns([3, 1, 2, 2, 2, 1])
            sc1.markdown("<div style='padding-top:8px'>VAT 7%</div>", unsafe_allow_html=True)
            # Dynamic key: เปลี่ยนตาม items hash → reset เป็น _auto_vat อัตโนมัติเมื่อ items เปลี่ยน
            _auto_vat = round(charges_subtotal * 0.07, 2) if data.get("vat_applicable") else 0.0
            _items_hash = hash(tuple(round(float(it.get("total") or 0), 2) for it in current_items))
            vat_7 = sc5.number_input("_", value=_auto_vat, min_value=0.0, step=0.01, format="%.2f", label_visibility="collapsed", key=f"vat_7_{_items_hash}")

            after_vat = charges_subtotal + vat_7
            av1, _, _, _, av5, _ = st.columns([3, 1, 2, 2, 2, 1])
            av1.markdown("<div style='padding-top:8px'>**รวมหลัง VAT**</div>", unsafe_allow_html=True)
            av5.markdown(f"<div style='padding-top:8px; text-align:right'><b>{after_vat:,.2f}</b></div>", unsafe_allow_html=True)

            wc1, _, _, _, wc5, _ = st.columns([3, 1, 2, 2, 2, 1])
            wc1.markdown("<div style='padding-top:8px'>WHT 1%</div>", unsafe_allow_html=True)
            wc5.markdown(f"<div style='padding-top:8px; text-align:right'>{calc_wht1:,.2f}</div>", unsafe_allow_html=True)

            w3c1, _, _, _, w3c5, _ = st.columns([3, 1, 2, 2, 2, 1])
            w3c1.markdown("<div style='padding-top:8px'>WHT 3%</div>", unsafe_allow_html=True)
            w3c5.markdown(f"<div style='padding-top:8px; text-align:right'>{calc_wht3:,.2f}</div>", unsafe_allow_html=True)

            after_wht = round(after_vat - calc_wht1 - calc_wht3, 2)
            aw1, _, _, _, aw5, _ = st.columns([3, 1, 2, 2, 2, 1])
            aw1.markdown("<div style='padding-top:8px'>**รวมหลังหัก WHT**</div>", unsafe_allow_html=True)
            aw5.markdown(f"<div style='padding-top:8px; text-align:right'><b>{after_wht:,.2f}</b></div>", unsafe_allow_html=True)

            # ── Save ──────────────────────────────────────────────
            if not hdr_due_date.strip():
                st.warning("⚠️ กรุณากรอก Due Date ก่อนบันทึก")
            if st.button("💾 บันทึก", use_container_width=True, key="lc_save", disabled=not hdr_due_date.strip()):
                header = {
                    "agent_invoice_no": hdr_agent_invoice_no or None,
                    "pay_to":        hdr_pay_to or None,
                    "tax_name":      hdr_tax_name or None,
                    "tax_id":        hdr_tax_id or None,
                    "delivery_port": hdr_delivery_port or None,
                    "etd":           hdr_etd or None,
                    "bl_no":         hdr_bl_no or None,
                    "due_date":      hdr_due_date or None,
                    "vat_7":         vat_7 if vat_7 else None,
                    "wht_1":         calc_wht1 if calc_wht1 else None,
                    "wht_3":         calc_wht3 if calc_wht3 else None,
                    "subtotal":      charges_subtotal if charges_subtotal else None,
                    "total":         after_wht if after_wht else None,
                    "source_file":   lc_file.name,
                    "booking_no":    selected_booking_no if selected_booking_no != "— เลือก Booking —" else None,
                    "ctc_invoice_no": hdr_ctc_invoice_no or None,
                    "remark":        hdr_remark or None,
                }
                save_items = [
                    {"description": it["description"], "category": it.get("category") or "other",
                     "wht_pct": it["wht_pct"],
                     "rate": it["rate"] or None, "qty": it["qty"] or None,
                     "total": it["total"] or None}
                    for it in current_items if it.get("description")
                ]
                if save_local_charge_v2(header, save_items, pdf_bytes=lc_file.getvalue(), filename=lc_file.name):
                    st.success("✅ บันทึกเรียบร้อยแล้ว")
                    del st.session_state["lc_items"]
                    del st.session_state["lc_items_source"]
                    st.session_state["lc_uploader_key"] += 1
                    st.rerun()

    # ── History table ────────────────────
    st.divider()
    st.subheader("📊 รายการ Local Charges ทั้งหมด")
    try:
        lc_res = (
            supabase.table(TBL_LOCAL_CHARGES_V2)
            .select("*")
            .order("created_at", desc=True)
            .execute()
        )
        if lc_res.data:
            df_lc = pd.DataFrame(lc_res.data)
            df_lc = bkk_time(df_lc, "created_at")
            st.dataframe(df_lc, use_container_width=True)

            # ── Edit รายการ ───────────────────────────────────────
            st.markdown("**✏️ แก้ไขรายการ**")
            _edit_options = {
                f"{r.get('agent_invoice_no') or r.get('ctc_invoice_no') or '-'} | {r.get('booking_no') or '-'} | {r.get('pay_to') or '-'} | {r.get('etd') or '-'}": r["id"]
                for r in lc_res.data
            }
            _edit_label = st.selectbox(
                "เลือกรายการที่ต้องการแก้ไข",
                options=["— เลือก —"] + list(_edit_options.keys()),
                key="lc_edit_select",
            )
            if _edit_label != "— เลือก —":
                _edit_id = _edit_options[_edit_label]
                _edit_row = next((r for r in lc_res.data if r["id"] == _edit_id), {})
                # Dynamic key suffix: เปลี่ยนตาม _edit_id → widget reset เมื่อสลับ record
                _ek = str(_edit_id)
                with st.form(f"lc_edit_form_{_ek}"):
                    _ec1, _ec2, _ec3 = st.columns(3)
                    _e_agent_inv = _ec1.text_input("Agent Invoice No.", value=str(_edit_row.get("agent_invoice_no") or ""), key=f"lc_edit_agent_inv_{_ek}")
                    _e_ctc_inv   = _ec2.text_input("CTC Invoice No.",   value=str(_edit_row.get("ctc_invoice_no") or ""),   key=f"lc_edit_ctc_inv_{_ek}")
                    _e_booking   = _ec3.text_input("Booking No.",       value=str(_edit_row.get("booking_no") or ""),       key=f"lc_edit_booking_{_ek}")

                    _ec4, _ec5, _ec6 = st.columns(3)
                    _e_pay_to    = _ec4.text_input("Pay To",     value=str(_edit_row.get("pay_to") or ""),    key=f"lc_edit_pay_to_{_ek}")
                    _e_tax_id    = _ec5.text_input("Tax ID No.", value=str(_edit_row.get("tax_id") or ""),    key=f"lc_edit_tax_id_{_ek}")
                    _e_bl_no     = _ec6.text_input("B/L No.",    value=str(_edit_row.get("bl_no") or ""),     key=f"lc_edit_bl_no_{_ek}")

                    _e_tax_name  = st.text_area("Tax Name & Address", value=str(_edit_row.get("tax_name") or ""), height=80, key=f"lc_edit_tax_name_{_ek}")

                    _ec7, _ec8, _ec9 = st.columns(3)
                    _e_delivery  = _ec7.text_input("Delivery Port",       value=str(_edit_row.get("delivery_port") or ""), key=f"lc_edit_delivery_{_ek}")
                    _e_etd       = _ec8.text_input("ETD (DD/MM/YYYY)",    value=str(_edit_row.get("etd") or ""),           key=f"lc_edit_etd_{_ek}")
                    _e_due       = _ec9.text_input("Due Date (DD/MM/YYYY)", value=str(_edit_row.get("due_date") or ""),    key=f"lc_edit_due_{_ek}")

                    _e_remark    = st.text_input("Remark", value=str(_edit_row.get("remark") or ""), key=f"lc_edit_remark_{_ek}")

                    _edit_submit = st.form_submit_button("💾 บันทึกการแก้ไข", use_container_width=True)

                if _edit_submit:
                    _upd_payload = {
                        "agent_invoice_no": _e_agent_inv or None,
                        "ctc_invoice_no":   _e_ctc_inv or None,
                        "booking_no":       _e_booking or None,
                        "pay_to":           _e_pay_to or None,
                        "tax_id":           _e_tax_id or None,
                        "tax_name":         _e_tax_name or None,
                        "delivery_port":    _e_delivery or None,
                        "etd":              _e_etd or None,
                        "bl_no":            _e_bl_no or None,
                        "due_date":         _e_due or None,
                        "remark":           _e_remark or None,
                    }
                    try:
                        supabase.table(TBL_LOCAL_CHARGES_V2).update(_upd_payload).eq("id", _edit_id).execute()
                        st.success("✅ บันทึกการแก้ไขเรียบร้อยแล้ว")
                        st.rerun()
                    except Exception as _e:
                        st.error(f"❌ บันทึกไม่สำเร็จ: {_e}")

            st.markdown("**ลบรายการ**")
            delete_options = {
                f"{r.get('agent_invoice_no') or r.get('ctc_invoice_no') or '-'} | {r.get('booking_no') or '-'} | {r.get('pay_to') or '-'} | {r.get('etd') or '-'}": r["id"]
                for r in lc_res.data
            }
            selected_label = st.selectbox("เลือกรายการที่ต้องการลบ", options=["— เลือก —"] + list(delete_options.keys()), key="lc_delete_select")
            if selected_label != "— เลือก —":
                if st.button("🗑️ ลบรายการนี้", type="primary", key="lc_delete_btn"):
                    del_id = delete_options[selected_label]
                    try:
                        supabase.table(TBL_LOCAL_CHARGE_ITEMS).delete().eq("local_charge_id", del_id).execute()
                        supabase.table(TBL_LOCAL_CHARGES_V2).delete().eq("id", del_id).execute()
                        st.success("✅ ลบเรียบร้อยแล้ว")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ ลบไม่สำเร็จ: {e}")
        else:
            st.info("ยังไม่มีข้อมูล")
    except Exception as e:
        st.error(f"❌ Load Error: {e}")

# ─────────────────────────────────────────
# 10. PAGE: EXPORT SUMMARY
# ─────────────────────────────────────────
if page == "📊 Export Summary":
    st.subheader("📊 Export Summary")

    # ── Load all booking numbers that have local charges ──
    try:
        from collections import defaultdict
        lc_res = supabase.table(TBL_LOCAL_CHARGES_V2).select("id,booking_no,ctc_invoice_no,exported_at").execute()
        bno_rows = defaultdict(list)
        for r in lc_res.data:
            if r.get("booking_no"):
                bno_rows[r["booking_no"]].append(r)
    except Exception as e:
        st.error(f"❌ Load Error: {e}")
        bno_rows = {}

    if not bno_rows:
        st.info("ยังไม่มีข้อมูล Local Charges ในระบบ")
    else:
        def _make_label(bno, rows):
            all_exported = all(r.get("exported_at") for r in rows)
            icon = "✅" if all_exported else "⚠️"
            ctc_list = ", ".join(filter(None, (r.get("ctc_invoice_no") for r in rows))) or "—"
            return f"{icon} {bno}  [{ctc_list}]"

        label_to_bno = {_make_label(bno, rows): bno for bno, rows in bno_rows.items()}
        sorted_labels = sorted(label_to_bno)

        selected_labels = st.multiselect(
            "เลือก Booking No.",
            options=sorted_labels,
            placeholder="เลือกได้หลาย Booking No.",
        )
        selected_bnos = [label_to_bno[l] for l in selected_labels]

        # เคลียร์ PDF cache เมื่อ selection เปลี่ยน (คงสถานะ checkbox ไว้)
        if st.session_state.get("_export_bnos") != selected_bnos:
            st.session_state.pop("export_pdf", None)
            st.session_state["_export_bnos"] = selected_bnos
        if "export_checks" not in st.session_state:
            st.session_state["export_checks"] = {}

        if selected_bnos:
            try:
                # ── โหลดข้อมูลทั้งหมดสำหรับ preview ──
                all_records = []   # [{header, items}, ...]
                preview_rows = []  # rows สำหรับ DataFrame

                # โหลด no_container / no_pallet / country จาก bookings table
                bk_res = (
                    supabase.table(TBL_BOOKINGS)
                    .select("booking_no,no_container,no_pallet,country")
                    .in_("booking_no", selected_bnos)
                    .execute()
                )
                bk_map = {r["booking_no"]: r for r in (bk_res.data or [])}

                for bno in selected_bnos:
                    hdrs = (
                        supabase.table(TBL_LOCAL_CHARGES_V2)
                        .select("*")
                        .eq("booking_no", bno)
                        .execute()
                    ).data
                    for hdr in hdrs:
                        its = (
                            supabase.table(TBL_LOCAL_CHARGE_ITEMS)
                            .select("*")
                            .eq("local_charge_id", hdr["id"])
                            .execute()
                        ).data
                        subtotal = sum(float(it.get("total") or 0) for it in its)
                        vat_7 = float(hdr.get("vat_7") or 0)
                        wht_1 = float(hdr.get("wht_1") or 0)
                        wht_3 = float(hdr.get("wht_3") or 0)
                        net   = subtotal + vat_7 - wht_1 - wht_3
                        bk    = bk_map.get(bno, {})
                        all_records.append({"header": hdr, "items": its, "bk": bk})
                        preview_rows.append({
                            "เลือก":          st.session_state["export_checks"].get(hdr["id"], True),
                            "Booking No.":    bno,
                            "CTC Invoice":    hdr.get("ctc_invoice_no") or "—",
                            "Pay To":         hdr.get("pay_to") or "—",
                            "ยอดสุทธิ (THB)": round(net, 2),
                            "Delivery Port":  hdr.get("delivery_port") or "—",
                            "No. Container":  bk.get("no_container") or "—",
                            "No. Pallet":     bk.get("no_pallet") or "—",
                            "สถานะ":          "✅ Exported" if hdr.get("exported_at") else "⚠️ ยังไม่ export",
                        })

                # ── ตารางตัวอย่าง + เลือก/ไม่เลือก ──
                st.markdown("**ตัวอย่างข้อมูล** — ติ๊กถูกรายการที่ต้องการ Export")
                df_preview = pd.DataFrame(preview_rows)
                edited = st.data_editor(
                    df_preview,
                    column_config={
                        "เลือก":          st.column_config.CheckboxColumn("เลือก", default=True),
                        "ยอดสุทธิ (THB)": st.column_config.NumberColumn("ยอดสุทธิ (THB)", format="%.2f"),
                    },
                    disabled=["Booking No.", "CTC Invoice", "Pay To", "ยอดสุทธิ (THB)", "Delivery Port", "No. Container", "No. Pallet", "สถานะ"],
                    hide_index=True,
                    use_container_width=True,
                    key="export_preview_editor",
                )

                # บันทึกสถานะ checkbox กลับใน session state
                for i, rec in enumerate(all_records):
                    st.session_state["export_checks"][rec["header"]["id"]] = bool(edited.iloc[i]["เลือก"])

                n_selected = int(edited["เลือก"].sum())
                st.caption(f"เลือกอยู่ {n_selected} / {len(edited)} รายการ")

                # ── Cover Page Preview & Editor ──────────────────────
                st.divider()
                st.markdown("**ตัวอย่าง Cover Page** — แก้ไข รายการ / Remark ได้ก่อน Export")
                cover_rows = []
                for i, rec in enumerate(all_records):
                    hdr = rec["header"]
                    bk  = rec.get("bk") or {}
                    its = rec["items"]
                    subtotal = sum(float(it.get("total") or 0) for it in its)
                    vat_7 = float(hdr.get("vat_7") or 0)
                    wht_1 = float(hdr.get("wht_1") or 0)
                    wht_3 = float(hdr.get("wht_3") or 0)
                    net   = round(subtotal + vat_7 - wht_1 - wht_3, 2)
                    cats  = sorted(set(it.get("category") or "other" for it in its))
                    default_part = "+ ".join(c.upper().replace("_", "") for c in cats)
                    cover_rows.append({
                        "_id":         hdr["id"],
                        "รายการ":      default_part,
                        "Invoice No.": hdr.get("ctc_invoice_no") or "-",
                        "Country":     bk.get("country") or "-",
                        "Pay To":      hdr.get("pay_to") or "-",
                        "Due Date":    hdr.get("due_date") or "-",
                        "Remark":      hdr.get("remark") or "",
                        "Amount":      net,
                    })

                df_cover = pd.DataFrame(cover_rows)
                edited_cover = st.data_editor(
                    df_cover,
                    column_config={
                        "_id":    st.column_config.Column("_id", disabled=True),
                        "Amount": st.column_config.NumberColumn("Amount (THB)", format="%.2f", disabled=True),
                    },
                    column_order=["รายการ", "Invoice No.", "Country", "Pay To", "Due Date", "Remark", "Amount"],
                    disabled=["_id", "Amount"],
                    hide_index=True,
                    use_container_width=True,
                    key="cover_page_editor",
                )

                # ฉีด cover fields เข้าใน all_records (ใช้เฉพาะ cover page PDF)
                cover_edit_map = {row["_id"]: row for _, row in edited_cover.iterrows()}
                for rec in all_records:
                    eid = rec["header"]["id"]
                    if eid in cover_edit_map:
                        rec["cover_part"]    = cover_edit_map[eid]["รายการ"]
                        rec["cover_inv"]     = cover_edit_map[eid]["Invoice No."]
                        rec["cover_country"] = cover_edit_map[eid]["Country"]
                        rec["cover_payto"]   = cover_edit_map[eid]["Pay To"]
                        rec["cover_due"]     = cover_edit_map[eid]["Due Date"]
                        rec["cover_remark"]  = cover_edit_map[eid]["Remark"]

                st.divider()
                st.markdown("**ผู้รับผิดชอบ**")
                _PREPARED_BY = {
                    "Sutida Suwantatree":       "098-584-2550",
                    "Nattawan Kerdpol":         "094-914-0449",
                    "Alyssa Daenglang":         "095-363-0707",
                    "Waraphon Praneetpolkrang": "062-396-4024",
                    "Wannisa Seesuksam":        "066-109-7538",
                    "Nisachol Pongmulee":       "083-110-3758",
                }
                _pc1, _pc2 = st.columns(2)
                prepared_name  = _pc1.selectbox("ชื่อผู้รับผิดชอบ", options=["— เลือก —"] + list(_PREPARED_BY), key="export_prepared_name")
                prepared_phone = _PREPARED_BY.get(prepared_name, "")
                _pc2.markdown("**เบอร์โทรศัพท์**")
                _pc2.write(prepared_phone if prepared_phone else "—")
                if prepared_name == "— เลือก —":
                    prepared_name = ""

                if st.button("📄 Generate PDF", use_container_width=True, disabled=(n_selected == 0)):
                    filtered = [all_records[i] for i, chk in enumerate(edited["เลือก"]) if chk]
                    if filtered:
                        st.session_state["export_pdf"]      = generate_expense_pdf(filtered, prepared_by=prepared_name, prepared_by_phone=prepared_phone)
                        st.session_state["export_filename"] = "expense_summary_" + "_".join(selected_bnos) + ".pdf"
                        # Mark exported rows in Supabase
                        from datetime import datetime, timezone
                        now_iso = datetime.now(timezone.utc).isoformat()
                        exported_ids = [all_records[i]["header"]["id"] for i, chk in enumerate(edited["เลือก"]) if chk]
                        for eid in exported_ids:
                            supabase.table(TBL_LOCAL_CHARGES_V2).update({"exported_at": now_iso}).eq("id", eid).execute()
                    else:
                        st.warning("กรุณาเลือกอย่างน้อย 1 รายการ")

            except Exception as e:
                st.error(f"❌ Load Error: {e}")

        if st.session_state.get("export_pdf"):
            st.download_button(
                label="⬇️ ดาวน์โหลด PDF",
                data=st.session_state["export_pdf"],
                file_name=st.session_state.get("export_filename", "export.pdf"),
                mime="application/pdf",
                use_container_width=True,
            )

# ─────────────────────────────────────────
# 11. PAGE: RUN NO. ITEMS
# ─────────────────────────────────────────
if page == "📝 Run No. Items":
    st.subheader("📝 ใส่เลขลำดับ Items ใน Invoice/Packing List")

    pdf_file = st.file_uploader(
        "อัปโหลด Invoice + Packing List PDF",
        type="pdf",
        key="num_items_up",
    )

    if pdf_file:
        if st.button("🚀 สร้างไฟล์ที่มีเลขลำดับ", use_container_width=True):
            with st.spinner("กำลังประมวลผล..."):
                try:
                    result_bytes = number_items_in_pdf(pdf_file.read())
                    st.success("🎉 สำเร็จ!")
                    st.download_button(
                        "📥 ดาวน์โหลด PDF",
                        data=result_bytes,
                        file_name=f"numbered_{pdf_file.name}",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"❌ Error: {e}")
                    import traceback; st.code(traceback.format_exc())

    # ══════════════════════════════════════════════════════════════
    #  ADDED FEATURE · Create Invoice/Packing List from Shipping Order
    #  (independent from the numbering tool above — touches nothing else)
    # ══════════════════════════════════════════════════════════════
    st.divider()
    st.subheader("🧾 สร้าง Invoice / Packing List จาก Shipping Order")

    from openpyxl import load_workbook as _ci_load_wb

    _SO_PROMPT = """You are a logistics document expert. The PDF is a Thai "SHIPPING ORDER REQUISITION" form (Carrier Air Conditioning Thailand).
Read it and return ONLY a JSON object (no markdown, no explanation). Schema:
{
 "invoice_no": "value of I/V NO. e.g. CTC26-S141",
 "currency": "USD or THB (unit shown in the item table PRICE / AMOUNT)",
 "consignee_name": "company name in CONSIGNEE'S NAME AND ADDRESS (empty string if only an address, no company, is written)",
 "consignee_address": "the address lines joined with a real newline; exclude company name, ATTN and TEL",
 "attn": "text after ATTN:",
 "tel": "TEL value",
 "email": "email if present else empty string",
 "forwarder": "FORWARDER value e.g. DHL, FEDEX (empty string if the line is blank)",
 "courier_account": "COURIER'S ACCOUNT CODE digits with no spaces",
 "freight_term": "COLLECT or PREPAID — whichever word is circled on the FREIGHT line (empty if unclear)",
 "trade_term": "the circled TRADE TERM among FOB/CIF/C&F/DDP/DDU/DAP, or empty string if none circled",
 "made_in": "the circled COUNTRY OF ORIGINAL (JAPAN or CHINA or THAILAND, or the text written after OTHER)",
 "is_sample": true or false (true if SAMPLE is circled/selected),
 "no_commercial_value": true or false (true ONLY if the bullet 'INVOICE - NO COMMERCIAL VALUE' is filled/selected),
 "destination_country": "COUNTRY field value, or the consignee country/city",
 "package_count": integer from TOTAL NO. OF PACKAGE (default 1),
 "net_weight": number from N.W. OF PRODUCT in kgs,
 "gross_weight": number from G.W. / PACKAGE in kgs,
 "dim_w": number, "dim_l": number, "dim_h": number,  // DIMENSION / PACKAGE  W x L x H  in cms
 "items": [ {"part_no": "MODEL/PART NO.", "description": "DESCRIPTION in English only (ignore Thai text)", "qty": number from QTY, "unit_price": number from PRICE (@)} ]
}
Rules: numbers must be numbers, not strings. Use "" for unknown text, 0 for unknown numbers, [] for no items. Never invent values. If PRICE looks inconsistent with AMOUNT for qty=1, trust AMOUNT."""

    def _so_extract(pdf_bytes):
        cfg = types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0, seed=42)
        res = genai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[types.Content(role="user", parts=[
                types.Part.from_text(text=_SO_PROMPT),
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            ])],
            config=cfg,
        )
        raw = res.text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        return json.loads(raw.strip())

    # Carrier Invoice/Packing List template (Invoice + packinglist sheets), base64
    _CARRIER_TPL_B64 = "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0uJLiy8tvrT40uJLiy8tvrX41uJbi28tvrX41uJbi28tvrX41uJbi28tvrX41uJbi28tvrX41uJbi+8ti984e3n28CeDnv2Gb355dveR51sOHb7M09XFYr7q3vnj1Le8PHu44ZbLx/n5os93jsbZWO3k6fnZ6fm4XJyudmZXR4ud2cjL0/Ob9ZlFr+aLd490fnuakb267POJh8z1Qz55OmEOrdFmd+3p1EkPLjtEW6zbwcQ3Ax3dtb3HE9+ds7r/ik6dpmx+tTfaGdpLtPP7z2HiuVzct4nncnnfJu55hXaN9grtNdoN2hu0MRRDMRVLsRVninPFQ8WF4hFe5XGseKK4VFwpniqeKb7cjB//42j53b98/c3bne//+t3f3j6a+gfaONf9BWJcKt6TePp4+plNH+7+0+MXj598vrf/5PODJy923v2fqWd8rcd9pfha8UbxDWIMxVBMxVK8VzX93TKKc8VDxYXikeKx4oniUnGlqH95hP7tEfrXR5wrXiheKl4pXiu+UnyteKP4BjGHYiimYilq8anFpxafWnxq8anFpxafWnxq8anFpxafWnxq8anFpxafWnxq8anFpxafWnxq8anFlxZfWnxp8aXFlxZfWnxp8aXFlxZfWnxp8aXFlxZfWnxp8aXFlxZfWnxp8aXFlxZfWnxp8aXFlxZfWnxr8a3FtxbfWnxr8a3FtxbfWnxr8a3FtxbfWnxr8a3FtxbfWnxr8a3FtxbfWxa/cWKzvz6xOfhtf15p//1HXkwcLMT+xoN//P7nyZP9nctRfTjOxz+/P4L5/OJw8c+330Ty5ur0eHx+3Kt/3nm6u//582d7u/98d2wzdQhz/0BT329z1yb/GFTjutldm/ojS3Ncdoi2WLfbHw/8YTtaP83Hz7Z8rbYeXU39dFw8idP97XM5Q3u5bk8nvo7naBdol/ftycftCu0a7RUe7zXaDdobtDEUA890pGIptuJMca54qLhQPFI8vovThza6cKm4UtTKh2Y+XiqeK14oXipeKV4rvlJ8rXij+AYxhmIopmIptuJMca54qLhQPFI8VjxRXCquFLX40OJDiw8tPrT40OJDiw8tPrT40OJDiw8tPrX41OJTi08tPrX41OJTi08tPrX41OJTi08tPrX41OJTi08tPrX41OJTi08tPrX41OJTi08tPrX41OJLiy8tvrT40uJLiy8t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WghaCFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmha6G0WNo9vXqyPb/Yf/abTmxfvPzR5dvMC//xXrHXc25s4LOmH+HGb3d914rr5w3UHH8fDh/j847jY/jkePXwak39uCp/jieJSt13pylPFM8WXiueKF4qXileK14qvFF8r3ii+URyDlXsfHPwo1madsc5ZD1kXrEesHP/g+seSlfsfBDAoYJDAoIFBBIMKBhkMOhiEMChhkMKghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWkhfTvg2ghaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpoWihaKH8poAWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFnqbhY2Tm6eP1yc3T37Lyc389vIfp0811mlvbyKeruPkEz9TfKl4rniheKl4pXit+ErxteKN4hvFMViDNVmLtVlnrHPWQ9YF6xHrMesJ65J1xUoAgwIGCQwaGEQwqGCQwaCDQQiDEgYpDFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaCFoIWghaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWkhaSFpIWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFooWihaKFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWmhaaFpoWuhtFjZPGHZ/3wnD3YcmTxju0pYTht3tT/xM8aXiueKF4qXileK14ivF14o3im8Ux2AN1mQt1madsc5ZD1kXrEesx6wnrEvWFSsBDAoYJDBoYBDBoIJBBoMOBiEMShikMGghaCFoIWgh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"

    def _ci_shrink(ws, addr, size):
        import copy as _cp
        c = ws[addr]
        if c.font is not None:
            f = _cp.copy(c.font); f.size = size; c.font = f

    def _ci_clone_style(ws, dst, src):
        import copy as _cp
        for a in ("font", "border", "fill", "alignment", "protection"):
            setattr(ws[dst], a, _cp.copy(getattr(ws[src], a)))
        ws[dst].number_format = ws[src].number_format

    def _ci_copy_row_style(ws, src_row, dst_row, ncols):
        import copy as _cp
        for c in range(1, ncols + 1):
            s = ws.cell(row=src_row, column=c); dc = ws.cell(row=dst_row, column=c)
            dc.font = _cp.copy(s.font); dc.border = _cp.copy(s.border)
            dc.fill = _cp.copy(s.fill); dc.alignment = _cp.copy(s.alignment)
            dc.number_format = s.number_format

    def _ci_fit_patch(xlsx_bytes):
        """inject <pageSetUpPr fitToPage=1> so Excel prints 1 page/sheet (openpyxl may omit it)."""
        import zipfile as _zip, re as _re2, io as _io2
        src = _io2.BytesIO(xlsx_bytes); out = _io2.BytesIO()
        with _zip.ZipFile(src) as zin, _zip.ZipFile(out, "w", _zip.ZIP_DEFLATED) as zout:
            for it in zin.namelist():
                dat = zin.read(it)
                if _re2.match(r"xl/worksheets/sheet\d+\.xml$", it):
                    t = dat.decode("utf-8")
                    if "<pageSetUpPr" not in t:
                        if _re2.search(r"<sheetPr[^>]*/>", t):
                            t = _re2.sub(r"(<sheetPr[^>]*)/>", r'\1><pageSetUpPr fitToPage="1"/></sheetPr>', t, 1)
                        elif "<sheetPr" in t:
                            t = _re2.sub(r"(<sheetPr[^>]*>)", r'\1<pageSetUpPr fitToPage="1"/>', t, 1)
                        else:
                            t = _re2.sub(r"(<dimension )", r'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ', t, 1)
                    dat = t.encode("utf-8")
                zout.writestr(it, dat)
        return out.getvalue()

    def _build_carrier_invoice(d):
        """Fill the Carrier template from a shipping-order dict → xlsx bytes (Invoice + packinglist only)."""
        from openpyxl.worksheet.properties import PageSetupProperties
        tpl = base64.b64decode(_CARRIER_TPL_B64)
        wb = _ci_load_wb(io.BytesIO(tpl))
        inv = wb["Invoice"]; pk = wb["packinglist"]
        if "Sheet1" in wb.sheetnames:
            del wb["Sheet1"]

        cur = (d.get("currency") or "USD").upper()
        items = list(d.get("items") or [])
        if not items:
            items = [{"part_no": "", "description": "", "qty": 0, "unit_price": 0}]
        N = len(items)
        pkg = int(d.get("package_count") or 1)
        nw = float(d.get("net_weight") or 0)
        gw = float(d.get("gross_weight") or 0)
        dw = float(d.get("dim_w") or 0); dl = float(d.get("dim_l") or 0); dh = float(d.get("dim_h") or 0)

        # ── header ──
        inv["A3"] = f"INVOICE NO.{d.get('invoice_no','')}"
        inv["J3"] = "BANGKOK :"
        name = (d.get("consignee_name") or "").strip()
        addr = [l.strip() for l in str(d.get("consignee_address") or "").replace("\\n", "\n").split("\n") if l.strip()]
        if name:
            inv["B5"] = name
            inv["B6"] = addr[0] if len(addr) > 0 else " "
            inv["B7"] = addr[1] if len(addr) > 1 else " "
        else:
            inv["B5"] = addr[0] if addr else " "
            inv["B6"] = addr[1] if len(addr) > 1 else " "
            inv["B7"] = " "
        _ci_clone_style(inv, "B7", "B6")
        inv["B9"]  = d.get("attn", "")
        inv["B10"] = ("MAIL : " + (d.get("email") or "")).strip()
        inv["B11"] = ("TEL: " + (d.get("tel") or "")).strip()
        fwd = (d.get("forwarder") or "").strip(); acct = (d.get("courier_account") or "").strip()
        inv["B13"] = (f"{fwd} / {acct}".strip(" /")) if (fwd or acct) else " "
        inv["B14"] = "BANGKOK, THAILAND "
        inv["F14"] = d.get("destination", "") or ""
        inv["A15"] = "SAILING ON/OR ABOUT :"

        mark_name = name or (addr[0] if addr else "")
        inv["J6"] = mark_name
        inv["J7"] = "SAMPLE" if d.get("is_sample") else " "
        inv["J8"] = ("MADE IN " + str(d.get("made_in") or "").upper()).strip()
        inv["J9"] = f"INV.NO.{d.get('invoice_no','')}"
        inv["J10"] = "C/NO.1" + (f"/{pkg}-{pkg}/{pkg}" if pkg > 1 else "")
        ft = (d.get("freight_term") or "").upper()
        inv["J13"] = f'"FREIGHT {ft}"' if ft else " "
        tt = (d.get("trade_term") or "").upper()
        inv["J14"] = f"TERM {tt}" if tt else " "
        _mark_sz = 9 if len(mark_name) > 34 else (11 if len(mark_name) > 24 else None)
        if _mark_sz:
            _ci_shrink(inv, "J6", _mark_sz)
        if not name and addr:
            _ci_shrink(inv, "B5", 13)

        if not d.get("no_commercial_value"):
            inv["H19"] = None; inv["H20"] = None
        inv["J21"] = cur; inv["M21"] = cur
        inv["B22"] = d.get("category_header") or "SAMPLE OF AIR CONDITIONER PART"

        # ── item rows (invoice from 23, packing from 22) ──
        if N > 5:
            inv.insert_rows(28, N - 5)
            pk.insert_rows(27, N - 5)
            for i in range(1, N):
                _ci_copy_row_style(inv, 23, 23 + i, 15)
                _ci_copy_row_style(pk, 22, 22 + i, 15)
        for i in range(N):
            it = items[i]
            r = 23 + i
            inv.cell(row=r, column=2, value=it.get("part_no", ""))
            inv.cell(row=r, column=3, value=it.get("description", ""))
            inv.cell(row=r, column=7, value=it.get("qty", 0))
            inv.cell(row=r, column=8, value="PCS")
            inv.cell(row=r, column=10, value=it.get("unit_price", 0))
            inv.cell(row=r, column=13, value=f"=+J{r}*G{r}")
            pr = 22 + i
            pk.cell(row=pr, column=2, value=it.get("part_no", ""))
            pk.cell(row=pr, column=3, value=it.get("description", ""))
            pk.cell(row=pr, column=5, value=f"=Invoice!G{r}")
            pk.cell(row=pr, column=6, value="PCS")
            for _wc in (7, 9, 11, 13, 14, 15):   # clear any leftover weight/dim from template
                pk.cell(row=pr, column=_wc).value = None
        # clear leftover template rows when N < 5
        for r in range(23 + N, 28):
            for col in (2, 3, 7, 8, 10, 13):
                inv.cell(row=r, column=col).value = None
        for r in range(22 + N, 27):
            for col in (2, 3, 5, 6, 7, 9, 11, 13, 14, 15):
                pk.cell(row=r, column=col).value = None

        cbm_total = round(dw * dl * dh / 1000000.0 * pkg, 6) if (dw and dl and dh) else 0
        # weights: single item -> on its row; else only on carton total row
        if N == 1:
            pk.cell(row=22, column=7, value=nw)
            pk.cell(row=22, column=9, value=gw)
            pk.cell(row=22, column=13, value=dw)
            pk.cell(row=22, column=14, value=dl)
            pk.cell(row=22, column=15, value=dh)
            pk.cell(row=22, column=11, value="=M22*N22*O22/1000000")

        inv_tot = 29 + max(0, N - 5)
        pk_tot = 29 + max(0, N - 5)
        carton_lbl = f"{pkg} CARTON" + ("S" if pkg > 1 else "")
        # invoice total
        inv.cell(row=inv_tot, column=1, value="TOTAL")
        inv.cell(row=inv_tot, column=2, value=carton_lbl)
        inv.cell(row=inv_tot, column=7, value=f"=SUM(G23:G{inv_tot-1})")
        inv.cell(row=inv_tot, column=8, value="PCS")
        inv.cell(row=inv_tot, column=13, value=f"=SUM(M23:M{inv_tot-1})")
        inv.cell(row=inv_tot + 2, column=3, value=nw)   # TOTAL N.W.
        inv.cell(row=inv_tot + 3, column=3, value=gw)   # TOTAL G.W.
        inv.cell(row=inv_tot + 2, column=15, value=f"=M{inv_tot}*10%")
        inv.cell(row=inv_tot + 3, column=15, value=f"=SUM(M{inv_tot}+O{inv_tot+2})*1%")
        inv.cell(row=inv_tot + 4, column=15, value=f"=M{inv_tot}-O{inv_tot+2}-O{inv_tot+3}")
        # packing total
        pk.cell(row=pk_tot, column=1, value=carton_lbl)
        pk.cell(row=pk_tot, column=5, value=f"=SUM(E22:E{pk_tot-1})")
        pk.cell(row=pk_tot, column=6, value="PCS")
        pk.cell(row=pk_tot, column=7, value=nw)
        pk.cell(row=pk_tot, column=9, value=gw)
        pk.cell(row=pk_tot, column=11, value=cbm_total)

        pk["B8"] = "=Invoice!B7"          # 2nd consignee address line
        _ci_clone_style(pk, "B8", "B7")
        if _mark_sz:
            _ci_shrink(pk, "I7", _mark_sz)

        for ws in (inv, pk):
            ws.page_setup.orientation = "portrait"
            ws.page_setup.scale = None
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.print_options.horizontalCentered = True
            ws.page_margins.left = 0.4; ws.page_margins.right = 0.3
            ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.4

        buf = io.BytesIO(); wb.save(buf); data_bytes = buf.getvalue()
        # inserted rows lose custom height in-session -> fix on a fresh reload
        if N > 5:
            wb2 = _ci_load_wb(io.BytesIO(data_bytes))
            for r in range(23, inv_tot + 5):        # items + total + N.W./G.W. rows
                wb2["Invoice"].row_dimensions[r].height = 19.25
            for r in range(22, pk_tot + 1):
                wb2["packinglist"].row_dimensions[r].height = 19.25
            buf2 = io.BytesIO(); wb2.save(buf2); data_bytes = buf2.getvalue()
        return _ci_fit_patch(data_bytes)

    _so_pdf = st.file_uploader("อัปโหลด Shipping Order Requisition (PDF)", type="pdf", key="so_up")
    if _so_pdf is not None and st.button("🔎 อ่านข้อมูลด้วย AI", key="so_read", use_container_width=True):
        with st.spinner("AI กำลังอ่านเอกสาร..."):
            try:
                st.session_state["so_data"] = _so_extract(_so_pdf.read())
                st.session_state.pop("so_xlsx", None)
                st.success("อ่านข้อมูลสำเร็จ — ตรวจ/แก้ไขด้านล่างแล้วกดสร้าง")
            except Exception as e:
                st.error(f"อ่านไม่สำเร็จ: {e}")

    if st.session_state.get("so_data"):
        _d = st.session_state["so_data"]
        st.markdown("**ตรวจ/แก้ไขข้อมูล** (ช่องที่ AI ไม่แน่ใจอาจว่าง โปรดเติมเอง)")
        c1, c2 = st.columns(2)
        _inv_no = c1.text_input("Invoice No.", _d.get("invoice_no", ""), key="so_invno")
        _cur    = c2.selectbox("Currency", ["USD", "THB"],
                               index=(0 if str(_d.get("currency", "USD")).upper() != "THB" else 1), key="so_cur")
        _cname  = st.text_input("Consignee (ชื่อบริษัท — เว้นว่างได้ถ้าไม่มี)", _d.get("consignee_name", ""), key="so_cname")
        _caddr  = st.text_area("Consignee Address", str(_d.get("consignee_address", "")).replace("\\n", "\n"),
                               key="so_caddr", height=70)
        c3, c4 = st.columns(2)
        _attn = c3.text_input("ATTN", _d.get("attn", ""), key="so_attn")
        _tel  = c4.text_input("TEL", _d.get("tel", ""), key="so_tel")
        c5, c6 = st.columns(2)
        _email = c5.text_input("Email", _d.get("email", ""), key="so_email")
        _dest  = c6.text_input("TO (เมือง/ประเทศปลายทาง)", _d.get("destination_country", ""), key="so_dest")
        c7, c8 = st.columns(2)
        _fwd  = c7.text_input("Forwarder", _d.get("forwarder", ""), key="so_fwd")
        _acct = c8.text_input("Courier Account", str(_d.get("courier_account", "")), key="so_acct")
        c9, c10, c11 = st.columns(3)
        _ft = c9.selectbox("Freight", ["COLLECT", "PREPAID"],
                           index=(1 if str(_d.get("freight_term", "")).upper() == "PREPAID" else 0), key="so_ft")
        _tt = c10.text_input("Trade Term (เว้นว่างได้)", _d.get("trade_term", ""), key="so_tt")
        _made = c11.text_input("Made In", _d.get("made_in", ""), key="so_made")
        c12, c13 = st.columns(2)
        _cat = c12.text_input("หัวข้อสินค้า (category)", _d.get("category_header", "SAMPLE OF AIR CONDITIONER PART"), key="so_cat")
        _ncv = c13.checkbox("NO COMMERCIAL VALUE", value=bool(_d.get("no_commercial_value", False)), key="so_ncv")
        c14, c15, c16, c17 = st.columns(4)
        _pkg = c14.number_input("Packages", min_value=1, value=int(_d.get("package_count") or 1), step=1, key="so_pkg")
        _nw  = c15.number_input("N.W. (kg)", min_value=0.0, value=float(_d.get("net_weight") or 0), step=0.1, key="so_nw")
        _gw  = c16.number_input("G.W. (kg)", min_value=0.0, value=float(_d.get("gross_weight") or 0), step=0.1, key="so_gw")
        c18, c19, c20 = st.columns(3)
        _dwv = c18.number_input("Dim W (cm)", min_value=0.0, value=float(_d.get("dim_w") or 0), step=1.0, key="so_dw")
        _dlv = c19.number_input("Dim L (cm)", min_value=0.0, value=float(_d.get("dim_l") or 0), step=1.0, key="so_dl")
        _dhv = c20.number_input("Dim H (cm)", min_value=0.0, value=float(_d.get("dim_h") or 0), step=1.0, key="so_dh")

        st.markdown("**รายการสินค้า**")
        _items_src = _d.get("items") or [{"part_no": "", "description": "", "qty": 0, "unit_price": 0.0}]
        _items_df = pd.DataFrame(_items_src)
        for _c in ["part_no", "description", "qty", "unit_price"]:
            if _c not in _items_df.columns:
                _items_df[_c] = "" if _c in ("part_no", "description") else 0
        _items_df = _items_df[["part_no", "description", "qty", "unit_price"]]
        _edited = st.data_editor(_items_df, num_rows="dynamic", use_container_width=True, key="so_items")

        if st.button("📥 สร้างไฟล์ Excel", use_container_width=True, key="so_make"):
            try:
                _items = []
                for _, _row in _edited.iterrows():
                    _pn = str(_row.get("part_no", "") or "").strip()
                    _ds = str(_row.get("description", "") or "").strip()
                    if not _pn and not _ds:
                        continue
                    _items.append({
                        "part_no": _pn, "description": _ds,
                        "qty": float(_row.get("qty", 0) or 0),
                        "unit_price": float(_row.get("unit_price", 0) or 0),
                    })
                _payload = {
                    "invoice_no": _inv_no, "currency": _cur,
                    "consignee_name": _cname, "consignee_address": _caddr,
                    "attn": _attn, "tel": _tel, "email": _email,
                    "destination": _dest, "forwarder": _fwd, "courier_account": _acct,
                    "freight_term": _ft, "trade_term": _tt, "made_in": _made,
                    "category_header": _cat, "no_commercial_value": _ncv,
                    "is_sample": True,
                    "package_count": int(_pkg), "net_weight": _nw, "gross_weight": _gw,
                    "dim_w": _dwv, "dim_l": _dlv, "dim_h": _dhv,
                    "items": _items,
                }
                st.session_state["so_xlsx"] = _build_carrier_invoice(_payload)
                st.session_state["so_xlsx_name"] = f"INV & PL_{_inv_no or 'CTC'}_WITH INS+FREIGHT.xlsx"
                st.success("🎉 สร้างไฟล์สำเร็จ!")
            except Exception as e:
                st.error(f"สร้างไม่สำเร็จ: {e}")

        if st.session_state.get("so_xlsx"):
            st.download_button(
                "⬇️ ดาวน์โหลด Excel",
                data=st.session_state["so_xlsx"],
                file_name=st.session_state.get("so_xlsx_name", "INV & PL.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="so_dl",
            )
