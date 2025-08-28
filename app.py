# -*- coding: utf-8 -*-
"""
IT Stock (Streamlit + Google Sheets)
v11:
- แก้ปัญหา PDF แสดงภาษาไทยเป็นสี่เหลี่ยม
  * รองรับฟอนต์ไทย (Sarabun / TH Sarabun New / Noto Sans Thai)
  * ค้นหาอัตโนมัติจากโฟลเดอร์ ./fonts, Windows Fonts, และตำแหน่งทั่วไปบน Linux/Mac
  * หัวตารางใช้ฟอนต์หนา ถ้ามีไฟล์ Bold
  * ถ้าไม่พบฟอนต์ จะแจ้งเตือนในหน้าเว็บให้ติดตั้ง และยังสร้าง PDF ได้ด้วยฟอนต์เริ่มต้น
- รวมทุกฟีเจอร์จาก v10 (Dashboard, Stock, เบิก/รับ, รายงาน, Users, นำเข้า/แก้ไข หมวดหมู่, Settings + Clear test data)
"""
import os, io, uuid, re
from datetime import datetime, timedelta, date, time as dtime
import pytz, pandas as pd, streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import gspread
from google.oauth2.service_account import Credentials
import bcrypt
import altair as alt

# ---- Compatibility helper for Streamlit rerun ----
def safe_rerun():
    import streamlit as st

# ----------------- Global UI Style -----------------
def inject_global_css():
    import streamlit as st
    st.markdown(
        """
        <style>
        html, body, [class*="css"]  {
            font-size: 14px !important;
        }
        /* Responsive adjustments for small screens */
        @media (max-width: 768px) {
            html, body, [class*="css"] {
                font-size: 13px !important;
            }
            h1 { font-size: 1.5rem !important; }
            h2 { font-size: 1.3rem !important; }
            h3 { font-size: 1.1rem !important; }
            .stButton>button {
                padding: 0.25rem 0.75rem;
                font-size: 0.9rem;
            }
            .stTextInput>div>div>input, .stTextArea>div>textarea {
                font-size: 0.9rem !important;
            }
            .stDataFrame { font-size: 0.85rem !important; }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        safe_rerun()


APP_TITLE = "IT Intelligent System"
APP_TAGLINE = "Minimal, Modern, and Practical"
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit?gid=0#gid=0"
CREDENTIALS_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")

SHEET_ITEMS     = "Items"
SHEET_TXNS      = "Transactions"
SHEET_USERS     = "Users"
SHEET_CATS      = "Categories"
SHEET_BRANCHES  = "Branches"
SHEET_TICKETS   = "Tickets"
SHEET_TICKET_CATS = "TicketCategories"

ITEMS_HEADERS   = ["รหัส","หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
TXNS_HEADERS    = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"]
USERS_HEADERS   = ["Username","DisplayName","Role","PasswordHash","Active"]
CATS_HEADERS    = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS      = ["รหัสสาขา","ชื่อสาขา"]

TICKETS_HEADERS = ["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]
TICKET_CAT_HEADERS = ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]

TZ = pytz.timezone("Asia/Bangkok")

MINIMAL_CSS = """
<style>
:root { --radius: 16px; }
section.main > div { padding-top: 8px; }
.block-card { background: #fff; border:1px solid #eee; border-radius:16px; padding:16px; }
.kpi { display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.danger { color:#b00020; }
</style>"""

RESPONSIVE_CSS = """
<style>
/* Base font sizing */
html, body, [data-testid="stAppViewContainer"] { font-size: 15px; }

/* Tablet and below */
@media (max-width: 768px){
  html, body, [data-testid="stAppViewContainer"]{ font-size: 14px; }
  h1{font-size:1.6rem;} h2{font-size:1.35rem;} h3{font-size:1.15rem;}
  .stMarkdown, label, .stTextInput label, .stSelectbox label, .stFileUploader label{ font-size:0.95rem; }
  .stTextInput input, .stTextArea textarea{ font-size:0.95rem; }
  .stButton button{ font-size:0.95rem; padding:0.45rem 0.8rem; border-radius:10px; }
  [data-testid="stDataFrame"] div, [data-testid="stDataFrame"] table{ font-size:0.92rem; }
}

/* Small phones */
@media (max-width: 480px){
  html, body, [data-testid="stAppViewContainer"]{ font-size: 13px; }
  h1{font-size:1.45rem;} h2{font-size:1.25rem;} h3{font-size:1.1rem;}
  .stMarkdown, label, .stTextInput label, .stSelectbox label, .stFileUploader label{ font-size:0.9rem; }
  .stTextInput input, .stTextArea textarea{ font-size:0.9rem; }
  .stButton button{ font-size:0.9rem; padding:0.4rem 0.75rem; border-radius:10px; }
  [data-testid="stDataFrame"] div, [data-testid="stDataFrame"] table{ font-size:0.88rem; }
}
</style>"""


def ensure_credentials_ui():
    if os.path.exists(CREDENTIALS_FILE): return True
    st.warning("ยังไม่พบไฟล์ service_account.json")
    up = st.file_uploader("อัปโหลดไฟล์ service_account.json", type=["json"])
    if up is not None:
        with open(CREDENTIALS_FILE, "wb") as f: f.write(up.read())
        st.success("บันทึกไฟล์แล้ว รีเฟรช..."); st.rerun()
    st.stop()

@st.cache_resource(show_spinner=False)
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return gspread.authorize(creds)

def open_sheet_by_url(sheet_url: str):
    gc = get_client()
    return gc.open_by_url(sheet_url)

def ensure_sheets_exist(sh):
    titles = [ws.title for ws in sh.worksheets()]
    if SHEET_ITEMS not in titles:
        ws = sh.add_worksheet(SHEET_ITEMS, 1000, len(ITEMS_HEADERS)+5); ws.append_row(ITEMS_HEADERS)
    if SHEET_TXNS not in titles:
        ws = sh.add_worksheet(SHEET_TXNS, 2000, len(TXNS_HEADERS)+5); ws.append_row(TXNS_HEADERS)
    if SHEET_USERS not in titles:
        ws = sh.add_worksheet(SHEET_USERS, 100, len(USERS_HEADERS)+2); ws.append_row(USERS_HEADERS)
        default_pwd = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        sh.worksheet(SHEET_USERS).append_row(["admin","Administrator","admin",default_pwd,"Y"])
    if SHEET_CATS not in titles:
        ws = sh.add_worksheet(SHEET_CATS, 200, len(CATS_HEADERS)+2); ws.append_row(CATS_HEADERS)
    if SHEET_BRANCHES not in titles:
        ws = sh.add_worksheet(SHEET_BRANCHES, 200, len(BR_HEADERS)+2); ws.append_row(BR_HEADERS)
    if SHEET_TICKETS not in titles:
        ws = sh.add_worksheet(SHEET_TICKETS, 1000, len(TICKETS_HEADERS)+5); ws.append_row(TICKETS_HEADERS)
    if SHEET_TICKET_CATS not in titles:
        ws = sh.add_worksheet(SHEET_TICKET_CATS, 200, len(TICKET_CAT_HEADERS)+2); ws.append_row(TICKET_CAT_HEADERS)

def read_df(sh, title, headers):
    ws = sh.worksheet(title); vals = ws.get_all_values()
    if not vals: return pd.DataFrame(columns=headers)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    return df if not df.empty else pd.DataFrame(columns=headers)

def write_df(sh, title, df):
    if title==SHEET_ITEMS: cols=ITEMS_HEADERS
    elif title==SHEET_TXNS: cols=TXNS_HEADERS
    elif title==SHEET_USERS: cols=USERS_HEADERS
    elif title==SHEET_CATS: cols=CATS_HEADERS
    elif title==SHEET_BRANCHES: cols=BR_HEADERS
    else: cols = df.columns.tolist()
    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols]
    ws = sh.worksheet(title)
    ws.clear(); ws.update([df.columns.values.tolist()] + df.values.tolist())

def append_row(sh, title, row): sh.worksheet(title).append_row(row)

def auth_block(sh):
    st.session_state.setdefault("user", None); st.session_state.setdefault("role", None)
    if st.session_state.get("user"):
        with st.sidebar:
            st.markdown(f"**👤 {st.session_state['user']}**"); st.caption(f"Role: {st.session_state['role']}")
            if st.button("ออกจากระบบ"): st.session_state["user"]=None; st.session_state["role"]=None; st.rerun()
        return True
    st.sidebar.subheader("เข้าสู่ระบบ")
    u = st.sidebar.text_input("Username"); p = st.sidebar.text_input("Password", type="password")
    if st.sidebar.button("Login", use_container_width=True):
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
        row = users[(users["Username"]==u) & (users["Active"].str.upper()=="Y")]
        if not row.empty:
            ok = False
            try: ok = bcrypt.checkpw(p.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8"))
            except: ok = False
            if ok: st.session_state["user"]=u; st.session_state["role"]=row.iloc[0]["Role"]; st.success("เข้าสู่ระบบสำเร็จ"); st.rerun()
            else: st.error("รหัสผ่านไม่ถูกต้อง")
        else: st.error("ไม่พบบัญชีหรือถูกปิดใช้งาน")
    return False

# -------- Thai font registration --------
def register_thai_fonts() -> dict:
    """Try to register a Thai TTF font. Return {'normal': name, 'bold': name_or_None}"""
    candidates = [
        # local project fonts folder
        ("ThaiFont", "./fonts/Sarabun-Regular.ttf", "./fonts/Sarabun-Bold.ttf"),
        ("ThaiFont", "./fonts/THSarabunNew.ttf", "./fonts/THSarabunNew-Bold.ttf"),
        ("ThaiFont", "./fonts/NotoSansThai-Regular.ttf", "./fonts/NotoSansThai-Bold.ttf"),
        # Windows
        ("ThaiFont", "C:/Windows/Fonts/Sarabun-Regular.ttf", "C:/Windows/Fonts/Sarabun-Bold.ttf"),
        ("ThaiFont", "C:/Windows/Fonts/THSarabunNew.ttf", "C:/Windows/Fonts/THSarabunNew-Bold.ttf"),
        ("ThaiFont", "C:/Windows/Fonts/NotoSansThai-Regular.ttf", "C:/Windows/Fonts/NotoSansThai-Bold.ttf"),
        # Linux common
        ("ThaiFont", "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf"),
        ("ThaiFont", "/usr/share/fonts/truetype/sarabun/Sarabun-Regular.ttf", "/usr/share/fonts/truetype/sarabun/Sarabun-Bold.ttf"),
        # macOS
        ("ThaiFont", "/Library/Fonts/NotoSansThai-Regular.ttf", "/Library/Fonts/NotoSansThai-Bold.ttf"),
    ]
    chosen = None
    for fam, normal_path, bold_path in candidates:
        if os.path.exists(normal_path):
            try:
                pdfmetrics.registerFont(TTFont(fam, normal_path))
                bold_name = None
                if os.path.exists(bold_path):
                    bold_name = fam + "-Bold"
                    pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                return {"normal": fam, "bold": bold_name}
            except Exception:
                continue
    return {"normal": None, "bold": None}

def df_to_pdf_bytes(df, title="รายงาน", subtitle=""):
    # Register Thai font (if available)
    f = register_thai_fonts()
    use_thai = f["normal"] is not None
    if not use_thai:
        st.warning("⚠️ ไม่พบฟอนต์ไทยสำหรับ PDF (Sarabun / TH Sarabun New / Noto Sans Thai). โปรดวางไฟล์ .ttf ไว้ในโฟลเดอร์ ./fonts แล้วลองใหม่อีกครั้ง.", icon="⚠️")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=14, rightMargin=14, topMargin=14, bottomMargin=14
    )
    styles = getSampleStyleSheet()

    # Override to Thai font
    if use_thai:
        styles["Normal"].fontName = f["normal"]
        styles["Normal"].fontSize = 11
        styles["Normal"].leading = 14
        styles["Normal"].wordWrap = 'CJK'
        # Create a Thai Title style
        styles.add(ParagraphStyle(name="ThaiTitle", parent=styles["Title"],
                                  fontName=f["bold"] or f["normal"],
                                  fontSize=18, leading=22, wordWrap='CJK'))
        styles.add(ParagraphStyle(name="ThaiHeader", parent=styles["Normal"],
                                  fontName=f["bold"] or f["normal"],
                                  fontSize=12, leading=15, wordWrap='CJK'))
        title_style = styles["ThaiTitle"]
        header_style = styles["ThaiHeader"]
    else:
        title_style = styles["Title"]
        header_style = styles["Heading4"]

    story = []
    story.append(Paragraph(title, title_style))
    if subtitle:
        story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1, 8))

    if df.empty:
        story.append(Paragraph("ไม่มีข้อมูล", styles["Normal"]))
    else:
        # Ensure string & Thai-compatible data
        data = [df.columns.astype(str).tolist()] + df.astype(str).values.tolist()
        table = Table(data, repeatRows=1)

        # Table style
        ts = [
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f2f2f2')),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('GRID', (0,0), (-1,-1), 0.25, colors.HexColor('#5a5a5a')),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]
        if use_thai:
            ts.append(('FONTNAME', (0,0), (-1,-1), f["normal"]))
            ts.append(('FONTNAME', (0,0), (-1,0), f["bold"] or f["normal"]))  # header row bold if available

        table.setStyle(TableStyle(ts))
        story.append(table)

    doc.build(story)
    pdf = buf.getvalue(); buf.close()
    return pdf

# ---------- (ส่วนที่เหลือเหมือน v10) ----------
def fmt_dt(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")

def get_now_str(): 
    return fmt_dt(datetime.now(TZ))

def combine_date_time(d: date, t: dtime) -> datetime:
    naive = datetime.combine(d, t)
    return TZ.localize(naive)

def generate_item_code(sh, cat_code: str) -> str:
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    pattern = re.compile(rf"^{re.escape(cat_code)}-(\d+)$")
    max_num = 0
    for code in items["รหัส"].dropna().astype(str):
        m = pattern.match(code.strip())
        if m:
            try:
                num = int(m.group(1))
                if num > max_num: max_num = num
            except: pass
    next_num = max_num + 1
    return f"{cat_code}-{next_num:03d}"

def ensure_item_row(items_df, code): return (items_df["รหัส"]==code).any()

def adjust_stock(sh, code, delta, actor, branch="", note="", txn_type="OUT", ts_str=None):
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    if items.empty or not ensure_item_row(items, code): st.error("ไม่พบรหัสอุปกรณ์นี้ในคลัง"); return False
    row = items[items["รหัส"]==code].iloc[0]
    cur = int(float(row["คงเหลือ"])) if str(row["คงเหลือ"]).strip()!="" else 0
    if txn_type=="OUT" and cur+delta < 0: st.error("สต็อกไม่เพียงพอ"); return False
    items.loc[items["รหัส"]==code, "คงเหลือ"] = cur+delta; write_df(sh, SHEET_ITEMS, items)
    ts = ts_str if ts_str else get_now_str()
    append_row(sh, SHEET_TXNS, [str(uuid.uuid4())[:8], ts, txn_type, code, row["ชื่ออุปกรณ์"], branch, abs(delta), actor, note])
    return True

def make_pie(df: pd.DataFrame, label_col: str, value_col: str, top_n: int, title: str):
    if df.empty or (value_col in df.columns and pd.to_numeric(df[value_col], errors="coerce").fillna(0).sum() == 0):
        st.info(f"ยังไม่มีข้อมูลสำหรับกราฟ: {title}")
        return
    work = df.copy()
    if value_col in work.columns:
        work[value_col] = pd.to_numeric(work[value_col], errors="coerce").fillna(0)
    work = work.groupby(label_col, dropna=False)[value_col].sum().reset_index().rename(columns={value_col: "sum_val"})
    work[label_col] = work[label_col].replace("", "ไม่ระบุ")
    work = work.sort_values("sum_val", ascending=False)
    if len(work) > top_n:
        top = work.head(top_n)
        others = pd.DataFrame({label_col:["อื่นๆ"], "sum_val":[work["sum_val"].iloc[top_n:].sum()]})
        work = pd.concat([top, others], ignore_index=True)
    total = work["sum_val"].sum()
    work["เปอร์เซ็นต์"] = (work["sum_val"] / total * 100).round(2) if total>0 else 0
    st.markdown(f"**{title}**")
    chart = alt.Chart(work).mark_arc(innerRadius=60).encode(
        theta="sum_val:Q",
        color=f"{label_col}:N",
        tooltip=[f"{label_col}:N","sum_val:Q","เปอร์เซ็นต์:Q"]
    )
    st.altair_chart(chart, use_container_width=True)

def parse_range(choice: str, d1: date=None, d2: date=None):
    today = datetime.now(TZ).date()
    if choice == "วันนี้":
        return today, today
    if choice == "7 วันล่าสุด":
        return today - timedelta(days=6), today
    if choice == "30 วันล่าสุด":
        return today - timedelta(days=29), today
    if choice == "90 วันล่าสุด":
        return today - timedelta(days=89), today
    if choice == "ปีนี้":
        return date(today.year, 1, 1), today
    if choice == "กำหนดเอง" and d1 and d2:
        return d1, d2
    return today - timedelta(days=29), today

def page_dashboard(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📊 Dashboard (ปรับแต่งได้)")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)

    total_items = len(items)
    total_qty = items["คงเหลือ"].apply(lambda x: int(float(x)) if str(x).strip() != "" else 0).sum() if not items.empty else 0
    low_df = items[(items["ใช้งาน"].str.upper() == "Y") & (items["คงเหลือ"].astype(str) != "")]
    if not low_df.empty:
        low_df = low_df[pd.to_numeric(low_df["คงเหลือ"], errors='coerce').fillna(0) <= pd.to_numeric(low_df["จุดสั่งซื้อ"], errors='coerce').fillna(0)]
    low_count = len(low_df) if not low_df.empty else 0

    c1, c2, c3 = st.columns(3)
    with c1: st.metric("จำนวนรายการ", f"{total_items:,}")
    with c2: st.metric("ยอดคงเหลือรวม", f"{total_qty:,}")
    with c3: st.metric("ใกล้หมดสต็อก", f"{low_count:,}")

    st.markdown("### 🎛️ ตัวเลือกการแสดงผล")
    colA, colB, colC = st.columns([2,1,1])
    with colA:
        chart_opts = st.multiselect(
            "เลือกกราฟวงกลมที่ต้องการแสดง",
            options=[
                "คงเหลือตามหมวดหมู่",
                "คงเหลือตามที่เก็บ",
                "จำนวนรายการตามหมวดหมู่",
                "เบิกตามสาขา (OUT)",
                "เบิกตามอุปกรณ์ (OUT)",
                "เบิกตามหมวดหมู่ (OUT)",
                "Ticket ตามสถานะ",
                "Ticket ตามสาขา",
            ],
            default=["คงเหลือตามหมวดหมู่","เบิกตามสาขา (OUT)","Ticket ตามสถานะ"]
        )
    with colB:
        top_n = st.slider("Top-N ต่อกราฟ", min_value=3, max_value=20, value=10, step=1)
    with colC:
        per_row = st.selectbox("จำนวนกราฟต่อแถว", [1,2,3,4], index=1)

    st.markdown("### ⏱️ ช่วงเวลา (ใช้กับกราฟประเภท 'เบิก ... (OUT)' เท่านั้น)")
    colR1, colR2, colR3 = st.columns(3)
    with colR1:
        range_choice = st.selectbox("เลือกช่วงเวลา", ["วันนี้","7 วันล่าสุด","30 วันล่าสุด","90 วันล่าสุด","ปีนี้","กำหนดเอง"], index=2)
    with colR2:
        d1 = st.date_input("วันที่เริ่ม", value=datetime.now(TZ).date()-timedelta(days=29))
    with colR3:
        d2 = st.date_input("วันที่สิ้นสุด", value=datetime.now(TZ).date())
    start_date, end_date = parse_range(range_choice, d1, d2)

    if not txns.empty:
        tx = txns.copy()
        tx["วันเวลา"] = pd.to_datetime(tx["วันเวลา"], errors='coerce')
        tx = tx.dropna(subset=["วันเวลา"])
        tx = tx[(tx["วันเวลา"].dt.date >= start_date) & (tx["วันเวลา"].dt.date <= end_date)]
        tx["จำนวน"] = pd.to_numeric(tx["จำนวน"], errors="coerce").fillna(0)
        tx_out = tx[tx["ประเภท"]=="OUT"]
    else:
        tx_out = pd.DataFrame(columns=TXNS_HEADERS)

    charts = []
    if "คงเหลือตามหมวดหมู่" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("หมวดหมู่")["คงเหลือ"].sum().reset_index()
        charts.append(("คงเหลือตามหมวดหมู่", tmp, "หมวดหมู่", "คงเหลือ"))

    if "คงเหลือตามที่เก็บ" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("ที่เก็บ")["คงเหลือ"].sum().reset_index()
        charts.append(("คงเหลือตามที่เก็บ", tmp, "ที่เก็บ", "คงเหลือ"))

    if "จำนวนรายการตามหมวดหมู่" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["count"] = 1
        tmp = tmp.groupby("หมวดหมู่")["count"].sum().reset_index()
        charts.append(("จำนวนรายการตามหมวดหมู่", tmp, "หมวดหมู่", "count"))

    if "เบิกตามสาขา (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("สาขา", dropna=False)["จำนวน"].sum().reset_index()
            charts.append((f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", tmp, "สาขา", "จำนวน"))
        else:
            charts.append((f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"สาขา":[], "จำนวน":[]}), "สาขา", "จำนวน"))

    if "เบิกตามอุปกรณ์ (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("ชื่ออุปกรณ์")["จำนวน"].sum().reset_index()
            charts.append((f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", tmp, "ชื่ออุปกรณ์", "จำนวน"))
        else:
            charts.append((f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"ชื่ออุปกรณ์":[], "จำนวน":[]}), "ชื่ออุปกรณ์", "จำนวน"))

    if "เบิกตามหมวดหมู่ (OUT)" in chart_opts:
        if not tx_out.empty and not items.empty:
            it = items[["รหัส","หมวดหมู่"]].copy()
            tmp = tx_out.merge(it, left_on="รหัส", right_on="รหัส", how="left")
            tmp = tmp.groupby("หมวดหมู่")["จำนวน"].sum().reset_index()
            charts.append((f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", tmp, "หมวดหมู่", "จำนวน"))
        else:
            charts.append((f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"หมวดหมู่":[], "จำนวน":[]}), "หมวดหมู่", "จำนวน"))


    # ----- Tickets Summary (use the same date range based on 'วันที่แจ้ง') -----
    tickets_df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    if not tickets_df.empty:
        tdf = tickets_df.copy()
        tdf["วันที่แจ้ง"] = pd.to_datetime(tdf["วันที่แจ้ง"], errors="coerce")
        tdf = tdf.dropna(subset=["วันที่แจ้ง"])
        tdf = tdf[(tdf["วันที่แจ้ง"].dt.date >= start_date) & (tdf["วันที่แจ้ง"].dt.date <= end_date)]
    else:
        tdf = tickets_df

    if "Ticket ตามสถานะ" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สถานะ")["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            charts.append((f"Ticket ตามสถานะ {start_date} ถึง {end_date}", tmp, "สถานะ", "จำนวน"))
        else:
            charts.append((f"Ticket ตามสถานะ {start_date} ถึง {end_date}", pd.DataFrame({"สถานะ":[], "จำนวน":[]}), "สถานะ", "จำนวน"))

    if "Ticket ตามสาขา" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สาขา", dropna=False)["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", tmp, "สาขา", "จำนวน"))
        else:
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", pd.DataFrame({"สาขา":[], "จำนวน":[]}), "สาขา", "จำนวน"))

    if len(charts)==0:
        st.info("โปรดเลือกกราฟที่ต้องการแสดงจากด้านบน")
    else:
        rows = (len(charts) + per_row - 1) // per_row
        idx = 0
        for r in range(rows):
            cols = st.columns(per_row)
            for c in range(per_row):
                if idx >= len(charts): break
                title, df, label_col, value_col = charts[idx]
                with cols[c]:
                    make_pie(df, label_col, value_col, top_n, title)
                idx += 1

    items_num = items.copy()
    if not items_num.empty:
        items_num["คงเหลือ"] = pd.to_numeric(items_num["คงเหลือ"], errors="coerce").fillna(0)
        items_num["จุดสั่งซื้อ"] = pd.to_numeric(items_num["จุดสั่งซื้อ"], errors="coerce").fillna(0)
        low_df2 = items_num[(items_num["ใช้งาน"].str.upper()=="Y") & (items_num["คงเหลือ"] <= items_num["จุดสั่งซื้อ"])]
    else:
        low_df2 = pd.DataFrame(columns=ITEMS_HEADERS)
    if not low_df2.empty:
        with st.expander("⚠️ อุปกรณ์ใกล้หมด (Reorder)", expanded=False):
            st.dataframe(low_df2[["รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]], use_container_width=True, height=240)
            pdf = df_to_pdf_bytes(low_df2[["รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]], title="อุปกรณ์ใกล้หมดสต็อก", subtitle=get_now_str())
            st.download_button("ดาวน์โหลด PDF รายการใกล้หมด", data=pdf, file_name="low_stock.pdf", mime="application/pdf")

    st.markdown("</div>", unsafe_allow_html=True)

def get_unit_options(items_df):
    opts = sorted([x for x in items_df["หน่วย"].dropna().astype(str).unique() if x.strip()!=""])
    if "ชิ้น" not in opts: opts = ["ชิ้น"] + opts
    return opts + ["พิมพ์เอง"]

def get_loc_options(items_df):
    opts = sorted([x for x in items_df["ที่เก็บ"].dropna().astype(str).unique() if x.strip()!=""])
    if "IT Room" not in opts: opts = ["IT Room"] + opts
    return opts + ["พิมพ์เอง"]


def generate_ticket_id() -> str:
    from datetime import datetime
    return "TCK-" + datetime.now(TZ).strftime("%Y%m%d-%H%M%S")

def page_tickets(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")

    # Load data
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    t_cats = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)

    # Filters
    st.markdown("### ตัวกรอง")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        statuses = ["ทั้งหมด","รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"]
        status_pick = st.selectbox("สถานะ", statuses, index=0, key="tk_status")
    with f2:
        br_opts = ["ทั้งหมด"] + ((branches["รหัสสาขา"] + " | " + branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
        branch_pick = st.selectbox("สาขา", br_opts, index=0, key="tk_branch")
    with f3:
        cat_opts = ["ทั้งหมด"] + ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else [])
        cat_pick = st.selectbox("หมวดหมู่ปัญหา", cat_opts, index=0, key="tk_cat")
    with f4:
        q = st.text_input("ค้นหา (ผู้แจ้ง/หมวด/รายละเอียด)", key="tk_query")

    # Date filter
    dcol1, dcol2 = st.columns(2)
    with dcol1:
        d1 = st.date_input("วันที่เริ่ม", key="tk_d1")
    with dcol2:
        d2 = st.date_input("วันที่สิ้นสุด", key="tk_d2")

    view = tickets.copy()
    if not view.empty:
        view["วันที่แจ้ง"] = pd.to_datetime(view["วันที่แจ้ง"], errors="coerce")
        view = view.dropna(subset=["วันที่แจ้ง"])
        if "tk_d1" in st.session_state and st.session_state.get("tk_d1"):
            view = view[view["วันที่แจ้ง"].dt.date >= st.session_state["tk_d1"]]
        if "tk_d2" in st.session_state and st.session_state.get("tk_d2"):
            view = view[view["วันที่แจ้ง"].dt.date <= st.session_state["tk_d2"]]
        if status_pick != "ทั้งหมด":
            view = view[view["สถานะ"] == status_pick]
        if branch_pick != "ทั้งหมด":
            view = view[view["สาขา"] == branch_pick]
        if "cat_pick" in locals() and cat_pick != "ทั้งหมด":
            view = view[view["หมวดหมู่"] == cat_pick]
        if q:
            mask = (view["ผู้แจ้ง"].str.contains(q, case=False, na=False) |
                    view["หมวดหมู่"].str.contains(q, case=False, na=False) |
                    view["รายละเอียด"].str.contains(q, case=False, na=False))
            view = view[mask]

    st.markdown("### รายการแจ้งปัญหา")
    st.dataframe(view.sort_values("วันที่แจ้ง", ascending=False) if not view.empty else view,
                 use_container_width=True, height=300)

    st.markdown("---")
    t_add, t_update = st.tabs(["➕ รับแจ้งใหม่","🔁 เปลี่ยนสถานะ/แก้ไข"])

    with t_add:
        with st.form("tk_new", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
                now_str = get_now_str()
                branch_sel = st.selectbox("สาขา", br_opts[1:] if len(br_opts)>1 else ["พิมพ์เอง"])
                if branch_sel == "พิมพ์เอง":
                    branch_sel = st.text_input("ระบุสาขา (พิมพ์เอง)", value="")
                reporter = st.text_input("ผู้แจ้ง", value="")
            with c2:
                tkc_opts = ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else []) + ["พิมพ์เอง"]
                pick_c = st.selectbox("หมวดหมู่ปัญหา", options=tkc_opts if tkc_opts else ["พิมพ์เอง"], key="tk_new_cat_sel")
                cate_custom = st.text_input("ระบุหมวด (ถ้าเลือกพิมพ์เอง)", value="" if pick_c!="พิมพ์เอง" else "", disabled=(pick_c!="พิมพ์เอง"))
                cate = pick_c if pick_c != "พิมพ์เอง" else cate_custom
                assignee = st.text_input("ผู้รับผิดชอบ (IT)", value=st.session_state.get("user",""))
            with c3:
                detail = st.text_area("รายละเอียด", height=100)
                note = st.text_input("หมายเหตุ", value="")
            s = st.form_submit_button("บันทึกการรับแจ้ง", use_container_width=True)
        if s:
            tid = generate_ticket_id()
            row = [tid, get_now_str(), branch_sel, reporter, cate, detail, "รับแจ้ง", assignee, get_now_str(), note]
            append_row(sh, SHEET_TICKETS, row)
            st.success(f"รับแจ้งเรียบร้อย (Ticket: {tid})")
            safe_rerun()

    with t_update:
        if tickets.empty:
            st.info("ยังไม่มีรายการในชีต Tickets")
        else:
            ids = tickets["TicketID"].tolist()
            pick = st.selectbox("เลือก Ticket", options=["-- เลือก --"]+ids, key="tk_pick")
            if pick != "-- เลือก --":
                row = tickets[tickets["TicketID"]==pick].iloc[0]
                with st.form("tk_edit", clear_on_submit=False):
                    c1,c2,c3 = st.columns(3)
                    with c1:
                        branch = st.text_input("สาขา", value=row["สาขา"])
                        reporter = st.text_input("ผู้แจ้ง", value=row["ผู้แจ้ง"])
                    with c2:
                        tkc_opts2 = ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else []) + ["พิมพ์เอง"]
                        default_index = tkc_opts2.index(row["หมวดหมู่"]) if (row["หมวดหมู่"] in tkc_opts2) else len(tkc_opts2)-1
                        pick_c2 = st.selectbox("หมวดหมู่", options=tkc_opts2, index=default_index, key="tk_edit_cat_sel")
                        cate_custom2 = st.text_input("ระบุหมวด (ถ้าเลือกพิมพ์เอง)", value="" if pick_c2!="พิมพ์เอง" else row["หมวดหมู่"], disabled=(pick_c2!="พิมพ์เอง"))
                        cate = pick_c2 if pick_c2 != "พิมพ์เอง" else cate_custom2
                        status = st.selectbox("สถานะ", ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"],
                                              index=["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"].index(row["สถานะ"]) if row["สถานะ"] in ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"] else 0)
                        assignee = st.text_input("ผู้รับผิดชอบ", value=row["ผู้รับผิดชอบ"])
                    with c3:
                        detail = st.text_area("รายละเอียด", value=row["รายละเอียด"], height=100)
                        note = st.text_input("หมายเหตุ", value=row["หมายเหตุ"])
                    colA, colB, colC = st.columns([2,1,1])
                    save = colA.form_submit_button("💾 บันทึกการแก้ไข", use_container_width=True)
                    done = colB.form_submit_button("✅ ปิดงาน (เสร็จ)", use_container_width=True)
                    delete = colC.form_submit_button("🗑️ ลบ Ticket", use_container_width=True)

                if save or done or delete:
                    df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
                    if delete:
                        df = df[df["TicketID"] != pick]
                        write_df(sh, SHEET_TICKETS, df)
                        st.success("ลบแล้ว"); safe_rerun()
                    else:
                        if done: status = "ดำเนินการเสร็จ"
                        df.loc[df["TicketID"]==pick, ["สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]] = \
                            [branch, reporter, cate, detail, status, assignee, get_now_str(), note]
                        write_df(sh, SHEET_TICKETS, df)
                        st.success("อัปเดตแล้ว"); safe_rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_categories_admin(sh):
    """UI จัดการหมวดหมู่ (Categories) สำหรับแอดมิน/สตาฟ
    - เพิ่ม/แก้ไข 1 รายการ
    - นำเข้าไฟล์หลายรายการ CSV/Excel
    - ค้นหา/แก้ไขแบบตาราง + ลบ (มีการป้องกันเมื่อถูกใช้งานใน Items)
    """
    st.markdown("#### 🏷️ จัดการหมวดหมู่")
    cats = read_df(sh, SHEET_CATS, CATS_HEADERS)

    tab1, tab2, tab3 = st.tabs(["✏️ เพิ่ม/แก้ไข 1 รายการ", "📥 นำเข้าไฟล์ (หลายรายการ)", "🔎 ค้นหา/แก้ไข/ลบ (ตาราง)"])

    # ---- TAB 1: Single add/update ----
    with tab1:
        c1, c2 = st.columns([1,2])
        with c1:
            code_in = st.text_input("รหัสหมวด", placeholder="เช่น PRT, KBD").upper().strip()
        with c2:
            name_in = st.text_input("ชื่อหมวด", placeholder="เช่น หมึกพิมพ์, คีย์บอร์ด").strip()
        if st.button("💾 บันทึก/แก้ไข 1 รายการ", use_container_width=True, key="cat_save_single"):
            if not code_in or not name_in:
                st.warning("กรุณากรอกรหัสและชื่อหมวดให้ครบ")
            else:
                df = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if (df["รหัสหมวด"] == code_in).any():
                    df.loc[df["รหัสหมวด"] == code_in, "ชื่อหมวด"] = name_in
                    msg = "อัปเดต"
                else:
                    df = pd.concat([df, pd.DataFrame([[code_in, name_in]], columns=CATS_HEADERS)], ignore_index=True)
                    msg = "เพิ่มใหม่"
                write_df(sh, SHEET_CATS, df); st.success(f"{msg}เรียบร้อย: {code_in} — {name_in}"); safe_rerun()

    # ---- TAB 2: Import many ----
    with tab2:
        with st.expander("วิธีใช้งาน/เทมเพลต (คลิกเพื่อดู)", expanded=False):
            st.markdown("""\
- รองรับไฟล์ .csv หรือ .xlsx ที่มีคอลัมน์ **รหัสหมวด, ชื่อหมวด**
- ระบบจะ **อัปเดตชื่อหมวด** หากพบรหัสซ้ำ และ **เพิ่มใหม่** เมื่อไม่พบรหัสเดิม
- ไม่ลบรายการเดิมโดยอัตโนมัติ เว้นแต่เลือกโหมด 'แทนที่ทั้งชีต'
            """)
            tpl = """รหัสหมวด,ชื่อหมวด
PRT,หมึกพิมพ์
KBD,คีย์บอร์ด
"""
            st.download_button("ดาวน์โหลดเทมเพลต (CSV)", data=tpl.encode("utf-8-sig"),
                               file_name="template_categories.csv", mime="text/csv")

        cA, cB = st.columns([2,1])
        with cA:
            up = st.file_uploader("เลือกไฟล์ (.csv, .xlsx)", type=["csv","xlsx","xls"], key="cat_uploader_stocktab")
        with cB:
            replace_all = st.checkbox("แทนที่ทั้งชีต (ล้างและใส่ใหม่)", value=False)

        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    df_up = pd.read_csv(up, dtype=str)
                else:
                    df_up = pd.read_excel(up, dtype=str)
                df_up = df_up.fillna("").applymap(lambda x: str(x).strip())
            except Exception as e:
                st.error(f"อ่านไฟล์ไม่สำเร็จ: {e}"); return

            rename_map = {"รหัสหมวดหมู่":"รหัสหมวด","ชื่อหมวดหมู่":"ชื่อหมวด","code":"รหัสหมวด","name":"ชื่อหมวด","category_code":"รหัสหมวด","category_name":"ชื่อหมวด"}
            df_up.columns = [rename_map.get(c.strip(), c.strip()) for c in df_up.columns]
            missing = [c for c in ["รหัสหมวด","ชื่อหมวด"] if c not in df_up.columns]
            if missing:
                st.error(f"ไฟล์ขาดคอลัมน์ที่บังคับ: {', '.join(missing)}"); return

            df_up["รหัสหมวด"] = df_up["รหัสหมวด"].str.upper()
            df_up = df_up[df_up["รหัสหมวด"]!=""]
            df_up = df_up.drop_duplicates(subset=["รหัสหมวด"], keep="last")

            st.success(f"พรีวิว {len(df_up):,} รายการ")
            st.dataframe(df_up, use_container_width=True, height=240)

            if st.button("🚀 ดำเนินการนำเข้า/อัปเดต", use_container_width=True, key="cat_do_import_stocktab"):
                base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if replace_all:
                    # ก่อนแทนที่ทั้งชีต ให้ตรวจว่ามีการอ้างอิงใน Items หรือไม่
                    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                    used_cats = set(items["หมวดหมู่"].tolist()) if not items.empty else set()
                    new_cats = set(df_up["รหัสหมวด"].tolist())
                    if used_cats - new_cats:
                        st.error("ไม่สามารถแทนที่ทั้งชีตได้: มีรหัสหมวดที่ถูกใช้งานใน Items แต่ไม่อยู่ในไฟล์ใหม่นี้"); return
                    write_df(sh, SHEET_CATS, df_up[CATS_HEADERS]); st.success(f"แทนที่ทั้งชีตสำเร็จ • บันทึก {len(df_up):,} รายการ"); safe_rerun()
                else:
                    added, updated = 0, 0
                    for _, r in df_up.iterrows():
                        cd, nm = str(r["รหัสหมวด"]).strip().upper(), str(r["ชื่อหมวด"]).strip()
                        if not cd or not nm: 
                            continue
                        if (base["รหัสหมวด"] == cd).any():
                            base.loc[base["รหัสหมวด"] == cd, "ชื่อหมวด"] = nm; updated += 1
                        else:
                            base = pd.concat([base, pd.DataFrame([[cd, nm]], columns=CATS_HEADERS)], ignore_index=True); added += 1
                    write_df(sh, SHEET_CATS, base); st.success(f"นำเข้าสำเร็จ • เพิ่มใหม่ {added:,} • อัปเดต {updated:,}"); safe_rerun()

    # ---- TAB 3: Search / inline edit / delete with protection ----
    with tab3:
        c1, c2 = st.columns([2,1])
        with c1:
            q = st.text_input("ค้นหา (รหัส/ชื่อ)", key="cat_search_stocktab")
        with c2:
            st.caption("แก้ไขค่าในตารางได้โดยตรง และสามารถลบหมวดที่ไม่ได้ใช้งาน")

        view = cats.copy()
        if not view.empty and q:
            mask = view["รหัสหมวด"].str.contains(q, case=False, na=False) | view["ชื่อหมวด"].str.contains(q, case=False, na=False)
            view = view[mask]

        edited = st.data_editor(
            view.sort_values("รหัสหมวด"),
            use_container_width=True,
            height=360,
            disabled=["รหัสหมวด"],
            key="cat_editor_stocktab"
        )

        cL, cM, cR = st.columns([1,1,1])
        with cL:
            if st.button("💾 บันทึกการแก้ไข", use_container_width=True, key="cat_save_table_stocktab"):
                base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                for _, r in edited.iterrows():
                    base.loc[base["รหัสหมวด"] == str(r["รหัสหมวด"]).strip().upper(), "ชื่อหมวด"] = str(r["ชื่อหมวด"]).strip()
                write_df(sh, SHEET_CATS, base); st.success("บันทึกการแก้ไขเรียบร้อย"); safe_rerun()
        with cM:
            # Quick add
            with st.popover("➕ เพิ่มใหม่"):
                q_code = st.text_input("รหัสหมวด (ใหม่)", key="cat_quick_code_stocktab").upper().strip()
                q_name = st.text_input("ชื่อหมวด", key="cat_quick_name_stocktab").strip()
                if st.button("เพิ่ม", key="cat_quick_add_stocktab"):
                    if not q_code or not q_name:
                        st.warning("กรุณากรอกให้ครบ")
                    else:
                        base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                        if (base["รหัสหมวด"] == q_code).any():
                            st.error("รหัสนี้มีอยู่แล้ว"); st.stop()
                        base = pd.concat([base, pd.DataFrame([[q_code, q_name]], columns=CATS_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_CATS, base); st.success("เพิ่มใหม่แล้ว"); safe_rerun()
        with cR:
            # Delete selection with protection
            base = read_df(sh, SHEET_CATS, CATS_HEADERS)
            opts = (base["รหัสหมวด"]+" | "+base["ชื่อหมวด"]).tolist() if not base.empty else []
            del_sel = st.multiselect("เลือกรายการที่จะลบ (เฉพาะหมวดที่ไม่ถูกใช้งาน)", options=opts, key="cat_del_sel_stocktab")
            if st.button("🗑️ ลบที่เลือก", type="secondary", use_container_width=True, key="cat_do_delete_stocktab"):
                items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                used = set(items["หมวดหมู่"].tolist()) if not items.empty else set()
                to_del_codes = {x.split(" | ")[0] for x in del_sel}
                blocked = sorted(list(used.intersection(to_del_codes)))
                if blocked:
                    st.error("ไม่สามารถลบได้ เพราะหมวดต่อไปนี้ถูกใช้งานใน Items: " + ", ".join(blocked))
                else:
                    base = base[~base["รหัสหมวด"].isin(list(to_del_codes))]
                    write_df(sh, SHEET_CATS, base); st.success(f"ลบแล้ว {len(to_del_codes):,} รายการ"); safe_rerun()
def page_stock(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("📦 คลังอุปกรณ์")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS); cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view_df = items.copy()
    if q and not items.empty:
        mask = items["รหัส"].str.contains(q, case=False, na=False) | items["ชื่ออุปกรณ์"].str.contains(q, case=False, na=False) | items["หมวดหมู่"].str.contains(q, case=False, na=False)
        view_df = items[mask]
    st.dataframe(view_df, use_container_width=True, height=320)

    unit_opts = get_unit_options(items)
    loc_opts  = get_loc_options(items)

    if st.session_state.get("role") in ("admin","staff"):
        t_add, t_edit, t_cat = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ (เลือกรายการเดิม)","🏷️ หมวดหมู่"])

        with t_cat:
            render_categories_admin(sh)

        with t_add:
            with st.form("item_add", clear_on_submit=True):
                c1,c2,c3 = st.columns(3)
                with c1:
                    if cats.empty: st.info("ยังไม่มีหมวดหมู่ในชีต Categories (ใช้เมนู นำเข้า/แก้ไข หมวดหมู่ เพื่อเพิ่ม)"); cat_opt=""
                    else:
                        opts = (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist(); selected = st.selectbox("หมวดหมู่", options=opts)
                        cat_opt = selected.split(" | ")[0]
                    name = st.text_input("ชื่ออุปกรณ์")
                with c2:
                    sel_unit = st.selectbox("หน่วย (เลือกจากรายการ)", options=unit_opts, index=0)
                    unit = st.text_input("ระบุหน่วยใหม่", value="", disabled=(sel_unit!="พิมพ์เอง"))
                    if sel_unit!="พิมพ์เอง": unit = sel_unit
                    qty = st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                    rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
                with c3:
                    sel_loc = st.selectbox("ที่เก็บ (เลือกจากรายการ)", options=loc_opts, index=0)
                    loc = st.text_input("ระบุที่เก็บใหม่", value="", disabled=(sel_loc!="พิมพ์เอง"))
                    if sel_loc!="พิมพ์เอง": loc = sel_loc
                    active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                    auto_code = st.checkbox("สร้างรหัสอัตโนมัติ", value=True)
                    code = st.text_input("รหัสอุปกรณ์ (ถ้าไม่ออโต้)", disabled=auto_code)
                    s_add = st.form_submit_button("บันทึก/อัปเดต", use_container_width=True)
            if s_add:
                if (auto_code and not cat_opt) or (not auto_code and code.strip()==""): st.error("กรุณาเลือกหมวด/ระบุรหัส")
                else:
                    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS); gen_code = generate_item_code(sh, cat_opt) if auto_code else code.strip().upper()
                    if (items["รหัส"]==gen_code).any():
                        items.loc[items["รหัส"]==gen_code, ITEMS_HEADERS] = [gen_code, cat_opt, name, unit, qty, rop, loc, active]
                    else:
                        items = pd.concat([items, pd.DataFrame([[gen_code, cat_opt, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_ITEMS, items); st.success(f"บันทึกเรียบร้อย (รหัส: {gen_code})"); safe_rerun()

        with t_edit:
            st.caption("เลือก 'รหัสอุปกรณ์' เพื่อโหลดขึ้นมาปรับแก้ หรือกดลบ")
            if items.empty:
                st.info("ยังไม่มีรายการให้แก้ไข")
            else:
                codes = items["รหัส"].tolist()
                pick = st.selectbox("เลือกรหัสอุปกรณ์", options=["-- เลือก --"]+codes)
                if pick != "-- เลือก --":
                    row = items[items["รหัส"]==pick].iloc[0]
                    unit_opts_edit = unit_opts[:-1]
                    if row["หน่วย"] not in unit_opts_edit and str(row["หน่วย"]).strip()!="":
                        unit_opts_edit = [row["หน่วย"]] + unit_opts_edit
                    unit_opts_edit = unit_opts_edit + ["พิมพ์เอง"]
                    loc_opts_edit = loc_opts[:-1]
                    if row["ที่เก็บ"] not in loc_opts_edit and str(row["ที่เก็บ"]).strip()!="":
                        loc_opts_edit = [row["ที่เก็บ"]] + loc_opts_edit
                    loc_opts_edit = loc_opts_edit + ["พิมพ์เอง"]

                    with st.form("item_edit", clear_on_submit=False):
                        c1,c2,c3 = st.columns(3)
                        with c1:
                            name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
                            sel_unit = st.selectbox("หน่วย (เลือกจากรายการ)", options=unit_opts_edit, index=0)
                            unit = st.text_input("ระบุหน่วยใหม่", value="", disabled=(sel_unit!="พิมพ์เอง"))
                            if sel_unit!="พิมพ์เอง": unit = sel_unit
                        with c2:
                            qty = st.number_input("คงเหลือ", min_value=0, value=int(float(row["คงเหลือ"]) if str(row["คงเหลือ"]).strip()!="" else 0), step=1)
                            rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=int(float(row["จุดสั่งซื้อ"]) if str(row["จุดสั่งซื้อ"]).strip()!="" else 0), step=1)
                        with c3:
                            sel_loc = st.selectbox("ที่เก็บ (เลือกจากรายการ)", options=loc_opts_edit, index=0)
                            loc = st.text_input("ระบุที่เก็บใหม่", value="", disabled=(sel_loc!="พิมพ์เอง"))
                            if sel_loc!="พิมพ์เอง": loc = sel_loc
                            active = st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(row["ใช้งาน"]).upper()=="Y" else 1)
                        col_save, col_delete = st.columns([3,1])
                        s_save = col_save.form_submit_button("💾 บันทึกการแก้ไข", use_container_width=True)
                        s_del  = col_delete.form_submit_button("🗑️ ลบรายการ", use_container_width=True)
                    if s_save:
                        items.loc[items["รหัส"]==pick, ITEMS_HEADERS] = [pick, row["หมวดหมู่"], name, unit, qty, rop, loc, "Y" if active=="Y" else "N"]
                        write_df(sh, SHEET_ITEMS, items); st.success("อัปเดตแล้ว"); safe_rerun()
                    if s_del:
                        items = items[items["รหัส"]!=pick]; write_df(sh, SHEET_ITEMS, items); st.success(f"ลบ {pick} แล้ว"); safe_rerun()

def group_period(df, period="ME"):
    dfx = df.copy(); dfx["วันเวลา"] = pd.to_datetime(dfx["วันเวลา"], errors='coerce'); dfx = dfx.dropna(subset=["วันเวลา"])
    return dfx.groupby([pd.Grouper(key="วันเวลา", freq=period), "ประเภท", "ชื่ออุปกรณ์"])['จำนวน'].sum().reset_index()

def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("🧾 เบิก/รับเข้า")
    if st.session_state.get("role") not in ("admin","staff"): st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกรายการได้"); st.markdown("</div>", unsafe_allow_html=True); return
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS); branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty: st.warning("ยังไม่มีรายการอุปกรณ์ในคลัง"); st.markdown("</div>", unsafe_allow_html=True); return
    t1,t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])

    with t1:
        with st.form("issue", clear_on_submit=True):
            c1,c2 = st.columns([2,1])
            with c1: item = st.selectbox("เลือกอุปกรณ์", options=items["รหัส"]+" | "+items["ชื่ออุปกรณ์"])
            with c2: qty = st.number_input("จำนวนที่เบิก", min_value=1, value=1, step=1)
            if branches.empty:
                branch_sel = st.text_input("สาขา/หน่วยงานผู้ขอ (ยังไม่มีรายการสาขา ให้พิมพ์เองหรือ นำเข้า/แก้ไข หมวดหมู่)")
            else:
                br_opts = (branches["รหัสสาขา"] + " | " + branches["ชื่อสาขา"]).tolist() + ["พิมพ์เอง"]
                br_pick = st.selectbox("สาขา/หน่วยงานผู้ขอ", options=br_opts)
                branch_sel = st.text_input("ถ้าพิมพ์เอง ใส่ที่นี่", value="" if br_pick!="พิมพ์เอง" else "")
                if br_pick!="พิมพ์เอง": branch_sel = br_pick
            note = st.text_input("หมายเหตุ", placeholder="เช่น งานซ่อมเครื่องพิมพ์")
            st.markdown("**วัน-เวลาการเบิก**")
            manual = st.checkbox("กำหนดวันเวลาเอง", value=False)
            if manual:
                d = st.date_input("วันที่", value=datetime.now(TZ).date(), key="out_d")
                t = st.time_input("เวลา", value=datetime.now(TZ).time().replace(microsecond=0), key="out_t")
                ts_str = fmt_dt(combine_date_time(d, t))
            else:
                ts_str = None
            s = st.form_submit_button("บันทึกการเบิก", use_container_width=True)
        if s:
            code = item.split(" | ")[0]
            ok = adjust_stock(sh, code, -qty, st.session_state.get("user","unknown"), branch_sel, note, "OUT", ts_str=ts_str)
            if ok: st.success("บันทึกแล้ว"); safe_rerun()

    with t2:
        with st.form("recv", clear_on_submit=True):
            c1,c2 = st.columns([2,1])
            with c1: item = st.selectbox("เลือกอุปกรณ์", options=items["รหัส"]+" | "+items["ชื่ออุปกรณ์"], key="recv_item")
            with c2: qty = st.number_input("จำนวนที่รับเข้า", min_value=1, value=1, step=1, key="recv_qty")
            branch = st.text_input("แหล่งที่มา/เลข PO", key="recv_branch"); note = st.text_input("หมายเหตุ", placeholder="เช่น ซื้อเข้า-เติมสต็อก", key="recv_note")
            st.markdown("**วัน-เวลารับเข้า**")
            manual_in = st.checkbox("กำหนดวันเวลาเอง ", value=False, key="in_manual")
            if manual_in:
                d = st.date_input("วันที่", value=datetime.now(TZ).date(), key="in_d")
                t = st.time_input("เวลา", value=datetime.now(TZ).time().replace(microsecond=0), key="in_t")
                ts_str = fmt_dt(combine_date_time(d, t))
            else:
                ts_str = None
            s = st.form_submit_button("บันทึกรับเข้า", use_container_width=True)
        if s:
            code = item.split(" | ")[0]
            ok = adjust_stock(sh, code, qty, st.session_state.get("user","unknown"), branch, note, "IN", ts_str=ts_str)
            if ok: st.success("บันทึกรับเข้าแล้ว"); safe_rerun()





def page_reports(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📑 รายงาน / ประวัติ")

    txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)

    # ---------- Quick range state ----------
    if "report_d1" not in st.session_state or "report_d2" not in st.session_state:
        today = datetime.now(TZ).date()
        st.session_state["report_d1"] = today - timedelta(days=30)
        st.session_state["report_d2"] = today

    def _set_range(days=None, today=False, this_month=False, this_year=False):
        nowd = datetime.now(TZ).date()
        if today:
            st.session_state["report_d1"] = nowd
            st.session_state["report_d2"] = nowd
        elif this_month:
            st.session_state["report_d1"] = nowd.replace(day=1)
            st.session_state["report_d2"] = nowd
        elif this_year:
            st.session_state["report_d1"] = date(nowd.year, 1, 1)
            st.session_state["report_d2"] = nowd
        elif days is not None:
            st.session_state["report_d1"] = nowd - timedelta(days=days-1)
            st.session_state["report_d2"] = nowd

    st.markdown("### ⏱️ เลือกช่วงวันที่อย่างรวดเร็ว")
    bcols = st.columns(6)
    with bcols[0]:
        st.button("วันนี้", on_click=_set_range, kwargs=dict(today=True), key="btn_today_r")
    with bcols[1]:
        st.button("7 วันล่าสุด", on_click=_set_range, kwargs=dict(days=7), key="btn_7d_r")
    with bcols[2]:
        st.button("30 วันล่าสุด", on_click=_set_range, kwargs=dict(days=30), key="btn_30d_r")
    with bcols[3]:
        st.button("90 วันล่าสุด", on_click=_set_range, kwargs=dict(days=90), key="btn_90d_r")
    with bcols[4]:
        st.button("เดือนนี้", on_click=_set_range, kwargs=dict(this_month=True), key="btn_month_r")
    with bcols[5]:
        st.button("ปีนี้", on_click=_set_range, kwargs=dict(this_year=True), key="btn_year_r")

    with st.expander("กำหนดช่วงวันที่เอง (เลือกแล้วกด 'ใช้ช่วงนี้')", expanded=False):
        d1m = st.date_input("วันที่เริ่ม (กำหนดเอง)", value=st.session_state["report_d1"], key="report_manual_d1_r")
        d2m = st.date_input("วันที่สิ้นสุด (กำหนดเอง)", value=st.session_state["report_d2"], key="report_manual_d2_r")
        st.button("ใช้ช่วงนี้", on_click=lambda: (st.session_state.__setitem__("report_d1", d1m),
                                                st.session_state.__setitem__("report_d2", d2m)),
                  key="btn_apply_manual_r")

    q = st.text_input("ค้นหา (ชื่อ/รหัส/สาขา/เรื่อง)", key="report_query_r")

    d1 = st.session_state["report_d1"]
    d2 = st.session_state["report_d2"]
    st.caption(f"ช่วงที่เลือก: **{d1} → {d2}**")

    # ---------- Transactions (filter for existing tabs) ----------
    if not txns.empty:
        df_f = txns.copy()
        df_f["วันเวลา"] = pd.to_datetime(df_f["วันเวลา"], errors="coerce")
        df_f = df_f.dropna(subset=["วันเวลา"])
        df_f = df_f[(df_f["วันเวลา"].dt.date >= d1) & (df_f["วันเวลา"].dt.date <= d2)]
        if q:
            mask_q = (
                df_f["ชื่ออุปกรณ์"].str.contains(q, case=False, na=False) |
                df_f["รหัส"].str.contains(q, case=False, na=False) |
                df_f["สาขา"].str.contains(q, case=False, na=False)
            )
            df_f = df_f[mask_q]
    else:
        df_f = pd.DataFrame(columns=TXNS_HEADERS)

    # ---------- Tickets (filtered by วันที่แจ้ง) ----------
    if not tickets.empty:
        tdf = tickets.copy()
        tdf["วันที่แจ้ง"] = pd.to_datetime(tdf["วันที่แจ้ง"], errors="coerce")
        tdf = tdf.dropna(subset=["วันที่แจ้ง"])
        tdf = tdf[(tdf["วันที่แจ้ง"].dt.date >= d1) & (tdf["วันที่แจ้ง"].dt.date <= d2)]
        if q:
            mask_t = (
                (tdf["รายละเอียด"].astype(str).str.contains(q, case=False, na=False)) |
                (tdf["สาขา"].astype(str).str.contains(q, case=False, na=False)) |
                (tdf["ผู้แจ้ง"].astype(str).str.contains(q, case=False, na=False))
            )
            # หากมีคอลัมน์ "เรื่อง" ก็ให้ค้นหาด้วย
            if "เรื่อง" in tdf.columns:
                mask_t = mask_t | tdf["เรื่อง"].astype(str).str.contains(q, case=False, na=False)
            tdf = tdf[mask_t]
        # สร้างคอลัมน์ "เรื่อง" อัตโนมัติจากบรรทัดแรกของรายละเอียด ถ้าไม่มี
        if "เรื่อง" not in tdf.columns:
            def _derive_subject(x):
                s = str(x or "").strip().splitlines()[0]
                return s[:60] if s else "ไม่ระบุเรื่อง"
            tdf["เรื่อง"] = tdf["รายละเอียด"].apply(_derive_subject)
    else:
        tdf = pd.DataFrame(columns=TICKETS_HEADERS + ["เรื่อง"])

    # ---------- Tabs ----------
    tOut, tTickets, tW, tM, tY = st.tabs(["รายละเอียดการเบิก (OUT)", "ประวัติการแจ้งปัญหา", "รายสัปดาห์", "รายเดือน", "รายปี"])

    # --- OUT detail ---
    with tOut:
        out_df = df_f[df_f["ประเภท"] == "OUT"].copy().sort_values("วันเวลา", ascending=False)
        cols = [c for c in ["วันเวลา", "ชื่ออุปกรณ์", "จำนวน", "สาขา", "ผู้ดำเนินการ", "หมายเหตุ", "รหัส"] if c in out_df.columns]
        st.dataframe(out_df[cols], use_container_width=True, height=320)
        pdf = df_to_pdf_bytes(
            out_df[cols].rename(columns={"วันเวลา":"วันที่-เวลา","ชื่ออุปกรณ์":"อุปกรณ์","จำนวน":"จำนวนที่เบิก","สาขา":"ปลายทาง"}),
            title="รายละเอียดการเบิก (OUT)", subtitle=f"ช่วง {d1} ถึง {d2}"
        )
        st.download_button("ดาวน์โหลด PDF รายละเอียดการเบิก", data=pdf, file_name="issue_detail_out.pdf", mime="application/pdf", key="dl_pdf_out_r")

    # --- Tickets detail + summary ---
    with tTickets:
        st.markdown("#### ตารางรายการแจ้งปัญหา")
        show_cols = [c for c in ["วันที่แจ้ง","เรื่อง","รายละเอียด","สาขา","ผู้แจ้ง","สถานะ","ผู้รับผิดชอบ","หมายเหตุ","TicketID"] if c in tdf.columns]
        tdf_sorted = tdf.sort_values("วันที่แจ้ง", ascending=False)
        st.dataframe(tdf_sorted[show_cols], use_container_width=True, height=320)

        st.markdown("#### สรุปจำนวนครั้งตาม 'เรื่อง' และ 'สาขา'")
        if not tdf.empty:
            agg = tdf.groupby(["เรื่อง","สาขา"])["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวนครั้ง"})
        else:
            agg = pd.DataFrame(columns=["เรื่อง","สาขา","จำนวนครั้ง"])
        st.dataframe(agg.sort_values(["จำนวนครั้ง","เรื่อง"], ascending=[False, True]), use_container_width=True, height=260)

        pdf_t = df_to_pdf_bytes(agg.rename(columns={"เรื่อง":"ชื่อเรื่อง"}), title="สรุปการแจ้งปัญหา: เรื่อง × สาขา", subtitle=f"ช่วง {d1} ถึง {d2}")
        st.download_button("ดาวน์โหลด PDF สรุปการแจ้งปัญหา", data=pdf_t, file_name="ticket_summary_subject_branch.pdf", mime="application/pdf", key="dl_pdf_ticket_r")

    # --- summaries by period (same as before) ---
    def group_period(df, period="ME"):
        dfx = df.copy()
        dfx["วันเวลา"] = pd.to_datetime(dfx["วันเวลา"], errors='coerce')
        dfx = dfx.dropna(subset=["วันเวลา"])
        return dfx.groupby([pd.Grouper(key="วันเวลา", freq=period), "ประเภท", "ชื่ออุปกรณ์"])["จำนวน"].sum().reset_index()

    with tW:
        g = group_period(df_f, "W")
        st.dataframe(g, use_container_width=True, height=220)
        st.download_button("ดาวน์โหลด PDF รายสัปดาห์", data=df_to_pdf_bytes(g, "สรุปรายสัปดาห์", f"ช่วง {d1} ถึง {d2}"), file_name="weekly_report.pdf", mime="application/pdf", key="dl_pdf_w_r")

    with tM:
        g = group_period(df_f, "ME")
        st.dataframe(g, use_container_width=True, height=220)
        st.download_button("ดาวน์โหลด PDF รายเดือน", data=df_to_pdf_bytes(g, "สรุปรายเดือน", f"ช่วง {d1} ถึง {d2}"), file_name="monthly_report.pdf", mime="application/pdf", key="dl_pdf_m_r")

    with tY:
        g = group_period(df_f, "YE")
        st.dataframe(g, use_container_width=True, height=220)
        st.download_button("ดาวน์โหลด PDF รายปี", data=df_to_pdf_bytes(g, "สรุปรายปี", f"ช่วง {d1} ถึง {d2}"), file_name="yearly_report.pdf", mime="application/pdf", key="dl_pdf_y_r")

    st.markdown("</div>", unsafe_allow_html=True)

def page_users_admin(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("👥 ผู้ใช้ & สิทธิ์ (Admin)")
    if st.session_state.get("role") != "admin": st.info("เฉพาะผู้ดูแลระบบ (admin)"); st.markdown("</div>", unsafe_allow_html=True); return
    users = read_df(sh, SHEET_USERS, USERS_HEADERS); st.dataframe(users, use_container_width=True, height=260)
    st.markdown("### เพิ่ม/แก้ไข ผู้ใช้")
    with st.form("user_form", clear_on_submit=True):
        c1,c2,c3 = st.columns(3)
        with c1: uname = st.text_input("Username"); dname = st.text_input("Display Name")
        with c2: role = st.selectbox("Role", ["admin","staff","viewer"], index=1); active = st.selectbox("Active", ["Y","N"], index=0)
        with c3: pwd = st.text_input("ตั้ง/รีเซ็ตรหัสผ่าน", type="password")
        s = st.form_submit_button("บันทึกผู้ใช้", use_container_width=True)
    if s:
        if uname.strip()=="": st.error("กรุณาใส่ Username")
        else:
            if pwd.strip(): hash_str = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            else:
                mask = users["Username"]==uname
                hash_str = users.loc[mask,"PasswordHash"].iloc[0] if mask.any() else bcrypt.hashpw("password123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            if (users["Username"]==uname).any():
                users.loc[users["Username"]==uname, USERS_HEADERS] = [uname, dname, role, hash_str, active]
            else:
                users = pd.concat([users, pd.DataFrame([[uname, dname, role, hash_str, active]], columns=USERS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_USERS, users); st.success("บันทึกแล้ว"); safe_rerun()

def is_test_text(s: str) -> bool:
    s = str(s).lower()
    return ("test" in s) or ("ทดสอบ" in s)

def page_settings():
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("⚙️ Settings"); st.caption("ตรวจสอบว่าได้แชร์ Google Sheet ให้ service account แล้ว")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ/ตรวจสอบชีตที่จำเป็น", use_container_width=True):
        try:
            sh = open_sheet_by_url(url); ensure_sheets_exist(sh); st.success("เชื่อมต่อสำเร็จ พร้อมใช้งาน")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")

    st.markdown("---")
    st.markdown("### 🧹 ล้างข้อมูลทดสอบ (เฉพาะ Admin)")
    role = st.session_state.get("role","viewer")
    if role != "admin":
        st.info("ต้องเป็นผู้ดูแลระบบ (admin) จึงจะใช้งานได้")
        st.markdown("</div>", unsafe_allow_html=True); return

    st.caption("เงื่อนไขข้อมูล 'ทดสอบ' ที่จะถูกลบ: Transactions ที่มีคำว่า **test/ทดสอบ** ในคอลัมน์ หมายเหตุ/สาขา/ชื่ออุปกรณ์/รหัส และ (ตัวเลือก) Items ที่ชื่อมี test/ทดสอบ หรือรหัสขึ้นต้น **TEST-/TST-**")
    include_items = st.checkbox("รวมการลบ Items ที่เป็นข้อมูลทดสอบ", value=False)
    with st.form("clear_test_confirm"):
        pwd = st.text_input("กรอกรหัสผ่านของผู้ใช้ที่กำลังล็อกอิน", type="password")
        confirm = st.text_input("พิมพ์คำว่า CLEAR เพื่อยืนยัน", placeholder="CLEAR")
        submitted = st.form_submit_button("ล้างข้อมูลทดสอบ", use_container_width=True)
    if submitted:
        try:
            sh = open_sheet_by_url(st.session_state["sheet_url"])
        except Exception as e:
            st.error(f"เชื่อมต่อชีตไม่สำเร็จ: {e}")
            st.markdown("</div>", unsafe_allow_html=True); return

        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
        row = users[users["Username"]==st.session_state.get("user")]
        if row.empty:
            st.error("ไม่พบผู้ใช้ปัจจุบันในชีต Users"); st.markdown("</div>", unsafe_allow_html=True); return
        if not pwd:
            st.error("กรุณากรอกรหัสผ่าน"); st.markdown("</div>", unsafe_allow_html=True); return
        try:
            if not bcrypt.checkpw(pwd.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8")):
                st.error("รหัสผ่านไม่ถูกต้อง"); st.markdown("</div>", unsafe_allow_html=True); return
        except Exception:
            st.error("ไม่สามารถตรวจสอบรหัสผ่านได้"); st.markdown("</div>", unsafe_allow_html=True); return
        if confirm.strip().upper() != "CLEAR":
            st.error("กรุณาพิมพ์คำว่า CLEAR ให้ถูกต้องเพื่อยืนยัน"); st.markdown("</div>", unsafe_allow_html=True); return

        removed_txn = 0; removed_items = 0
        tx = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
        if not tx.empty:
            mask = (
                tx["หมายเหตุ"].apply(is_test_text) |
                tx["สาขา"].apply(is_test_text) |
                tx["ชื่ออุปกรณ์"].apply(is_test_text) |
                tx["รหัส"].apply(is_test_text)
            )
            removed_txn = int(mask.sum())
            tx = tx[~mask]
            write_df(sh, SHEET_TXNS, tx)

        if include_items:
            it = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
            if not it.empty:
                mask_items = (
                    it["ชื่ออุปกรณ์"].apply(is_test_text) |
                    it["รหัส"].str.upper().str.startswith("TEST-", na=False) |
                    it["รหัส"].str.upper().str.startswith("TST-", na=False)
                )
                removed_items = int(mask_items.sum())
                it = it[~mask_items]
                write_df(sh, SHEET_ITEMS, it)

        st.success(f"ลบข้อมูลทดสอบเรียบร้อย • Transactions: {removed_txn} แถว • Items: {removed_items} แถว")
    st.markdown("</div>", unsafe_allow_html=True)
# ---------- นำเข้า/แก้ไข หมวดหมู่ Page (Categories / Branches / Items) ----------
def _read_upload_df(file):
    if file is None: return None, "ยังไม่ได้เลือกไฟล์"
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(file, dtype=str).fillna("")
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            df = pd.read_excel(file, dtype=str).fillna("")
        else:
            return None, "รองรับเฉพาะ .csv หรือ .xlsx"
        # strip spaces
        df = df.applymap(lambda x: str(x).strip())
        return df, None
    except Exception as e:
        return None, f"อ่านไฟล์ไม่สำเร็จ: {e}"




def page_import(sh):
    """
    นำเข้า/แก้ไข หมวดหมู่
    - จัดเป็น 3 แท็บ: (1) เพิ่ม/แก้ไข 1 รายการ, (2) นำเข้าไฟล์หลายรายการ, (3) ค้นหา/แก้ไขแบบตาราง
    - ใช้ชีต SHEET_CATS และคอลัมน์ CATS_HEADERS = ["รหัสหมวด","ชื่อหมวด"]
    """
    st.subheader("นำเข้า/แก้ไข หมวดหมู่")
    cats = read_df(sh, SHEET_CATS, CATS_HEADERS)

    tab1, tab2, tab3 = st.tabs(["✏️ เพิ่ม/แก้ไข 1 รายการ", "📥 นำเข้าไฟล์ (หลายรายการ)", "🔎 ค้นหา/แก้ไขแบบตาราง"])

    # ---------------- TAB 1: Single add/update ----------------
    with tab1:
        st.caption("ใส่รหัสและชื่อหมวด แล้วกดบันทึก ระบบจะ 'อัปเดต' ถ้ารหัสมีอยู่แล้ว หรือ 'เพิ่มใหม่' ถ้าไม่พบ")
        c1, c2 = st.columns([1,2])
        with c1:
            code = st.text_input("รหัสหมวด", placeholder="เช่น PRT, KBD").upper().strip()
        with c2:
            name = st.text_input("ชื่อหมวด", placeholder="เช่น หมึกพิมพ์, คีย์บอร์ด").strip()
        if st.button("💾 บันทึก/แก้ไข 1 รายการ", use_container_width=True, key="save_single"):
            if not code or not name:
                st.warning("กรุณากรอกรหัสและชื่อหมวดให้ครบ")
            else:
                df = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if (df["รหัสหมวด"] == code).any():
                    df.loc[df["รหัสหมวด"] == code, "ชื่อหมวด"] = name
                    msg = "อัปเดต"
                else:
                    df = pd.concat([df, pd.DataFrame([[code, name]], columns=CATS_HEADERS)], ignore_index=True)
                    msg = "เพิ่มใหม่"
                write_df(sh, SHEET_CATS, df)
                st.success(f"{msg}เรียบร้อย: {code} — {name}")
                safe_rerun()

    # ---------------- TAB 2: Import many ----------------
    with tab2:
        with st.expander("วิธีใช้งาน/เทมเพลต (คลิกเพื่อดู)", expanded=False):
            st.markdown("""\
- รองรับไฟล์ .csv หรือ .xlsx ที่มีคอลัมน์ **รหัสหมวด, ชื่อหมวด**
- ระบบจะ **อัปเดตชื่อหมวด** หากพบรหัสซ้ำ และ **เพิ่มใหม่** เมื่อไม่พบรหัสเดิม
- ไม่ลบรายการเดิมโดยอัตโนมัติ เว้นแต่เลือกโหมด 'แทนที่ทั้งชีต'
            """)
            tpl = """รหัสหมวด,ชื่อหมวด
PRT,หมึกพิมพ์
KBD,คีย์บอร์ด
"""
            st.download_button("ดาวน์โหลดเทมเพลต (CSV)", data=tpl.encode("utf-8-sig"),
                               file_name="template_categories.csv", mime="text/csv")

        cA, cB = st.columns([2,1])
        with cA:
            up = st.file_uploader("เลือกไฟล์ (.csv, .xlsx)", type=["csv","xlsx","xls"], key="cat_uploader_v2")
        with cB:
            replace_all = st.checkbox("แทนที่ทั้งชีต (ล้างและใส่ใหม่)", value=False)

        if up is not None:
            try:
                if up.name.lower().endswith(".csv"):
                    df_up = pd.read_csv(up, dtype=str)
                else:
                    df_up = pd.read_excel(up, dtype=str)
                df_up = df_up.fillna("").applymap(lambda x: str(x).strip())
            except Exception as e:
                st.error(f"อ่านไฟล์ไม่สำเร็จ: {e}")
                return

            rename_map = {"รหัสหมวดหมู่":"รหัสหมวด","ชื่อหมวดหมู่":"ชื่อหมวด","code":"รหัสหมวด","name":"ชื่อหมวด","category_code":"รหัสหมวด","category_name":"ชื่อหมวด"}
            df_up.columns = [rename_map.get(c.strip(), c.strip()) for c in df_up.columns]
            missing = [c for c in ["รหัสหมวด","ชื่อหมวด"] if c not in df_up.columns]
            if missing:
                st.error(f"ไฟล์ขาดคอลัมน์ที่บังคับ: {', '.join(missing)}"); return

            df_up["รหัสหมวด"] = df_up["รหัสหมวด"].str.upper()
            df_up = df_up[df_up["รหัสหมวด"]!=""]
            df_up = df_up.drop_duplicates(subset=["รหัสหมวด"], keep="last")

            st.success(f"พรีวิว {len(df_up):,} รายการ")
            st.dataframe(df_up, use_container_width=True, height=240)

            if st.button("🚀 ดำเนินการนำเข้า/อัปเดต", use_container_width=True, key="do_import"):
                base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if replace_all:
                    write_df(sh, SHEET_CATS, df_up[CATS_HEADERS])
                    st.success(f"แทนที่ทั้งชีตสำเร็จ • บันทึก {len(df_up):,} รายการ"); safe_rerun()
                else:
                    added, updated = 0, 0
                    for _, r in df_up.iterrows():
                        code, name = str(r["รหัสหมวด"]).strip().upper(), str(r["ชื่อหมวด"]).strip()
                        if not code or not name: 
                            continue
                        if (base["รหัสหมวด"] == code).any():
                            base.loc[base["รหัสหมวด"] == code, "ชื่อหมวด"] = name; updated += 1
                        else:
                            base = pd.concat([base, pd.DataFrame([[code, name]], columns=CATS_HEADERS)], ignore_index=True); added += 1
                    write_df(sh, SHEET_CATS, base)
                    st.success(f"นำเข้าสำเร็จ • เพิ่มใหม่ {added:,} • อัปเดต {updated:,}"); safe_rerun()

    # ---------------- TAB 3: Search & inline edit ----------------
    with tab3:
        c1, c2 = st.columns([2,1])
        with c1:
            q = st.text_input("ค้นหา (รหัส/ชื่อ)", key="cat_search_v2")
        with c2:
            st.caption("โหมดตารางสามารถแก้ไขค่าได้ แล้วกดบันทึกด้านล่าง")

        view = cats.copy()
        if not view.empty and q:
            mask = view["รหัสหมวด"].str.contains(q, case=False, na=False) | view["ชื่อหมวด"].str.contains(q, case=False, na=False)
            view = view[mask]

        # เปิดแก้ไขเฉพาะคอลัมน์ 'ชื่อหมวด' เพื่อกันรหัสเปลี่ยนโดยไม่ตั้งใจ
        edited = st.data_editor(
            view.sort_values("รหัสหมวด"),
            use_container_width=True,
            height=360,
            disabled=["รหัสหมวด"],
            key="cat_editor"
        )

        cL, cR = st.columns([1,1])
        with cL:
            if st.button("💾 บันทึกการแก้ไขในตาราง", use_container_width=True, key="save_table"):
                base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                # sync: update name for matching codes; ignore codesที่ไม่มีในฐาน
                for _, r in edited.iterrows():
                    base.loc[base["รหัสหมวด"] == str(r["รหัสหมวด"]).strip().upper(), "ชื่อหมวด"] = str(r["ชื่อหมวด"]).strip()
                write_df(sh, SHEET_CATS, base)
                st.success("บันทึกการแก้ไขเรียบร้อย")
                safe_rerun()
        with cR:
            # เพิ่มรายการใหม่แบบฟอร์มเล็กที่แท็บนี้
            with st.popover("➕ เพิ่มใหม่อย่างเร็ว"):
                code2 = st.text_input("รหัสหมวด (ใหม่)", key="quick_code").upper().strip()
                name2 = st.text_input("ชื่อหมวด", key="quick_name").strip()
                if st.button("เพิ่ม", key="quick_add"):
                    if not code2 or not name2:
                        st.warning("กรุณากรอกให้ครบ")
                    else:
                        base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                        if (base["รหัสหมวด"] == code2).any():
                            st.error("รหัสนี้มีอยู่แล้ว"); st.stop()
                        base = pd.concat([base, pd.DataFrame([[code2, name2]], columns=CATS_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_CATS, base)
                        st.success("เพิ่มใหม่แล้ว")
                        safe_rerun()


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide"); st.markdown(MINIMAL_CSS, unsafe_allow_html=True); st.markdown(RESPONSIVE_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE); st.caption(APP_TAGLINE)
    ensure_credentials_ui()
    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"): st.session_state["sheet_url"] = DEFAULT_SHEET_URL
    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", ["Dashboard","Stock","แจ้งปัญหา","เบิก/รับเข้า","รายงาน","ผู้ใช้","Settings"], index=0)
    if page == "Settings":
        page_settings(); st.caption("© 2025 IT Stock · Streamlit + Google Sheets"); return
    sheet_url = st.session_state.get("sheet_url", DEFAULT_SHEET_URL)
    if not sheet_url:
        st.info("ไปที่เมนู **Settings** แล้ววาง Google Sheet URL ที่คุณเป็นเจ้าของ จากนั้นกดปุ่มทดสอบเชื่อมต่อ"); return
    try:
        sh = open_sheet_by_url(sheet_url)
    except Exception as e:
        st.error(f"เปิดชีตไม่สำเร็จ: {e}"); return
    ensure_sheets_exist(sh)
    auth_block(sh)
    if page=="Dashboard": page_dashboard(sh)
    elif page=="Stock": page_stock(sh)
    elif page=="แจ้งปัญหา": page_tickets(sh)
    elif page=="เบิก/รับเข้า": page_issue_receive(sh)
    elif page=="รายงาน": page_reports(sh)
    elif page=="ผู้ใช้": page_users_admin(sh)
    elif page=="นำเข้า/แก้ไข หมวดหมู่": page_import(sh)
    st.caption("© 2025 IT Stock · Streamlit + Google Sheets")

if __name__ == "__main__":
    main()
