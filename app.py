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

def add_reload_button():
    col1, col2 = st.columns([1, 12])
    with col1:
        if st.button("🔁 รีโหลดข้อมูล", help="ล้าง cache และรีเฟรช"):
            try:
                st.cache_data.clear()
            except Exception:
                pass
            try:
                st.cache_resource.clear()
            except Exception:
                pass
            try:
                st.rerun()
            except Exception:
                st.experimental_rerun()
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


def page_stock(sh):
    add_reload_button()
    st.subheader("📦 คลังอุปกรณ์")

    import pandas as pd
    df = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)

    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)", "")
    if q.strip():
        ql = q.strip().lower()
        df = df[df.apply(lambda r: ql in str(r["รหัส"]).lower()
                                   or ql in str(r["ชื่ออุปกรณ์"]).lower()
                                   or ql in str(r["รหัสหมวด"]).lower(), axis=1)]

    df_show = df.copy()
    if "เลือก" not in df_show.columns:
        df_show.insert(0, "เลือก", False)

    st.caption("ติ๊กเลือก 1 แถวจากตารางเพื่อแก้ไขรายละเอียดด้านล่าง")
    edited_table = st.data_editor(
        df_show,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "เลือก": st.column_config.CheckboxColumn(help="เลือก 1 แถวเพื่อแก้ไขด้านล่าง"),
            "รหัส": st.column_config.TextColumn(disabled=True),
        },
        disabled=[c for c in df_show.columns if c not in ["เลือก"]],
        key="items_picker",
    )

    selected = edited_table[edited_table["เลือก"] == True]
    if len(selected) != 1:
        st.info("เลือก 1 แถวจากตารางด้านบนเพื่อแก้ไขรายละเอียด", icon="ℹ️")
        return

    row = selected.iloc[0]
    st.markdown("### ✏️ แก้ไขอุปกรณ์")

    col1, col2 = st.columns(2)
    code_id = col1.text_input("รหัส", value=row["รหัส"], disabled=True)
    cat  = col2.text_input("รหัสหมวด", value=row["รหัสหมวด"])
    name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
    unit = st.text_input("หน่วย", value=row["หน่วย"])
    bal  = st.number_input("คงเหลือ", min_value=0, step=1, value=int(pd.to_numeric(row["คงเหลือ"], errors="coerce") or 0))
    rop  = st.number_input("จุดสั่งซื้อ", min_value=0, step=1, value=int(pd.to_numeric(row["จุดสั่งซื้อ"], errors="coerce") or 0))
    loc  = st.text_input("ที่เก็บ", value=row["ที่เก็บ"])
    use  = st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(row["ใช้งาน"]).upper()=="Y" else 1)

    if st.button("บันทึกการแก้ไข", type="primary"):
        df2 = df.copy()
        df2.loc[df2["รหัส"] == code_id, ["รหัสหมวด","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]] = \
            [cat, name, unit, str(bal), str(rop), loc, use]
        write_df(sh, SHEET_ITEMS, df2.astype(str))
        st.success("บันทึกการแก้ไขเรียบร้อย", icon="✅")
        st.rerun()


def group_period(df, period="ME"):
    dfx = df.copy(); dfx["วันเวลา"] = pd.to_datetime(dfx["วันเวลา"], errors='coerce'); dfx = dfx.dropna(subset=["วันเวลา"])
    return dfx.groupby([pd.Grouper(key="วันเวลา", freq=period), "ประเภท", "ชื่ออุปกรณ์"])['จำนวน'].sum().reset_index()


def page_issue_out_multi5(sh):
    """เบิก (OUT): เลือกสาขาก่อน แล้วกรอกได้ 5 รายการในครั้งเดียว"""
    import pandas as pd, uuid
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    if items.empty:
        st.info("ยังไม่มีรายการอุปกรณ์", icon="ℹ️"); return

    bopt = st.selectbox("สาขา/หน่วยงานผู้ขอ", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
    branch_code = bopt.split(" | ")[0] if bopt else ""

    st.write("")
    st.markdown("**เลือกรายการที่ต้องการเบิก (ได้สูงสุด 5 รายการต่อครั้ง)**")

    opts = []
    for _, r in items.iterrows():
        remain = int(pd.to_numeric(r["คงเหลือ"], errors="coerce") or 0)
        opts.append(f'{r["รหัส"]} | {r["ชื่ออุปกรณ์"]} (คงเหลือ {remain})')

    df_template = pd.DataFrame({"รายการ": ["", "", "", "", ""], "จำนวน": [1, 1, 1, 1, 1]})
    ed = st.data_editor(
        df_template,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "รายการ": st.column_config.SelectboxColumn(options=opts, required=False),
            "จำนวน": st.column_config.NumberColumn(min_value=1, step=1)
        },
        key="issue_out_multi5",
    )

    note = st.text_input("หมายเหตุ (ถ้ามี)", value="")

    if st.button("บันทึกการเบิก (1 ครั้ง/หลายรายการ)", type="primary", disabled=(not branch_code)):
        txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
        errors = []
        processed = 0
        items_local = items.copy()

        for _, r in ed.iterrows():
            sel = str(r.get("รายการ","") or "").strip()
            qty = int(pd.to_numeric(r.get("จำนวน", 0), errors="coerce") or 0)
            if not sel or qty <= 0:
                continue

            code_sel = sel.split(" | ")[0]
            row_sel = items_local[items_local["รหัส"]==code_sel]
            if row_sel.empty:
                errors.append(f"{code_sel}: ไม่พบในคลัง")
                continue
            row_sel = row_sel.iloc[0]
            remain = int(pd.to_numeric(row_sel["คงเหลือ"], errors="coerce") or 0)
            if qty > remain:
                errors.append(f"{code_sel}: เกินคงเหลือ ({remain})")
                continue

            new_remain = remain - qty
            items_local.loc[items_local["รหัส"]==code_sel, "คงเหลือ"] = str(new_remain)

            from datetime import datetime
            txn = [str(uuid.uuid4())[:8], datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
                   "OUT", code_sel, row_sel["ชื่ออุปกรณ์"], branch_code, str(qty), get_username(), note]
            txns = pd.concat([txns, pd.DataFrame([txn], columns=TXNS_HEADERS)], ignore_index=True)
            processed += 1

        if processed > 0:
            write_df(sh, SHEET_ITEMS, items_local)
            write_df(sh, SHEET_TXNS, txns)
            st.success(f"บันทึกการเบิกแล้ว {processed} รายการ", icon="✅")
            st.rerun()
        else:
            st.warning("ยังไม่มีบรรทัดที่สมบูรณ์ให้บันทึก", icon="⚠️")

def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("🧾 เบิก/รับเข้า")
    if st.session_state.get("role") not in ("admin","staff"): st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกรายการได้"); st.markdown("</div>", unsafe_allow_html=True); return
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS); branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty: st.warning("ยังไม่มีรายการอุปกรณ์ในคลัง"); st.markdown("</div>", unsafe_allow_html=True); return
    t1,t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])

    with t1:
        page_issue_out_multi5(sh)
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
    st.subheader("นำเข้า/แก้ไข หมวดหมู่")

    # โหลดข้อมูลหมวดหมู่จากชีต
    try:
        cats = read_df(sh, SHEET_CATEGORIES)
    except Exception:
        import pandas as pd
        cats = pd.DataFrame(columns=["รหัสหมวดหมู่","ชื่อหมวดหมู่"])
    if "รหัสหมวดหมู่" not in cats.columns or "ชื่อหมวดหมู่" not in cats.columns:
        # สร้างคอลัมน์เริ่มต้นถ้ายังไม่มี
        if "รหัสหมวดหมู่" not in cats.columns: cats["รหัสหมวดหมู่"] = ""
        if "ชื่อหมวดหมู่" not in cats.columns: cats["ชื่อหมวดหมู่"] = ""
        st.dataframe(cats)

    with st.form("edit_category_form", clear_on_submit=False):
        cat_code = st.text_input("รหัสหมวดหมู่")
        cat_name = st.text_input("ชื่อหมวดหมู่")
        submitted = st.form_submit_button("บันทึก/แก้ไข")
    if submitted:
        if cat_code.strip() != "" and cat_name.strip() != "":
            # ถ้ามีรหัสหมวดหมู่เดิมแล้ว ให้แก้ไขชื่อแทน
            mask = cats["รหัสหมวดหมู่"] == cat_code
            if mask.any():
                cats.loc[mask, "ชื่อหมวดหมู่"] = cat_name
            else:
                cats.loc[len(cats)] = [cat_code, cat_name]
            write_df(sh, SHEET_CATEGORIES, cats)
            st.success("อัปเดตหมวดหมู่แล้ว")
            safe_rerun()
        else:
            st.warning("กรุณากรอกรหัสและชื่อหมวดหมู่ให้ครบ")
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📥 นำเข้า/แก้ไข หมวดหมู่ / เพิ่มข้อมูล (หมวดหมู่ / สาขา / อุปกรณ์ / หมวดหมู่ปัญหา)")
    st.caption("อัปโหลด CSV/Excel หรือ เพิ่มเองหลายรายการพร้อมการตรวจสอบความถูกต้อง")

    if st.session_state.get("role") not in ("admin","staff"):
        st.info("เฉพาะ admin/staff เท่านั้น")
        st.markdown("</div>", unsafe_allow_html=True)

    # ===== ปุ่มดาวน์โหลดเทมเพลต =====
    t1, t2, t3, t4 = st.columns(4)
    with t1:
        cat_csv = """รหัสหมวด,ชื่อหมวด
PRT,หมึกพิมพ์
KBD,คีย์บอร์ด
"""
        st.download_button("เทมเพลต หมวดหมู่ (CSV)", data=cat_csv.encode("utf-8-sig"),
                           file_name="template_categories.csv", mime="text/csv", use_container_width=True)
    with t2:
        br_csv = """รหัสสาขา,ชื่อสาขา
HQ,สำนักงานใหญ่
BKK1,สาขาบางนา
"""
        st.download_button("เทมเพลต สาขา (CSV)", data=br_csv.encode("utf-8-sig"),
                           file_name="template_branches.csv", mime="text/csv", use_container_width=True)
    with t3:
        it_csv = ",".join(ITEMS_HEADERS) + "\n" + "PRT-001,PRT,ตลับหมึก HP 206A,ตลับ,5,2,IT Room,Y\n"
        st.download_button("เทมเพลต อุปกรณ์ (CSV)", data=it_csv.encode("utf-8-sig"),
                           file_name="template_items.csv", mime="text/csv", use_container_width=True)
    with t4:
        tkc_csv = "รหัสหมวดปัญหา,ชื่อหมวดปัญหา\nNW,Network\nPRN,Printer\nSW,Software\n"
        st.download_button("เทมเพลต หมวดหมู่ปัญหา (CSV)", data=tkc_csv.encode("utf-8-sig"),
                           file_name="template_ticket_categories.csv", mime="text/csv", use_container_width=True)

    # ===== Tabs =====
    tab_cat, tab_br, tab_it, tab_tkc = st.tabs(["หมวดหมู่","สาขา","อุปกรณ์","หมวดหมู่ปัญหา"])

    # ---------- utils ----------
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
            df = df.applymap(lambda x: str(x).strip())
            return df, None
        except Exception as e:
            return None, f"อ่านไฟล์ไม่สำเร็จ: {e}"

    # ===== หมวดหมู่ =====
    with tab_cat:
        st.markdown("##### อัปโหลดไฟล์")
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ (CSV/Excel)", type=["csv","xlsx"], key="up_cat")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), use_container_width=True, height=200)
                if not set(["รหัสหมวด","ชื่อหมวด"]).issubset(df.columns):
                    st.error("หัวตารางต้องประกอบด้วย: รหัสหมวด, ชื่อหมวด")
                else:
                    if st.button("นำเข้า/อัปเดต หมวดหมู่", use_container_width=True, key="btn_imp_cat"):
                        cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
                        for _, r in df.iterrows():
                            code_c = str(r["รหัสหมวด"]).strip()
                            name_c = str(r["ชื่อหมวด"]).strip()
                            if code_c == "": continue
                            if (cur["รหัสหมวด"]==code_c).any():
                                cur.loc[cur["รหัสหมวด"]==code_c, ["รหัสหมวด","ชื่อหมวด"]] = [code_c, name_c]
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code_c, name_c]], columns=CATS_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_CATS, cur); st.success("นำเข้าหมวดหมู่สำเร็จ")

        st.markdown("##### เพิ่มทีละรายการ")
        with st.form("form_add_cat", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1: code_c = st.text_input("รหัสหมวด*", max_chars=10)
            with col2: name_c = st.text_input("ชื่อหมวด*")
            s = st.form_submit_button("เพิ่มหมวดหมู่", use_container_width=True)
        if s:
            if not code_c or not name_c:
                st.warning("กรอกข้อมูลให้ครบ")
            else:
                cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if (cur["รหัสหมวด"]==code_c).any():
                    st.error("มีรหัสนี้อยู่แล้ว")
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_c.strip(), name_c.strip()]], columns=CATS_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_CATS, cur); st.success("เพิ่มสำเร็จ")

        st.markdown("##### เพิ่มหลายรายการ")
        n_cat = st.number_input("จำนวนบรรทัด", min_value=1, max_value=100, value=10, step=1, key="cat_rows")
        df_multi = pd.DataFrame({"รหัสหมวด":[""]*n_cat, "ชื่อหมวด":[""]*n_cat})
        edited = st.data_editor(df_multi, use_container_width=True, num_rows="dynamic", key="cat_editor")
        if st.button("บันทึกหลายรายการ (หมวดหมู่)", use_container_width=True, key="save_cats_multi"):
            cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
            errs = []
            add = 0; upd = 0
            seen = set()
            for i, r in edited.iterrows():
                code_c = str(r.get("รหัสหมวด","")).strip()
                name_c = str(r.get("ชื่อหมวด","")).strip()
                if code_c=="" and name_c=="": continue
                if code_c=="":
                    errs.append({"row":i+1,"error":"รหัสหมวดว่าง","code":code_c})
                    continue
                if code_c in seen:
                    errs.append({"row":i+1,"error":"รหัสซ้ำในตาราง","code":code_c}); continue
                seen.add(code_c)
                if (cur["รหัสหมวด"]==code_c).any():
                    cur.loc[cur["รหัสหมวด"]==code_c, ["รหัสหมวด","ชื่อหมวด"]] = [code_c, name_c]; upd+=1
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_c, name_c]], columns=CATS_HEADERS)], ignore_index=True); add+=1
            write_df(sh, SHEET_CATS, cur)
            st.success(f"เพิ่ม {add} ราย / อัปเดต {upd} ราย")
            if errs: st.warning(pd.DataFrame(errs))

    # ===== สาขา =====
    with tab_br:
        st.markdown("##### อัปโหลดไฟล์")
        up = st.file_uploader("อัปโหลดไฟล์ สาขา (CSV/Excel)", type=["csv","xlsx"], key="up_br")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), use_container_width=True, height=200)
                if not set(["รหัสสาขา","ชื่อสาขา"]).issubset(df.columns):
                    st.error("หัวตารางต้องประกอบด้วย: รหัสสาขา, ชื่อสาขา")
                else:
                    if st.button("นำเข้า/อัปเดต สาขา", use_container_width=True, key="btn_imp_br"):
                        cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
                        for _, r in df.iterrows():
                            code_b = str(r["รหัสสาขา"]).strip()
                            name_b = str(r["ชื่อสาขา"]).strip()
                            if code_b == "": continue
                            if (cur["รหัสสาขา"]==code_b).any():
                                cur.loc[cur["รหัสสาขา"]==code_b, ["รหัสสาขา","ชื่อสาขา"]] = [code_b, name_b]
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code_b, name_b]], columns=BR_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_BRANCHES, cur); st.success("นำเข้าสาขาสำเร็จ")

        st.markdown("##### เพิ่มทีละรายการ")
        with st.form("form_add_branch", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1: code_b = st.text_input("รหัสสาขา*", max_chars=10, key="br_code_m")
            with col2: name_b = st.text_input("ชื่อสาขา*", key="br_name_m")
            s2 = st.form_submit_button("เพิ่มสาขา", use_container_width=True)
        if s2:
            if not code_b or not name_b:
                st.warning("กรอกข้อมูลให้ครบ")
            else:
                cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
                if (cur["รหัสสาขา"]==code_b).any():
                    st.error("มีรหัสนี้อยู่แล้ว")
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_b.strip(), name_b.strip()]], columns=BR_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_BRANCHES, cur); st.success("เพิ่มสำเร็จ")

        st.markdown("##### เพิ่มหลายรายการ")
        n_br = st.number_input("จำนวนบรรทัด", min_value=1, max_value=200, value=10, step=1, key="br_rows")
        df_multi = pd.DataFrame({"รหัสสาขา":[""]*n_br, "ชื่อสาขา":[""]*n_br})
        edited = st.data_editor(df_multi, use_container_width=True, num_rows="dynamic", key="br_editor")
        if st.button("บันทึกหลายรายการ (สาขา)", use_container_width=True, key="save_br_multi"):
            cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
            errs = []; add=0; upd=0; seen=set()
            for i, r in edited.iterrows():
                code_b = str(r.get("รหัสสาขา","")).strip()
                name_b = str(r.get("ชื่อสาขา","")).strip()
                if code_b=="" and name_b=="": continue
                if code_b=="":
                    errs.append({"row":i+1,"error":"รหัสสาขาว่าง"}); continue
                if code_b in seen: errs.append({"row":i+1,"error":"รหัสซ้ำในตาราง","code":code_b}); continue
                seen.add(code_b)
                if (cur["รหัสสาขา"]==code_b).any():
                    cur.loc[cur["รหัสสาขา"]==code_b, ["รหัสสาขา","ชื่อสาขา"]] = [code_b, name_b]; upd+=1
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_b, name_b]], columns=BR_HEADERS)], ignore_index=True); add+=1
            write_df(sh, SHEET_BRANCHES, cur); st.success(f"เพิ่ม {add} / อัปเดต {upd}")
            if errs: st.warning(pd.DataFrame(errs))

    # ===== อุปกรณ์ =====
    with tab_it:
        st.markdown("##### อัปโหลดไฟล์")
        up = st.file_uploader("อัปโหลดไฟล์ อุปกรณ์ (CSV/Excel)", type=["csv","xlsx"], key="up_it")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), use_container_width=True, height=260)
                missing_cols = [c for c in ["หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"] if c not in df.columns]
                if missing_cols:
                    st.error("หัวตารางต้องประกอบด้วยอย่างน้อย: หมวดหมู่, ชื่ออุปกรณ์, หน่วย, คงเหลือ, จุดสั่งซื้อ, ที่เก็บ (รหัส, ใช้งาน เป็นออปชัน)")
                else:
                    if st.button("นำเข้า/อัปเดต อุปกรณ์", use_container_width=True, key="btn_imp_items"):
                        cur = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                        cats_df = read_df(sh, SHEET_CATS, CATS_HEADERS)
                        valid_cats = set(cats_df["รหัสหมวด"].tolist()) if not cats_df.empty else set()
                        errs=[]; add=0; upd=0; seen=set()
                        for i, r in df.iterrows():
                            code_i = str(r.get("รหัส","")).strip().upper()
                            cat  = str(r.get("หมวดหมู่","")).strip()
                            name = str(r.get("ชื่ออุปกรณ์","")).strip()
                            unit = str(r.get("หน่วย","")).strip()
                            qty  = str(r.get("คงเหลือ","")).strip()
                            rop  = str(r.get("จุดสั่งซื้อ","")).strip()
                            loc  = str(r.get("ที่เก็บ","")).strip()
                            active = str(r.get("ใช้งาน","Y")).strip().upper() or "Y"
                            if name=="" or unit=="":
                                errs.append({"row":i+1,"error":"ชื่อ/หน่วย ว่าง"}); continue
                            if cat not in valid_cats:
                                errs.append({"row":i+1,"error":"หมวดไม่มีในระบบ","cat":cat}); continue
                            try: qty = int(float(qty)); 
                            except: qty = 0
                            try: rop = int(float(rop)); 
                            except: rop = 0
                            qty = max(0, qty); rop = max(0, rop) # ไม่ติดลบ
                            if code_i=="": code_i = generate_item_code(sh, cat)
                            if code_i in seen: errs.append({"row":i+1,"error":"รหัสซ้ำในไฟล์/ตาราง","code":code_i}); continue
                            seen.add(code_i)
                            if (cur["รหัส"]==code_i).any():
                                cur.loc[cur["รหัส"]==code_i, ITEMS_HEADERS] = [code_i, cat, name, unit, qty, rop, loc, active]; upd+=1
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code_i, cat, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True); add+=1
                        write_df(sh, SHEET_ITEMS, cur)
                        st.success(f"เพิ่ม {add} ราย / อัปเดต {upd} ราย")
                        if errs: st.warning(pd.DataFrame(errs))

        st.markdown("##### เพิ่มหลายรายการ (แก้ไขในตาราง)")
        cats_df = read_df(sh, SHEET_CATS, CATS_HEADERS)
        cat_opts = (cats_df["รหัสหมวด"].tolist() if not cats_df.empty else [])
        n_item = st.number_input("จำนวนบรรทัด", min_value=1, max_value=200, value=10, step=1, key="it_rows")
        df_multi = pd.DataFrame({
            "หมวดหมู่":[""]*n_item,
            "รหัส":[""]*n_item,
            "ชื่ออุปกรณ์":[""]*n_item,
            "หน่วย":[""]*n_item,
            "คงเหลือ":[0]*n_item,
            "จุดสั่งซื้อ":[0]*n_item,
            "ที่เก็บ":[""]*n_item,
            "ใช้งาน":["Y"]*n_item,
        })
        cfg = {
            "หมวดหมู่": st.column_config.SelectboxColumn(options=cat_opts if cat_opts else ["กรอกเอง"], required=False),
            "ใช้งาน": st.column_config.SelectboxColumn(options=["Y","N"]),
            "คงเหลือ": st.column_config.NumberColumn(min_value=0, step=1),
            "จุดสั่งซื้อ": st.column_config.NumberColumn(min_value=0, step=1),
        }
        edited = st.data_editor(df_multi, use_container_width=True, num_rows="dynamic", column_config=cfg, key="it_editor")
        mode = st.selectbox("ถ้าพบรหัสซ้ำ", ["อัปเดต","ข้าม"], index=0, key="dup_mode_items")
        if st.button("บันทึกหลายรายการ (อุปกรณ์)", use_container_width=True, key="save_items_multi"):
            cur = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
            valid_cats = set(cat_opts)
            errs=[]; add=0; upd=0; seen=set()
            for i, r in edited.iterrows():
                cat  = str(r.get("หมวดหมู่","")).strip()
                code_i = str(r.get("รหัส","")).strip().upper()
                name = str(r.get("ชื่ออุปกรณ์","")).strip()
                unit = str(r.get("หน่วย","")).strip()
                qty  = r.get("คงเหลือ",0); rop = r.get("จุดสั่งซื้อ",0)
                loc  = str(r.get("ที่เก็บ","")).strip()
                active = str(r.get("ใช้งาน","Y")).strip().upper() or "Y"
                if (cat=="" and name=="" and unit==""): continue
                if name=="" or unit=="":
                    errs.append({"row":i+1,"error":"ชื่อ/หน่วย ว่าง","code":code_i}); continue
                if cat not in valid_cats:
                    errs.append({"row":i+1,"error":"หมวดไม่มีในระบบ","cat":cat}); continue
                try: qty = int(qty)
                except: qty = 0
                try: rop = int(rop)
                except: rop = 0
                qty = max(0, qty); rop = max(0, rop)
                if code_i=="": code_i = generate_item_code(sh, cat)
                if code_i in seen:
                    errs.append({"row":i+1,"error":"รหัสซ้ำในตาราง","code":code_i}); continue
                seen.add(code_i)
                if (cur["รหัส"]==code_i).any():
                    if mode=="อัปเดต":
                        cur.loc[cur["รหัส"]==code_i, ITEMS_HEADERS] = [code_i, cat, name, unit, qty, rop, loc, active]; upd+=1
                    else:
                        errs.append({"row":i+1,"error":"รหัสชนกับระบบ (ข้าม)","code":code_i}); continue
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_i, cat, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True); add+=1
            write_df(sh, SHEET_ITEMS, cur)
            st.success(f"เพิ่ม {add} ราย / อัปเดต {upd} ราย")
            if errs:
                err_df = pd.DataFrame(errs)
                st.warning(err_df)
                st.download_button("ดาวน์โหลดรายการที่ผิดพลาด (CSV)", data=err_df.to_csv(index=False).encode("utf-8-sig"),
                                   file_name="item_batch_errors.csv", mime="text/csv")

    # ===== หมวดหมู่ปัญหา =====
    with tab_tkc:
        st.markdown("##### อัปโหลดไฟล์")
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ปัญหา (CSV/Excel)", type=["csv","xlsx"], key="up_tkc")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), use_container_width=True, height=200)
                if not set(["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]).issubset(df.columns):
                    st.error("หัวตารางต้องประกอบด้วย: รหัสหมวดปัญหา, ชื่อหมวดปัญหา")
                else:
                    if st.button("นำเข้า/อัปเดต หมวดหมู่ปัญหา", use_container_width=True, key="btn_imp_tkc"):
                        cur = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
                        for _, r in df.iterrows():
                            code_t = str(r["รหัสหมวดปัญหา"]).strip()
                            name_t = str(r["ชื่อหมวดปัญหา"]).strip()
                            if code_t == "": continue
                            if (cur["รหัสหมวดปัญหา"]==code_t).any():
                                cur.loc[cur["รหัสหมวดปัญหา"]==code_t, ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]] = [code_t, name_t]
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code_t, name_t]], columns=TICKET_CAT_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_TICKET_CATS, cur); st.success("นำเข้าหมวดหมู่ปัญหาสำเร็จ")

        st.markdown("##### เพิ่มทีละรายการ")
        with st.form("form_add_tkc", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1: code_t = st.text_input("รหัสหมวดปัญหา*", max_chars=10, key="tkc_code_m")
            with col2: name_t = st.text_input("ชื่อหมวดปัญหา*", key="tkc_name_m")
            s4 = st.form_submit_button("เพิ่มหมวดหมู่ปัญหา", use_container_width=True)
        if s4:
            if not code_t or not name_t:
                st.warning("กรอกข้อมูลให้ครบ")
            else:
                cur = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
                if (cur["รหัสหมวดปัญหา"]==code_t).any():
                    st.error("มีรหัสนี้อยู่แล้ว")
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_t.strip(), name_t.strip()]], columns=TICKET_CAT_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_TICKET_CATS, cur); st.success("เพิ่มสำเร็จ")

        st.markdown("##### เพิ่มหลายรายการ")
        n_tkc = st.number_input("จำนวนบรรทัด", min_value=1, max_value=200, value=10, step=1, key="tkc_rows")
        df_multi = pd.DataFrame({"รหัสหมวดปัญหา":[""]*n_tkc, "ชื่อหมวดปัญหา":[""]*n_tkc})
        edited = st.data_editor(df_multi, use_container_width=True, num_rows="dynamic", key="tkc_editor")
        if st.button("บันทึกหลายรายการ (หมวดหมู่ปัญหา)", use_container_width=True, key="save_tkc_multi"):
            cur = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
            errs=[]; add=0; upd=0; seen=set()
            for i, r in edited.iterrows():
                code_t = str(r.get("รหัสหมวดปัญหา","")).strip()
                name_t = str(r.get("ชื่อหมวดปัญหา","")).strip()
                if code_t=="" and name_t=="": continue
                if code_t=="": errs.append({"row":i+1,"error":"รหัสว่าง"}); continue
                if code_t in seen: errs.append({"row":i+1,"error":"รหัสซ้ำในตาราง","code":code_t}); continue
                seen.add(code_t)
                if (cur["รหัสหมวดปัญหา"]==code_t).any():
                    cur.loc[cur["รหัสหมวดปัญหา"]==code_t, ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]] = [code_t, name_t]; upd+=1
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_t, name_t]], columns=TICKET_CAT_HEADERS)], ignore_index=True); add+=1
            write_df(sh, SHEET_TICKET_CATS, cur); st.success(f"เพิ่ม {add} / อัปเดต {upd}")
            if errs: st.warning(pd.DataFrame(errs))

    st.markdown("</div>", unsafe_allow_html=True)
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide"); st.markdown(MINIMAL_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE); st.caption(APP_TAGLINE)
    ensure_credentials_ui()
    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"): st.session_state["sheet_url"] = DEFAULT_SHEET_URL
    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", ["Dashboard","คลังอุปกรณ์","แจ้งซ่อม / แจ้งปัญหา (Tickets)","เบิก/รับเข้า","รายงาน","ผู้ใช้","นำเข้า/แก้ไข หมวดหมู่","Settings"], index=0)
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
    elif page=="คลังอุปกรณ์": page_stock(sh)
    elif page=="แจ้งซ่อม / แจ้งปัญหา (Tickets)": page_tickets(sh)
    elif page=="เบิก/รับเข้า": page_issue_receive(sh)
    elif page=="รายงาน": page_reports(sh)
    elif page=="ผู้ใช้": page_users_admin(sh)
    elif page=="นำเข้า/แก้ไข หมวดหมู่": page_import(sh)
    st.caption("© 2025 IT Stock · Streamlit + Google Sheets")

if __name__ == "__main__":
    main()
