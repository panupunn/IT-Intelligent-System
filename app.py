#!/usr/bin/env python3
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
import os, io, uuid, re, time
from datetime import datetime, date, timedelta, timedelta, date, time as dtime
import pytz, pandas as pd, streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import gspread
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials
import bcrypt
import altair as alt

# ---- Compatibility helper for Streamlit rerun ----

# -------------------- User helper --------------------
def get_username():
    """
    ดึงชื่อผู้ใช้จาก session_state ให้รองรับหลาย key
    ถ้าไม่พบจะคืนค่า "unknown" เพื่อกัน NameError/KeyError
    """
    import streamlit as st

def setup_responsive():
    # Global CSS for better smartphone experience
    st.markdown("""
    <style>
    /* Reduce paddings on narrow screens */
    @media (max-width: 640px) {
        .block-container { padding: 0.6rem 0.7rem !important; }
        /* Stack columns (Streamlit columns are flex items) */
        [data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; padding-right: 0 !important; }
        /* Make buttons fill width for easier tapping */
        .stButton > button { width: 100% !important; }
        /* Make selects and inputs fill width */
        .stSelectbox, .stTextInput, .stTextArea, .stDateInput { width: 100% !important; }
        /* Dataframe should use container width; let it be scrollable horizontally */
        .stDataFrame { width: 100% !important; }
        /* Smaller chart margins */
        .js-plotly-plot, .vega-embed { width: 100% !important; }
    }
    </style>
    """, unsafe_allow_html=True)
    return (
        st.session_state.get("user")
        or st.session_state.get("username")
        or st.session_state.get("display_name")
        or "unknown"
    )
# -----------------------------------------------------

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
    """
    Make sure all required worksheets exist.
    More resilient:
      - Retries listing worksheets (handles intermittent API errors/quotas)
      - Falls back to per-sheet check to avoid hard failure
    """
    import time
    from gspread.exceptions import APIError, WorksheetNotFound

    # Try listing worksheets up to 3 times
    titles = []
    for attempt in range(3):
        try:
            titles = [ws.title for ws in sh.worksheets()]
            break
        except APIError as e:
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            # Fallback will check per-sheet below
            titles = None

    required = [
        (SHEET_ITEMS, ITEMS_HEADERS, 1000, len(ITEMS_HEADERS)+5),
        (SHEET_TXNS, TXNS_HEADERS, 2000, len(TXNS_HEADERS)+5),
        (SHEET_USERS, USERS_HEADERS, 100, len(USERS_HEADERS)+2),
        (SHEET_CATS, CATS_HEADERS, 200, len(CATS_HEADERS)+2),
        (SHEET_BRANCHES, BR_HEADERS, 200, len(BR_HEADERS)+2),
        (SHEET_TICKETS, TICKETS_HEADERS, 1000, len(TICKETS_HEADERS)+5),
        (SHEET_TICKET_CATS, TICKET_CAT_HEADERS, 200, len(TICKET_CAT_HEADERS)+2),
    ]

    def ensure_one(name, headers, rows, cols):
        try:
            if titles is not None:
                if name in titles:
                    return
                # when titles are known and sheet missing -> create
                ws = sh.add_worksheet(name, rows, cols)
                ws.append_row(headers)
            else:
                # Fallback: check directly
                try:
                    sh.worksheet(name)  # exists
                except WorksheetNotFound:
                    ws = sh.add_worksheet(name, rows, cols)
                    ws.append_row(headers)
        except APIError as e:
            # Surface a user-friendly error but don't crash the entire app
            st.warning(f"ไม่สามารถตรวจสอบ/สร้างชีต '{name}' ได้ชั่วคราว: {e}. ลองรีเฟรชใหม่อีกครั้ง")

    for name, headers, r, c in required:
        ensure_one(name, headers, r, c)

    # Seed default admin user when USERS sheet was newly created (or empty)
    try:
        ws_users = sh.worksheet(SHEET_USERS)
        values = ws_users.get_all_values()
        if len(values) <= 1:  # only header
            default_pwd = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            ws_users.append_row(["admin","Administrator","admin",default_pwd,"Y"])
    except Exception:
        pass

_READ_CACHE = {}

def clear_read_cache():
    _READ_CACHE.clear()

def _get_all_values_with_retry(ws, max_attempts: int = 5):
    # Call ws.get_all_values() with simple exponential backoff for 429/5xx errors.
    for attempt in range(max_attempts):
        try:
            return ws.get_all_values()
        except Exception as e:
            status = getattr(getattr(e, 'response', None), 'status_code', None)
            message = str(e)
            retryable = (status in (429, 500, 503)) or ('429' in message) or ('Quota exceeded' in message)
            if not retryable or attempt == max_attempts - 1:
                raise
            sleep_s = min(2 ** attempt, 16)
            time.sleep(sleep_s)

def read_df(sh, title, headers, _ttl_seconds: int = 15):
    # Read a worksheet into DataFrame with retry + short-term caching.
    try:
        sh_id = getattr(sh, 'id', None) or getattr(sh, 'spreadsheet_id', None) or 'unknown'
    except Exception:
        sh_id = 'unknown'
    key = (str(sh_id), str(title), tuple(headers))
    now = time.time()
    entry = _READ_CACHE.get(key)
    if entry and (now - entry['ts'] < _ttl_seconds):
        return entry['df'].copy()

    ws = sh.worksheet(title)
    vals = _get_all_values_with_retry(ws)
    if not vals:
        df = pd.DataFrame(columns=headers)
    else:
        df = pd.DataFrame(vals[1:], columns=vals[0])
        if df.empty:
            df = pd.DataFrame(columns=headers)

    _READ_CACHE[key] = {'df': df.copy(), 'ts': now}
    return df

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
    clear_read_cache()

def append_row(sh, title, row):
    sh.worksheet(title).append_row(row)
    clear_read_cache()

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

def make_bar(df: pd.DataFrame, label_col: str, value_col: str, top_n: int, title: str):
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
        work = work.head(top_n)
    st.markdown(f"**{title}**")
    chart = alt.Chart(work).mark_bar().encode(
        x=alt.X(f"{label_col}:N", sort='-y'),
        y=alt.Y("sum_val:Q"),
        tooltip=[f"{label_col}:N","sum_val:Q"]
    )
    st.altair_chart(chart.properties(height=320), use_container_width=True)
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

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

def export_charts_to_pdf(charts, selected_titles, chart_kind):
    """Build a PDF (bytes) of selected charts. charts: list of (title, df, label_col, value_col)."""
    import pandas as pd
    from io import BytesIO

    # Use DejaVu Sans which supports Thai well
    try:
        matplotlib.rcParams['font.family'] = 'DejaVu Sans'
    except Exception:
        pass

    buf = BytesIO()
    with PdfPages(buf) as pdf:

        for title, df, label_col, value_col in charts:
            if title not in selected_titles:
                continue
            data = df.copy()
            # ensure numeric
            if value_col in data.columns:
                data[value_col] = pd.to_numeric(data[value_col], errors="coerce").fillna(0)

            plt.figure()
            if chart_kind.endswith("(Bar)"):
                # bar
                plt.bar(data[label_col].astype(str), data[value_col])
                plt.xticks(rotation=45, ha="right")
                plt.ylabel(value_col)
            else:
                # pie
                vals = data[value_col]
                labels = data[label_col].astype(str)
                if vals.sum() > 0:
                    plt.pie(vals, labels=labels, autopct="%1.1f%%")
                else:
                    # avoid zero-sum pie
                    plt.bar(labels, vals)
                    plt.xticks(rotation=45, ha="right")
                    plt.ylabel(value_col)
            plt.title(title)
            plt.tight_layout()
            pdf.savefig()  # saves the current figure
            plt.close()

    buf.seek(0)
    return buf.getvalue()

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
    cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    cat_map = {str(r['รหัสหมวด']).strip(): str(r['ชื่อหมวด']).strip() for _, r in cats.iterrows()} if not cats.empty else {}
    br_map = {str(r['รหัสสาขา']).strip(): f"{str(r['รหัสสาขา']).strip()} | {str(r['ชื่อสาขา']).strip()}" for _, r in branches.iterrows()} if not branches.empty else {}

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
    chart_kind = st.radio("ชนิดกราฟ", ["กราฟวงกลม (Pie)", "กราฟแท่ง (Bar)"], horizontal=True)

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
        tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
        charts.append(("คงเหลือตามหมวดหมู่", tmp, "หมวดหมู่ชื่อ", "คงเหลือ"))

    if "คงเหลือตามที่เก็บ" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("ที่เก็บ")["คงเหลือ"].sum().reset_index()
        charts.append(("คงเหลือตามที่เก็บ", tmp, "ที่เก็บ", "คงเหลือ"))

    if "จำนวนรายการตามหมวดหมู่" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["count"] = 1
        tmp = tmp.groupby("หมวดหมู่")["count"].sum().reset_index()
        tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
        charts.append(("จำนวนรายการตามหมวดหมู่", tmp, "หมวดหมู่ชื่อ", "count"))

    if "เบิกตามสาขา (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("สาขา", dropna=False)["จำนวน"].sum().reset_index()
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda x: br_map.get(str(x).split(" | ")[0], str(x) if "|" in str(x) else str(x)))
            charts.append((f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน"))
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
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda x: br_map.get(str(x).split(" | ")[0], str(x) if "|" in str(x) else str(x)))
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน"))
        else:
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", pd.DataFrame({"สาขา":[], "จำนวน":[]}), "สาขา", "จำนวน"))

    if len(charts)==0:
        st.info("โปรดเลือกกราฟที่ต้องการแสดงจากด้านบน")
    else:

        # ====== พิมพ์/ดาวน์โหลดกราฟเป็น PDF ======
        titles_all = [t for t,_,_,_ in charts]
        if len(titles_all) > 0:
            with st.expander("พิมพ์/ดาวน์โหลดกราฟเป็น PDF", expanded=False):
                sel = st.multiselect("เลือกกราฟที่จะพิมพ์เป็น PDF", options=titles_all, default=titles_all[:min(2,len(titles_all))])
                if sel:
                    pdf_bytes = export_charts_to_pdf(charts, sel, chart_kind)
                    st.download_button("ดาวน์โหลด PDF กราฟที่เลือก", data=pdf_bytes, file_name="dashboard_charts.pdf", mime="application/pdf")
        # =========================================
else:
        rows = (len(charts) + per_row - 1) // per_row
        idx = 0
        # ====== พิมพ์/ดาวน์โหลดกราฟเป็น PDF ======
        titles_all = [t for t,_,_,_ in charts]
        if len(titles_all) > 0:
            with st.expander("พิมพ์/ดาวน์โหลดกราฟเป็น PDF", expanded=False):
                sel = st.multiselect("เลือกกราฟที่จะพิมพ์เป็น PDF", options=titles_all, default=titles_all[:min(2,len(titles_all))])
                if sel:
                    pdf_bytes = export_charts_to_pdf(charts, sel, chart_kind)
                    st.download_button("ดาวน์โหลด PDF กราฟที่เลือก", data=pdf_bytes, file_name="dashboard_charts.pdf", mime="application/pdf")
        # =========================================

        for r in range(rows):
            cols = st.columns(per_row)
            for c in range(per_row):
                if idx >= len(charts): break
                title, df, label_col, value_col = charts[idx]
                with cols[c]:
                    make_bar(df, label_col, value_col, top_n, title) if chart_kind.endswith('(Bar)') else make_pie(df, label_col, value_col, top_n, title)
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
            st.dataframe(low_df2[["รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]], height=240, use_container_width=True)
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
    from datetime import datetime, date, timedelta
    return "TCK-" + datetime.now(TZ).strftime("%Y%m%d-%H%M%S")
