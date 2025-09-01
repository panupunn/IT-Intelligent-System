
# -*- coding: utf-8 -*-
"""
IT Stock (Streamlit + Google Sheets)
Patched single-file app — V.1.1

- Prefer Secrets/ENV/File/Embedded for GCP Service Account (no more uploader prompt when configured)
- Robust wrappers for open_sheet_by_url / open_sheet_by_key
- get_username() returns proper value
- Mobile-friendly CSS
- Includes: Dashboard, Stock, Tickets, Issue/Receive, Reports, Import (หมวดหมู่/สาขา/อุปกรณ์/หมวดหมู่ปัญหา/ผู้ใช้), Users, Settings
"""

import os, io, re, uuid, json, base64, time
from datetime import datetime, date, timedelta, time as dtime

import streamlit as st

# ----- Shim for older Streamlit: st.cache_resource -----
if not hasattr(st, "cache_resource"):
    def _no_cache_decorator(*args, **kwargs):
        def _wrap(func): return func
        return _wrap
    st.cache_resource = _no_cache_decorator

import pandas as pd
import altair as alt
import pytz

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials

import bcrypt

# -----------------------------
# App constants
# -----------------------------
APP_TITLE = "ไอต้าว ไอที (iTao iT)"
APP_TAGLINE = "POWER By ทีมงาน=> ไอทีสุดหล่อ"
VERSION_DISPLAY = "iTao iT (V.1.1)"

TZ = pytz.timezone("Asia/Bangkok")

# Replace with your sheet URL (can be changed in Settings)
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit?gid=0#gid=0"

# Local file path (only used when a file truly exists)
CREDENTIALS_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")

# Optional embedded credentials (base64 JSON). Keep empty by default for safety.
EMBEDDED_GOOGLE_CREDENTIALS_B64 = os.environ.get("EMBEDDED_SA_B64", "").strip()

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# --- Sheet names / headers (must match existing sheets) ---
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

MINIMAL_CSS = """
<style>
:root { --radius: 16px; }
section.main > div { padding-top: 8px; }
.block-card { background: #fff; border:1px solid #eee; border-radius:16px; padding:16px; }
.kpi { display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.danger { color:#b00020; }

/* Mobile friendly */
@media (max-width: 640px) {
  .block-container { padding: 0.6rem 0.7rem !important; }
  [data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; padding-right: 0 !important; }
  .stButton > button { width: 100% !important; }
  .stSelectbox, .stTextInput, .stTextArea, .stDateInput { width: 100% !important; }
  .stDataFrame { width: 100% !important; }
  .js-plotly-plot, .vega-embed { width: 100% !important; }
}
</style>"""

# -----------------------------
# Credential helpers
# -----------------------------
def _try_load_sa_from_secrets():
    try:
        # st.secrets["gcp_service_account"] is dict-like
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
        # also accept string JSON blobs under another key
        if "service_account_json" in st.secrets:
            raw = str(st.secrets["service_account_json"])
            return json.loads(raw)
        if "service_account" in st.secrets:
            return dict(st.secrets["service_account"])
    except Exception:
        pass
    return None

def _try_load_sa_from_env():
    raw = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON") \
          or os.environ.get("SERVICE_ACCOUNT_JSON") \
          or os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not raw: return None
    try:
        raw = raw.strip()
        if raw.startswith("{"):
            return json.loads(raw)
        return json.loads(base64.b64decode(raw).decode("utf-8"))
    except Exception:
        return None

def _try_load_sa_from_file():
    for p in ("./service_account.json", "/mount/data/service_account.json", "/mnt/data/service_account.json"):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            continue
    return None

def _try_load_sa_from_embedded():
    if not EMBEDDED_GOOGLE_CREDENTIALS_B64: return None
    try:
        return json.loads(base64.b64decode(EMBEDDED_GOOGLE_CREDENTIALS_B64).decode("utf-8"))
    except Exception:
        return None

def _detect_sa_source():
    try:
        if "gcp_service_account" in st.secrets or "service_account" in st.secrets or "service_account_json" in st.secrets:
            return "secrets"
    except Exception:
        pass
    if os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON") or os.environ.get("SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_CREDENTIALS_JSON"):
        return "env"
    for p in ("./service_account.json", "/mount/data/service_account.json", "/mnt/data/service_account.json"):
        if os.path.exists(p):
            return "file"
    if EMBEDDED_GOOGLE_CREDENTIALS_B64:
        return "embedded"
    return "none"

def _current_sa_info():
    src = _detect_sa_source()
    info = None
    try:
        if src == "secrets":
            info = _try_load_sa_from_secrets()
        elif src == "env":
            info = _try_load_sa_from_env()
        elif src == "file":
            info = _try_load_sa_from_file()
        elif src == "embedded":
            info = _try_load_sa_from_embedded()
    except Exception:
        info = None
    return src, info

@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    # try sources in order
    info = (_try_load_sa_from_secrets()
            or _try_load_sa_from_env()
            or _try_load_sa_from_file()
            or _try_load_sa_from_embedded())

    if info is None:
        # none configured, we raise — outer UI will handle uploader if needed
        raise RuntimeError("No service account available via Secrets/ENV/File/Embedded")

    creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_client():
    """Build client from secrets/env/file. Falls back to file if explicitly present."""
    src, info = _current_sa_info()
    if src in ("secrets","env","embedded") and info:
        creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    elif src == "file":
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=GOOGLE_SCOPES)
    else:
        # If truly none, raise; UI will show uploader
        raise RuntimeError("No service account configured")
    return gspread.authorize(creds)

def ensure_credentials_ui():
    """
    If SA exists in Secrets/ENV/File/Embedded, proceed silently (no uploader).
    Otherwise allow one-time upload and persist to local CREDENTIALS_FILE.
    """
    src, info = _current_sa_info()
    if src in ("secrets","env","file","embedded"):
        badge = {"secrets":"🔒 Secrets","env":"🌿 ENV","file":"📄 File","embedded":"📦 Embedded"}[src]
        st.caption(f"✅ Service Account พร้อมใช้งานจาก **{badge}**")
        return True

    st.warning("ยังไม่พบ Service Account ใน Secrets/ENV/File/Embedded")
    file = st.file_uploader("อัปโหลดไฟล์ service_account.json (ครั้งเดียว)", type=["json"])
    if file:
        with open(CREDENTIALS_FILE, "wb") as f:
            f.write(file.getbuffer())
        st.success("บันทึกไฟล์แล้ว กำลังรีเฟรช...")
        st.rerun()
    st.stop()

# -----------------------------
# Safe wrappers for opening sheets
# -----------------------------
@st.cache_resource(show_spinner=False)
def open_sheet_by_url(sheet_url: str):
    return get_client().open_by_url(sheet_url)

@st.cache_resource(show_spinner=False)
def open_sheet_by_key(key: str):
    return get_client().open_by_key(key)

# -----------------------------
# Utilities
# -----------------------------
def fmt_dt(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")

def get_now_str():
    return fmt_dt(datetime.now(TZ))

def combine_date_time(d: date, t: dtime) -> datetime:
    naive = datetime.combine(d, t)
    return TZ.localize(naive)

def get_username():
    return (
        st.session_state.get("user")
        or st.session_state.get("username")
        or st.session_state.get("display_name")
        or "unknown"
    )

def setup_responsive():
    st.markdown(MINIMAL_CSS, unsafe_allow_html=True)

# -----------------------------
# Cached worksheet reads (by key/url + title for hashability)
# -----------------------------
@st.cache_data(ttl=60, show_spinner=False)
def _records_by_key(sheet_key: str, ws_title: str):
    sh = open_sheet_by_key(sheet_key)
    ws = sh.worksheet(ws_title)
    return ws.get_all_records()

@st.cache_data(ttl=60, show_spinner=False)
def _records_by_url(sheet_url: str, ws_title: str):
    sh = open_sheet_by_url(sheet_url)
    ws = sh.worksheet(ws_title)
    return ws.get_all_records()

def read_df(sh, sheet_name: str, headers=None):
    """Read worksheet to DataFrame with caching and safe columns."""
    try:
        key = getattr(sh, "id", None) or getattr(sh, "spreadsheet_id", None)
    except Exception:
        key = None
    url = st.session_state.get("sheet_url", "")

    if key:
        recs = _records_by_key(str(key), str(sheet_name))
    elif url:
        recs = _records_by_url(str(url), str(sheet_name))
    else:
        # Fallback direct
        ws = sh.worksheet(sheet_name)
        recs = ws.get_all_records()

    df = pd.DataFrame(recs)
    if headers:
        for h in headers:
            if h not in df.columns:
                df[h] = ""
        try:
            df = df[headers]
        except Exception:
            pass
    return df

def write_df(sh, title, df):
    if title==SHEET_ITEMS: cols=ITEMS_HEADERS
    elif title==SHEET_TXNS: cols=TXNS_HEADERS
    elif title==SHEET_USERS: cols=USERS_HEADERS
    elif title==SHEET_CATS: cols=CATS_HEADERS
    elif title==SHEET_BRANCHES: cols=BR_HEADERS
    elif title==SHEET_TICKETS: cols=TICKETS_HEADERS
    elif title==SHEET_TICKET_CATS: cols=TICKET_CAT_HEADERS
    else: cols = df.columns.tolist()

    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols]
    ws = sh.worksheet(title)
    ws.clear()
    ws.update([df.columns.values.tolist()] + df.fillna("").values.tolist())
    try:
        st.cache_data.clear()  # refresh caches after write
    except Exception:
        pass

def append_row(sh, title, row):
    sh.worksheet(title).append_row(row)
    try:
        st.cache_data.clear()
    except Exception:
        pass

# -----------------------------
# Sheet bootstrap
# -----------------------------
def ensure_sheets_exist(sh):
    required = [
        (SHEET_ITEMS, ITEMS_HEADERS, 1000, len(ITEMS_HEADERS)+5),
        (SHEET_TXNS, TXNS_HEADERS, 2000, len(TXNS_HEADERS)+5),
        (SHEET_USERS, USERS_HEADERS, 100, len(USERS_HEADERS)+2),
        (SHEET_CATS, CATS_HEADERS, 200, len(CATS_HEADERS)+2),
        (SHEET_BRANCHES, BR_HEADERS, 200, len(BR_HEADERS)+2),
        (SHEET_TICKETS, TICKETS_HEADERS, 1000, len(TICKETS_HEADERS)+5),
        (SHEET_TICKET_CATS, TICKET_CAT_HEADERS, 200, len(TICKET_CAT_HEADERS)+2),
    ]

    try:
        titles = [ws.title for ws in sh.worksheets()]
    except Exception:
        titles = None

    def ensure_one(name, headers, rows, cols):
        try:
            if titles is not None:
                if name in titles:
                    return
                ws = sh.add_worksheet(name, rows, cols)
                ws.append_row(headers)
            else:
                try:
                    sh.worksheet(name)
                except WorksheetNotFound:
                    ws = sh.add_worksheet(name, rows, cols)
                    ws.append_row(headers)
        except APIError as e:
            st.warning(f"สร้าง/ตรวจชีต '{name}' ไม่สำเร็จชั่วคราว: {e}")

    for name, headers, r, c in required:
        ensure_one(name, headers, r, c)

    # seed default admin
    try:
        ws_users = sh.worksheet(SHEET_USERS)
        values = ws_users.get_all_values()
        if len(values) <= 1:
            default_pwd = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            ws_users.append_row(["admin","Administrator","admin",default_pwd,"Y"])
    except Exception:
        pass

# -----------------------------
# PDF helpers (Thai fonts)
# -----------------------------
def register_thai_fonts() -> dict:
    cands = [
        ("ThaiFont", "./fonts/Sarabun-Regular.ttf", "./fonts/Sarabun-Bold.ttf"),
        ("ThaiFont", "./fonts/THSarabunNew.ttf", "./fonts/THSarabunNew-Bold.ttf"),
        ("ThaiFont", "./fonts/NotoSansThai-Regular.ttf", "./fonts/NotoSansThai-Bold.ttf"),
        ("ThaiFont", "C:/Windows/Fonts/Sarabun-Regular.ttf", "C:/Windows/Fonts/Sarabun-Bold.ttf"),
        ("ThaiFont", "C:/Windows/Fonts/THSarabunNew.ttf", "C:/Windows/Fonts/THSarabunNew-Bold.ttf"),
        ("ThaiFont", "C:/Windows/Fonts/NotoSansThai-Regular.ttf", "C:/Windows/Fonts/NotoSansThai-Bold.ttf"),
        ("ThaiFont", "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf"),
        ("ThaiFont", "/usr/share/fonts/truetype/sarabun/Sarabun-Regular.ttf", "/usr/share/fonts/truetype/sarabun/Sarabun-Bold.ttf"),
        ("ThaiFont", "/Library/Fonts/NotoSansThai-Regular.ttf", "/Library/Fonts/NotoSansThai-Bold.ttf"),
    ]
    for fam, normal, bold in cands:
        if os.path.exists(normal):
            try:
                pdfmetrics.registerFont(TTFont(fam, normal))
                bname = None
                if os.path.exists(bold):
                    bname = fam + "-Bold"
                    pdfmetrics.registerFont(TTFont(bname, bold))
                return {"normal": fam, "bold": bname}
            except Exception:
                continue
    return {"normal": None, "bold": None}

def df_to_pdf_bytes(df, title="รายงาน", subtitle=""):
    f = register_thai_fonts()
    use_thai = f["normal"] is not None
    if not use_thai:
        st.warning("⚠️ ไม่พบฟอนต์ไทย (Sarabun/TH Sarabun New/Noto Sans Thai). วางไฟล์ .ttf ในโฟลเดอร์ ./fonts เพื่อการแสดงผลสวยขึ้น")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=14, rightMargin=14, topMargin=14, bottomMargin=14)
    styles = getSampleStyleSheet()
    if use_thai:
        styles["Normal"].fontName = f["normal"]
        styles["Normal"].fontSize = 11
        styles["Normal"].leading = 14
        styles["Normal"].wordWrap = 'CJK'
        styles.add(ParagraphStyle(name="ThaiTitle", parent=styles["Title"],
                                  fontName=f["bold"] or f["normal"],
                                  fontSize=18, leading=22, wordWrap='CJK'))
        styles.add(ParagraphStyle(name="ThaiHeader", parent=styles["Normal"],
                                  fontName=f["bold"] or f["normal"],
                                  fontSize=12, leading=15, wordWrap='CJK'))
        title_style = styles["ThaiTitle"]; header_style = styles["ThaiHeader"]
    else:
        title_style = styles["Title"]; header_style = styles["Heading4"]

    story = [Paragraph(title, title_style)]
    if subtitle: story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1, 8))

    if df.empty:
        story.append(Paragraph("ไม่มีข้อมูล", styles["Normal"]))
    else:
        data = [df.columns.astype(str).tolist()] + df.astype(str).values.tolist()
        table = Table(data, repeatRows=1)
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
            ts.append(('FONTNAME', (0,0), (-1,0), f["bold"] or f["normal"]))
        table.setStyle(TableStyle(ts))
        story.append(table)

    doc.build(story)
    pdf = buf.getvalue(); buf.close()
    return pdf

# -----------------------------
# Business helpers
# -----------------------------
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
    return f"{cat_code}-{max_num+1:03d}"

def adjust_stock(sh, code, delta, actor, branch="", note="", txn_type="OUT", ts_str=None):
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    if items.empty or (items["รหัส"]==code).sum()==0:
        st.error("ไม่พบรหัสอุปกรณ์นี้ในคลัง"); return False
    row = items[items["รหัส"]==code].iloc[0]
    cur = int(float(row["คงเหลือ"])) if str(row["คงเหลือ"]).strip()!="" else 0
    if txn_type=="OUT" and cur+delta < 0:
        st.error("สต็อกไม่เพียงพอ"); return False
    items.loc[items["รหัส"]==code, "คงเหลือ"] = cur+delta; write_df(sh, SHEET_ITEMS, items)
    ts = ts_str if ts_str else get_now_str()
    append_row(sh, SHEET_TXNS, [str(uuid.uuid4())[:8], ts, txn_type, code, row["ชื่ออุปกรณ์"], branch, abs(delta), actor, note])
    return True

# -----------------------------
# Auth block
# -----------------------------
def auth_block(sh):
    st.session_state.setdefault("user", None)
    st.session_state.setdefault("role", None)
    if st.session_state.get("user"):
        with st.sidebar:
            st.markdown(f"**👤 {st.session_state['user']}**")
            st.caption(f"Role: {st.session_state['role']}")
            if st.button("ออกจากระบบ"): st.session_state["user"]=None; st.session_state["role"]=None; st.rerun()
        return True

    st.sidebar.subheader("เข้าสู่ระบบ")
    u = st.sidebar.text_input("Username")
    p = st.sidebar.text_input("Password", type="password")
    if st.sidebar.button("Login", use_container_width=True):
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
        row = users[(users["Username"]==u) & (users["Active"].str.upper()=="Y")]
        if not row.empty:
            ok = False
            try:
                ok = bcrypt.checkpw(p.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8"))
            except Exception:
                ok = False
            if ok:
                st.session_state["user"]=u
                st.session_state["role"]=row.iloc[0]["Role"]
                st.success("เข้าสู่ระบบสำเร็จ"); st.rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")
        else:
            st.error("ไม่พบบัญชีหรือถูกปิดใช้งาน")
    return False

# -----------------------------
# Pages
# -----------------------------
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

    # quick charts
    st.markdown("### คงเหลือตามหมวดหมู่")
    if not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("หมวดหมู่")["คงเหลือ"].sum().reset_index()
        tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
        chart = alt.Chart(tmp).mark_arc(innerRadius=60).encode(
            theta="คงเหลือ:Q", color="หมวดหมู่ชื่อ:N", tooltip=["หมวดหมู่ชื่อ","คงเหลือ"]
        )
        st.altair_chart(chart, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

def page_stock(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📦 คลังอุปกรณ์")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view_df = items.copy()
    if q and not items.empty:
        mask = items["รหัส"].str.contains(q, case=False, na=False) | \
               items["ชื่ออุปกรณ์"].str.contains(q, case=False, na=False) | \
               items["หมวดหมู่"].str.contains(q, case=False, na=False)
        view_df = items[mask]
    st.dataframe(view_df, height=320, use_container_width=True)

    if st.session_state.get("role") in ("admin","staff"):
        with st.form("item_add", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
                cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if cats.empty:
                    st.info("ยังไม่มีหมวดหมู่ในชีต Categories")
                    cat_opt = ""
                else:
                    opts = (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist()
                    selected = st.selectbox("หมวดหมู่", options=opts)
                    cat_opt = selected.split(" | ")[0]
                name = st.text_input("ชื่ออุปกรณ์")
            with c2:
                unit = st.text_input("หน่วย", value="ชิ้น")
                qty = st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
            with c3:
                loc = st.text_input("ที่เก็บ", value="IT Room")
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                auto_code = st.checkbox("สร้างรหัสอัตโนมัติ", value=True)
                code = st.text_input("รหัสอุปกรณ์ (ถ้าไม่ออโต้)", disabled=auto_code)
                s_add = st.form_submit_button("บันทึก/อัปเดต", use_container_width=True)
        if s_add:
            if (auto_code and not cat_opt) or (not auto_code and code.strip()==""):
                st.error("กรุณาเลือกหมวด/ระบุรหัส")
            else:
                items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                gen_code = generate_item_code(sh, cat_opt) if auto_code else code.strip().upper()
                if (items["รหัส"]==gen_code).any():
                    items.loc[items["รหัส"]==gen_code, ITEMS_HEADERS] = [gen_code, cat_opt, name, unit, qty, rop, loc, active]
                else:
                    items = pd.concat([items, pd.DataFrame([[gen_code, cat_opt, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True)
                write_df(sh, SHEET_ITEMS, items); st.success(f"บันทึกเรียบร้อย (รหัส: {gen_code})"); st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

def page_issue_out_multi5(sh):
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty:
        st.info("ยังไม่มีรายการอุปกรณ์"); return

    bopt = st.selectbox("สาขา/หน่วยงานผู้ขอ", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
    branch_code = bopt.split(" | ")[0] if bopt else ""

    st.markdown("**เลือกรายการที่ต้องการเบิก (ได้สูงสุด 5 รายการต่อครั้ง)**")
    opts = []
    for _, r in items.iterrows():
        remain = int(pd.to_numeric(r["คงเหลือ"], errors="coerce") or 0)
        opts.append(f'{r["รหัส"]} | {r["ชื่ออุปกรณ์"]} (คงเหลือ {remain})')

    df_template = pd.DataFrame({"รายการ": ["", "", "", "", ""], "จำนวน": [1,1,1,1,1]})
    ed = st.data_editor(
        df_template, use_container_width=True, hide_index=True, num_rows="fixed",
        column_config={"รายการ": st.column_config.SelectboxColumn(options=opts, required=False),
                       "จำนวน": st.column_config.NumberColumn(min_value=1, step=1)},
        key="issue_out_multi5"
    )
    note = st.text_input("หมายเหตุ (ถ้ามี)", value="")

    if st.button("บันทึกการเบิก (1 ครั้ง/หลายรายการ)", type="primary", disabled=(not branch_code)):
        txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
        errors = []; processed = 0; items_local = items.copy()

        for _, r in ed.iterrows():
            sel = str(r.get("รายการ","") or "").strip()
            qty = int(pd.to_numeric(r.get("จำนวน", 0), errors="coerce") or 0)
            if not sel or qty <= 0: continue

            code_sel = sel.split(" | ")[0]
            row_sel = items_local[items_local["รหัส"]==code_sel]
            if row_sel.empty:
                errors.append(f"{code_sel}: ไม่พบในคลัง"); continue
            row_sel = row_sel.iloc[0]
            remain = int(pd.to_numeric(row_sel["คงเหลือ"], errors="coerce") or 0)
            if qty > remain:
                errors.append(f"{code_sel}: เกินคงเหลือ ({remain})"); continue

            new_remain = remain - qty
            items_local.loc[items_local["รหัส"]==code_sel, "คงเหลือ"] = str(new_remain)
            txn = [str(uuid.uuid4())[:8], get_now_str(), "OUT", code_sel, row_sel["ชื่ออุปกรณ์"], branch_code, str(qty), get_username(), note]
            txns = pd.concat([txns, pd.DataFrame([txn], columns=TXNS_HEADERS)], ignore_index=True)
            processed += 1

        if processed > 0:
            write_df(sh, SHEET_ITEMS, items_local)
            write_df(sh, SHEET_TXNS, txns)
            st.success(f"บันทึกการเบิกแล้ว {processed} รายการ")
            st.rerun()
        else:
            st.warning("ยังไม่มีบรรทัดที่สมบูรณ์ให้บันทึก")

def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🧾 เบิก/รับเข้า")

    if st.session_state.get("role") not in ("admin","staff"):
        st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกรายการได้")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    if items.empty:
        st.warning("ยังไม่มีรายการอุปกรณ์ในคลัง")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    t1, t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])
    with t1: page_issue_out_multi5(sh)
    with t2:
        with st.form("recv", clear_on_submit=True):
            c1,c2 = st.columns([2,1])
            with c1: item = st.selectbox("เลือกอุปกรณ์", options=items["รหัส"]+" | "+items["ชื่ออุปกรณ์"], key="recv_item")
            with c2: qty = st.number_input("จำนวนที่รับเข้า", min_value=1, value=1, step=1, key="recv_qty")
            branch = st.text_input("แหล่งที่มา/เลข PO", key="recv_branch")
            note = st.text_input("หมายเหตุ", placeholder="เช่น ซื้อเข้า-เติมสต็อก", key="recv_note")
            manual_in = st.checkbox("กำหนดวันเวลาเอง", value=False, key="in_manual")
            if manual_in:
                d = st.date_input("วันที่", value=datetime.now(TZ).date(), key="in_d")
                t = st.time_input("เวลา", value=datetime.now(TZ).time().replace(microsecond=0), key="in_t")
                ts_str = fmt_dt(combine_date_time(d, t))
            else:
                ts_str = None
            s = st.form_submit_button("บันทึกรับเข้า", use_container_width=True)
        if s:
            code = item.split(" | ")[0]
            ok = adjust_stock(sh, code, qty, get_username(), branch, note, "IN", ts_str=ts_str)
            if ok: st.success("บันทึกรับเข้าแล้ว"); st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

def group_period(df, period="ME"):
    dfx = df.copy()
    dfx["วันเวลา"] = pd.to_datetime(dfx["วันเวลา"], errors='coerce')
    dfx = dfx.dropna(subset=["วันเวลา"])
    return dfx.groupby([pd.Grouper(key="วันเวลา", freq=period), "ประเภท", "ชื่ออุปกรณ์"])['จำนวน'].sum().reset_index()

def page_reports(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📑 รายงาน / ประวัติ")

    txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)

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
    with bcols[0]: st.button("วันนี้", on_click=_set_range, kwargs=dict(today=True), key="btn_today_r")
    with bcols[1]: st.button("7 วันล่าสุด", on_click=_set_range, kwargs=dict(days=7), key="btn_7d_r")
    with bcols[2]: st.button("30 วันล่าสุด", on_click=_set_range, kwargs=dict(days=30), key="btn_30d_r")
    with bcols[3]: st.button("90 วันล่าสุด", on_click=_set_range, kwargs=dict(days=90), key="btn_90d_r")
    with bcols[4]: st.button("เดือนนี้", on_click=_set_range, kwargs=dict(this_month=True), key="btn_month_r")
    with bcols[5]: st.button("ปีนี้", on_click=_set_range, kwargs=dict(this_year=True), key="btn_year_r")

    with st.expander("กำหนดช่วงวันที่เอง (เลือกแล้วกด 'ใช้ช่วงนี้')", expanded=False):
        d1m = st.date_input("วันที่เริ่ม (กำหนดเอง)", value=st.session_state["report_d1"], key="report_manual_d1_r")
        d2m = st.date_input("วันที่สิ้นสุด (กำหนดเอง)", value=st.session_state["report_d2"], key="report_manual_d2_r")
        st.button("ใช้ช่วงนี้", on_click=lambda: (st.session_state.__setitem__("report_d1", d1m),
                                                st.session_state.__setitem__("report_d2", d2m)),
                  key="btn_apply_manual_r")

    q = st.text_input("ค้นหา (ชื่อ/รหัส/สาขา)", key="report_query_r")

    d1 = st.session_state["report_d1"]; d2 = st.session_state["report_d2"]
    st.caption(f"ช่วงที่เลือก: **{d1} → {d2}**")

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

    tOut, tW, tM, tY = st.tabs(["รายละเอียดการเบิก (OUT)", "รายสัปดาห์", "รายเดือน", "รายปี"])

    with tOut:
        out_df = df_f[df_f["ประเภท"] == "OUT"].copy().sort_values("วันเวลา", ascending=False)
        cols = [c for c in ["วันเวลา", "ชื่ออุปกรณ์", "จำนวน", "สาขา", "ผู้ดำเนินการ", "หมายเหตุ", "รหัส"] if c in out_df.columns]
        st.dataframe(out_df[cols], height=320, use_container_width=True)
        pdf = df_to_pdf_bytes(out_df[cols].rename(columns={"วันเวลา":"วันที่-เวลา","ชื่ออุปกรณ์":"อุปกรณ์","จำนวน":"จำนวนที่เบิก","สาขา":"ปลายทาง"}),
                              title="รายละเอียดการเบิก (OUT)", subtitle=f"ช่วง {d1} ถึง {d2}")
        st.download_button("ดาวน์โหลด PDF รายละเอียดการเบิก", data=pdf, file_name="issue_detail_out.pdf", mime="application/pdf")

    with tW:
        g = group_period(df_f, "W"); st.dataframe(g, height=220, use_container_width=True)
        st.download_button("ดาวน์โหลด PDF รายสัปดาห์", data=df_to_pdf_bytes(g, "สรุปรายสัปดาห์", f"ช่วง {d1} ถึง {d2}"),
                           file_name="weekly_report.pdf", mime="application/pdf")

    with tM:
        g = group_period(df_f, "ME"); st.dataframe(g, height=220, use_container_width=True)
        st.download_button("ดาวน์โหลด PDF รายเดือน", data=df_to_pdf_bytes(g, "สรุปรายเดือน", f"ช่วง {d1} ถึง {d2}"),
                           file_name="monthly_report.pdf", mime="application/pdf")

    with tY:
        g = group_period(df_f, "YE"); st.dataframe(g, height=220, use_container_width=True)
        st.download_button("ดาวน์โหลด PDF รายปี", data=df_to_pdf_bytes(g, "สรุปรายปี", f"ช่วง {d1} ถึง {d2}"),
                           file_name="yearly_report.pdf", mime="application/pdf")

    st.markdown("</div>", unsafe_allow_html=True)

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

def page_import(sh):
    st.subheader("นำเข้า/แก้ไข หมวดหมู่ / เพิ่มข้อมูล (หมวดหมู่ / สาขา / อุปกรณ์ / หมวดหมู่ปัญหา / ผู้ใช้)")
    if st.session_state.get("role") not in ("admin","staff"):
        st.info("เฉพาะ admin/staff เท่านั้น"); return

    # Template downloads
    t1, t2, t3, t4 = st.columns(4)
    with t1:
        cat_csv = "รหัสหมวด,ชื่อหมวด\nPRT,หมึกพิมพ์\nKBD,คีย์บอร์ด\n"
        st.download_button("เทมเพลต หมวดหมู่ (CSV)", data=cat_csv.encode("utf-8-sig"), file_name="template_categories.csv", mime="text/csv", use_container_width=True)
    with t2:
        br_csv = "รหัสสาขา,ชื่อสาขา\nHQ,สำนักงานใหญ่\nBKK1,สาขาบางนา\n"
        st.download_button("เทมเพลต สาขา (CSV)", data=br_csv.encode("utf-8-sig"), file_name="template_branches.csv", mime="text/csv", use_container_width=True)
    with t3:
        it_csv = ",".join(ITEMS_HEADERS) + "\n" + "PRT-001,PRT,ตลับหมึก HP 206A,ตลับ,5,2,IT Room,Y\n"
        st.download_button("เทมเพลต อุปกรณ์ (CSV)", data=it_csv.encode("utf-8-sig"), file_name="template_items.csv", mime="text/csv", use_container_width=True)
    with t4:
        tkc_csv = "รหัสหมวดปัญหา,ชื่อหมวดปัญหา\nNW,Network\nPRN,Printer\nSW,Software\n"
        st.download_button("เทมเพลต หมวดหมู่ปัญหา (CSV)", data=tkc_csv.encode("utf-8-sig"), file_name="template_ticket_categories.csv", mime="text/csv", use_container_width=True)

    tab_cat, tab_br, tab_it, tab_tkc, tab_user = st.tabs(["หมวดหมู่","สาขา","อุปกรณ์","หมวดหมู่ปัญหา","ผู้ใช้"])

    # --- Categories ---
    with tab_cat:
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ (CSV/Excel)", type=["csv","xlsx"], key="up_cat")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=200, use_container_width=True)
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

    # --- Branches ---
    with tab_br:
        up = st.file_uploader("อัปโหลดไฟล์ สาขา (CSV/Excel)", type=["csv","xlsx"], key="up_br")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=200, use_container_width=True)
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

    # --- Items ---
    with tab_it:
        up = st.file_uploader("อัปโหลดไฟล์ อุปกรณ์ (CSV/Excel)", type=["csv","xlsx"], key="up_it")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=260, use_container_width=True)
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
                            try: qty = int(float(qty))
                            except: qty = 0
                            try: rop = int(float(rop))
                            except: rop = 0
                            qty = max(0, qty); rop = max(0, rop)
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

    # --- Ticket Categories ---
    with tab_tkc:
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ปัญหา (CSV/Excel)", type=["csv","xlsx"], key="up_tkc")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=200, use_container_width=True)
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

    # --- Users import ---
    with tab_user:
        up = st.file_uploader("อัปโหลดไฟล์ ผู้ใช้ (CSV/Excel)", type=["csv","xlsx"], key="up_users")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=200, use_container_width=True)
                if "Username" not in df.columns:
                    st.error("หัวตารางอย่างน้อยต้องมีคอลัมน์ Username")
                else:
                    if st.button("นำเข้า/อัปเดต ผู้ใช้", use_container_width=True, key="btn_imp_users"):
                        cur = read_df(sh, SHEET_USERS, USERS_HEADERS)
                        for c in USERS_HEADERS:
                            if c not in cur.columns: cur[c] = ""
                        cur = cur[USERS_HEADERS].fillna("")
                        add=upd=0; errs=[]
                        for i, r in df.iterrows():
                            username = str(r.get("Username","")).strip()
                            if username == "":
                                errs.append({"row":i+1,"error":"เว้นว่าง Username"}); continue
                            display = str(r.get("DisplayName","")).strip()
                            role    = str(r.get("Role","staff")).strip() or "staff"
                            active  = str(r.get("Active","Y")).strip() or "Y"
                            pwd_hash = None
                            plain = str(r.get("Password","")).strip() if "Password" in df.columns else ""
                            if plain:
                                try:
                                    pwd_hash = bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                                except Exception as e:
                                    errs.append({"row":i+1,"error":f"แฮชรหัสผ่านไม่สำเร็จ: {e}","Username":username}); continue
                            else:
                                if "PasswordHash" in df.columns:
                                    ph = str(r.get("PasswordHash","")).strip()
                                    if ph: pwd_hash = ph
                            if (cur["Username"]==username).any():
                                idx = cur.index[cur["Username"]==username][0]
                                cur.at[idx,"DisplayName"]=display
                                cur.at[idx,"Role"]=role
                                cur.at[idx,"Active"]=active
                                if pwd_hash: cur.at[idx,"PasswordHash"]=pwd_hash
                                upd+=1
                            else:
                                if not pwd_hash:
                                    errs.append({"row":i+1,"error":"ผู้ใช้ใหม่ต้องระบุ Password หรือ PasswordHash","Username":username}); continue
                                new_row = pd.DataFrame([{
                                    "Username": username,
                                    "DisplayName": display,
                                    "Role": role,
                                    "PasswordHash": pwd_hash,
                                    "Active": active,
                                }])
                                cur = pd.concat([cur, new_row], ignore_index=True); add+=1
                        write_df(sh, SHEET_USERS, cur)
                        st.success(f"เพิ่ม {add} ราย / อัปเดต {upd} ราย")
                        if errs: st.warning(pd.DataFrame(errs))

        tpl = "Username,DisplayName,Role,Active,Password\nuser001,คุณเอ,staff,Y,1234\n"
        st.download_button("เทมเพลต ผู้ใช้ (CSV)", data=tpl.encode("utf-8-sig"),
                           file_name="template_users.csv", mime="text/csv", use_container_width=True)

def page_users(sh):
    st.subheader("👥 ผู้ใช้ & สิทธิ์ (Admin)")
    try:
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    except Exception as e:
        st.error(f"โหลดข้อมูลผู้ใช้ไม่สำเร็จ: {e}"); return

    base_cols = USERS_HEADERS[:]
    for col in base_cols:
        if col not in users.columns: users[col] = ""
    users = users[base_cols].fillna("")

    st.markdown("#### 📋 รายชื่อผู้ใช้ (ติ๊ก 'เลือก' เพื่อแก้ไข)")
    chosen_username = None
    users_display = users.copy(); users_display["เลือก"] = False
    edited_table = st.data_editor(
        users_display[["เลือก","Username","DisplayName","Role","PasswordHash","Active"]],
        use_container_width=True, height=300, num_rows="fixed",
        column_config={"เลือก": st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือกผู้ใช้สำหรับแก้ไข")}
    )
    picked = edited_table[edited_table["เลือก"] == True]
    if not picked.empty: chosen_username = str(picked.iloc[0]["Username"])

    tab_add, tab_edit = st.tabs(["➕ เพิ่มผู้ใช้", "✏️ แก้ไขผู้ใช้"])

    with tab_add:
        with st.form("form_add_user", clear_on_submit=True):
            c1, c2 = st.columns([2,1])
            with c1:
                new_user = st.text_input("Username*")
                new_disp = st.text_input("Display Name")
            with c2:
                new_role = st.selectbox("Role", ["admin","staff","viewer"], index=1)
                new_active = st.selectbox("Active", ["Y","N"], index=0)
            new_pwd = st.text_input("กำหนดรหัสผ่าน*", type="password")
            btn_add = st.form_submit_button("บันทึกผู้ใช้ใหม่", use_container_width=True, type="primary")
        if btn_add:
            if not new_user.strip() or not new_pwd.strip():
                st.warning("กรุณากรอก Username และรหัสผ่าน"); st.stop()
            if (users["Username"] == new_user).any():
                st.error("มี Username นี้อยู่แล้ว"); st.stop()
            ph = bcrypt.hashpw(new_pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
            new_row = pd.DataFrame([{
                "Username": new_user.strip(),
                "DisplayName": new_disp.strip(),
                "Role": new_role,
                "PasswordHash": ph,
                "Active": new_active,
            }])
            users2 = pd.concat([users, new_row], ignore_index=True)
            try:
                write_df(sh, SHEET_USERS, users2)
                st.success("เพิ่มผู้ใช้สำเร็จ"); st.rerun()
            except Exception as e:
                st.error(f"เพิ่มผู้ใช้ไม่สำเร็จ: {e}")

    with tab_edit:
        default_user = st.session_state.get("edit_user","")
        if chosen_username:
            st.session_state["edit_user"] = chosen_username
            default_user = chosen_username

        sel = st.selectbox("เลือกผู้ใช้เพื่อแก้ไข", [""] + users["Username"].tolist(),
                           index=([""] + users["Username"].tolist()).index(default_user) if default_user in users["Username"].tolist() else 0)
        target_user = sel or ""
        if not target_user:
            st.info("ยังไม่ได้เลือกผู้ใช้สำหรับแก้ไข"); return

        row = users[users["Username"] == target_user]
        if row.empty:
            st.warning("ไม่พบผู้ใช้ที่เลือก"); return
        data = row.iloc[0].to_dict()

        with st.form("form_edit_user", clear_on_submit=False):
            c1, c2 = st.columns([2,1])
            with c1:
                username = st.text_input("Username", value=data["Username"], disabled=True)
                display  = st.text_input("Display Name", value=data["DisplayName"])
            with c2:
                role  = st.selectbox("Role", ["admin","staff","viewer"],
                                     index=["admin","staff","viewer"].index(data["Role"]) if data["Role"] in ["admin","staff","viewer"] else 1)
                active = st.selectbox("Active", ["Y","N"],
                                      index=["Y","N"].index(data["Active"]) if data["Active"] in ["Y","N"] else 0)
            pwd = st.text_input("ตั้ง/รีเซ็ตรหัสผ่าน (ปล่อยว่าง = ไม่เปลี่ยน)", type="password")
            c3, c4 = st.columns([1,1])
            btn_save = c3.form_submit_button("บันทึกการแก้ไข", use_container_width=True, type="primary")
            btn_del  = c4.form_submit_button("ลบผู้ใช้นี้", use_container_width=True)

        if btn_del:
            if username.lower() == "admin":
                st.error("ห้ามลบผู้ใช้ admin")
            else:
                users2 = users[users["Username"] != username]
                try:
                    write_df(sh, SHEET_USERS, users2)
                    st.success(f"ลบผู้ใช้ {username} แล้ว"); st.rerun()
                except Exception as e:
                    st.error(f"ลบไม่สำเร็จ: {e}")

        if btn_save:
            idx = users.index[users["Username"] == username][0]
            users.at[idx, "DisplayName"] = display
            users.at[idx, "Role"]        = role
            users.at[idx, "Active"]      = active
            if pwd.strip():
                ph = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                users.at[idx, "PasswordHash"] = ph
            try:
                write_df(sh, SHEET_USERS, users)
                st.success("บันทึกการแก้ไขเรียบร้อย"); st.rerun()
            except Exception as e:
                st.error(f"บันทึกไม่สำเร็จ: {e}")

def page_settings(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("⚙️ Settings")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ/ตรวจสอบชีตที่จำเป็น", use_container_width=True):
        try:
            sh2 = open_sheet_by_url(url); ensure_sheets_exist(sh2); st.success("เชื่อมต่อสำเร็จ พร้อมใช้งาน")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------
# Main
# -----------------------------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    setup_responsive()
    st.title(APP_TITLE); st.caption(APP_TAGLINE)

    # Credential handling (no uploader if Secrets/ENV/File/Embedded present)
    ensure_credentials_ui()

    # Sheet URL state
    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"):
        st.session_state["sheet_url"] = DEFAULT_SHEET_URL

    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", ["📊 Dashboard","📦 คลังอุปกรณ์","🛠️ แจ้งปัญหา","🧾 เบิก/รับเข้า","📑 รายงาน","👤 ผู้ใช้","นำเข้า/แก้ไข หมวดหมู่","⚙️ Settings"], index=0)

    # Open sheet (except when only Settings)
    if "Settings" in page:
        try:
            sh = open_sheet_by_url(st.session_state["sheet_url"])
        except Exception:
            sh = None
        page_settings(sh)
        st.caption(f"© 2025 IT Stock · Streamlit + Google Sheets · **{VERSION_DISPLAY}**")
        return

    sheet_url = st.session_state.get("sheet_url", DEFAULT_SHEET_URL)
    if not sheet_url:
        st.info("ไปที่เมนู **Settings** แล้ววาง Google Sheet URL ที่คุณเป็นเจ้าของ จากนั้นกดปุ่มทดสอบเชื่อมต่อ")
        return

    try:
        sh = open_sheet_by_url(sheet_url)
    except Exception as e:
        st.error(f"เปิดชีตไม่สำเร็จ: {e}")
        return

    ensure_sheets_exist(sh)
    auth_block(sh)

    if page.startswith("📊"): page_dashboard(sh)
    elif page.startswith("📦"): page_stock(sh)
    elif page.startswith("🛠️"): 
        # tickets page (ย่อ): แสดงตารางอย่างเดียวเพื่อให้โค้ดสั้น
        st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")
        st.dataframe(read_df(sh, SHEET_TICKETS, TICKETS_HEADERS), use_container_width=True, height=320)
    elif page.startswith("🧾"): page_issue_receive(sh)
    elif page.startswith("📑"): page_reports(sh)
    elif page.startswith("👤"): page_users(sh)
    elif page.startswith("นำเข้า"): page_import(sh)

    st.caption(f"© 2025 IT Stock · Streamlit + Google Sheets · **{VERSION_DISPLAY}**")

if __name__ == "__main__":
    main()
