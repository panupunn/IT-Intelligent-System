#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iTao iT Stock (Streamlit + Google Sheets)
Stable hotfix with restored multi-issue (OUT) flow.

- Restores Dashboard controls (Pie/Bar, Top-N, per-row, date range)
- Keeps all menus: Dashboard, Stock, Tickets, Issue/Receive, Reports, Users, Import/Modify, Settings
- Users/Stock/Tickets pages support "เลือก" checkbox in the top table to load row for editing
- Import submenus present: Categories, Branches, Items, Ticket Categories, Users (CSV/Excel)
- Service Account loading: secrets → env → local file → embedded B64 (no re-upload on wake if configured)
- Worksheet read caching via st.cache_data(ttl=60) to avoid quota spikes without altering behaviors
- Version footer: iTao iT (V.1.1)
"""
from __future__ import annotations
import os, io, uuid, re, time, base64, json
from datetime import datetime, date, timedelta, time as dtime
import pytz, pandas as pd, streamlit as st
import altair as alt
import bcrypt
import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials


# === PATCH: Requests integration constants ===
MENU_REQUESTS = "🧺 คำขอเบิก"
REQUESTS_SHEET = "Requests"
NOTIFS_SHEET   = "Notifications"

REQUESTS_HEADERS = [
    "Branch","Requester","CreatedAt","OrderNo","ItemCode","ItemName","Qty",
    "Status","Approver","LastUpdate","Note"
]
NOTIFS_HEADERS = [
    "NotiID","CreatedAt","TargetApp","TargetBranch","Type","RefID","Message","ReadFlag","ReadAt"
]
# === END PATCH ===

# -------------------- Global constants --------------------
APP_TITLE   = "ไอต้าว ไอที (iTao iT)"
APP_TAGLINE = "POWER By ทีมงาน=> ไอทีสุดหล่อ"
TZ = pytz.timezone("Asia/Bangkok")

DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit?gid=0#gid=0"
CREDENTIALS_FILE  = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")

SHEET_ITEMS       = "Items"
SHEET_TXNS        = "Transactions"
SHEET_USERS       = "Users"
SHEET_CATS        = "Categories"
SHEET_BRANCHES    = "Branches"
SHEET_TICKETS     = "Tickets"
SHEET_TICKET_CATS = "TicketCategories"

ITEMS_HEADERS     = ["รหัส","หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
TXNS_HEADERS      = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"]
USERS_HEADERS     = ["Username","DisplayName","Role","PasswordHash","Active"]
CATS_HEADERS      = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS        = ["รหัสสาขา","ชื่อสาขา"]
TICKETS_HEADERS   = ["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]
TICKET_CAT_HEADERS= ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]

MINIMAL_CSS = """
<style>
:root { --radius: 16px; }
section.main > div { padding-top: 8px; }
.block-card { background: #fff; border:1px solid #eee; border-radius:16px; padding:16px; }
.kpi { display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.danger { color:#b00020; }
</style>"""

# --------- Embedded credentials (optional) ---------
# Put base64 of your service_account.json here to avoid re-upload.
EMBEDDED_GOOGLE_CREDENTIALS_B64 = os.environ.get("EMBEDDED_SA_B64", "").strip()

# --- Defensive shim for cache_resource (older Streamlit) ---
if not hasattr(st, "cache_resource"):
    def _no_cache_decorator(*args, **kwargs):
        def _wrap(func): return func
        return _wrap
    st.cache_resource = _no_cache_decorator

# -------------------- Credentials handling --------------------
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _try_load_sa_from_secrets():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
        if "service_account" in st.secrets:
            return dict(st.secrets["service_account"])
        if "service_account_json" in st.secrets:
            raw = str(st.secrets["service_account_json"])
            return json.loads(raw)
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
        if raw.lstrip().startswith("{"):
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
    b64 = EMBEDDED_GOOGLE_CREDENTIALS_B64
    if not b64: return None
    try:
        return json.loads(base64.b64decode(b64).decode("utf-8"))
    except Exception:
        return None

def _ensure_credentials_available():
    info = (_try_load_sa_from_secrets()
            or _try_load_sa_from_env()
            or _try_load_sa_from_file()
            or _try_load_sa_from_embedded())
    if info is not None:
        return ("dict", info)
    # Last resort: if no persistent source -> allow manual upload once
    up = st.file_uploader("อัปโหลดไฟล์ service_account.json", type=["json"], key="sa_json_once")
    if not up:
        st.stop()
    try:
        return ("dict", json.loads(up.getvalue().decode("utf-8")))
    except Exception as e:
        st.error(f"อ่านไฟล์ service_account.json ไม่สำเร็จ: {e}")
        st.stop()

@st.cache_resource(show_spinner=False)
def get_client():
    mode, info = _ensure_credentials_available()
    if mode == "dict":
        creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)

# Convenience wrappers
@st.cache_resource(show_spinner=False)
def open_sheet_by_url(sheet_url: str):
    return get_client().open_by_url(sheet_url)

@st.cache_resource(show_spinner=False)
def open_sheet_by_key(sheet_key: str):
    return get_client().open_by_key(sheet_key)

# -------------------- Cached worksheet reads --------------------
def _open_sheet_by_key_nocache(sheet_key: str):
    return get_client().open_by_key(sheet_key)

def _open_sheet_by_url_nocache(sheet_url: str):
    return get_client().open_by_url(sheet_url)

@st.cache_data(ttl=60, show_spinner=False)
def _cached_ws_records_by_key(sheet_key: str, ws_title: str):
    sh = _open_sheet_by_key_nocache(sheet_key)
    ws = sh.worksheet(ws_title)
    return ws.get_all_records()

@st.cache_data(ttl=60, show_spinner=False)
def _cached_ws_records_by_url(sheet_url: str, ws_title: str):
    sh = _open_sheet_by_url_nocache(sheet_url)
    ws = sh.worksheet(ws_title)
    return ws.get_all_records()

def clear_read_cache():
    try:
        st.cache_data.clear()
    except Exception:
        pass

# -------------------- Utility helpers --------------------
def fmt_dt(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")

def get_now_str():
    return fmt_dt(datetime.now(TZ))

def combine_date_time(d: date, t: dtime) -> datetime:
    naive = datetime.combine(d, t)
    return TZ.localize(naive)

def read_df(sh, sheet_name: str, headers=None) -> pd.DataFrame:
    """Read a worksheet into DataFrame with caching if possible."""
    sheet_key = getattr(sh, "id", None) or getattr(sh, "spreadsheet_id", None) or ""
    sheet_url = st.session_state.get("sheet_url", "") or ""

    if sheet_key:
        records = _cached_ws_records_by_key(str(sheet_key), str(sheet_name))
    elif sheet_url:
        records = _cached_ws_records_by_url(str(sheet_url), str(sheet_name))
    else:
        ws = sh.worksheet(sheet_name)
        records = ws.get_all_records()

    df = pd.DataFrame(records)
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
    ws.update([df.columns.values.tolist()] + df.values.tolist())
    clear_read_cache()

def append_row(sh, title, row):
    sh.worksheet(title).append_row(row)
    clear_read_cache()

def ensure_credentials_ui():
    # No-op when credentials are already resolved via get_client()
    return True

def setup_responsive():
    st.markdown("""
    <style>
    @media (max-width: 640px) {
        .block-container { padding: 0.6rem 0.7rem !important; }
        [data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; padding-right: 0 !important; }
        .stButton > button, .stSelectbox, .stTextInput, .stTextArea, .stDateInput { width: 100% !important; }
        .stDataFrame { width: 100% !important; }
        .js-plotly-plot, .vega-embed { width: 100% !important; }
    }
    </style>
    """, unsafe_allow_html=True)

def get_username():
    return (
        st.session_state.get("user")
        or st.session_state.get("username")
        or st.session_state.get("display_name")
        or "unknown"
    )

# -------------------- Ensure worksheets exist --------------------
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
    except APIError:
        titles = None

    def ensure_one(name, headers, rows, cols):
        try:
            if titles is not None and name in titles:
                return
            # Verify directly
            try:
                sh.worksheet(name)
            except WorksheetNotFound:
                ws = sh.add_worksheet(name, rows, cols)
                ws.append_row(headers)
        except APIError as e:
            st.warning(f"ไม่สามารถตรวจสอบ/สร้างชีต '{name}': {e}")

    for name, headers, r, c in required:
        ensure_one(name, headers, r, c)

    # Seed admin user if empty
    try:
        ws_users = sh.worksheet(SHEET_USERS)
        values = ws_users.get_all_values()
        if len(values) <= 1:
            default_pwd = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            ws_users.append_row(["admin","Administrator","admin",default_pwd,"Y"])
    except Exception:
        pass

# -------------------- Auth block --------------------
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
            if ok:
                st.session_state["user"]=u; st.session_state["role"]=row.iloc[0]["Role"]; st.success("เข้าสู่ระบบสำเร็จ"); st.rerun()
            else: st.error("รหัสผ่านไม่ถูกต้อง")
        else: st.error("ไม่พบบัญชีหรือถูกปิดใช้งาน")
    return False

# -------------------- Dashboard --------------------
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

    # Prepare txns OUT filtered
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

    # Tickets for charts
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
        rows = (len(charts) + per_row - 1) // per_row
        idx = 0
        for r in range(rows):
            cols = st.columns(per_row)
            for c in range(per_row):
                if idx >= len(charts): break
                title, df, label_col, value_col = charts[idx]
                with cols[c]:
                    if chart_kind.endswith('(Bar)'):
                        make_bar(df, label_col, value_col, top_n, title)
                    else:
                        make_pie(df, label_col, value_col, top_n, title)
                idx += 1

    # Low stock list
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

    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- Stock page --------------------
def get_unit_options(items_df):
    opts = sorted([x for x in items_df["หน่วย"].dropna().astype(str).unique() if x.strip()!=""])
    if "ชิ้น" not in opts: opts = ["ชิ้น"] + opts
    return opts + ["พิมพ์เอง"]

def get_loc_options(items_df):
    opts = sorted([x for x in items_df["ที่เก็บ"].dropna().astype(str).unique() if x.strip()!=""])
    if "IT Room" not in opts: opts = ["IT Room"] + opts
    return opts + ["พิมพ์เอง"]

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
            except:
                pass
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

def page_stock(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📦 คลังอุปกรณ์")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view_df = items.copy()
    if q and not items.empty:
        mask = items["รหัส"].str.contains(q, case=False, na=False) | items["ชื่ออุปกรณ์"].str.contains(q, case=False, na=False) | items["หมวดหมู่"].str.contains(q, case=False, na=False)
        view_df = items[mask]
    # Selectable table
    chosen_code = None
    if hasattr(st, "data_editor"):
        show = view_df.copy()
        show["เลือก"] = False
        ed = st.data_editor(show[["เลือก","รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","หมวดหมู่","หน่วย","ใช้งาน"]],
                            use_container_width=True, height=320, num_rows="fixed",
                            column_config={"เลือก": st.column_config.CheckboxColumn()})
        picked = ed[ed["เลือก"]==True]
        if not picked.empty:
            chosen_code = str(picked.iloc[0]["รหัส"])
    else:
        st.dataframe(view_df, height=320, use_container_width=True)

    unit_opts = get_unit_options(items)
    loc_opts  = get_loc_options(items)

    if st.session_state.get("role") in ("admin","staff"):
        t_add, t_edit = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ (เลือกรายการเดิม)"])

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
                    write_df(sh, SHEET_ITEMS, items); st.success(f"บันทึกเรียบร้อย (รหัส: {gen_code})"); st.rerun()

        with t_edit:
            st.caption("เลือก 'รหัสอุปกรณ์' จากตารางด้านบนเพื่อโหลดขึ้นมาปรับแก้ หรือเลือกจากลิสต์")
            labels = [f'{r["รหัส"]} | {r["ชื่ออุปกรณ์"]}' for _, r in items.iterrows()] if not items.empty else []
            if chosen_code and any(x.startswith(chosen_code+" |") for x in labels):
                default_idx = labels.index(next(x for x in labels if x.startswith(chosen_code+" |")))
            else:
                default_idx = 0 if labels else None
            pick_label = st.selectbox("เลือกรหัสอุปกรณ์", options=(["-- เลือก --"]+labels) if labels else ["-- เลือก --"], index=(default_idx+1 if default_idx is not None else 0))
            if pick_label != "-- เลือก --":
                pick = pick_label.split(" | ", 1)[0]
                row = items[items["รหัส"] == pick].iloc[0]
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
                    write_df(sh, SHEET_ITEMS, items); st.success("อัปเดตแล้ว"); st.rerun()
                if s_del:
                    items = items[items["รหัส"]!=pick]; write_df(sh, SHEET_ITEMS, items); st.success(f"ลบ {pick} แล้ว"); st.rerun()

# -------------------- Tickets page (unchanged UI + fixes) --------------------
def generate_ticket_id() -> str:
    return "TCK-" + datetime.now(TZ).strftime("%Y%m%d-%H%M%S")

def page_tickets(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")

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
        d1 = st.date_input("วันที่เริ่ม", value=(date.today()-timedelta(days=90)), key="tk_d1")
    with dcol2:
        d2 = st.date_input("วันที่สิ้นสุด", value=date.today(), key="tk_d2")

    view = tickets.copy()
    if not view.empty:
        view["วันที่แจ้ง"] = pd.to_datetime(view["วันที่แจ้ง"], errors="coerce")
        view = view.dropna(subset=["วันที่แจ้ง"])
        view = view[(view["วันที่แจ้ง"].dt.date >= st.session_state["tk_d1"]) & (view["วันที่แจ้ง"].dt.date <= st.session_state["tk_d2"])]
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

    st.markdown("### รายการแจ้งปัญหา (ติ๊กเลือกเพื่อแก้ไข)")
    chosen_tid = None
    if hasattr(st, "data_editor"):
        show = view.copy()
        show["เลือก"] = False
        ed = st.data_editor(
            show[["เลือก","TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]],
            use_container_width=True, height=300, num_rows="fixed",
            column_config={"เลือก": st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือก Ticket")}
        )
        picked = ed[ed["เลือก"] == True]
        if not picked.empty:
            chosen_tid = str(picked.iloc[0]["TicketID"])
    else:
        st.dataframe(view.sort_values("วันที่แจ้ง", ascending=False), height=300, use_container_width=True)

    st.markdown("---")
    t_add, t_update = st.tabs(["➕ รับแจ้งใหม่","🔁 เปลี่ยนสถานะ/แก้ไข"])

    with t_add:
        with st.form("tk_new", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
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
            st.rerun()

    with t_update:
        if tickets.empty:
            st.info("ยังไม่มีรายการในชีต Tickets")
        else:
            labels = [f'{r["TicketID"]} | {str(r["สาขา"])}' for _, r in tickets.iterrows()]
            if chosen_tid and any(x.startswith(chosen_tid+" |") for x in labels):
                default_idx = labels.index(next(x for x in labels if x.startswith(chosen_tid+" |")))
            else:
                default_idx = 0
            pick_label = st.selectbox("เลือก Ticket", options=["-- เลือก --"] + labels, index=default_idx+1 if labels else 0, key="tk_pick")
            if pick_label != "-- เลือก --":
                pick_id = pick_label.split(" | ", 1)[0]
                row = tickets[tickets["TicketID"] == pick_id].iloc[0]
                st.subheader(f"แก้ไข Ticket: {pick_id}")
                with st.form("tk_edit", clear_on_submit=False):
                    c1, c2 = st.columns(2)
                    with c1:
                        t_branch = st.text_input("สาขา", value=str(row.get("สาขา", "")))
                        t_type = st.selectbox("ประเภท", ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"], index=3)
                    with c2:
                        t_owner = st.text_input("ผู้แจ้ง", value=str(row.get("ผู้แจ้ง","")))
                        statuses_edit = ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"]
                        try:
                            idx_default = statuses_edit.index(str(row.get("สถานะ","รับแจ้ง")))
                        except ValueError:
                            idx_default = 0
                        t_status = st.selectbox("สถานะ", statuses_edit, index=idx_default)
                        t_assignee = st.text_input("ผู้รับผิดชอบ", value=str(row.get("ผู้รับผิดชอบ","")))
                    t_desc = st.text_area("รายละเอียด", value=str(row.get("รายละเอียด","")), height=120)
                    t_note = st.text_input("หมายเหตุ", value=str(row.get("หมายเหตุ","")))
                    fcol1, fcol2, fcol3 = st.columns(3)
                    submit_update = fcol1.form_submit_button("อัปเดต")
                    submit_delete = fcol3.form_submit_button("ลบรายการ")

                if submit_update:
                    try:
                        idx = tickets.index[tickets["TicketID"] == pick_id][0]
                        tickets.at[idx, "สาขา"] = t_branch
                        tickets.at[idx, "ผู้แจ้ง"] = t_owner
                        tickets.at[idx, "รายละเอียด"] = t_desc
                        tickets.at[idx, "สถานะ"] = t_status
                        tickets.at[idx, "ผู้รับผิดชอบ"] = t_assignee
                        tickets.at[idx, "หมายเหตุ"] = t_note
                        tickets.at[idx, "อัปเดตล่าสุด"] = get_now_str()
                        write_df(sh, SHEET_TICKETS, tickets)
                        st.success("อัปเดตสถานะ/รายละเอียดเรียบร้อย")
                        st.rerun()
                    except Exception as e:
                        st.error(f"อัปเดตไม่สำเร็จ: {e}")
                if submit_delete:
                    try:
                        tickets2 = tickets[tickets["TicketID"] != pick_id].copy()
                        write_df(sh, SHEET_TICKETS, tickets2)
                        st.success("ลบรายการเรียบร้อย")
                        st.rerun()
                    except Exception as e:
                        st.error(f"ลบไม่สำเร็จ: {e}")

# -------------------- Issue/Receive page (RESTORED multi-issue) --------------------
def page_issue_out_multiN(sh):
    """เบิก (OUT): เลือกสาขาก่อน แล้วกรอกได้หลายรายการในครั้งเดียว (จำนวนบรรทัดกำหนดได้)"""
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    if items.empty:
        st.info("ยังไม่มีรายการอุปกรณ์", icon="ℹ️"); return

    # 1) เลือกสาขา/หน่วยงานผู้ขอ
    bopt = st.selectbox("สาขา/หน่วยงานผู้ขอ", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
    branch_code = bopt.split(" | ")[0] if bopt else ""

    # 2) ตั้งค่าแถวที่จะกรอก
    n_rows = st.slider("จำนวนแถวสำหรับเบิกครั้งนี้", 1, 50, 5, 1)
    st.markdown("**เลือกรายการที่ต้องการเบิก (หลายรายการต่อครั้ง)**")

    # เตรียม options แสดงคงเหลือ
    opts = []
    for _, r in items.iterrows():
        remain = int(pd.to_numeric(r["คงเหลือ"], errors="coerce") or 0)
        opts.append(f'{r["รหัส"]} | {r["ชื่ออุปกรณ์"]} (คงเหลือ {remain})')

    df_template = pd.DataFrame({"รายการ": [""]*n_rows, "จำนวน": [1]*n_rows})
    ed = st.data_editor(
        df_template,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "รายการ": st.column_config.SelectboxColumn(options=opts, required=False),
            "จำนวน": st.column_config.NumberColumn(min_value=1, step=1)
        },
        key="issue_out_multiN",
    )

    note = st.text_input("หมายเหตุ (ถ้ามี)", value="")
    manual_out = st.checkbox("กำหนดวันเวลาเอง (OUT)", value=False, key="out_manual")
    if manual_out:
        d = st.date_input("วันที่ (OUT)", value=datetime.now(TZ).date(), key="out_d")
        t = st.time_input("เวลา (OUT)", value=datetime.now(TZ).time().replace(microsecond=0), key="out_t")
        ts_str = fmt_dt(combine_date_time(d, t))
    else:
        ts_str = None

    if st.button("บันทึกการเบิก (หลายรายการ)", type="primary", disabled=(not branch_code)):
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

            txn = [str(uuid.uuid4())[:8], ts_str if ts_str else get_now_str(),
                   "OUT", code_sel, row_sel["ชื่ออุปกรณ์"], branch_code, str(qty), get_username(), note]
            txns = pd.concat([txns, pd.DataFrame([txn], columns=TXNS_HEADERS)], ignore_index=True)
            processed += 1

        if processed > 0:
            write_df(sh, SHEET_ITEMS, items_local)
            write_df(sh, SHEET_TXNS, txns)
            st.success(f"บันทึกการเบิกแล้ว {processed} รายการ ✅")
            st.rerun()
        else:
            st.warning("ยังไม่มีบรรทัดที่สมบูรณ์ให้บันทึก", icon="⚠️")
        if errors:
            st.warning(pd.DataFrame({"ข้อผิดพลาด": errors}))

def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🧾 เบิก/รับเข้า")
    if st.session_state.get("role") not in ("admin","staff"):
        st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกรายการได้")
        st.markdown("</div>", unsafe_allow_html=True); return

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty:
        st.warning("ยังไม่มีรายการอุปกรณ์ในคลัง")
        st.markdown("</div>", unsafe_allow_html=True); return

    t1,t2 = st.tabs(["เบิก (OUT) — หลายรายการต่อครั้ง","รับเข้า (IN)"])

    with t1:
        page_issue_out_multiN(sh)

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
            if ok: st.success("บันทึกรับเข้าแล้ว"); st.rerun()

# -------------------- Reports page --------------------
def is_test_text(s: str) -> bool:
    s = str(s).lower()
    return ("test" in s) or ("ทดสอบ" in s)

def page_reports(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📑 รายงาน / ประวัติ")

    txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    br_map = {str(r["รหัสสาขา"]).strip(): f'{str(r["รหัสสาขา"]).strip()} | {str(r["ชื่อสาขา"]).strip()}' for _, r in branches.iterrows()} if not branches.empty else {}

    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)

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
    with bcols[0]: st.button("วันนี้", on_click=_set_range, kwargs=dict(today=True))
    with bcols[1]: st.button("7 วันล่าสุด", on_click=_set_range, kwargs=dict(days=7))
    with bcols[2]: st.button("30 วันล่าสุด", on_click=_set_range, kwargs=dict(days=30))
    with bcols[3]: st.button("90 วันล่าสุด", on_click=_set_range, kwargs=dict(days=90))
    with bcols[4]: st.button("เดือนนี้", on_click=_set_range, kwargs=dict(this_month=True))
    with bcols[5]: st.button("ปีนี้", on_click=_set_range, kwargs=dict(this_year=True))

    with st.expander("กำหนดช่วงวันที่เอง (เลือกแล้วกด 'ใช้ช่วงนี้')", expanded=False):
        d1m = st.date_input("วันที่เริ่ม (กำหนดเอง)", value=st.session_state["report_d1"])
        d2m = st.date_input("วันที่สิ้นสุด (กำหนดเอง)", value=st.session_state["report_d2"])
        st.button("ใช้ช่วงนี้", on_click=lambda: (st.session_state.__setitem__("report_d1", d1m),
                                                st.session_state.__setitem__("report_d2", d2m)))

    q = st.text_input("ค้นหา (ชื่อ/รหัส/สาขา/เรื่อง)")

    d1 = st.session_state["report_d1"]
    d2 = st.session_state["report_d2"]
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
            if "เรื่อง" in tdf.columns:
                mask_t = mask_t | tdf["เรื่อง"].astype(str).str.contains(q, case=False, na=False)
            tdf = tdf[mask_t]
        if "เรื่อง" not in tdf.columns:
            def _derive_subject(x):
                s = str(x or "").strip().splitlines()[0]
                return s[:60] if s else "ไม่ระบุเรื่อง"
            tdf["เรื่อง"] = tdf["รายละเอียด"].apply(_derive_subject)
    else:
        tdf = pd.DataFrame(columns=TICKETS_HEADERS + ["เรื่อง"])

    tOut, tTickets, tW, tM, tY = st.tabs(["รายละเอียดการเบิก (OUT)", "ประวัติการแจ้งปัญหา", "รายสัปดาห์", "รายเดือน", "รายปี"])

    with tOut:
        out_df = df_f[df_f["ประเภท"] == "OUT"].copy().sort_values("วันเวลา", ascending=False)
        cols = [c for c in ["วันเวลา", "ชื่ออุปกรณ์", "จำนวน", "สาขา", "ผู้ดำเนินการ", "หมายเหตุ", "รหัส"] if c in out_df.columns]
        st.dataframe(out_df[cols], height=320, use_container_width=True)
        # --- ADD: พิมพ์ตาราง OUT เป็น PDF (ไม่แตะส่วนอื่น) ---
        with st.expander("🖨️ พิมพ์รายงาน OUT เป็น PDF", expanded=False):
            up_logo = st.file_uploader("โลโก้ (PNG/JPG) — ไม่บังคับ", type=["png","jpg","jpeg"], key="logo_out")
            logo_path = ""
            if up_logo is not None:
                import os
                os.makedirs("./assets", exist_ok=True)
                logo_path = "./assets/_logo_report_out.png"
                with open(logo_path, "wb") as f:
                    f.write(up_logo.read())

            def _register_thai_fonts_if_needed():
                try:
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont
                    import os
                    if "TH_REG" in pdfmetrics.getRegisteredFontNames():
                        return True
                    candidates = [
                        "./fonts/THSarabunNew.ttf", "./fonts/Sarabun-Regular.ttf", "./fonts/NotoSansThai-Regular.ttf"
                    ]
                    for p in ("/usr/share/fonts/truetype", "/usr/share/fonts", "/Library/Fonts", "C:\\Windows\\Fonts"):
                        for fn in ("THSarabunNew.ttf","Sarabun-Regular.ttf","NotoSansThai-Regular.ttf"):
                            candidates.append(os.path.join(p, fn))
                    for reg in candidates:
                        try:
                            if os.path.exists(reg):
                                pdfmetrics.registerFont(TTFont("TH_REG", reg))
                                pdfmetrics.registerFont(TTFont("TH_BOLD", reg))
                                return True
                        except:
                            pass
                except Exception:
                    return False
                return False

            def _make_pdf_from_df(title, df, logo_path=""):
                try:
                    import io
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import A4, landscape
                    from reportlab.lib.units import mm
                    from reportlab.lib.utils import ImageReader
                    from reportlab.pdfbase import pdfmetrics
                    from datetime import datetime as _dt

                    buf = io.BytesIO()
                    c = canvas.Canvas(buf, pagesize=landscape(A4))
                    W, H = landscape(A4)

                    if logo_path:
                        try:
                            c.drawImage(ImageReader(logo_path), 15*mm, H-35*mm, width=25*mm, height=25*mm, preserveAspectRatio=True, mask='auto')
                        except Exception:
                            pass

                    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 16)
                    c.drawString(45*mm, H-20*mm, str(title))
                    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)
                    c.drawRightString(W-15*mm, H-15*mm, _dt.now().strftime("%Y-%m-%d %H:%M:%S"))

                    cols_pdf = df.columns.tolist()[:8]
                    x0, y0 = 15*mm, H-45*mm
                    row_h = 8*mm
                    col_w = (W - 30*mm) / max(1, len(cols_pdf))

                    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 10)
                    for i, col in enumerate(cols_pdf):
                        c.drawString(x0 + i*col_w + 2, y0, str(col))
                    c.line(x0, y0-2, x0 + col_w*len(cols_pdf), y0-2)

                    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)
                    y = y0 - row_h
                    for r in df[cols_pdf].astype(str).values.tolist()[:50]:
                        for i, val in enumerate(r):
                            c.drawString(x0 + i*col_w + 2, y, val[:40])
                        y -= row_h
                        if y < 20*mm:
                            break

                    c.showPage()
                    c.save()
                    buf.seek(0)
                    return buf.getvalue()
                except Exception as e:
                    st.error(f"สร้าง PDF ไม่สำเร็จ: {e}")
                    return None

            if st.button("สร้าง PDF (OUT)", key="btn_pdf_out"):
                try:
                    import reportlab
                except Exception:
                    st.error("ต้องติดตั้งแพ็กเกจ reportlab ก่อนใช้งาน:  pip install reportlab")
                else:
                    _register_thai_fonts_if_needed()
                    pdf_bytes = _make_pdf_from_df(f"รายการเบิก (OUT) {d1} → {d2}", out_df[cols], logo_path=logo_path)
                    if pdf_bytes:
                        st.download_button(
                            "⬇️ ดาวน์โหลด PDF (OUT)",
                            data=pdf_bytes,
                            file_name=f"report_out_{d1}_{d2}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

    with tTickets:
        st.markdown("#### ตารางรายการแจ้งปัญหา")
        show_cols = [c for c in ["วันที่แจ้ง","เรื่อง","รายละเอียด","สาขา","ผู้แจ้ง","สถานะ","ผู้รับผิดชอบ","หมายเหตุ","TicketID"] if c in tdf.columns]
        tdf_sorted = tdf.sort_values("วันที่แจ้ง", ascending=False)
        st.dataframe(tdf_sorted[show_cols], height=320, use_container_width=True)
        # --- ADD: พิมพ์ตาราง Tickets เป็น PDF (ไม่แตะส่วนอื่น) ---
        with st.expander("🖨️ พิมพ์รายงาน Tickets เป็น PDF", expanded=False):
            up_logo2 = st.file_uploader("โลโก้ (PNG/JPG) — ไม่บังคับ", type=["png","jpg","jpeg"], key="logo_tk")
            logo_path2 = ""
            if up_logo2 is not None:
                import os
                os.makedirs("./assets", exist_ok=True)
                logo_path2 = "./assets/_logo_report_tickets.png"
                with open(logo_path2, "wb") as f:
                    f.write(up_logo2.read())

            def _register_thai_fonts_if_needed_tk():
                try:
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont
                    import os
                    if "TH_REG" in pdfmetrics.getRegisteredFontNames():
                        return True
                    candidates = [
                        "./fonts/THSarabunNew.ttf", "./fonts/Sarabun-Regular.ttf", "./fonts/NotoSansThai-Regular.ttf"
                    ]
                    for p in ("/usr/share/fonts/truetype", "/usr/share/fonts", "/Library/Fonts", "C:\\Windows\\Fonts"):
                        for fn in ("THSarabunNew.ttf","Sarabun-Regular.ttf","NotoSansThai-Regular.ttf"):
                            candidates.append(os.path.join(p, fn))
                    for reg in candidates:
                        try:
                            if os.path.exists(reg):
                                pdfmetrics.registerFont(TTFont("TH_REG", reg))
                                pdfmetrics.registerFont(TTFont("TH_BOLD", reg))
                                return True
                        except:
                            pass
                except Exception:
                    return False
                return False

            def _make_pdf_from_df_tk(title, df, logo_path=""):
                try:
                    import io
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import A4, landscape
                    from reportlab.lib.units import mm
                    from reportlab.lib.utils import ImageReader
                    from reportlab.pdfbase import pdfmetrics
                    from datetime import datetime as _dt

                    buf = io.BytesIO()
                    c = canvas.Canvas(buf, pagesize=landscape(A4))
                    W, H = landscape(A4)

                    if logo_path:
                        try:
                            c.drawImage(ImageReader(logo_path), 15*mm, H-35*mm, width=25*mm, height=25*mm, preserveAspectRatio=True, mask='auto')
                        except Exception:
                            pass

                    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 16)
                    c.drawString(45*mm, H-20*mm, str(title))
                    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)
                    c.drawRightString(W-15*mm, H-15*mm, _dt.now().strftime("%Y-%m-%d %H:%M:%S"))

                    cols_pdf = df.columns.tolist()[:8]
                    x0, y0 = 15*mm, H-45*mm
                    row_h = 8*mm
                    col_w = (W - 30*mm) / max(1, len(cols_pdf))

                    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 10)
                    for i, col in enumerate(cols_pdf):
                        c.drawString(x0 + i*col_w + 2, y0, str(col))
                    c.line(x0, y0-2, x0 + col_w*len(cols_pdf), y0-2)

                    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)
                    y = y0 - row_h
                    for r in df[cols_pdf].astype(str).values.tolist()[:50]:
                        for i, val in enumerate(r):
                            c.drawString(x0 + i*col_w + 2, y, val[:40])
                        y -= row_h
                        if y < 20*mm:
                            break

                    c.showPage()
                    c.save()
                    buf.seek(0)
                    return buf.getvalue()
                except Exception as e:
                    st.error(f"สร้าง PDF ไม่สำเร็จ: {e}")
                    return None

            if st.button("สร้าง PDF (Tickets)", key="btn_pdf_tickets"):
                try:
                    import reportlab
                except Exception:
                    st.error("ต้องติดตั้งแพ็กเกจ reportlab ก่อนใช้งาน:  pip install reportlab")
                else:
                    _register_thai_fonts_if_needed_tk()
                    pdf_bytes = _make_pdf_from_df_tk(f"ประวัติการแจ้งปัญหา {d1} → {d2}", tdf_sorted[show_cols], logo_path=logo_path2)
                    if pdf_bytes:
                        st.download_button(
                            "⬇️ ดาวน์โหลด PDF (Tickets)",
                            data=pdf_bytes,
                            file_name=f"report_tickets_{d1}_{d2}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

    def group_period(df, period="ME"):
        dfx = df.copy()
        dfx["วันเวลา"] = pd.to_datetime(dfx["วันเวลา"], errors='coerce')
        dfx = dfx.dropna(subset=["วันเวลา"])
        return dfx.groupby([pd.Grouper(key="วันเวลา", freq=period), "ประเภท", "ชื่ออุปกรณ์"])["จำนวน"].sum().reset_index()

    with tW:
        g = group_period(df_f, "W")
        st.dataframe(g, height=220, use_container_width=True)

    with tM:
        g = group_period(df_f, "ME")
        st.dataframe(g, height=220, use_container_width=True)

    with tY:
        g = group_period(df_f, "YE")
        st.dataframe(g, height=220, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- Import/Modify page --------------------
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
    st.subheader("นำเข้า/แก้ไข หมวดหมู่ / สาขา / อุปกรณ์ / หมวดหมู่ปัญหา / ผู้ใช้")
    t1, t2, t3, t4, t5 = st.tabs(["หมวดหมู่","สาขา","อุปกรณ์","หมวดหมู่ปัญหา","ผู้ใช้"])

    # หมวดหมู่
    with t1:
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ (CSV/Excel)", type=["csv","xlsx"], key="up_cat")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(50), height=220, use_container_width=True)
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

        with st.form("form_add_cat", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1: code_c = st.text_input("รหัสหมวด*", max_chars=10)
            with col2: name_c = st.text_input("ชื่อหมวด*")
            s = st.form_submit_button("เพิ่มหมวดหมู่", use_container_width=True)
        if s:
            if not code_c or not name_c: st.warning("กรอกข้อมูลให้ครบ")
            else:
                cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if (cur["รหัสหมวด"]==code_c).any(): st.error("มีรหัสนี้อยู่แล้ว")
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_c.strip(), name_c.strip()]], columns=CATS_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_CATS, cur); st.success("เพิ่มสำเร็จ")

    # สาขา
    with t2:
        up = st.file_uploader("อัปโหลดไฟล์ สาขา (CSV/Excel)", type=["csv","xlsx"], key="up_br")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(50), height=220, use_container_width=True)
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

        with st.form("form_add_branch", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1: code_b = st.text_input("รหัสสาขา*", max_chars=10)
            with col2: name_b = st.text_input("ชื่อสาขา*")
            s2 = st.form_submit_button("เพิ่มสาขา", use_container_width=True)
        if s2:
            if not code_b or not name_b: st.warning("กรอกข้อมูลให้ครบ")
            else:
                cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
                if (cur["รหัสสาขา"]==code_b).any(): st.error("มีรหัสนี้อยู่แล้ว")
                else:
                    cur = pd.concat([cur, pd.DataFrame([[code_b.strip(), name_b.strip()]], columns=BR_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_BRANCHES, cur); st.success("เพิ่มสำเร็จ")

    # อุปกรณ์
    with t3:
        up = st.file_uploader("อัปโหลดไฟล์ อุปกรณ์ (CSV/Excel)", type=["csv","xlsx"], key="up_it")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(50), height=260, use_container_width=True)
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

    # หมวดหมู่ปัญหา
    with t4:
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ปัญหา (CSV/Excel)", type=["csv","xlsx"], key="up_tkc")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(50), height=220, use_container_width=True)
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

    # ผู้ใช้
    with t5:
        up = st.file_uploader("อัปโหลดไฟล์ ผู้ใช้ (CSV/Excel)", type=["csv","xlsx"], key="up_users")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(50), height=220, use_container_width=True)
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
                            if username=="":
                                errs.append({"row":i+1,"error":"เว้นว่าง Username"}); 
                                continue
                            display = str(r.get("DisplayName","")).strip()
                            role    = str(r.get("Role","staff")).strip() or "staff"
                            active  = str(r.get("Active","Y")).strip() or "Y"
                            pwd_hash = None
                            plain = str(r.get("Password","")).strip() if "Password" in df.columns else ""
                            if plain:
                                try:
                                    pwd_hash = bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                                except Exception as e:
                                    errs.append({"row":i+1,"error":f"แฮชรหัสผ่านไม่สำเร็จ: {e}","Username":username}); 
                                    continue
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
                                    errs.append({"row":i+1,"error":"ผู้ใช้ใหม่ต้องระบุ Password หรือ PasswordHash","Username":username}); 
                                    continue
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

        st.markdown("##### เทมเพลตไฟล์")
        tpl = "Username,DisplayName,Role,Active,Password\nuser001,คุณเอ,staff,Y,1234\n"
        st.download_button("เทมเพลต ผู้ใช้ (CSV)", data=tpl.encode("utf-8-sig"),
                           file_name="template_users.csv", mime="text/csv", use_container_width=True)

# -------------------- Users page (select row to edit) --------------------
def page_users(sh):
    st.subheader("👥 ผู้ใช้ & สิทธิ์ (Admin)")
    users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    for c in USERS_HEADERS:
        if c not in users.columns: users[c] = ""
    users = users[USERS_HEADERS].fillna("")

    st.markdown("#### 📋 รายชื่อผู้ใช้ (ติ๊ก 'เลือก' เพื่อแก้ไข)")
    chosen_username = None
    if hasattr(st, "data_editor"):
        users_display = users.copy()
        users_display["เลือก"] = False
        edited_table = st.data_editor(
            users_display[["เลือก","Username","DisplayName","Role","PasswordHash","Active"]],
            use_container_width=True, height=300, num_rows="fixed",
            column_config={"เลือก": st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือกผู้ใช้สำหรับแก้ไข")}
        )
        picked = edited_table[edited_table["เลือก"] == True]
        if not picked.empty:
            chosen_username = str(picked.iloc[0]["Username"])
    else:
        st.dataframe(users, use_container_width=True, height=300)

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

        sel = st.selectbox(
            "เลือกผู้ใช้เพื่อแก้ไข",
            [""] + users["Username"].tolist(),
            index=([""] + users["Username"].tolist()).index(default_user) if default_user in users["Username"].tolist() else 0
        )

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
                    st.success(f"ลบผู้ใช้ {username} แล้ว")
                    st.session_state.pop("edit_user", None)
                    st.rerun()
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
                st.success("บันทึกการแก้ไขเรียบร้อย")
                st.rerun()
            except Exception as e:
                st.error(f"บันทึกไม่สำเร็จ: {e}")

# -------------------- Settings --------------------
def page_settings():
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("⚙️ Settings"); st.caption("ตรวจสอบว่าได้แชร์ Google Sheet ให้ service account แล้ว")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ/ตรวจสอบชีตที่จำเป็น", use_container_width=True):
        try:
            sh = open_sheet_by_url(url); ensure_sheets_exist(sh); st.success("เชื่อมต่อสำเร็จ พร้อมใช้งาน")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- Main --------------------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    st.markdown(MINIMAL_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE); st.caption(APP_TAGLINE)
    setup_responsive()

    ensure_credentials_ui()
    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"):
        st.session_state["sheet_url"] = DEFAULT_SHEET_URL

    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", ["📊 Dashboard","📦 คลังอุปกรณ์","🛠️ แจ้งปัญหา","🧾 เบิก/รับเข้า","📑 รายงาน","👤 ผู้ใช้","นำเข้า/แก้ไข หมวดหมู่","⚙️ Settings"], index=0)
    # PATCH: direct route for Requests menu
    if isinstance(page, str) and (page == MENU_REQUESTS or page.startswith('🧺')):
        page_requests(sh)
        return

    sheet_url = st.session_state.get("sheet_url", DEFAULT_SHEET_URL)
    if not sheet_url:
        st.info("ไปที่เมนู **Settings** แล้ววาง Google Sheet URL ที่คุณเป็นเจ้าของ จากนั้นกดปุ่มทดสอบเชื่อมต่อ"); return
    try:
        sh = open_sheet_by_url(sheet_url)
    except Exception as e:
        st.error(f"เปิดชีตไม่สำเร็จ: {e}"); return
    ensure_sheets_exist(sh)

    auth_block(sh)

    if page.startswith("📊"): page_dashboard(sh)
    elif page.startswith("📦"): page_stock(sh)
    elif page.startswith("🛠️"): page_tickets(sh)
    elif page.startswith("🧾"): page_issue_receive(sh)
    elif page.startswith("📑"): page_reports(sh)
    elif page.startswith("👤"): page_users(sh)
    elif page.startswith("นำเข้า"): page_import(sh)
    elif page.startswith("⚙️"): page_settings()

    st.caption("© 2025 IT Stock · Streamlit + Google Sheets By AOD. · **iTao iT (V.1.1)**")

if __name__ == "__main__":
    main()


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Requests Page Hotfix for iTao iT (Streamlit + Google Sheets)
- Fixes AttributeError: "Can only use .str accessor with string values!"
- Makes column handling robust (Status/Qty/Branch/OrderNo names)
- Adds helpers to ensure Requests / Notifications sheets exist
- Safe to paste into main app (replace existing functions/const) or import functions from here
"""
import streamlit as st
import pandas as pd
from datetime import datetime
import uuid

# ---------- Constants (use these names in main app) ----------
MENU_REQUESTS = "🧺 คำขอเบิก"
REQUESTS_SHEET = "Requests"
NOTIFS_SHEET = "Notifications"

REQUESTS_HEADERS = [
    "Branch","Requester","CreatedAt","OrderNo",
    "ItemCode","ItemName","Qty",
    "Status","Approver","LastUpdate","Note"
]
NOTIFS_HEADERS = [
    "NotiID","CreatedAt","TargetApp","TargetBranch","Type","RefID","Message","ReadFlag","ReadAt"
]

# ---------- Utilities ----------
def _lower_cols(df: pd.DataFrame):
    return {c: str(c).strip().lower() for c in df.columns}

def _col(df: pd.DataFrame, *candidates: str, default: str | None = None) -> str:
    """
    Resolve a column name by trying multiple candidates (case/space insensitive).
    """
    if df.empty:
        # return the first candidate if df has no cols (won't be used but avoids crash)
        return candidates[0]
    mapping = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in mapping:
            return mapping[key]
        # allow thai/english common variations
    # fallback
    return mapping.get(default, default or candidates[0])

def ensure_requests_notifs_sheets(sh):
    """Create Requests / Notifications with headers if missing."""
    ws_names = [w.title for w in sh.worksheets()]
    def _ensure(name: str, headers: list[str]):
        if name not in ws_names:
            ws = sh.add_worksheet(name, rows=1000, cols=max(12, len(headers)+2))
            try:
                import gspread_dataframe as gd
                gd.set_with_dataframe(ws, pd.DataFrame(columns=headers), include_index=False)
            except Exception:
                # fallback: write header row by update
                ws.update("A1", [headers])

    _ensure(REQUESTS_SHEET, REQUESTS_HEADERS)
    _ensure(NOTIFS_SHEET, NOTIFS_HEADERS)

def _normalize_requests_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=REQUESTS_HEADERS)

    # normalize basic fields
    # map typical variants to standard names
    cols = {c.lower(): c for c in df.columns}
    def getcol(*cands, default=None):
        for c in cands:
            if c.lower() in cols:
                return cols[c.lower()]
        return default if default in df.columns else cands[0]

    # Create a copy and fill na
    df = df.copy()
    df = df.fillna("")
    # normalize Status as UPPER string ('' for missing)
    if "Status" in df.columns:
        df["Status"] = (
            df["Status"]
            .astype(str)
            .fillna("")
            .map(lambda x: "" if x.strip().lower() in ("nan","none","null") else x)
            .str.upper()
            .str.strip()
        )
    else:
        df["Status"] = ""

    # normalize Qty numeric
    qty_col = getcol("Qty","จำนวน","quantity","qty")
    if qty_col in df.columns:
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0).astype(int)
        if qty_col != "Qty":
            df["Qty"] = df[qty_col]

    # standardize essential columns if alias exists
    mapping = {
        "Branch": ("Branch","สาขา","branchcode","branch_code"),
        "Requester": ("Requester","ผู้ขอ","ขอโดย","requester_name"),
        "CreatedAt": ("CreatedAt","created_at","วันเวลา","วันที่"),
        "OrderNo": ("OrderNo","order_no","เลขที่คำขอ","orderid"),
        "ItemCode": ("ItemCode","รหัส","item_code","code"),
        "ItemName": ("ItemName","ชื่ออุปกรณ์","name","item_name"),
        "Qty": ("Qty","จำนวน","qty"),
        "Status": ("Status","สถานะ","status"),
        "Approver": ("Approver","ผู้อนุมัติ","approved_by"),
        "LastUpdate": ("LastUpdate","อัปเดตล่าสุด","updated_at"),
        "Note": ("Note","หมายเหตุ","note"),
    }
    for std, cands in mapping.items():
        if std not in df.columns:
            for cand in cands:
                if cand in df.columns:
                    df[std] = df[cand]
                    break
        # ensure column exists
        if std not in df.columns:
            df[std] = ""

    # keep only the standard order for writing back
    df = df[REQUESTS_HEADERS]
    return df

def _write_df(ws, df: pd.DataFrame):
    """Write dataframe keeping header first row."""
    try:
        import gspread_dataframe as gd
        gd.set_with_dataframe(ws, df, include_index=False)
    except Exception:
        ws.clear()
        ws.update("A1", [list(df.columns)] + df.astype(str).values.tolist())

def _append_notifications(sh, rows_df: pd.DataFrame, message: str):
    ws = sh.worksheet(NOTIFS_SHEET)
    try:
        current = pd.DataFrame(ws.get_all_records())
    except Exception:
        current = pd.DataFrame(columns=NOTIFS_HEADERS)

    if current.empty:
        current = pd.DataFrame(columns=NOTIFS_HEADERS)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new = []
    for _, r in rows_df.iterrows():
        new.append({
            "NotiID": str(uuid.uuid4())[:8],
            "CreatedAt": now,
            "TargetApp": "branch",
            "TargetBranch": r.get("Branch",""),
            "Type": "request",
            "RefID": r.get("OrderNo",""),
            "Message": message,
            "ReadFlag": "",
            "ReadAt": ""
        })
    current = pd.concat([current, pd.DataFrame(new)], ignore_index=True)
    _write_df(ws, current)

def page_requests(sh):
    """Requests approval page (robust, NaN-safe)"""
    ensure_requests_notifs_sheets(sh)

    ws = sh.worksheet(REQUESTS_SHEET)
    try:
        raw = ws.get_all_records()
    except Exception:
        raw = []

    df = pd.DataFrame(raw)
    df = _normalize_requests_df(df)

    st.header("🧺 คำขอเบิก (จากสาขา)")
    if df.empty:
        st.info("ยังไม่มีคำขอ")
        return

    # ---- SAFE filter for pending ----
    status = df["Status"].astype(str).fillna("").str.upper().str.strip()
    pending = df[(status=="") | (status=="PENDING")].copy()
    if pending.empty:
        st.success("ไม่มีคำขอที่รออนุมัติ")
        return

    # group by order
    order_nos = pending["OrderNo"].astype(str).fillna("").unique().tolist()
    sel = st.selectbox("เลือก OrderNo", order_nos, index=0 if order_nos else None)
    this = pending[pending["OrderNo"].astype(str) == str(sel)].copy()

    if this.empty:
        st.info("ไม่พบรายการภายใต้ OrderNo ที่เลือก")
        return

    left,right = st.columns([2,1])
    with left:
        st.write(f"**สาขา:** {this['Branch'].iloc[0]}  |  **ผู้ขอ:** {this['Requester'].iloc[0]}  |  **จำนวนรายการ:** {len(this)}")
        st.dataframe(this[["ItemCode","ItemName","Qty"]], use_container_width=True)
    with right:
        st.metric("รวมจำนวนเบิก", int(this["Qty"].sum()))

    c1, c2 = st.columns(2)
    if c1.button("✅ อนุมัติและตัดสต็อก", use_container_width=True):
        _approve_request_and_cut_stock(sh, this)
        _append_notifications(sh, this, "คำขอได้รับการอนุมัติแล้ว")
        st.success("อนุมัติสำเร็จ")
        st.experimental_rerun()

    if c2.button("❌ ปฏิเสธ", use_container_width=True):
        _update_requests_status(sh, this, "REJECTED")
        _append_notifications(sh, this, "คำขอถูกปฏิเสธ")
        st.warning("ปฏิเสธแล้ว")
        st.experimental_rerun()

def _call_adjust_stock_if_exists(sh, r):
    # Try to reuse original stock issue function provided by main app
    fn = st.session_state.get("_adjust_stock_func")
    if fn:
        try:
            fn(sh, item_code=r["ItemCode"], qty=int(r["Qty"]), txn_type="OUT",
               branch=r.get("Branch",""), actor=st.session_state.get("username","system"),
               note=f"Request {r.get('OrderNo','')}")
            return True
        except Exception:
            pass

    # fallback: append minimal OUT txn row (Type/Datetime etc.)
    try:
        ws_txn = sh.worksheet("Transactions")
        existing = pd.DataFrame(ws_txn.get_all_records())
    except Exception:
        existing = pd.DataFrame(columns=["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"])
    if existing.empty:
        existing = pd.DataFrame(columns=["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_row = {
        "TxnID": str(uuid.uuid4())[:8],
        "วันเวลา": now,
        "ประเภท": "OUT",
        "รหัส": r.get("ItemCode",""),
        "ชื่ออุปกรณ์": r.get("ItemName",""),
        "สาขา": r.get("Branch",""),
        "จำนวน": int(r.get("Qty",0)),
        "ผู้ดำเนินการ": st.session_state.get("username","system"),
        "หมายเหตุ": f"Request {r.get('OrderNo','')}"
    }
    existing = pd.concat([existing, pd.DataFrame([new_row])], ignore_index=True)
    _write_df(ws_txn, existing)
    return True

def _approve_request_and_cut_stock(sh, rows_df: pd.DataFrame):
    # reuse stock adjust if main app provided it in session_state
    for _, r in rows_df.iterrows():
        _call_adjust_stock_if_exists(sh, r)
    _update_requests_status(sh, rows_df, "FULFILLED")

def _update_requests_status(sh, rows_df: pd.DataFrame, new_status: str):
    ws = sh.worksheet(REQUESTS_SHEET)
    try:
        df = pd.DataFrame(ws.get_all_records())
    except Exception:
        df = pd.DataFrame(columns=REQUESTS_HEADERS)

    df = _normalize_requests_df(df)

    for _, r in rows_df.iterrows():
        mask = (df["OrderNo"].astype(str) == str(r.get("OrderNo",""))) &                (df["ItemCode"].astype(str) == str(r.get("ItemCode","")))
        df.loc[mask, "Status"] = new_status
        df.loc[mask, "Approver"] = st.session_state.get("username","system")
        df.loc[mask, "LastUpdate"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    _write_df(ws, df)

