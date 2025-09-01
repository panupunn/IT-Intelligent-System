#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
iTao iT – IT Stock (Streamlit + Google Sheets)
Restored full menu + dashboard controls
Version: V.1.1
"""
import os, io, re, base64, json, uuid, time
from datetime import datetime, date, timedelta, time as dtime

import streamlit as st
import pandas as pd
import altair as alt
import bcrypt
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

# ========== Compatibility shims ==========
if not hasattr(st, "cache_resource"):
    def _no_cache_decorator(*args, **kwargs):
        def _wrap(fn): return fn
        return _wrap
    st.cache_resource = _no_cache_decorator

# ---------- App constants ----------
APP_TITLE = "ไอต้าว ไอที (iTao iT)"
APP_TAGLINE = "POWER By ทีมงาน=> ไอทีสุดหล่อ"
DEFAULT_SHEET_URL = st.secrets.get("google_sheet_url", "") or \
    "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit#gid=0"
TZ = pytz.timezone("Asia/Bangkok")

# Sheets & headers
SHEET_ITEMS      = "Items"
SHEET_TXNS       = "Transactions"
SHEET_USERS      = "Users"
SHEET_CATS       = "Categories"
SHEET_BRANCHES   = "Branches"
SHEET_TICKETS    = "Tickets"
SHEET_TICKET_CATS= "TicketCategories"

ITEMS_HEADERS = ["รหัส","หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
TXNS_HEADERS  = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"]
USERS_HEADERS = ["Username","DisplayName","Role","PasswordHash","Active"]
CATS_HEADERS  = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS    = ["รหัสสาขา","ชื่อสาขา"]
TICKETS_HEADERS = ["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]
TICKET_CAT_HEADERS = ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]

MINIMAL_CSS = """
<style>
:root { --radius: 14px; }
section.main > div { padding-top: 6px; }
.block-card { background: #fff; border:1px solid #eee; border-radius:16px; padding:16px; }
.kpi { display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
</style>
"""

# ========== Credentials loader (secrets -> env -> file; NO uploader) ==========
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _try_secrets():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
        if "service_account" in st.secrets:
            sa = st.secrets["service_account"]
            if isinstance(sa, dict):
                return dict(sa)
            if isinstance(sa, str):
                return json.loads(sa)
    except Exception:
        pass
    return None

def _try_env():
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

def _try_file():
    for p in ("./service_account.json","/mnt/data/service_account.json","/mount/data/service_account.json"):
        try:
            with open(p,"r",encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            continue
    return None

@st.cache_resource(show_spinner=False)
def get_client():
    info = _try_secrets() or _try_env() or _try_file()
    if not info:
        st.error("ไม่พบ Service Account ใน st.secrets / ENV หรือไฟล์ service_account.json")
        st.stop()
    creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)

# Open helpers
@st.cache_resource(show_spinner=False)
def open_sheet_by_url(url: str):
    return get_client().open_by_url(url)

@st.cache_resource(show_spinner=False)
def open_sheet_by_key(key: str):
    return get_client().open_by_key(key)

# ========== Sheet bootstrap ==========
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
        current = [ws.title for ws in sh.worksheets()]
    except Exception:
        current = []
    for name, headers, r, c in required:
        if name not in current:
            try:
                ws = sh.add_worksheet(name, r, c)
                ws.append_row(headers)
            except Exception:
                pass
    # seed admin
    try:
        ws_u = sh.worksheet(SHEET_USERS)
        vals = ws_u.get_all_values()
        if len(vals)<=1:
            ph = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            ws_u.append_row(["admin","Administrator","admin",ph,"Y"])
    except Exception:
        pass

# ========== Data helpers ==========
@st.cache_data(ttl=60, show_spinner=False)
def _records_by_sheet(url_or_id: str, ws_name: str):
    # url_or_id can be URL or key
    gc = get_client()
    sh = gc.open_by_url(url_or_id) if url_or_id.startswith("http") else gc.open_by_key(url_or_id)
    ws = sh.worksheet(ws_name)
    return ws.get_all_records()

def read_df(sh, ws_name: str, headers=None) -> pd.DataFrame:
    key = getattr(sh,"id",None) or getattr(sh,"spreadsheet_id",None) or st.session_state.get("sheet_url","")
    try:
        rows = _records_by_sheet(str(key if key else st.session_state.get("sheet_url","")), ws_name)
    except Exception:
        ws = sh.worksheet(ws_name)
        rows = ws.get_all_records()
    df = pd.DataFrame(rows)
    if headers:
        for h in headers:
            if h not in df.columns: df[h]=""
        try: df = df[headers]
        except Exception: pass
    return df

def write_df(sh, ws_name: str, df: pd.DataFrame):
    ws = sh.worksheet(ws_name)
    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").values.tolist())
    try: st.cache_data.clear()
    except Exception: pass

def append_row(sh, ws_name: str, row: list):
    sh.worksheet(ws_name).append_row(row)
    try: st.cache_data.clear()
    except Exception: pass

# ========== Common utils ==========
def fmt_dt(dt: datetime): return dt.strftime("%Y-%m-%d %H:%M:%S")
def now_str(): return fmt_dt(datetime.now(TZ))

def register_thai_fonts():
    cand = [
        ("ThaiFont","./fonts/Sarabun-Regular.ttf","./fonts/Sarabun-Bold.ttf"),
        ("ThaiFont","./fonts/THSarabunNew.ttf","./fonts/THSarabunNew-Bold.ttf"),
        ("ThaiFont","/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf","/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf"),
    ]
    for fam,n,b in cand:
        if os.path.exists(n):
            try:
                pdfmetrics.registerFont(TTFont(fam, n))
                bname=None
                if os.path.exists(b):
                    bname=fam+"-Bold"
                    pdfmetrics.registerFont(TTFont(bname, b))
                return fam, bname
            except Exception:
                pass
    return None, None

def df_to_pdf_bytes(df, title="รายงาน", subtitle=""):
    fam, bold = register_thai_fonts()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=14,rightMargin=14,topMargin=14,bottomMargin=14)
    styles = getSampleStyleSheet()
    if fam:
        styles["Normal"].fontName = fam
        styles["Normal"].fontSize = 11
        styles.add(ParagraphStyle(name="ThaiTitle", parent=styles["Title"], fontName=bold or fam, fontSize=18))
        title_style = styles["ThaiTitle"]
    else:
        title_style = styles["Title"]
    story=[Paragraph(title, title_style)]
    if subtitle: story.append(Paragraph(subtitle, styles["Normal"]))
    story.append(Spacer(1,8))
    data=[df.columns.astype(str).tolist()] + df.astype(str).values.tolist() if not df.empty else [["ไม่มีข้อมูล"]]
    table=Table(data, repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#f3f3f3')),
                               ('GRID',(0,0),(-1,-1),0.25,colors.grey),
                               ('ALIGN',(0,0),(-1,-1),'CENTER')]))
    story.append(table)
    doc.build(story)
    pdf=buf.getvalue(); buf.close(); return pdf

# ========== Auth block ==========
def auth_block(sh) -> bool:
    st.session_state.setdefault("user", None)
    st.session_state.setdefault("role", None)
    if st.session_state["user"]:
        with st.sidebar:
            st.markdown(f"**👤 {st.session_state['user']}**")
            st.caption(f"Role: {st.session_state['role']}")
            if st.button("ออกจากระบบ"): st.session_state["user"]=None; st.session_state["role"]=None; st.experimental_rerun()
        return True
    st.sidebar.subheader("เข้าสู่ระบบ")
    u = st.sidebar.text_input("Username")
    p = st.sidebar.text_input("Password", type="password")
    if st.sidebar.button("Login", use_container_width=True):
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
        row = users[(users["Username"]==u) & (users["Active"].str.upper()=="Y")]
        if not row.empty:
            ok=False
            try: ok=bcrypt.checkpw(p.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8"))
            except Exception: ok=False
            if ok:
                st.session_state["user"]=u
                st.session_state["role"]=row.iloc[0]["Role"]
                st.success("เข้าสู่ระบบสำเร็จ"); st.experimental_rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")
        else:
            st.error("ไม่พบบัญชีหรือถูกปิดใช้งาน")
    return False

# ========== Pages ==========
def page_dashboard(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📊 Dashboard")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    # KPIs
    total_items = len(items)
    total_qty = pd.to_numeric(items.get("คงเหลือ", pd.Series(dtype=float)), errors="coerce").fillna(0).sum() if not items.empty else 0
    low_df = pd.DataFrame(columns=ITEMS_HEADERS)
    if not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp["จุดสั่งซื้อ"] = pd.to_numeric(tmp["จุดสั่งซื้อ"], errors="coerce").fillna(0)
        low_df = tmp[(tmp["ใช้งาน"].str.upper()=="Y") & (tmp["คงเหลือ"] <= tmp["จุดสั่งซื้อ"])]

    c1,c2,c3 = st.columns(3)
    with c1: st.metric("จำนวนรายการ", f"{total_items:,}")
    with c2: st.metric("ยอดคงเหลือรวม", f"{int(total_qty):,}")
    with c3: st.metric("ใกล้หมดสต็อก", f"{len(low_df):,}")

    # Controls
    st.markdown("### เลือกกราฟที่จะแสดง")
    chart_opts = st.multiselect(
        " ",
        [
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
    left, right = st.columns([1,1])
    with left:
        chart_kind = st.radio("ชนิดกราฟ", ["กราฟวงกลม (Pie)","กราฟแท่ง (Bar)"], horizontal=True, index=0)
    with right:
        top_n = st.slider("Top-N", 3, 20, 10, 1)

    c_per = st.selectbox("กราฟต่อแถว", [1,2,3,4], index=1)

    st.markdown("### ช่วงเวลา (ใช้กับกราฟ OUT/Tickets)")
    colD1, colD2, colD3 = st.columns(3)
    with colD1:
        range_choice = st.selectbox("เลือกช่วง", ["วันนี้","7 วันล่าสุด","30 วันล่าสุด","90 วันล่าสุด","ปีนี้","กำหนดเอง"], index=2)
    with colD2:
        d1 = st.date_input("วันที่เริ่ม", value=(date.today()-timedelta(days=29)))
    with colD3:
        d2 = st.date_input("วันที่สิ้นสุด", value=date.today())

    def parse_range(choice, d1, d2):
        today = date.today()
        if choice=="วันนี้": return today, today
        if choice=="7 วันล่าสุด": return today-timedelta(days=6), today
        if choice=="30 วันล่าสุด": return today-timedelta(days=29), today
        if choice=="90 วันล่าสุด": return today-timedelta(days=89), today
        if choice=="ปีนี้": return date(today.year,1,1), today
        return d1, d2
    start_date, end_date = parse_range(range_choice, d1, d2)

    # Pre-calc maps
    cat_map = {str(r["รหัสหมวด"]).strip(): str(r["ชื่อหมวด"]).strip() for _,r in cats.iterrows()} if not cats.empty else {}
    br_map = {str(r['รหัสสาขา']).strip(): f"{str(r['รหัสสาขา']).strip()} | {str(r['ชื่อสาขา']).strip()}" for _,r in branches.iterrows()} if not branches.empty else {}

    # Prepare tx filtered
    if not txns.empty:
        tx = txns.copy()
        tx["วันเวลา"] = pd.to_datetime(tx["วันเวลา"], errors="coerce")
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
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda x: br_map.get(str(x).split(" | ")[0], str(x)))
            charts.append((f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน"))
        else:
            charts.append((f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"สาขาแสดง":[], "จำนวน":[]}), "สาขาแสดง", "จำนวน"))

    if "เบิกตามอุปกรณ์ (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("ชื่ออุปกรณ์")["จำนวน"].sum().reset_index()
            charts.append((f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", tmp, "ชื่ออุปกรณ์", "จำนวน"))
        else:
            charts.append((f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"ชื่ออุปกรณ์":[], "จำนวน":[]}), "ชื่ออุปกรณ์", "จำนวน"))

    if "เบิกตามหมวดหมู่ (OUT)" in chart_opts:
        if not tx_out.empty and not items.empty:
            it = items[["รหัส","หมวดหมู่"]]
            tmp = tx_out.merge(it, on="รหัส", how="left")
            tmp = tmp.groupby("หมวดหมู่")["จำนวน"].sum().reset_index()
            charts.append((f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", tmp, "หมวดหมู่", "จำนวน"))
        else:
            charts.append((f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"หมวดหมู่":[], "จำนวน":[]}), "หมวดหมู่", "จำนวน"))

    # Tickets
    tdf = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    if not tdf.empty:
        tdf["วันที่แจ้ง"] = pd.to_datetime(tdf["วันที่แจ้ง"], errors="coerce")
        tdf = tdf.dropna(subset=["วันที่แจ้ง"])
        tdf = tdf[(tdf["วันที่แจ้ง"].dt.date >= start_date) & (tdf["วันที่แจ้ง"].dt.date <= end_date)]
    if "Ticket ตามสถานะ" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สถานะ")["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            charts.append((f"Ticket ตามสถานะ {start_date} ถึง {end_date}", tmp, "สถานะ", "จำนวน"))
        else:
            charts.append((f"Ticket ตามสถานะ {start_date} ถึง {end_date}", pd.DataFrame({"สถานะ":[], "จำนวน":[]}), "สถานะ", "จำนวน"))
    if "Ticket ตามสาขา" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สาข", dropna=False)["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
        # fix correct column
        if not tdf.empty:
            tmp = tdf.groupby("สาขา", dropna=False)["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda v: br_map.get(str(v).split(" | ")[0], str(v)))
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน"))
        else:
            charts.append((f"Ticket ตามสาขา {start_date} ถึง {end_date}", pd.DataFrame({"สาขาแสดง":[], "จำนวน":[]}), "สาขาแสดง", "จำนวน"))

    # Render charts
    if not charts:
        st.info("โปรดเลือกกราฟ")
    else:
        rows = (len(charts) + c_per - 1)//c_per
        idx=0
        for _ in range(rows):
            cols = st.columns(c_per)
            for c in range(c_per):
                if idx>=len(charts): break
                title, df, label_col, val_col = charts[idx]
                with cols[c]:
                    df_show = df.copy()
                    if val_col in df_show.columns:
                        df_show[val_col] = pd.to_numeric(df_show[val_col], errors="coerce").fillna(0)
                        df_show = df_show.sort_values(val_col, ascending=False)
                    if len(df_show)>top_n:
                        top = df_show.head(top_n)
                        others = pd.DataFrame({label_col:["อื่นๆ"], val_col:[df_show[val_col].iloc[top_n:].sum()]})
                        df_show = pd.concat([top, others], ignore_index=True)
                    st.markdown(f"**{title}**")
                    if chart_kind.startswith("กราฟแท่ง"):
                        chart = alt.Chart(df_show).mark_bar().encode(
                            x=alt.X(f"{label_col}:N", sort='-y'),
                            y=alt.Y(f"{val_col}:Q"),
                            tooltip=[label_col, val_col]
                        )
                        st.altair_chart(chart.properties(height=300), use_container_width=True)
                    else:
                        chart = alt.Chart(df_show).mark_arc(innerRadius=60).encode(
                            theta=f"{val_col}:Q",
                            color=f"{label_col}:N",
                            tooltip=[label_col, val_col]
                        )
                        st.altair_chart(chart, use_container_width=True)
                idx+=1

    if not low_df.empty:
        with st.expander("⚠️ อุปกรณ์ใกล้หมด (Reorder)", expanded=False):
            show = low_df[["รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]]
            st.dataframe(show, height=240, use_container_width=True)
            st.download_button("ดาวน์โหลด PDF", data=df_to_pdf_bytes(show, "อุปกรณ์ใกล้หมดสต็อก", now_str()),
                               file_name="low_stock.pdf", mime="application/pdf")
    st.markdown("</div>", unsafe_allow_html=True)

def generate_item_code(sh, cat_code: str) -> str:
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    mx=0
    pat = re.compile(rf"^{re.escape(cat_code)}-(\d+)$")
    for code in items["รหัส"].astype(str):
        m = pat.match(code.strip())
        if m:
            try: mx=max(mx,int(m.group(1)))
            except: pass
    return f"{cat_code}-{mx+1:03d}"

def page_stock(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📦 คลังอุปกรณ์")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view = items.copy()
    if q and not items.empty:
        mask = items["รหัส"].str.contains(q, case=False, na=False) | \
               items["ชื่ออุปกรณ์"].str.contains(q, case=False, na=False) | \
               items["หมวดหมู่"].str.contains(q, case=False, na=False)
        view = items[mask]
    st.dataframe(view, height=320, use_container_width=True)

    if st.session_state.get("role") not in ("admin","staff"):
        st.info("สิทธิ์ผู้ชมไม่สามารถบันทึก/แก้ไขได้")
        st.markdown("</div>", unsafe_allow_html=True); return

    tabs = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ"])
    # add
    with tabs[0]:
        cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
        with st.form("add_item", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
                cat_opt = st.selectbox("หมวดหมู่", (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist() if not cats.empty else [])
                cat_code = cat_opt.split(" | ")[0] if cat_opt else ""
                name = st.text_input("ชื่ออุปกรณ์")
            with c2:
                unit = st.text_input("หน่วย", value="ชิ้น")
                qty = st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
            with c3:
                loc = st.text_input("ที่เก็บ", value="IT Room")
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                auto = st.checkbox("สร้างรหัสอัตโนมัติ", value=True)
                code = st.text_input("รหัส (ถ้าไม่ออโต้)", disabled=auto)
            s = st.form_submit_button("บันทึก", use_container_width=True)
        if s:
            if (auto and not cat_code) or (not auto and not code.strip()):
                st.error("กรุณาเลือกหมวดหรือระบุรหัส"); st.stop()
            code_final = generate_item_code(sh, cat_code) if auto else code.strip().upper()
            cur = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
            if (cur["รหัส"]==code_final).any():
                cur.loc[cur["รหัส"]==code_final, ITEMS_HEADERS] = [code_final, cat_code, name, unit, qty, rop, loc, active]
            else:
                cur = pd.concat([cur, pd.DataFrame([[code_final, cat_code, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_ITEMS, cur); st.success(f"บันทึกเรียบร้อย ({code_final})"); st.experimental_rerun()
    # edit
    with tabs[1]:
        if items.empty:
            st.info("ยังไม่มีรายการ")
        else:
            label = st.selectbox("เลือกรหัส", ["-- เลือก --"] + (items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
            if label != "-- เลือก --":
                code = label.split(" | ")[0]
                row = items[items["รหัส"]==code].iloc[0]
                with st.form("edit_item", clear_on_submit=False):
                    c1,c2,c3 = st.columns(3)
                    with c1:
                        name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
                        unit = st.text_input("หน่วย", value=row["หน่วย"])
                    with c2:
                        qty = st.number_input("คงเหลือ", min_value=0, value=int(pd.to_numeric(row["คงเหลือ"], errors="coerce") or 0), step=1)
                        rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=int(pd.to_numeric(row["จุดสั่งซื้อ"], errors="coerce") or 0), step=1)
                    with c3:
                        loc = st.text_input("ที่เก็บ", value=row["ที่เก็บ"])
                        active = st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(row["ใช้งาน"]).upper()=="Y" else 1)
                    col1,col2 = st.columns([3,1])
                    s1 = col1.form_submit_button("บันทึกการแก้ไข", use_container_width=True)
                    s2 = col2.form_submit_button("ลบ", use_container_width=True)
                if s1:
                    items.loc[items["รหัส"]==code, ITEMS_HEADERS] = [code, row["หมวดหมู่"], name, unit, qty, rop, loc, active]
                    write_df(sh, SHEET_ITEMS, items); st.success("อัปเดตแล้ว"); st.experimental_rerun()
                if s2:
                    items2 = items[items["รหัส"]!=code]
                    write_df(sh, SHEET_ITEMS, items2); st.success("ลบแล้ว"); st.experimental_rerun()
    st.markdown("</div>", unsafe_allow_html=True)

def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🧾 เบิก/รับเข้า")
    if st.session_state.get("role") not in ("admin","staff"):
        st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกรายการได้"); st.markdown("</div>", unsafe_allow_html=True); return
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty:
        st.warning("ยังไม่มีรายการอุปกรณ์"); st.markdown("</div>", unsafe_allow_html=True); return
    tab_out, tab_in = st.tabs(["เบิก (OUT) หลายรายการ","รับเข้า (IN)"])
    with tab_out:
        bopt = st.selectbox("สาขา/หน่วยงานผู้ขอ", (branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
        branch_code = bopt.split(" | ")[0] if bopt else ""
        opts = [f'{r["รหัส"]} | {r["ชื่ออุปกรณ์"]} (คงเหลือ {int(pd.to_numeric(r["คงเหลือ"], errors="coerce") or 0)})' for _,r in items.iterrows()]
        df_template = pd.DataFrame({"รายการ":[""]*5, "จำนวน":[1]*5})
        ed = st.data_editor(df_template, use_container_width=True, hide_index=True, num_rows="fixed",
                            column_config={"รายการ": st.column_config.SelectboxColumn(options=opts),
                                           "จำนวน": st.column_config.NumberColumn(min_value=1, step=1)})
        note = st.text_input("หมายเหตุ")
        if st.button("บันทึกการเบิก", type="primary", disabled=(not branch_code)):
            tx = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
            items_local = items.copy()
            ok_count=0; errs=[]
            for _,r in ed.iterrows():
                sel = str(r.get("รายการ","")).strip()
                qty = int(pd.to_numeric(r.get("จำนวน",0), errors="coerce") or 0)
                if not sel or qty<=0: continue
                code = sel.split(" | ")[0]
                row = items_local[items_local["รหัส"]==code]
                if row.empty: errs.append(f"{code}: ไม่พบ"); continue
                remain = int(pd.to_numeric(row.iloc[0]["คงเหลือ"], errors="coerce") or 0)
                if qty>remain: errs.append(f"{code}: เกินคงเหลือ ({remain})"); continue
                items_local.loc[items_local["รหัส"]==code,"คงเหลือ"]=remain-qty
                tx = pd.concat([tx, pd.DataFrame([[str(uuid.uuid4())[:8], now_str(),"OUT", code, row.iloc[0]["ชื่ออุปกรณ์"], branch_code, qty, st.session_state.get("user","unknown"), note]], columns=TXNS_HEADERS)], ignore_index=True)
                ok_count+=1
            if ok_count>0:
                write_df(sh, SHEET_ITEMS, items_local)
                write_df(sh, SHEET_TXNS, tx)
                st.success(f"บันทึกการเบิก {ok_count} รายการ"); st.experimental_rerun()
            else:
                st.warning("ยังไม่มีบรรทัดที่สมบูรณ์")
    with tab_in:
        item = st.selectbox("เลือกอุปกรณ์", (items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
        qty = st.number_input("จำนวนรับเข้า", min_value=1, value=1, step=1)
        src = st.text_input("แหล่งที่มา/เลข PO")
        note = st.text_input("หมายเหตุ", value="ซื้อเข้า")
        if st.button("บันทึกรับเข้า", type="primary"):
            code = item.split(" | ")[0]
            items2 = items.copy()
            row = items2[items2["รหัส"]==code].iloc[0]
            remain = int(pd.to_numeric(row["คงเหลือ"], errors="coerce") or 0)
            items2.loc[items2["รหัส"]==code,"คงเหลือ"]=remain+qty
            tx = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
            tx = pd.concat([tx, pd.DataFrame([[str(uuid.uuid4())[:8], now_str(),"IN", code, row["ชื่ออุปกรณ์"], src, qty, st.session_state.get("user","unknown"), note]], columns=TXNS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_ITEMS, items2); write_df(sh, SHEET_TXNS, tx)
            st.success("บันทึกรับเข้าแล้ว"); st.experimental_rerun()
    st.markdown("</div>", unsafe_allow_html=True)

def page_tickets(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    # list
    st.markdown("### รายการ")
    st.dataframe(tickets.sort_values("วันที่แจ้ง", ascending=False) if not tickets.empty else tickets, height=300, use_container_width=True)
    # add
    st.markdown("---")
    with st.form("tk_new", clear_on_submit=True):
        c1,c2,c3 = st.columns(3)
        with c1:
            branch = st.selectbox("สาขา", (branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
            reporter = st.text_input("ผู้แจ้ง")
        with c2:
            cate = st.text_input("หมวดหมู่")
            assignee = st.text_input("ผู้รับผิดชอบ (IT)", value=st.session_state.get("user",""))
        with c3:
            detail = st.text_area("รายละเอียด", height=100)
            note = st.text_input("หมายเหตุ")
        s = st.form_submit_button("บันทึกการรับแจ้ง", use_container_width=True)
    if s:
        tid = "TCK-"+datetime.now(TZ).strftime("%Y%m%d-%H%M%S")
        row=[tid, now_str(), branch, reporter, cate, detail, "รับแจ้ง", assignee, now_str(), note]
        append_row(sh, SHEET_TICKETS, row); st.success(f"รับแจ้งแล้ว ({tid})"); st.experimental_rerun()
    st.markdown("</div>", unsafe_allow_html=True)

def page_reports(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("📑 รายงาน / ประวัติ")
    tx = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    d1 = st.date_input("วันที่เริ่ม", value=(date.today()-timedelta(days=30)))
    d2 = st.date_input("วันที่สิ้นสุด", value=date.today())
    if not tx.empty:
        df = tx.copy()
        df["วันเวลา"]=pd.to_datetime(df["วันเวลา"], errors="coerce")
        df = df.dropna(subset=["วันเวลา"])
        df = df[(df["วันเวลา"].dt.date >= d1) & (df["วันเวลา"].dt.date <= d2)]
    else:
        df = tx
    st.dataframe(df.sort_values("วันเวลา", ascending=False) if not df.empty else df, height=320, use_container_width=True)
    st.download_button("ดาวน์โหลด PDF รายละเอียด", data=df_to_pdf_bytes(df, "รายละเอียดการเคลื่อนไหว", f"{d1} ถึง {d2}"),
                       file_name="transactions.pdf", mime="application/pdf")
    st.markdown("</div>", unsafe_allow_html=True)

def page_users(sh):
    st.subheader("👥 ผู้ใช้ & สิทธิ์ (Admin)")
    users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    for c in USERS_HEADERS:
        if c not in users.columns: users[c] = ""
    users = users[USERS_HEADERS].fillna("")
    st.dataframe(users, height=260, use_container_width=True)
    tab_add, tab_edit = st.tabs(["➕ เพิ่มผู้ใช้","✏️ แก้ไขผู้ใช้"])
    with tab_add:
        with st.form("add_user", clear_on_submit=True):
            c1,c2 = st.columns([2,1])
            with c1:
                un = st.text_input("Username*")
                disp = st.text_input("Display Name")
            with c2:
                role = st.selectbox("Role", ["admin","staff","viewer"], index=1)
                act = st.selectbox("Active", ["Y","N"], index=0)
            pwd = st.text_input("กำหนดรหัสผ่าน*", type="password")
            s = st.form_submit_button("บันทึกผู้ใช้ใหม่", use_container_width=True, type="primary")
        if s:
            if not un or not pwd: st.warning("กรอก Username/Password"); st.stop()
            if (users["Username"]==un).any(): st.error("มี Username นี้แล้ว"); st.stop()
            ph = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
            new_row = pd.DataFrame([[un, disp, role, ph, act]], columns=USERS_HEADERS)
            write_df(sh, SHEET_USERS, pd.concat([users, new_row], ignore_index=True))
            st.success("เพิ่มผู้ใช้สำเร็จ"); st.experimental_rerun()
    with tab_edit:
        sel = st.selectbox("เลือกผู้ใช้", [""]+users["Username"].tolist())
        if sel:
            row = users[users["Username"]==sel].iloc[0]
            with st.form("edit_user", clear_on_submit=False):
                c1,c2 = st.columns([2,1])
                with c1:
                    disp = st.text_input("Display Name", value=row["DisplayName"])
                with c2:
                    role = st.selectbox("Role", ["admin","staff","viewer"], index=["admin","staff","viewer"].index(row["Role"]) if row["Role"] in ["admin","staff","viewer"] else 1)
                    act = st.selectbox("Active", ["Y","N"], index=["Y","N"].index(row["Active"]) if row["Active"] in ["Y","N"] else 0)
                pwd = st.text_input("ตั้ง/รีเซ็ตรหัสผ่าน (เว้นว่าง = ไม่เปลี่ยน)", type="password")
                col1,col2 = st.columns([3,1])
                s1 = col1.form_submit_button("บันทึก", use_container_width=True)
                s2 = col2.form_submit_button("ลบ", use_container_width=True)
            if s1:
                idx = users.index[users["Username"]==sel][0]
                users.at[idx,"DisplayName"]=disp
                users.at[idx,"Role"]=role
                users.at[idx,"Active"]=act
                if pwd.strip():
                    users.at[idx,"PasswordHash"]=bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                write_df(sh, SHEET_USERS, users); st.success("บันทึกแล้ว"); st.experimental_rerun()
            if s2:
                if sel.lower()=="admin": st.error("ห้ามลบ admin")
                else:
                    write_df(sh, SHEET_USERS, users[users["Username"]!=sel]); st.success("ลบแล้ว"); st.experimental_rerun()

def page_import(sh):
    st.subheader("นำเข้า/แก้ไข หมวดหมู่ / เพิ่มข้อมูล (หมวดหมู่ / สาขา / อุปกรณ์ / หมวดหมู่ปัญหา / ผู้ใช้)")
    t1,t2,t3,t4,t5 = st.tabs(["หมวดหมู่","สาขา","อุปกรณ์","หมวดหมู่ปัญหา","ผู้ใช้"])
    def _read_upload(file):
        if not file: return None, "ยังไม่ได้เลือกไฟล์"
        name=file.name.lower()
        try:
            if name.endswith(".csv"): df=pd.read_csv(file, dtype=str).fillna("")
            else: df=pd.read_excel(file, dtype=str).fillna("")
            df=df.applymap(lambda x: str(x).strip())
            return df,None
        except Exception as e:
            return None, f"อ่านไฟล์ไม่สำเร็จ: {e}"
    # categories
    with t1:
        up=st.file_uploader("อัปโหลด หมวดหมู่ (CSV/Excel)", type=["csv","xlsx"], key="up_cat")
        if up:
            df,err=_read_upload(up)
            if err: st.error(err)
            else:
                st.dataframe(df, height=200, use_container_width=True)
                if not set(["รหัสหมวด","ชื่อหมวด"]).issubset(df.columns):
                    st.error("ต้องมีคอลัมน์ รหัสหมวด, ชื่อหมวด")
                elif st.button("นำเข้า/อัปเดต หมวดหมู่", use_container_width=True):
                    cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
                    for _,r in df.iterrows():
                        code=str(r["รหัสหมวด"]).strip(); name=str(r["ชื่อหมวด"]).strip()
                        if not code: continue
                        if (cur["รหัสหมวด"]==code).any():
                            cur.loc[cur["รหัสหมวด"]==code, ["รหัสหมวด","ชื่อหมวด"]]=[code,name]
                        else:
                            cur=pd.concat([cur, pd.DataFrame([[code,name]], columns=CATS_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_CATS, cur); st.success("สำเร็จ")
    # branches
    with t2:
        up=st.file_uploader("อัปโหลด สาขา (CSV/Excel)", type=["csv","xlsx"], key="up_br")
        if up:
            df,err=_read_upload(up)
            if err: st.error(err)
            else:
                st.dataframe(df, height=200, use_container_width=True)
                if not set(["รหัสสาขา","ชื่อสาขา"]).issubset(df.columns):
                    st.error("ต้องมีคอลัมน์ รหัสสาขา, ชื่อสาขา")
                elif st.button("นำเข้า/อัปเดต สาขา", use_container_width=True):
                    cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
                    for _,r in df.iterrows():
                        code=str(r["รหัสสาขา"]).strip(); name=str(r["ชื่อสาขา"]).strip()
                        if not code: continue
                        if (cur["รหัสสาขา"]==code).any():
                            cur.loc[cur["รหัสสาขา"]==code, ["รหัสสาขา","ชื่อสาขา"]]=[code,name]
                        else:
                            cur=pd.concat([cur, pd.DataFrame([[code,name]], columns=BR_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_BRANCHES, cur); st.success("สำเร็จ")
    # items
    with t3:
        up=st.file_uploader("อัปโหลด อุปกรณ์ (CSV/Excel)", type=["csv","xlsx"], key="up_items")
        if up:
            df,err=_read_upload(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=240, use_container_width=True)
                need = ["หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]
                if any(c not in df.columns for c in need):
                    st.error("หัวตารางขาดคอลัมน์ที่จำเป็น")
                elif st.button("นำเข้า/อัปเดต อุปกรณ์", use_container_width=True):
                    cur = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                    cats = set(read_df(sh, SHEET_CATS, CATS_HEADERS)["รหัสหมวด"].tolist())
                    add=upd=0
                    for _,r in df.iterrows():
                        code = str(r.get("รหัส","")).strip().upper()
                        cat  = str(r.get("หมวดหมู่","")).strip()
                        if cat not in cats: continue
                        name = str(r.get("ชื่ออุปกรณ์","")).strip()
                        unit = str(r.get("หน่วย","")).strip() or "ชิ้น"
                        qty  = int(pd.to_numeric(r.get("คงเหลือ",0), errors="coerce") or 0)
                        rop  = int(pd.to_numeric(r.get("จุดสั่งซื้อ",0), errors="coerce") or 0)
                        loc  = str(r.get("ที่เก็บ","")).strip() or "IT Room"
                        active = str(r.get("ใช้งาน","Y")).strip().upper() or "Y"
                        if not code: code = generate_item_code(sh, cat)
                        if (cur["รหัส"]==code).any():
                            cur.loc[cur["รหัส"]==code, ITEMS_HEADERS]=[code,cat,name,unit,qty,rop,loc,active]; upd+=1
                        else:
                            cur=pd.concat([cur, pd.DataFrame([[code,cat,name,unit,qty,rop,loc,active]], columns=ITEMS_HEADERS)], ignore_index=True); add+=1
                    write_df(sh, SHEET_ITEMS, cur); st.success(f"เพิ่ม {add} / อัปเดต {upd}")
    # ticket cats
    with t4:
        up=st.file_uploader("อัปโหลด หมวดหมู่ปัญหา (CSV/Excel)", type=["csv","xlsx"], key="up_tkc")
        if up:
            df,err=_read_upload(up)
            if err: st.error(err)
            else:
                st.dataframe(df, height=200, use_container_width=True)
                if not set(["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]).issubset(df.columns):
                    st.error("ต้องมีคอลัมน์ รหัสหมวดปัญหา, ชื่อหมวดปัญหา")
                elif st.button("นำเข้า/อัปเดต หมวดหมู่ปัญหา", use_container_width=True):
                    cur = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
                    for _,r in df.iterrows():
                        code=str(r["รหัสหมวดปัญหา"]).strip(); name=str(r["ชื่อหมวดปัญหา"]).strip()
                        if not code: continue
                        if (cur["รหัสหมวดปัญหา"]==code).any():
                            cur.loc[cur["รหัสหมวดปัญหา"]==code, TICKET_CAT_HEADERS]=[code,name]
                        else:
                            cur=pd.concat([cur, pd.DataFrame([[code,name]], columns=TICKET_CAT_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_TICKET_CATS, cur); st.success("สำเร็จ")
    # users import
    with t5:
        up=st.file_uploader("อัปโหลด ผู้ใช้ (CSV/Excel)", type=["csv","xlsx"], key="up_users")
        if up:
            df,err=_read_upload(up)
            if err: st.error(err)
            else:
                st.dataframe(df.head(20), height=220, use_container_width=True)
                if "Username" not in df.columns:
                    st.error("อย่างน้อยต้องมีคอลัมน์ Username")
                elif st.button("นำเข้า/อัปเดต ผู้ใช้", use_container_width=True):
                    cur = read_df(sh, SHEET_USERS, USERS_HEADERS)
                    add=upd=0; errs=[]
                    for _,r in df.iterrows():
                        username = str(r.get("Username","")).strip()
                        if not username: continue
                        display = str(r.get("DisplayName","")).strip()
                        role = str(r.get("Role","staff")).strip() or "staff"
                        active = str(r.get("Active","Y")).strip() or "Y"
                        pwd_hash=None
                        if "Password" in df.columns and str(r.get("Password","")).strip():
                            pwd_hash=bcrypt.hashpw(str(r.get("Password")).encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                        elif "PasswordHash" in df.columns and str(r.get("PasswordHash","")).strip():
                            pwd_hash=str(r.get("PasswordHash")).strip()
                        if (cur["Username"]==username).any():
                            idx = cur.index[cur["Username"]==username][0]
                            cur.at[idx,"DisplayName"]=display
                            cur.at[idx,"Role"]=role
                            cur.at[idx,"Active"]=active
                            if pwd_hash: cur.at[idx,"PasswordHash"]=pwd_hash
                            upd+=1
                        else:
                            if not pwd_hash: continue
                            cur=pd.concat([cur, pd.DataFrame([[username,display,role,pwd_hash,active]], columns=USERS_HEADERS)], ignore_index=True)
                            add+=1
                    write_df(sh, SHEET_USERS, cur); st.success(f"เพิ่ม {add} / อัปเดต {upd}")

def page_settings(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    st.subheader("⚙️ Settings")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ/ตรวจสอบชีต", use_container_width=True):
        try:
            sh2 = open_sheet_by_url(url); ensure_sheets_exist(sh2); st.success("เชื่อมต่อสำเร็จ พร้อมใช้งาน")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# ========== Main ==========
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    st.markdown(MINIMAL_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE); st.caption(APP_TAGLINE)

    # keep URL in session
    if "sheet_url" not in st.session_state or not st.session_state["sheet_url"]:
        st.session_state["sheet_url"] = DEFAULT_SHEET_URL

    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", [
            "📊 Dashboard",
            "📦 คลังอุปกรณ์",
            "🛠️ แจ้งปัญหา",
            "🧾 เบิก/รับเข้า",
            "📑 รายงาน",
            "👤 ผู้ใช้",
            "นำเข้า/แก้ไข หมวดหมู่",
            "⚙️ Settings",
        ], index=0)

    # open sheet
    sheet_url = st.session_state.get("sheet_url", DEFAULT_SHEET_URL)
    try:
        sh = open_sheet_by_url(sheet_url)
    except Exception as e:
        st.error(f"เปิดชีตไม่สำเร็จ: {e}")
        st.stop()

    ensure_sheets_exist(sh)
    auth_block(sh)

    if page.startswith("📊"): page_dashboard(sh)
    elif page.startswith("📦"): page_stock(sh)
    elif page.startswith("🛠️"): page_tickets(sh)
    elif page.startswith("🧾"): page_issue_receive(sh)
    elif page.startswith("📑"): page_reports(sh)
    elif page.startswith("👤"): page_users(sh)
    elif page.startswith("นำเข้า"): page_import(sh)
    elif page.startswith("⚙️"): page_settings(sh)

    st.caption("© 2025 IT Stock · Streamlit + Google Sheets · **iTao iT (V.1.1)**")

if __name__ == "__main__":
    main()
