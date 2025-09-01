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
    st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets) — เวอร์ชันปรับปรุงสมบูรณ์")

    # โหลดข้อมูล + ทำให้แน่ใจว่ามีคอลัมน์พื้นฐานครบ
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    for col in ["ประเภท"]:
        if col not in tickets.columns:
            tickets[col] = ""
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    t_cats   = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)

    # ----- ตัวกรอง -----
    st.markdown("### ตัวกรอง")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        status_pick = st.selectbox("สถานะ", ["ทั้งหมด","รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"], index=0)
    with f2:
        br_opts = ["ทั้งหมด"] + ((branches["รหัสสาขา"] + " | " + branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
        branch_pick = st.selectbox("สาขา", br_opts, index=0)
    with f3:
        cat_opts = ["ทั้งหมด"] + ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else [])
        cat_pick = st.selectbox("หมวดหมู่ปัญหา", cat_opts, index=0)
    with f4:
        q = st.text_input("ค้นหา (ผู้แจ้ง/รายละเอียด/หมายเหตุ)")

    d1c, d2c = st.columns(2)
    with d1c:
        d1 = st.date_input("วันที่เริ่ม", value=(date.today()-timedelta(days=30)))
    with d2c:
        d2 = st.date_input("วันที่สิ้นสุด", value=date.today())

    view = tickets.copy()
    if not view.empty:
        view["วันที่แจ้ง"] = pd.to_datetime(view["วันที่แจ้ง"], errors="coerce")
        view = view.dropna(subset=["วันที่แจ้ง"])
        view = view[(view["วันที่แจ้ง"].dt.date >= d1) & (view["วันที่แจ้ง"].dt.date <= d2)]
        if status_pick != "ทั้งหมด":
            view = view[view["สถานะ"] == status_pick]
        if branch_pick != "ทั้งหมด":
            view = view[view["สาขา"] == branch_pick]
        if cat_pick != "ทั้งหมด":
            view = view[view["หมวดหมู่"] == cat_pick]
        if q:
            mask = (
                view["ผู้แจ้ง"].astype(str).str.contains(q, case=False, na=False) |
                view["รายละเอียด"].astype(str).str.contains(q, case=False, na=False) |
                view["หมายเหตุ"].astype(str).str.contains(q, case=False, na=False)
            )
            view = view[mask]

    st.markdown("### ตารางรายการ (กรองแล้ว)")
    st.dataframe(view.sort_values("วันที่แจ้ง", ascending=False), height=320, use_container_width=True)
    if not view.empty:
        st.download_button("ดาวน์โหลด CSV", data=view.to_csv(index=False).encode("utf-8-sig"),
                           file_name="tickets_filtered.csv", mime="text/csv")

    st.markdown("---")
    tab_add, tab_edit, tab_bulk = st.tabs(["➕ รับแจ้งใหม่","✏️ เปลี่ยนสถานะ/แก้ไข","✅ ปิดงานแบบกลุ่ม"])

    # ===== รับแจ้งใหม่ =====
    with tab_add:
        with st.form("tk_new", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                br_sel = st.selectbox("สาขา", options=(br_opts[1:] if len(br_opts)>1 else ["พิมพ์เอง"]))
                if br_sel == "พิมพ์เอง":
                    br_sel = st.text_input("ระบุสาขา (พิมพ์เอง)", value="")
                reporter = st.text_input("ผู้แจ้ง")
                t_type = st.selectbox("ประเภท", ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"], index=0)
            with c2:
                cat_sel_opts = ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else []) + ["พิมพ์เอง"]
                cat_sel = st.selectbox("หมวดหมู่ปัญหา", options=cat_sel_opts if cat_sel_opts else ["พิมพ์เอง"])
                cat_custom = st.text_input("ระบุหมวด (ถ้าเลือกพิมพ์เอง)", value="" if cat_sel!="พิมพ์เอง" else "", disabled=(cat_sel!="พิมพ์เอง"))
                cate_val = cat_sel if cat_sel!="พิมพ์เอง" else cat_custom
                assignee = st.text_input("ผู้รับผิดชอบ (IT)", value=st.session_state.get("user",""))
            detail = st.text_area("รายละเอียด", height=120)
            note = st.text_input("หมายเหตุ")
            s = st.form_submit_button("บันทึกการรับแจ้ง", use_container_width=True, type="primary")
        if s:
            if not br_sel or not reporter or not cate_val or not detail:
                st.warning("กรุณากรอกข้อมูลให้ครบ")
            else:
                tid = "TCK-" + datetime.now(TZ).strftime("%Y%m%d-%H%M%S") if hasattr(datetime, "now") else "TCK-" + str(int(time.time()))
                row = [tid, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       br_sel, reporter, cate_val, detail, "รับแจ้ง", assignee,
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S"), note]
                df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
                for c in ["ประเภท"]:
                    if c not in df.columns: df[c] = ""
                df = pd.concat([df, pd.DataFrame([row], columns=TICKETS_HEADERS)], ignore_index=True)
                if "ประเภท" in df.columns:
                    df.loc[df["TicketID"]==tid, "ประเภท"] = t_type
                write_df(sh, SHEET_TICKETS, df[[c for c in df.columns]])
                st.success(f"รับแจ้งเรียบร้อย (Ticket: {tid})"); st.rerun()

    # ===== แก้ไข =====
    with tab_edit:
        if tickets.empty:
            st.info("ยังไม่มีรายการ")
        else:
            labels = (tickets["TicketID"] + " | " + tickets["สาขา"].astype(str)).tolist()
            pick_label = st.selectbox("เลือก Ticket", options=["-- เลือก --"] + labels)
            if pick_label != "-- เลือก --":
                pick_id = pick_label.split(" | ", 1)[0]
                row = tickets[tickets["TicketID"] == pick_id].iloc[0]

                with st.form("tk_edit", clear_on_submit=False):
                    c1, c2 = st.columns(2)
                    with c1:
                        t_branch = st.text_input("สาขา", value=str(row.get("สาขา","")))
                        t_owner  = st.text_input("ผู้แจ้ง", value=str(row.get("ผู้แจ้ง","")))
                        t_type   = st.selectbox("ประเภท", ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"],
                                                index=0 if str(row.get("ประเภท","")) not in ["ซอฟต์แวร์","เครือข่าย","อื่นๆ"] else ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"].index(str(row.get("ประเภท",""))))
                    with c2:
                        statuses_edit = ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"]
                        try:
                            idx_default = statuses_edit.index(str(row.get("สถานะ","รับแจ้ง")))
                        except ValueError:
                            idx_default = 0
                        t_status = st.selectbox("สถานะ", statuses_edit, index=idx_default)
                        t_assignee = st.text_input("ผู้รับผิดชอบ", value=str(row.get("ผู้รับผิดชอบ","")))
                    t_cate = st.text_input("หมวดหมู่", value=str(row.get("หมวดหมู่","")))
                    t_desc = st.text_area("รายละเอียด", value=str(row.get("รายละเอียด","")), height=120)
                    t_note = st.text_input("หมายเหตุ", value=str(row.get("หมายเหตุ","")))
                    colA, colB, colC = st.columns(3)
                    s_update = colA.form_submit_button("อัปเดต", use_container_width=True, type="primary")
                    s_close  = colB.form_submit_button("ปิดงาน (เสร็จ)", use_container_width=True)
                    s_delete = colC.form_submit_button("ลบรายการ", use_container_width=True)

                if s_delete:
                    df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
                    df = df[df["TicketID"] != pick_id].copy()
                    write_df(sh, SHEET_TICKETS, df)
                    st.success("ลบเรียบร้อย"); st.rerun()

                if s_update or s_close:
                    if s_close: t_status = "ดำเนินการเสร็จ"
                    df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
                    for c in ["ประเภท"]:
                        if c not in df.columns: df[c] = ""
                    idx = df.index[df["TicketID"] == pick_id]
                    if len(idx)==1:
                        i = idx[0]
                        df.at[i,"สาขา"] = t_branch
                        df.at[i,"ผู้แจ้ง"] = t_owner
                        df.at[i,"หมวดหมู่"] = t_cate
                        df.at[i,"รายละเอียด"] = t_desc
                        df.at[i,"สถานะ"] = t_status
                        df.at[i,"ผู้รับผิดชอบ"] = t_assignee
                        df.at[i,"อัปเดตล่าสุด"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        df.at[i,"หมายเหตุ"] = t_note
                        df.at[i,"ประเภท"] = t_type
                        write_df(sh, SHEET_TICKETS, df[[c for c in df.columns]])
                        st.success("อัปเดตเรียบร้อย"); st.rerun()

    # ===== ปิดงานแบบกลุ่ม =====
    with tab_bulk:
        st.caption("เลือก Ticket หลายรายการแล้วปิดงานทีเดียว (เปลี่ยนสถานะเป็น 'ดำเนินการเสร็จ')")
        if tickets.empty:
            st.info("ยังไม่มีรายการ")
        else:
            # แสดงเฉพาะที่ยังไม่เสร็จ
            open_df = tickets[tickets["สถานะ"] != "ดำเนินการเสร็จ"].copy()
            if open_df.empty:
                st.success("ไม่มีงานค้าง")
            else:
                open_df["เลือก"] = False
                ed = st.data_editor(open_df[["เลือก","TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","หมายเหตุ"]],
                                    use_container_width=True, height=260, num_rows="fixed",
                                    column_config={"เลือก": st.column_config.CheckboxColumn()})
                if st.button("ปิดงานที่เลือก"):
                    picked = ed[ed["เลือก"] == True]["TicketID"].tolist()
                    if not picked:
                        st.warning("ยังไม่ได้เลือกรายการ")
                    else:
                        df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
                        df.loc[df["TicketID"].isin(picked), "สถานะ"] = "ดำเนินการเสร็จ"
                        df.loc[df["TicketID"].isin(picked), "อัปเดตล่าสุด"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        write_df(sh, SHEET_TICKETS, df)
                        st.success(f"ปิดงาน {len(picked)} รายการเรียบร้อย"); st.rerun()

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



# ======================= OVERRIDES: checkbox-select edit for Stock / Tickets / Users =======================
import pandas as _pd
import streamlit as _st

def page_stock(sh):
    _st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    _st.subheader("📦 คลังอุปกรณ์")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)

    # --- Search & Data table with checkbox select ---
    q = _st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view_df = items.copy()
    if q and not items.empty:
        mask = (
            items["รหัส"].astype(str).str.contains(q, case=False, na=False) |
            items["ชื่ออุปกรณ์"].astype(str).str.contains(q, case=False, na=False) |
            items["หมวดหมู่"].astype(str).str.contains(q, case=False, na=False)
        )
        view_df = items[mask]

    chosen_code = None
    if hasattr(_st, "data_editor"):
        view_display = view_df.copy()
        view_display.insert(0, "เลือก", False)
        edited_table = _st.data_editor(
            view_display[["เลือก"] + ITEMS_HEADERS],
            use_container_width=True, height=320, num_rows="fixed",
            column_config={"เลือก": _st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือกสำหรับแก้ไข")}
        )
        picked = edited_table[edited_table["เลือก"] == True]
        if not picked.empty:
            chosen_code = str(picked.iloc[0]["รหัส"]).strip()
    else:
        _st.dataframe(view_df, height=320, use_container_width=True)

    unit_opts = get_unit_options(items)
    loc_opts  = get_loc_options(items)

    if _st.session_state.get("role") in ("admin","staff"):
        t_add, t_edit = _st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ (เลือกจากตาราง)"])

        with t_add:
            with _st.form("item_add", clear_on_submit=True):
                c1,c2,c3 = _st.columns(3)
                with c1:
                    if cats.empty:
                        _st.info("ยังไม่มีหมวดหมู่ในชีต Categories (ใช้เมนู นำเข้า/แก้ไข หมวดหมู่ เพื่อเพิ่ม)")
                        cat_opt=""
                    else:
                        opts = (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist()
                        selected = _st.selectbox("หมวดหมู่", options=opts)
                        cat_opt = selected.split(" | ")[0]
                    name = _st.text_input("ชื่ออุปกรณ์")
                with c2:
                    sel_unit = _st.selectbox("หน่วย (เลือกจากรายการ)", options=unit_opts, index=0)
                    unit = _st.text_input("ระบุหน่วยใหม่", value="", disabled=(sel_unit!="พิมพ์เอง"))
                    if sel_unit!="พิมพ์เอง": unit = sel_unit
                    qty = _st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                    rop = _st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
                with c3:
                    sel_loc = _st.selectbox("ที่เก็บ (เลือกจากรายการ)", options=loc_opts, index=0)
                    loc = _st.text_input("ระบุที่เก็บใหม่", value="", disabled=(sel_loc!="พิมพ์เอง"))
                    if sel_loc!="พิมพ์เอง": loc = sel_loc
                    active = _st.selectbox("ใช้งาน", ["Y","N"], index=0)
                    auto_code = _st.checkbox("สร้างรหัสอัตโนมัติ", value=True)
                    code = _st.text_input("รหัสอุปกรณ์ (ถ้าไม่ออโต้)", disabled=auto_code)
                    s_add = _st.form_submit_button("บันทึก/อัปเดต", use_container_width=True)
            if s_add:
                if (auto_code and not cat_opt) or (not auto_code and code.strip()==""):
                    _st.error("กรุณาเลือกหมวด/ระบุรหัส")
                else:
                    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                    gen_code = generate_item_code(sh, cat_opt) if auto_code else code.strip().upper()
                    if (items["รหัส"]==gen_code).any():
                        items.loc[items["รหัส"]==gen_code, ITEMS_HEADERS] = [gen_code, cat_opt, name, unit, qty, rop, loc, active]
                    else:
                        items = _pd.concat([items, _pd.DataFrame([[gen_code, cat_opt, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True)
                    write_df(sh, SHEET_ITEMS, items)
                    _st.success(f"บันทึกเรียบร้อย (รหัส: {gen_code})")
                    safe_rerun()

        with t_edit:
            if not chosen_code:
                _st.info("ติ๊กเลือกรายการในตารางด้านบนเพื่อแก้ไข หรือเลือกรายการเดิมด้วยตัวเลือกด้านล่าง")
            # Fallback dropdown for editing
            if items.empty:
                _st.info("ยังไม่มีรายการให้แก้ไข")
            else:
                labels = (items["รหัส"] + " | " + items["ชื่ออุปกรณ์"]).tolist()
                default_index = 0
                if chosen_code:
                    try:
                        default_index = ["-- เลือก --"] + labels
                        default_index = default_index.index(f"{chosen_code} | {items[items['รหัส']==chosen_code].iloc[0]['ชื่ออุปกรณ์']}")
                    except Exception:
                        default_index = 0
                pick_label = _st.selectbox("เลือกรหัสอุปกรณ์", options=["-- เลือก --"] + labels)
                if chosen_code:
                    # override by chosen
                    for lab in labels:
                        if lab.startswith(chosen_code + " "):
                            pick_label = lab
                            break
                if pick_label != "-- เลือก --":
                    pick = pick_label.split(" | ", 1)[0]
                    row = items[items["รหัส"] == pick].iloc[0]
                    # Build option lists that include current values
                    unit_opts_edit = [u for u in unit_opts if u != "พิมพ์เอง"]
                    if row["หน่วย"] not in unit_opts_edit and str(row["หน่วย"]).strip()!="":
                        unit_opts_edit = [row["หน่วย"]] + unit_opts_edit
                    unit_opts_edit = unit_opts_edit + ["พิมพ์เอง"]

                    loc_opts_edit = [l for l in loc_opts if l != "พิมพ์เอง"]
                    if row["ที่เก็บ"] not in loc_opts_edit and str(row["ที่เก็บ"]).strip()!="":
                        loc_opts_edit = [row["ที่เก็บ"]] + loc_opts_edit
                    loc_opts_edit = loc_opts_edit + ["พิมพ์เอง"]

                    with _st.form("item_edit", clear_on_submit=False):
                        c1,c2,c3 = _st.columns(3)
                        with c1:
                            name = _st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
                            sel_unit = _st.selectbox("หน่วย (เลือกจากรายการ)", options=unit_opts_edit, index=0)
                            unit = _st.text_input("ระบุหน่วยใหม่", value="", disabled=(sel_unit!="พิมพ์เอง"))
                            if sel_unit!="พิมพ์เอง": unit = sel_unit
                        with c2:
                            qty = _st.number_input("คงเหลือ", min_value=0, value=int(float(row["คงเหลือ"]) if str(row["คงเหลือ"]).strip()!="" else 0), step=1)
                            rop = _st.number_input("จุดสั่งซื้อ", min_value=0, value=int(float(row["จุดสั่งซื้อ"]) if str(row["จุดสั่งซื้อ"]).strip()!="" else 0), step=1)
                        with c3:
                            sel_loc = _st.selectbox("ที่เก็บ (เลือกจากรายการ)", options=loc_opts_edit, index=0)
                            loc = _st.text_input("ระบุที่เก็บใหม่", value="", disabled=(sel_loc!="พิมพ์เอง"))
                            if sel_loc!="พิมพ์เอง": loc = sel_loc
                            active = _st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(row["ใช้งาน"]).upper()=="Y" else 1)
                        col_save, col_delete = _st.columns([3,1])
                        s_save = col_save.form_submit_button("💾 บันทึกการแก้ไข", use_container_width=True)
                        s_del  = col_delete.form_submit_button("🗑️ ลบรายการ", use_container_width=True)
                    if s_save:
                        items.loc[items["รหัส"]==pick, ITEMS_HEADERS] = [pick, row["หมวดหมู่"], name, unit, qty, rop, loc, "Y" if active=="Y" else "N"]
                        write_df(sh, SHEET_ITEMS, items)
                        _st.success("อัปเดตแล้ว"); safe_rerun()
                    if s_del:
                        items = items[items["รหัส"]!=pick]; write_df(sh, SHEET_ITEMS, items)
                        _st.success(f"ลบ {pick} แล้ว"); safe_rerun()

    _st.markdown("</div>", unsafe_allow_html=True)


def page_tickets(sh):
    _st.markdown("<div class='block-card'>", unsafe_allow_html=True)
    _st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")

    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    t_cats = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)

    # --- Filters ---
    _st.markdown("### ตัวกรอง")
    f1, f2, f3, f4 = _st.columns(4)
    with f1:
        statuses = ["ทั้งหมด","รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"]
        status_pick = _st.selectbox("สถานะ", statuses, index=0, key="tk_status2")
    with f2:
        br_opts = ["ทั้งหมด"] + ((branches["รหัสสาขา"] + " | " + branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
        branch_pick = _st.selectbox("สาขา", br_opts, index=0, key="tk_branch2")
    with f3:
        cat_opts = ["ทั้งหมด"] + ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else [])
        cat_pick = _st.selectbox("หมวดหมู่ปัญหา", cat_opts, index=0, key="tk_cat2")
    with f4:
        q = _st.text_input("ค้นหา (ผู้แจ้ง/หมวด/รายละเอียด)", key="tk_query2")

    dcol1, dcol2 = _st.columns(2)
    with dcol1:
        d1 = _st.date_input("วันที่เริ่ม", value=(date.today()-timedelta(days=90)), key="tk_d1_2")
    with dcol2:
        d2 = _st.date_input("วันที่สิ้นสุด", value=date.today(), key="tk_d2_2")

    view = tickets.copy()
    if not view.empty:
        view["วันที่แจ้ง"] = _pd.to_datetime(view["วันที่แจ้ง"], errors="coerce")
        view = view.dropna(subset=["วันที่แจ้ง"])
        view = view[(view["วันที่แจ้ง"].dt.date >= d1) & (view["วันที่แจ้ง"].dt.date <= d2)]
        if status_pick != "ทั้งหมด":
            view = view[view["สถานะ"] == status_pick]
        if branch_pick != "ทั้งหมด":
            view = view[view["สาขา"] == branch_pick]
        if "cat_pick" in locals() and cat_pick != "ทั้งหมด":
            view = view[view["หมวดหมู่"] == cat_pick]
        if q:
            mask = (view["ผู้แจ้ง"].astype(str).str.contains(q, case=False, na=False) |
                    view["หมวดหมู่"].astype(str).str.contains(q, case=False, na=False) |
                    view["รายละเอียด"].astype(str).str.contains(q, case=False, na=False))
            view = view[mask]

    # --- Table with checkbox select ---
    _st.markdown("### รายการแจ้งปัญหา")
    chosen_id = None
    picked_ids = []
    if hasattr(_st, "data_editor"):
        tdisp = view.copy()
        tdisp.insert(0, "เลือก", False)
        ed = _st.data_editor(
            tdisp[["เลือก"] + TICKETS_HEADERS],
            use_container_width=True, height=300, num_rows="fixed",
            column_config={"เลือก": _st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือก")}
        )
        pick = ed[ed["เลือก"]==True]
        if not pick.empty:
            picked_ids = pick["TicketID"].astype(str).tolist()
            chosen_id = picked_ids[0]  # first for edit form
    else:
        _st.dataframe(view, height=300, use_container_width=True)

    # --- Bulk close ---
    if picked_ids:
        if _st.button(f"ปิดงาน (เสร็จ) {len(picked_ids)} รายการ"):
            cur = tickets.copy()
            cur.loc[cur["TicketID"].astype(str).isin(picked_ids), "สถานะ"] = "ดำเนินการเสร็จ"
            cur.loc[cur["TicketID"].astype(str).isin(picked_ids), "อัปเดตล่าสุด"] = get_now_str()
            write_df(sh, SHEET_TICKETS, cur)
            _st.success("อัปเดตสถานะเรียบร้อย")
            safe_rerun()

    _st.markdown("---")
    t_add, t_update = _st.tabs(["➕ รับแจ้งใหม่","🔁 เปลี่ยนสถานะ/แก้ไข"])

    with t_add:
        with _st.form("tk_new2", clear_on_submit=True):
            c1,c2,c3 = _st.columns(3)
            with c1:
                now_str = get_now_str()
                branch_sel = _st.selectbox("สาขา", br_opts[1:] if len(br_opts)>1 else ["พิมพ์เอง"])
                if branch_sel == "พิมพ์เอง":
                    branch_sel = _st.text_input("ระบุสาขา (พิมพ์เอง)", value="")
                reporter = _st.text_input("ผู้แจ้ง", value="")
            with c2:
                tkc_opts = ((t_cats["รหัสหมวดปัญหา"] + " | " + t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else []) + ["พิมพ์เอง"]
                pick_c = _st.selectbox("หมวดหมู่ปัญหา", options=tkc_opts if tkc_opts else ["พิมพ์เอง"], key="tk_new_cat_sel2")
                cate_custom = _st.text_input("ระบุหมวด (ถ้าเลือกพิมพ์เอง)", value="" if pick_c!="พิมพ์เอง" else "", disabled=(pick_c!="พิมพ์เอง"))
                cate = pick_c if pick_c != "พิมพ์เอง" else cate_custom
                t_type = _st.selectbox("ประเภท", ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"], index=0)
                assignee = _st.text_input("ผู้รับผิดชอบ (IT)", value=_st.session_state.get("user",""))
            with c3:
                detail = _st.text_area("รายละเอียด", height=100)
                note = _st.text_input("หมายเหตุ", value="")
            s = _st.form_submit_button("บันทึกการรับแจ้ง", use_container_width=True)
        if s:
            tid = generate_ticket_id()
            row = [tid, get_now_str(), branch_sel, reporter, cate, detail, "รับแจ้ง", assignee, get_now_str(), note]
            # ensure columns
            cur = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
            if "ประเภท" not in cur.columns:
                cur["ประเภท"] = ""
            new = _pd.DataFrame([row], columns=TICKETS_HEADERS)
            new["ประเภท"] = t_type
            cur = _pd.concat([cur, new], ignore_index=True)
            write_df(sh, SHEET_TICKETS, cur)
            _st.success(f"รับแจ้งเรียบร้อย (Ticket: {tid})")
            safe_rerun()

    with t_update:
        target_id = chosen_id or _st.selectbox("เลือก Ticket (สำรอง)", [""] + tickets["TicketID"].astype(str).tolist())
        if not target_id:
            _st.info("ติ๊กเลือกรายการในตารางด้านบน หรือเลือกจากกล่องด้านบนเพื่อแก้ไข")
        else:
            row = tickets[tickets["TicketID"].astype(str) == str(target_id)]
            if row.empty:
                _st.warning("ไม่พบ Ticket ที่เลือก")
            else:
                data = row.iloc[0].to_dict()
                with _st.form("tk_edit2", clear_on_submit=False):
                    c1, c2 = _st.columns(2)
                    with c1:
                        t_branch = _st.text_input("สาขา", value=str(data.get("สาขา","")))
                        t_type   = _st.selectbox("ประเภท", ["ฮาร์ดแวร์","ซอฟต์แวร์","เครือข่าย","อื่นๆ"],
                                                 index=0 if str(data.get("ประเภท",""))=="" else 3)
                    with c2:
                        t_owner = _st.text_input("ผู้แจ้ง", value=str(data.get("ผู้แจ้ง","")))
                        statuses_edit = ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"]
                        try:
                            idx_default = statuses_edit.index(str(data.get("สถานะ","รับแจ้ง")))
                        except ValueError:
                            idx_default = 0
                        t_status = _st.selectbox("สถานะ", statuses_edit, index=idx_default)
                        t_assignee = _st.text_input("ผู้รับผิดชอบ", value=str(data.get("ผู้รับผิดชอบ","")))
                    t_desc = _st.text_area("รายละเอียด", value=str(data.get("รายละเอียด","")), height=120)
                    t_note = _st.text_input("หมายเหตุ", value=str(data.get("หมายเหตุ","")))
                    fcol1, fcol2, fcol3 = _st.columns(3)
                    submit_update = fcol1.form_submit_button("อัปเดต")
                    submit_close  = fcol2.form_submit_button("ปิดงาน (เสร็จ)")
                    submit_delete = fcol3.form_submit_button("ลบรายการ")

                if submit_update or submit_close:
                    tickets2 = tickets.copy()
                    idx = tickets2.index[tickets2["TicketID"].astype(str) == str(target_id)]
                    if len(idx) == 1:
                        idx0 = idx[0]
                        tickets2.at[idx0, "สาขา"] = t_branch
                        tickets2.at[idx0, "ผู้แจ้ง"] = t_owner
                        tickets2.at[idx0, "รายละเอียด"] = t_desc
                        tickets2.at[idx0, "สถานะ"] = "ดำเนินการเสร็จ" if submit_close else t_status
                        tickets2.at[idx0, "ผู้รับผิดชอบ"] = t_assignee
                        tickets2.at[idx0, "หมายเหตุ"] = t_note
                        # ensure 'ประเภท'
                        if "ประเภท" not in tickets2.columns:
                            tickets2["ประเภท"] = ""
                        tickets2.at[idx0, "ประเภท"] = t_type
                        tickets2.at[idx0, "อัปเดตล่าสุด"] = get_now_str()
                        write_df(sh, SHEET_TICKETS, tickets2)
                        _st.success("บันทึกเรียบร้อย")
                        safe_rerun()
                    else:
                        _st.error("ไม่พบรายการหรือพบหลายรายการ")

                if submit_delete:
                    tickets2 = tickets[tickets["TicketID"].astype(str) != str(target_id)].copy()
                    write_df(sh, SHEET_TICKETS, tickets2)
                    _st.success("ลบรายการแล้ว")
                    safe_rerun()


# Keep Users page as checkbox-select (redefine to ensure present)
def page_users(sh):
    _st.subheader("👥 ผู้ใช้ & สิทธิ์ (Admin)")

    try:
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    except Exception as e:
        _st.error(f"โหลดข้อมูลผู้ใช้ไม่สำเร็จ: {e}")
        return

    base_cols = ["Username","DisplayName","Role","PasswordHash","Active"]
    for col in base_cols:
        if col not in users.columns:
            users[col] = ""
    users = users[base_cols].fillna("")

    _st.markdown("#### 📋 รายชื่อผู้ใช้ (ติ๊ก 'เลือก' เพื่อแก้ไข)")
    chosen_username = None
    if hasattr(_st, "data_editor"):
        users_display = users.copy()
        users_display.insert(0, "เลือก", False)
        edited_table = _st.data_editor(
            users_display[["เลือก","Username","DisplayName","Role","PasswordHash","Active"]],
            use_container_width=True, height=300, num_rows="fixed",
            column_config={"เลือก": _st.column_config.CheckboxColumn(help="ติ๊กเพื่อเลือกผู้ใช้สำหรับแก้ไข")}
        )
        picked = edited_table[edited_table["เลือก"] == True]
        if not picked.empty:
            chosen_username = str(picked.iloc[0]["Username"])

    tab_add, tab_edit = _st.tabs(["➕ เพิ่มผู้ใช้", "✏️ แก้ไขผู้ใช้"])

    with tab_add:
        with _st.form("form_add_user2", clear_on_submit=True):
            c1, c2 = _st.columns([2,1])
            with c1:
                new_user = _st.text_input("Username*")
                new_disp = _st.text_input("Display Name")
            with c2:
                new_role = _st.selectbox("Role", ["admin","staff","viewer"], index=1)
                new_active = _st.selectbox("Active", ["Y","N"], index=0)
            new_pwd = _st.text_input("กำหนดรหัสผ่าน*", type="password")
            btn_add = _st.form_submit_button("บันทึกผู้ใช้ใหม่", use_container_width=True, type="primary")

        if btn_add:
            if not new_user.strip() or not new_pwd.strip():
                _st.warning("กรุณากรอก Username และรหัสผ่าน"); _st.stop()
            if (users["Username"] == new_user).any():
                _st.error("มี Username นี้อยู่แล้ว"); _st.stop()
            ph = bcrypt.hashpw(new_pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
            new_row = _pd.DataFrame([{
                "Username": new_user.strip(),
                "DisplayName": new_disp.strip(),
                "Role": new_role,
                "PasswordHash": ph,
                "Active": new_active,
            }])
            users2 = _pd.concat([users, new_row], ignore_index=True)
            try:
                write_df(sh, SHEET_USERS, users2)
                try: _st.cache_data.clear()
                except Exception: pass
                _st.success("เพิ่มผู้ใช้สำเร็จ"); _st.rerun()
            except Exception as e:
                _st.error(f"เพิ่มผู้ใช้ไม่สำเร็จ: {e}")

    with tab_edit:
        default_user = _st.session_state.get("edit_user","")
        if chosen_username:
            _st.session_state["edit_user"] = chosen_username
            default_user = chosen_username

        sel = _st.selectbox(
            "เลือกผู้ใช้เพื่อแก้ไข",
            [""] + users["Username"].tolist(),
            index=([""] + users["Username"].tolist()).index(default_user) if default_user in users["Username"].tolist() else 0
        )

        target_user = sel or ""
        if not target_user:
            _st.info("ยังไม่ได้เลือกผู้ใช้สำหรับแก้ไข"); return

        row = users[users["Username"] == target_user]
        if row.empty:
            _st.warning("ไม่พบผู้ใช้ที่เลือก"); return
        data = row.iloc[0].to_dict()

        with _st.form("form_edit_user2", clear_on_submit=False):
            c1, c2 = _st.columns([2,1])
            with c1:
                username = _st.text_input("Username", value=data["Username"], disabled=True)
                display  = _st.text_input("Display Name", value=data["DisplayName"])
            with c2:
                role  = _st.selectbox("Role", ["admin","staff","viewer"],
                                     index=["admin","staff","viewer"].index(data["Role"]) if data["Role"] in ["admin","staff","viewer"] else 1)
                active = _st.selectbox("Active", ["Y","N"],
                                      index=["Y","N"].index(data["Active"]) if data["Active"] in ["Y","N"] else 0)
            pwd = _st.text_input("ตั้ง/รีเซ็ตรหัสผ่าน (ปล่อยว่าง = ไม่เปลี่ยน)", type="password")

            c3, c4 = _st.columns([1,1])
            btn_save = c3.form_submit_button("บันทึกการแก้ไข", use_container_width=True, type="primary")
            btn_del  = c4.form_submit_button("ลบผู้ใช้นี้", use_container_width=True)

        if btn_del:
            if username.lower() == "admin":
                _st.error("ห้ามลบผู้ใช้ admin")
            else:
                users2 = users[users["Username"] != username]
                try:
                    write_df(sh, SHEET_USERS, users2)
                    try: _st.cache_data.clear()
                    except Exception: pass
                    _st.success(f"ลบผู้ใช้ {username} แล้ว")
                    _st.session_state.pop("edit_user", None)
                    _st.rerun()
                except Exception as e:
                    _st.error(f"ลบไม่สำเร็จ: {e}")

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
                try: _st.cache_data.clear()
                except Exception: pass
                _st.success("บันทึกการแก้ไขเรียบร้อย")
                _st.rerun()
            except Exception as e:
                _st.error(f"บันทึกไม่สำเร็จ: {e}")
# ======================= END OVERRIDES =======================
