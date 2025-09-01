#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# =============== iTao iT (V.1.1) — Dashboard Restored ===============
# Single-file Streamlit app focused on a configurable Dashboard.
# Includes stable Service Account loading (secrets → env → file → embedded).

import os, json, base64, re, uuid, time
from datetime import datetime, date, timedelta, time as dtime

import pandas as pd
import altair as alt
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound

# ---------------- Streamlit compatibility shims ----------------
if not hasattr(st, "cache_resource"):
    def _no_cache_decorator(*args, **kwargs):
        def _wrap(func): return func
        return _wrap
    st.cache_resource = _no_cache_decorator

# --------------- App meta ----------------
APP_TITLE = "ไอต้าว ไอที (iTao iT)"
APP_TAGLINE = "POWER By ทีมงาน=> ไอทีสุดหล่อ"
VERSION_TEXT = "© 2025 IT Stock · Streamlit + Google Sheets By AOD. · iTao iT (V.1.1)"

# ---- Replace this URL with your production Sheet URL (user provided) ----
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit?gid=0#gid=0"

# --------------- Google Sheets config ----------------
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

# Optional embedded service account (base64 of JSON). Leave empty if unused.
EMBEDDED_SA_B64 = os.environ.get("EMBEDDED_SA_B64", "").strip()

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ---------------- Service Account loading ----------------
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
    raw = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON") or os.environ.get("SERVICE_ACCOUNT_JSON")
    if not raw: 
        return None
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
    if not EMBEDDED_SA_B64:
        return None
    try:
        return json.loads(base64.b64decode(EMBEDDED_SA_B64).decode("utf-8"))
    except Exception:
        return None

@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    info = (_try_load_sa_from_secrets() or
            _try_load_sa_from_env() or
            _try_load_sa_from_file() or
            _try_load_sa_from_embedded())
    if info is None:
        st.error("ไม่พบ Service Account (ใช้ st.secrets / ENV / ไฟล์ / ฝัง base64)")
        st.stop()
    creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)

# convenience wrappers
@st.cache_resource(show_spinner=False)
def open_sheet_by_url(sheet_url: str):
    return _get_gspread_client().open_by_url(sheet_url)

@st.cache_resource(show_spinner=False)
def open_sheet_by_key(key: str):
    return _get_gspread_client().open_by_key(key)

# ---------------- Sheet utilities ----------------
def ensure_sheets_exist(sh):
    """Create required worksheets (if not exist) and seed header rows."""
    required = [
        (SHEET_ITEMS, ITEMS_HEADERS, 1000, len(ITEMS_HEADERS)+5),
        (SHEET_TXNS, TXNS_HEADERS, 2000, len(TXNS_HEADERS)+5),
        (SHEET_USERS, USERS_HEADERS, 100, len(USERS_HEADERS)+2),
        (SHEET_CATS, CATS_HEADERS, 200, len(CATS_HEADERS)+2),
        (SHEET_BRANCHES, BR_HEADERS, 200, len(BR_HEADERS)+2),
        (SHEET_TICKETS, TICKETS_HEADERS, 1000, len(TICKETS_HEADERS)+5),
        (SHEET_TICKET_CATS, TICKET_CAT_HEADERS, 200, len(TICKET_CAT_HEADERS)+2),
    ]
    titles = []
    try:
        titles = [ws.title for ws in sh.worksheets()]
    except Exception:
        titles = []

    for name, headers, rows, cols in required:
        try:
            if name not in titles:
                ws = sh.add_worksheet(name, rows, cols)
                ws.append_row(headers)
        except Exception:
            pass

@st.cache_data(ttl=60, show_spinner=False)
def _cached_records(sheet_url: str, ws_title: str):
    sh = open_sheet_by_url(sheet_url)
    ws = sh.worksheet(ws_title)
    return ws.get_all_records()

def read_df_by_url(sheet_url: str, sheet_name: str, headers=None) -> pd.DataFrame:
    try:
        records = _cached_records(sheet_url, sheet_name)
    except Exception as e:
        st.error(f"อ่านชีตไม่สำเร็จ ({sheet_name}): {e}")
        records = []
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

# ---------------- Dashboard (configurable, restored) ----------------
def page_dashboard(sheet_url: str):
    st.subheader("📊 Dashboard")
    items    = read_df_by_url(sheet_url, SHEET_ITEMS, ITEMS_HEADERS)
    txns     = read_df_by_url(sheet_url, SHEET_TXNS, TXNS_HEADERS)
    cats     = read_df_by_url(sheet_url, SHEET_CATS, CATS_HEADERS)
    branches = read_df_by_url(sheet_url, SHEET_BRANCHES, BR_HEADERS)
    cat_map = {str(r['รหัสหมวด']).strip(): str(r['ชื่อหมวด']).strip() for _, r in cats.iterrows()} if not cats.empty else {}
    br_map = {str(r['รหัสสาขา']).strip(): f"{str(r['รหัสสาขา']).strip()} | {str(r['ชื่อสาขา']).strip()}" for _, r in branches.iterrows()} if not branches.empty else {}

    # KPIs
    total_items = len(items)
    total_qty = 0
    if not items.empty:
        total_qty = pd.to_numeric(items["คงเหลือ"], errors="coerce").fillna(0).astype(int).sum()
    low_df = items.copy()
    if not low_df.empty:
        low_df["คงเหลือ"] = pd.to_numeric(low_df["คงเหลือ"], errors="coerce").fillna(0)
        low_df["จุดสั่งซื้อ"] = pd.to_numeric(low_df["จุดสั่งซื้อ"], errors="coerce").fillna(0)
        low_df = low_df[(low_df["ใช้งาน"].astype(str).str.upper()=="Y") & (low_df["คงเหลือ"] <= low_df["จุดสั่งซื้อ"])]
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("จำนวนรายการ", f"{total_items:,}")
    with c2: st.metric("ยอดคงเหลือรวม", f"{total_qty:,}")
    with c3: st.metric("ใกล้หมดสต็อก", f"{len(low_df):,}")

    st.markdown("---")
    # ------- Control panel (restored) -------
    colA, colB, colC = st.columns([2,1,1])
    with colA:
        chart_opts = st.multiselect(
            "เลือกกราฟที่จะแสดง",
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
            default=["คงเหลือตามหมวดหมู่","เบิกตามสาขา (OUT)","Ticket ตามสถานะ"],
        )
    with colB:
        top_n = st.slider("Top-N", min_value=3, max_value=20, value=10, step=1)
    with colC:
        per_row = st.selectbox("กราฟต่อแถว", [1,2,3,4], index=1)
    chart_kind = st.radio("ชนิดกราฟ", ["กราฟวงกลม (Pie)", "กราฟแท่ง (Bar)"], horizontal=True)

    # Date range for OUT/Tickets
    st.markdown("### ⏱️ ช่วงเวลา (สำหรับ OUT และ Tickets)")
    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        range_choice = st.selectbox("ช่วงเวลา", ["วันนี้","7 วันล่าสุด","30 วันล่าสุด","90 วันล่าสุด","ปีนี้","กำหนดเอง"], index=2)
    with rc2:
        d1 = st.date_input("วันที่เริ่ม", value=date.today()-timedelta(days=29))
    with rc3:
        d2 = st.date_input("วันที่สิ้นสุด", value=date.today())
    def parse_range(choice: str, d1: date=None, d2: date=None):
        today = date.today()
        if choice == "วันนี้": return today, today
        if choice == "7 วันล่าสุด": return today-timedelta(days=6), today
        if choice == "30 วันล่าสุด": return today-timedelta(days=29), today
        if choice == "90 วันล่าสุด": return today-timedelta(days=89), today
        if choice == "ปีนี้": return date(today.year, 1, 1), today
        if choice == "กำหนดเอง" and d1 and d2: return d1, d2
        return today-timedelta(days=29), today
    start_date, end_date = parse_range(range_choice, d1, d2)

    # Prepare transactions (OUT only within range)
    if not txns.empty:
        tx = txns.copy()
        tx["วันเวลา"] = pd.to_datetime(tx["วันเวลา"], errors="coerce")
        tx = tx.dropna(subset=["วันเวลา"])
        tx = tx[(tx["วันเวลา"].dt.date >= start_date) & (tx["วันเวลา"].dt.date <= end_date)]
        tx["จำนวน"] = pd.to_numeric(tx["จำนวน"], errors="coerce").fillna(0)
        tx_out = tx[tx["ประเภท"]=="OUT"]
    else:
        tx_out = pd.DataFrame(columns=TXNS_HEADERS)

    # Build chart list
    charts = []

    def add_chart(title, df, label_col, value_col):
        charts.append((title, df.copy(), label_col, value_col))

    if "คงเหลือตามหมวดหมู่" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("หมวดหมู่")["คงเหลือ"].sum().reset_index()
        tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
        add_chart("คงเหลือตามหมวดหมู่", tmp, "หมวดหมู่ชื่อ", "คงเหลือ")

    if "คงเหลือตามที่เก็บ" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["คงเหลือ"] = pd.to_numeric(tmp["คงเหลือ"], errors="coerce").fillna(0)
        tmp = tmp.groupby("ที่เก็บ")["คงเหลือ"].sum().reset_index()
        add_chart("คงเหลือตามที่เก็บ", tmp, "ที่เก็บ", "คงเหลือ")

    if "จำนวนรายการตามหมวดหมู่" in chart_opts and not items.empty:
        tmp = items.copy()
        tmp["count"] = 1
        tmp = tmp.groupby("หมวดหมู่")["count"].sum().reset_index()
        tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
        add_chart("จำนวนรายการตามหมวดหมู่", tmp, "หมวดหมู่ชื่อ", "count")

    if "เบิกตามสาขา (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("สาขา", dropna=False)["จำนวน"].sum().reset_index()
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda x: br_map.get(str(x).split(" | ")[0], str(x) if "|" in str(x) else str(x)))
            add_chart(f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน")
        else:
            add_chart(f"เบิกตามสาขา (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"สาขาแสดง":[], "จำนวน":[]}), "สาขาแสดง", "จำนวน")

    if "เบิกตามอุปกรณ์ (OUT)" in chart_opts:
        if not tx_out.empty:
            tmp = tx_out.groupby("ชื่ออุปกรณ์")["จำนวน"].sum().reset_index()
            add_chart(f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", tmp, "ชื่ออุปกรณ์", "จำนวน")
        else:
            add_chart(f"เบิกตามอุปกรณ์ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"ชื่ออุปกรณ์":[], "จำนวน":[]}), "ชื่ออุปกรณ์", "จำนวน")

    if "เบิกตามหมวดหมู่ (OUT)" in chart_opts:
        if not tx_out.empty and not items.empty:
            it = items[["รหัส","หมวดหมู่"]].copy()
            tmp = tx_out.merge(it, left_on="รหัส", right_on="รหัส", how="left")
            tmp = tmp.groupby("หมวดหมู่")["จำนวน"].sum().reset_index()
            tmp["หมวดหมู่ชื่อ"] = tmp["หมวดหมู่"].map(cat_map).fillna(tmp["หมวดหมู่"])
            add_chart(f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", tmp, "หมวดหมู่ชื่อ", "จำนวน")
        else:
            add_chart(f"เบิกตามหมวดหมู่ (OUT) {start_date} ถึง {end_date}", pd.DataFrame({"หมวดหมู่ชื่อ":[], "จำนวน":[]}), "หมวดหมู่ชื่อ", "จำนวน")

    # Tickets (read + filter by date if column exists)
    tickets_df = read_df_by_url(sheet_url, SHEET_TICKETS, TICKETS_HEADERS)
    if not tickets_df.empty and "วันที่แจ้ง" in tickets_df.columns:
        tdf = tickets_df.copy()
        tdf["วันที่แจ้ง"] = pd.to_datetime(tdf["วันที่แจ้ง"], errors="coerce")
        tdf = tdf.dropna(subset=["วันที่แจ้ง"])
        tdf = tdf[(tdf["วันที่แจ้ง"].dt.date >= start_date) & (tdf["วันที่แจ้ง"].dt.date <= end_date)]
    else:
        tdf = pd.DataFrame(columns=TICKETS_HEADERS)

    if "Ticket ตามสถานะ" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สถานะ")["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            add_chart(f"Ticket ตามสถานะ {start_date} ถึง {end_date}", tmp, "สถานะ", "จำนวน")
        else:
            add_chart(f"Ticket ตามสถานะ {start_date} ถึง {end_date}", pd.DataFrame({"สถานะ":[], "จำนวน":[]}), "สถานะ", "จำนวน")

    if "Ticket ตามสาขา" in chart_opts:
        if not tdf.empty:
            tmp = tdf.groupby("สาขา", dropna=False)["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
            tmp["สาขาแสดง"] = tmp["สาขา"].apply(lambda x: br_map.get(str(x).split(" | ")[0], str(x) if "|" in str(x) else str(x)))
            add_chart(f"Ticket ตามสาขา {start_date} ถึง {end_date}", tmp, "สาขาแสดง", "จำนวน")
        else:
            add_chart(f"Ticket ตามสาขา {start_date} ถึง {end_date}", pd.DataFrame({"สาขาแสดง":[], "จำนวน":[]}), "สาขาแสดง", "จำนวน")

    # ---------- Render charts ----------
    def show_pie(df: pd.DataFrame, label_col: str, value_col: str, title: str):
        if df.empty or value_col not in df.columns:
            st.info(f"ยังไม่มีข้อมูล: {title}")
            return
        work = df.copy()
        work[value_col] = pd.to_numeric(work[value_col], errors="coerce").fillna(0)
        work = work.groupby(label_col, dropna=False)[value_col].sum().reset_index().rename(columns={value_col:"sum_val"})
        work[label_col] = work[label_col].replace("", "ไม่ระบุ")
        work = work.sort_values("sum_val", ascending=False)
        if len(work) > top_n:
            top = work.head(top_n)
            others = pd.DataFrame({label_col:["อื่นๆ"], "sum_val":[work["sum_val"].iloc[top_n:].sum()]})
            work = pd.concat([top, others], ignore_index=True)
        chart = alt.Chart(work).mark_arc(innerRadius=60).encode(
            theta="sum_val:Q",
            color=f"{label_col}:N",
            tooltip=[f"{label_col}:N","sum_val:Q"]
        ).properties(title=title)
        st.altair_chart(chart, use_container_width=True)

    def show_bar(df: pd.DataFrame, label_col: str, value_col: str, title: str):
        if df.empty or value_col not in df.columns:
            st.info(f"ยังไม่มีข้อมูล: {title}")
            return
        work = df.copy()
        work[value_col] = pd.to_numeric(work[value_col], errors="coerce").fillna(0)
        work = work.groupby(label_col, dropna=False)[value_col].sum().reset_index().rename(columns={value_col:"sum_val"})
        work[label_col] = work[label_col].replace("", "ไม่ระบุ")
        work = work.sort_values("sum_val", ascending=False)
        if len(work) > top_n:
            work = work.head(top_n)
        chart = alt.Chart(work).mark_bar().encode(
            x=alt.X(f"{label_col}:N", sort='-y'),
            y=alt.Y("sum_val:Q"),
            tooltip=[f"{label_col}:N","sum_val:Q"]
        ).properties(title=title, height=320)
        st.altair_chart(chart, use_container_width=True)

    if len(charts) == 0:
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
                    if chart_kind.endswith("(Bar)"):
                        show_bar(df, label_col, value_col, title)
                    else:
                        show_pie(df, label_col, value_col, title)
                idx += 1

    # Low stock table
    if not low_df.empty:
        with st.expander("⚠️ อุปกรณ์ใกล้หมด (Reorder)", expanded=False):
            st.dataframe(low_df[["รหัส","ชื่ออุปกรณ์","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ"]], height=240, use_container_width=True)

# ---------------- Settings page ----------------
def page_settings():
    st.subheader("⚙️ Settings")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ", type="primary"):
        try:
            sh = open_sheet_by_url(url)
            ensure_sheets_exist(sh)
            st.success("เชื่อมต่อสำเร็จ พร้อมใช้งาน")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")

# ---------------- Main ----------------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    st.title(APP_TITLE); st.caption(APP_TAGLINE)

    # Keep URL in session
    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"):
        st.session_state["sheet_url"] = DEFAULT_SHEET_URL

    with st.sidebar:
        page = st.radio("เมนู", ["📊 Dashboard", "⚙️ Settings"], index=0)
        st.caption(VERSION_TEXT)

    sheet_url = st.session_state["sheet_url"]

    if page.startswith("📊"):
        try:
            # open once to ensure exists and cache connections
            sh = open_sheet_by_url(sheet_url)
            ensure_sheets_exist(sh)
        except Exception as e:
            st.error(f"เปิดชีตไม่สำเร็จ: {e}")
            st.info("ไปที่ Settings แล้วตรวจสอบ URL / สิทธิ์ share ให้ service account")
            return
        page_dashboard(sheet_url)
    else:
        page_settings()

    st.markdown(f"<div style='text-align:center;color:#666;padding-top:8px'>{VERSION_TEXT}</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
