#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, json, base64, io, re, uuid, time
from datetime import datetime, date, timedelta, time as dtime

import streamlit as st
import pandas as pd
import altair as alt
import pytz
import bcrypt

# =========================
# Safe rerun shim (fixes AttributeError for experimental_rerun)
# =========================
def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()
    else:
        # fallback: ask the user to click a small button to refresh
        st.toast("รีเฟรชหน้าเพื่อแสดงผลล่าสุด", icon="🔁")

# for legacy code paths
if not hasattr(st, "experimental_rerun"):
    st.experimental_rerun = safe_rerun  # alias for old calls

# =========================
# App Constants
# =========================
APP_TITLE = "ไอต้าว ไอที (iTao iT)"
APP_TAGLINE = "POWER By ทีมงาน=> ไอทีสุดหล่อ"
TZ = pytz.timezone("Asia/Bangkok")

# Sheet / Worksheet names
SHEET_ITEMS       = "Items"
SHEET_TXNS        = "Transactions"
SHEET_USERS       = "Users"
SHEET_CATS        = "Categories"
SHEET_BRANCHES    = "Branches"
SHEET_TICKETS     = "Tickets"
SHEET_TICKET_CATS = "TicketCategories"

# Headers
ITEMS_HEADERS   = ["รหัส","หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
TXNS_HEADERS    = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"]
USERS_HEADERS   = ["Username","DisplayName","Role","PasswordHash","Active"]
CATS_HEADERS    = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS      = ["รหัสสาขา","ชื่อสาขา"]
TICKETS_HEADERS = ["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]
TICKET_CAT_HEADERS = ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]

# Default Google Sheet URL (เปลี่ยนได้ใน Settings)
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1SGKzZ9WKkRtcmvN3vZj9w2yeM6xNoB6QV3-gtnJY-Bw/edit?gid=0#gid=0"

# =========================
# Credentials loader (secrets → env → file). No upload prompt.
# =========================
from google.oauth2.service_account import Credentials
import gspread

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
            raw = st.secrets["service_account_json"]
            return json.loads(str(raw))
    except Exception:
        pass
    return None

def _try_load_sa_from_env():
    raw = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON") or os.environ.get("SERVICE_ACCOUNT_JSON") or os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not raw: 
        return None
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

@st.cache_resource(show_spinner=False)
def get_client():
    info = _try_load_sa_from_secrets() or _try_load_sa_from_env() or _try_load_sa_from_file()
    if info is None:
        st.error("ไม่พบ Service Account ใน secrets/env/file\n\nโปรดเพิ่มใน **Secrets** ชื่อ `gcp_service_account` (ทั้ง object) หรือ ENV `GOOGLE_APPLICATION_CREDENTIALS_JSON`.")
        st.stop()
    creds = Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def open_sheet_by_url(url: str):
    gc = get_client()
    return gc.open_by_url(url)

@st.cache_resource(show_spinner=False)
def open_sheet_by_key(key: str):
    gc = get_client()
    return gc.open_by_key(key)

# =========================
# Utils
# =========================
def fmt_dt(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S")

def get_now_str():
    return fmt_dt(datetime.now(TZ))

def combine_date_time(d: date, t: dtime) -> datetime:
    naive = datetime.combine(d, t)
    return TZ.localize(naive)

def clear_cache():
    try:
        st.cache_data.clear()
    except Exception:
        pass

# Data cache helpers
@st.cache_data(ttl=60, show_spinner=False)
def _cached_ws_records(sheet_url_or_key: str, ws_title: str, by="url"):
    try:
        sh = open_sheet_by_url(sheet_url_or_key) if by=="url" else open_sheet_by_key(sheet_url_or_key)
        ws = sh.worksheet(ws_title)
        return ws.get_all_records()
    except Exception as e:
        raise

def read_df(sh, sheet_name: str, headers=None):
    # try cached by url if available
    sheet_url = st.session_state.get("sheet_url", "") or ""
    try:
        if sheet_url:
            recs = _cached_ws_records(sheet_url, sheet_name, by="url")
        else:
            ws = sh.worksheet(sheet_name)
            recs = ws.get_all_records()
    except Exception:
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
    ws = sh.worksheet(title)
    ws.clear()
    # ensure columns order
    if title==SHEET_ITEMS: cols=ITEMS_HEADERS
    elif title==SHEET_TXNS: cols=TXNS_HEADERS
    elif title==SHEET_USERS: cols=USERS_HEADERS
    elif title==SHEET_CATS: cols=CATS_HEADERS
    elif title==SHEET_BRANCHES: cols=BR_HEADERS
    elif title==SHEET_TICKETS: cols=TICKETS_HEADERS
    elif title==SHEET_TICKET_CATS: cols=TICKET_CAT_HEADERS
    else: cols = list(df.columns)
    for c in cols:
        if c not in df.columns: df[c] = ""
    df = df[cols]
    ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
    clear_cache()

def append_row(sh, title, row):
    sh.worksheet(title).append_row([str(x) for x in row])
    clear_cache()

def ensure_sheets_exist(sh):
    required = [
        (SHEET_ITEMS, ITEMS_HEADERS, 1000, len(ITEMS_HEADERS)+4),
        (SHEET_TXNS, TXNS_HEADERS, 1000, len(TXNS_HEADERS)+4),
        (SHEET_USERS, USERS_HEADERS, 100, len(USERS_HEADERS)+2),
        (SHEET_CATS, CATS_HEADERS, 200, len(CATS_HEADERS)+2),
        (SHEET_BRANCHES, BR_HEADERS, 200, len(BR_HEADERS)+2),
        (SHEET_TICKETS, TICKETS_HEADERS, 500, len(TICKETS_HEADERS)+4),
        (SHEET_TICKET_CATS, TICKET_CAT_HEADERS, 200, len(TICKET_CAT_HEADERS)+2),
    ]
    try:
        titles = [ws.title for ws in sh.worksheets()]
    except Exception:
        titles = []
    for name, headers, r, c in required:
        if name not in titles:
            ws = sh.add_worksheet(name, r, c)
            ws.append_row(headers)
    # seed default admin
    try:
        users = sh.worksheet(SHEET_USERS).get_all_records()
        if len(users)==0:
            ph = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            sh.worksheet(SHEET_USERS).append_row(["admin","Administrator","admin",ph,"Y"])
    except Exception:
        pass

# =========================
# Auth
# =========================
def auth_block(sh):
    st.session_state.setdefault("user", None)
    st.session_state.setdefault("role", None)

    if st.session_state.get("user"):
        with st.sidebar:
            st.markdown(f"**👤 {st.session_state['user']}**")
            st.caption(f"Role: {st.session_state['role']}")
            if st.button("ออกจากระบบ"):
                st.session_state["user"] = None
                st.session_state["role"] = None
                safe_rerun()
        return True

    st.sidebar.subheader("เข้าสู่ระบบ")
    u = st.sidebar.text_input("Username", key="login_u")
    p = st.sidebar.text_input("Password", type="password", key="login_p")
    if st.sidebar.button("Login", use_container_width=True):
        users = read_df(sh, SHEET_USERS, USERS_HEADERS)
        row = users[(users["Username"]==u) & (users["Active"].astype(str).str.upper()=="Y")]
        if not row.empty:
            ok = False
            try:
                ok = bcrypt.checkpw(p.encode("utf-8"), row.iloc[0]["PasswordHash"].encode("utf-8"))
            except Exception:
                ok = False
            if ok:
                st.session_state["user"] = u
                st.session_state["role"] = row.iloc[0]["Role"]
                st.success("เข้าสู่ระบบสำเร็จ")
                safe_rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")
        else:
            st.error("ไม่พบบัญชีหรือถูกปิดใช้งาน")
    return False

# =========================
# Business helpers
# =========================
def generate_item_code(sh, cat_code: str) -> str:
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    pattern = re.compile(rf"^{re.escape(cat_code)}-(\d+)$")
    max_num = 0
    for code in items["รหัส"].dropna().astype(str):
        m = pattern.match(code.strip())
        if m:
            try:
                max_num = max(max_num, int(m.group(1)))
            except Exception:
                pass
    return f"{cat_code}-{max_num+1:03d}"

def adjust_stock(sh, code, delta, actor, branch="", note="", txn_type="OUT", ts_str=None):
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    row = items[items["รหัส"]==code]
    if row.empty:
        st.error("ไม่พบรหัสอุปกรณ์นี้"); return False
    cur = int(pd.to_numeric(row.iloc[0]["คงเหลือ"], errors="coerce").fillna(0))
    if txn_type=="OUT" and cur + delta < 0:
        st.error("สต็อกไม่พอ"); return False
    items.loc[items["รหัส"]==code, "คงเหลือ"] = str(cur + delta)
    write_df(sh, SHEET_ITEMS, items)
    ts = ts_str if ts_str else get_now_str()
    append_row(sh, SHEET_TXNS, [str(uuid.uuid4())[:8], ts, txn_type, code, row.iloc[0]["ชื่ออุปกรณ์"], branch, abs(delta), actor, note])
    return True

# =========================
# Pages
# =========================
def page_dashboard(sh):
    st.subheader("📊 Dashboard")

    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)

    total_items = len(items)
    total_qty = pd.to_numeric(items.get("คงเหลือ", pd.Series(dtype=str)), errors="coerce").fillna(0).sum() if not items.empty else 0
    low_df = items.copy()
    if not low_df.empty:
        low_df["คงเหลือ"] = pd.to_numeric(low_df["คงเหลือ"], errors="coerce").fillna(0)
        low_df["จุดสั่งซื้อ"] = pd.to_numeric(low_df["จุดสั่งซื้อ"], errors="coerce").fillna(0)
        low_df = low_df[(low_df["ใช้งาน"].astype(str).str.upper()=="Y") & (low_df["คงเหลือ"] <= low_df["จุดสั่งซื้อ"])]
    low_count = len(low_df)

    c1, c2, c3 = st.columns(3)
    c1.metric("จำนวนรายการ", f"{total_items:,}")
    c2.metric("ยอดคงเหลือรวม", f"{int(total_qty):,}")
    c3.metric("ใกล้หมดสต็อก", f"{low_count:,}")

    st.markdown("### เลือกกราฟที่ต้องการแสดง")
    chart_opts = st.multiselect(
        "กรุณาเลือก",
        ["คงเหลือตามหมวดหมู่","คงเหลือตามที่เก็บ","เบิกตามสาขา (OUT)","เบิกตามอุปกรณ์ (OUT)","Ticket ตามสถานะ"],
        default=["คงเหลือตามหมวดหมู่","เบิกตามสาขา (OUT)","Ticket ตามสถานะ"]
    )
    top_n = st.slider("Top-N", 3, 20, 10, 1)
    per_row = st.selectbox("กราฟต่อแถว", [1,2,3,4], index=1)
    kind = st.radio("ชนิดกราฟ", ["กราฟวงกลม (Pie)","กราฟแท่ง (Bar)"], horizontal=True)

    # ช่วงเวลาใช้กับ OUT เท่านั้น
    st.markdown("#### ⏱️ ช่วงเวลา (สำหรับ OUT และ Tickets)")
    colR1, colR2, colR3 = st.columns(3)
    with colR1:
        ranges = ["วันนี้","7 วันล่าสุด","30 วันล่าสุด","90 วันล่าสุด","ปีนี้"]
        pick = st.selectbox("ช่วงเวลา", ranges, index=2)
    today = datetime.now(TZ).date()
    if pick=="วันนี้": d1, d2 = today, today
    elif pick=="7 วันล่าสุด": d1, d2 = today - timedelta(days=6), today
    elif pick=="30 วันล่าสุด": d1, d2 = today - timedelta(days=29), today
    elif pick=="90 วันล่าสุด": d1, d2 = today - timedelta(days=89), today
    else: d1, d2 = date(today.year,1,1), today

    # Build charts
    charts = []
    if "คงเหลือตามหมวดหมู่" in chart_opts and not items.empty:
        df = items.copy()
        df["คงเหลือ"] = pd.to_numeric(df["คงเหลือ"], errors="coerce").fillna(0)
        work = df.groupby("หมวดหมู่")["คงเหลือ"].sum().reset_index()
        charts.append(("คงเหลือตามหมวดหมู่", work, "หมวดหมู่", "คงเหลือ"))
    if "คงเหลือตามที่เก็บ" in chart_opts and not items.empty:
        df = items.copy()
        df["คงเหลือ"] = pd.to_numeric(df["คงเหลือ"], errors="coerce").fillna(0)
        work = df.groupby("ที่เก็บ")["คงเหลือ"].sum().reset_index()
        charts.append(("คงเหลือตามที่เก็บ", work, "ที่เก็บ", "คงเหลือ"))
    if "เบิกตามสาขา (OUT)" in chart_opts:
        if not txns.empty:
            df = txns.copy()
            df["วันเวลา"] = pd.to_datetime(df["วันเวลา"], errors="coerce")
            df = df.dropna(subset=["วันเวลา"])
            df = df[(df["วันเวลา"].dt.date >= d1) & (df["วันเวลา"].dt.date <= d2) & (df["ประเภท"]=="OUT")]
            df["จำนวน"] = pd.to_numeric(df["จำนวน"], errors="coerce").fillna(0)
            work = df.groupby("สาขา")["จำนวน"].sum().reset_index()
        else:
            work = pd.DataFrame({"สาขา":[],"จำนวน":[]})
        charts.append((f"เบิกตามสาขา (OUT) {d1} ถึง {d2}", work, "สาขา", "จำนวน"))
    if "เบิกตามอุปกรณ์ (OUT)" in chart_opts:
        if not txns.empty:
            df = txns.copy()
            df["วันเวลา"] = pd.to_datetime(df["วันเวลา"], errors="coerce")
            df = df.dropna(subset=["วันเวลา"])
            df = df[(df["วันเวลา"].dt.date >= d1) & (df["วันเวลา"].dt.date <= d2) & (df["ประเภท"]=="OUT")]
            df["จำนวน"] = pd.to_numeric(df["จำนวน"], errors="coerce").fillna(0)
            work = df.groupby("ชื่ออุปกรณ์")["จำนวน"].sum().reset_index()
        else:
            work = pd.DataFrame({"ชื่ออุปกรณ์":[],"จำนวน":[]})
        charts.append((f"เบิกตามอุปกรณ์ (OUT) {d1} ถึง {d2}", work, "ชื่ออุปกรณ์", "จำนวน"))
    if "Ticket ตามสถานะ" in chart_opts:
        if not tickets.empty:
            df = tickets.copy()
            df["วันที่แจ้ง"] = pd.to_datetime(df["วันที่แจ้ง"], errors="coerce")
            df = df.dropna(subset=["วันที่แจ้ง"])
            df = df[(df["วันที่แจ้ง"].dt.date >= d1) & (df["วันที่แจ้ง"].dt.date <= d2)]
            work = df.groupby("สถานะ")["TicketID"].count().reset_index().rename(columns={"TicketID":"จำนวน"})
        else:
            work = pd.DataFrame({"สถานะ":[],"จำนวน":[]})
        charts.append((f"Ticket ตามสถานะ {d1} ถึง {d2}", work, "สถานะ", "จำนวน"))

    def show_chart(title, df, label, value):
        if df.empty or (value in df.columns and pd.to_numeric(df[value], errors="coerce").fillna(0).sum()==0):
            st.info(f"ยังไม่มีข้อมูลสำหรับ: {title}")
            return
        work = df.copy()
        work[value] = pd.to_numeric(work[value], errors="coerce").fillna(0)
        work = work.sort_values(value, ascending=False)
        if len(work) > top_n:
            work = work.head(top_n)
        st.markdown(f"**{title}**")
        if kind.startswith("กราฟแท่ง"):
            ch = alt.Chart(work).mark_bar().encode(
                x=alt.X(f"{label}:N", sort='-y'),
                y=alt.Y(f"{value}:Q"),
                tooltip=[label, value]
            )
            st.altair_chart(ch.properties(height=320), use_container_width=True)
        else:
            ch = alt.Chart(work).mark_arc(innerRadius=60).encode(
                theta=f"{value}:Q", color=f"{label}:N", tooltip=[label, value]
            )
            st.altair_chart(ch, use_container_width=True)

    rows = (len(charts) + per_row - 1) // per_row
    idx = 0
    for r in range(rows):
        cols = st.columns(per_row)
        for c in range(per_row):
            if idx >= len(charts): break
            title, df, label, value = charts[idx]
            with cols[c]:
                show_chart(title, df, label, value)
            idx += 1

def page_stock(sh):
    st.subheader("📦 คลังอุปกรณ์")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view = items.copy()
    if q and not items.empty:
        mask = (
            items["รหัส"].astype(str).str.contains(q, case=False, na=False) |
            items["ชื่ออุปกรณ์"].astype(str).str.contains(q, case=False, na=False) |
            items["หมวดหมู่"].astype(str).str.contains(q, case=False, na=False)
        )
        view = items[mask]
    st.dataframe(view, height=300, use_container_width=True)

    tab_add, tab_edit = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ (ติ๊กจากตาราง)"])

    with tab_add:
        cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
        branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
        unit_default = "ชิ้น"
        with st.form("add_item", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                if cats.empty:
                    st.info("ยังไม่มีหมวดหมู่ (ไปเพิ่มในเมนู นำเข้า/แก้ไข หมวดหมู่)")
                    cat_code = st.text_input("หมวดหมู่ (รหัส)", "")
                else:
                    pick = st.selectbox("หมวดหมู่ (รหัส | ชื่อ)", options=(cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist())
                    cat_code = pick.split(" | ")[0]
                name = st.text_input("ชื่ออุปกรณ์")
            with c2:
                auto = st.checkbox("สร้างรหัสอัตโนมัติ", value=True)
                code = st.text_input("รหัสอุปกรณ์ (ถ้าไม่ออโต้)", disabled=auto)
                unit = st.text_input("หน่วย", value=unit_default)
                qty = st.number_input("คงเหลือ", 0, 10**9, 0, 1)
            with c3:
                rop = st.number_input("จุดสั่งซื้อ", 0, 10**9, 0, 1)
                loc = st.text_input("ที่เก็บ", "IT Room")
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                s = st.form_submit_button("บันทึก", use_container_width=True)
        if s:
            if auto and cat_code.strip():
                code_gen = generate_item_code(sh, cat_code.strip())
            else:
                code_gen = code.strip().upper()
            if not code_gen:
                st.error("กรุณาระบุรหัส"); st.stop()
            items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
            if (items["รหัส"]==code_gen).any():
                # update row
                items.loc[items["รหัส"]==code_gen, ITEMS_HEADERS] = [code_gen, cat_code, name, unit, qty, rop, loc, active]
            else:
                items = pd.concat([items, pd.DataFrame([[code_gen, cat_code, name, unit, qty, rop, loc, active]], columns=ITEMS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_ITEMS, items)
            st.success(f"บันทึกแล้ว: {code_gen}")
            safe_rerun()

    with tab_edit:
        if items.empty:
            st.info("ยังไม่มีรายการให้แก้ไข")
            return
        # แสดงตารางพร้อมเช็คบ็อกซ์ 'เลือก'
        if hasattr(st, "data_editor"):
            items_editable = items.copy()
            items_editable.insert(0, "เลือก", False)
            edited = st.data_editor(items_editable, hide_index=True, num_rows="fixed", height=360, use_container_width=True,
                                    column_config={"เลือก": st.column_config.CheckboxColumn(required=False)})
            picked = edited[edited["เลือก"]==True]
            if picked.empty:
                st.info("ติ๊กเลือกรายการจากคอลัมน์ 'เลือก' เพื่อแก้ไข/ลบ")
                return
            target_code = str(picked.iloc[0]["รหัส"])
        else:
            st.dataframe(items, height=360, use_container_width=True)
            target_code = st.text_input("ระบุรหัสที่จะปรับแก้")
            if not target_code: return

        row = items[items["รหัส"]==target_code]
        if row.empty:
            st.warning("ไม่พบรหัสนี้"); return
        data = row.iloc[0]

        with st.form("edit_item", clear_on_submit=False):
            c1,c2,c3 = st.columns(3)
            with c1:
                code = st.text_input("รหัส", value=data["รหัส"], disabled=True)
                cat  = st.text_input("หมวดหมู่", value=str(data["หมวดหมู่"]))
                name = st.text_input("ชื่ออุปกรณ์", value=str(data["ชื่ออุปกรณ์"]))
            with c2:
                unit = st.text_input("หน่วย", value=str(data["หน่วย"]))
                qty = st.number_input("คงเหลือ", 0, 10**9, int(pd.to_numeric(data["คงเหลือ"], errors="coerce") or 0), 1)
                rop = st.number_input("จุดสั่งซื้อ", 0, 10**9, int(pd.to_numeric(data["จุดสั่งซื้อ"], errors="coerce") or 0), 1)
            with c3:
                loc = st.text_input("ที่เก็บ", value=str(data["ที่เก็บ"]))
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(data["ใช้งาน"]).upper()=="Y" else 1)
                b1, b2 = st.columns([2,1])
                save_btn = b1.form_submit_button("บันทึกการแก้ไข", use_container_width=True)
                del_btn  = b2.form_submit_button("ลบรายการ", use_container_width=True)

        if save_btn:
            items.loc[items["รหัส"]==code, ITEMS_HEADERS] = [code, cat, name, unit, qty, rop, loc, active]
            write_df(sh, SHEET_ITEMS, items); st.success("บันทึกแล้ว"); safe_rerun()
        if del_btn:
            items2 = items[items["รหัส"]!=code]
            write_df(sh, SHEET_ITEMS, items2); st.success("ลบแล้ว"); safe_rerun()

def page_issue_receive(sh):
    st.subheader("🧾 เบิก/รับเข้า")
    if st.session_state.get("role") not in ("admin","staff"):
        st.info("สิทธิ์ผู้ชมไม่สามารถบันทึกได้"); return
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    if items.empty:
        st.info("ยังไม่มีสินค้า"); return
    t1, t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])
    with t1:
        with st.form("out", clear_on_submit=True):
            item = st.selectbox("อุปกรณ์", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
            qty  = st.number_input("จำนวน", 1, 10**9, 1, 1)
            branch = st.text_input("สาขา/ปลายทาง", "")
            note = st.text_input("หมายเหตุ", "")
            s = st.form_submit_button("บันทึกเบิกออก", use_container_width=True)
        if s:
            code = item.split(" | ")[0]
            ok = adjust_stock(sh, code, -qty, st.session_state.get("user","unknown"), branch, note, "OUT")
            if ok: st.success("บันทึกแล้ว"); safe_rerun()
    with t2:
        with st.form("in", clear_on_submit=True):
            item = st.selectbox("อุปกรณ์", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist(), key="in_pick")
            qty  = st.number_input("จำนวน", 1, 10**9, 1, 1, key="in_qty")
            src  = st.text_input("แหล่งที่มา/เลข PO", "", key="in_src")
            note = st.text_input("หมายเหตุ", "", key="in_note")
            s = st.form_submit_button("บันทึกรับเข้า", use_container_width=True)
        if s:
            code = item.split(" | ")[0]
            ok = adjust_stock(sh, code, qty, st.session_state.get("user","unknown"), src, note, "IN")
            if ok: st.success("บันทึกแล้ว"); safe_rerun()

def page_tickets(sh):
    st.subheader("🛠️ แจ้งปัญหา (Tickets)")
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    t_cats = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)

    st.markdown("### รายการล่าสุด")
    if hasattr(st, "data_editor"):
        tdisp = tickets.copy()
        if not tdisp.empty:
            tdisp.insert(0, "เลือก", False)
        edited = st.data_editor(tdisp if not tickets.empty else pd.DataFrame(columns=["เลือก"]+TICKETS_HEADERS),
                                hide_index=True, use_container_width=True, height=300,
                                column_config={"เลือก": st.column_config.CheckboxColumn()})
        picked = edited[edited["เลือก"]==True]
    else:
        st.dataframe(tickets, use_container_width=True, height=300)
        picked = pd.DataFrame()

    t_add, t_edit = st.tabs(["➕ รับแจ้งใหม่","✏️ แก้ไข/ลบ (ติ๊กเลือก)"])
    with t_add:
        with st.form("tk_new", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
                b_opts = (branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else []
                branch = st.selectbox("สาขา", b_opts)
                reporter = st.text_input("ผู้แจ้ง", "")
            with c2:
                cat_opts = (t_cats["รหัสหมวดปัญหา"]+" | "+t_cats["ชื่อหมวดปัญหา"]).tolist() if not t_cats.empty else []
                cate = st.selectbox("หมวดหมู่ปัญหา", cat_opts)
                assignee = st.text_input("ผู้รับผิดชอบ (IT)", st.session_state.get("user",""))
            with c3:
                detail = st.text_area("รายละเอียด", height=100)
                note = st.text_input("หมายเหตุ", "")
            s = st.form_submit_button("บันทึกการรับแจ้ง", use_container_width=True)
        if s:
            tid = "TCK-" + datetime.now(TZ).strftime("%Y%m%d-%H%M%S")
            row = [tid, get_now_str(), branch, reporter, cate, detail, "รับแจ้ง", assignee, get_now_str(), note]
            append_row(sh, SHEET_TICKETS, row)
            st.success(f"รับแจ้งแล้ว: {tid}")
            safe_rerun()

    with t_edit:
        if picked.empty:
            st.info("ติ๊กเลือกรายการจากตารางด้านบนก่อน")
            return
        pick_id = str(picked.iloc[0]["TicketID"]) if "TicketID" in picked.columns else ""
        row = tickets[tickets["TicketID"]==pick_id]
        if row.empty:
            st.warning("ไม่พบ Ticket ที่เลือก"); return
        data = row.iloc[0]

        with st.form("tk_edit", clear_on_submit=False):
            c1,c2 = st.columns(2)
            with c1:
                branch = st.text_input("สาขา", value=str(data.get("สาขา","")))
                owner  = st.text_input("ผู้แจ้ง", value=str(data.get("ผู้แจ้ง","")))
            with c2:
                status = st.selectbox("สถานะ", ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"],
                                      index=["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"].index(str(data.get("สถานะ","รับแจ้ง"))) if str(data.get("สถานะ","รับแจ้ง")) in ["รับแจ้ง","กำลังดำเนินการ","ดำเนินการเสร็จ"] else 0)
                assignee = st.text_input("ผู้รับผิดชอบ", value=str(data.get("ผู้รับผิดชอบ","")))
            desc = st.text_area("รายละเอียด", value=str(data.get("รายละเอียด","")), height=120)
            note = st.text_input("หมายเหตุ", value=str(data.get("หมายเหตุ","")))
            b1,b2 = st.columns([2,1])
            up_btn = b1.form_submit_button("อัปเดต", use_container_width=True)
            del_btn= b2.form_submit_button("ลบ", use_container_width=True)

        if up_btn:
            idx = tickets.index[tickets["TicketID"]==pick_id][0]
            tickets.at[idx,"สาขา"] = branch
            tickets.at[idx,"ผู้แจ้ง"] = owner
            tickets.at[idx,"รายละเอียด"] = desc
            tickets.at[idx,"สถานะ"] = status
            tickets.at[idx,"ผู้รับผิดชอบ"] = assignee
            tickets.at[idx,"หมายเหตุ"] = note
            tickets.at[idx,"อัปเดตล่าสุด"] = get_now_str()
            write_df(sh, SHEET_TICKETS, tickets); st.success("อัปเดตแล้ว"); safe_rerun()
        if del_btn:
            tickets2 = tickets[tickets["TicketID"]!=pick_id]
            write_df(sh, SHEET_TICKETS, tickets2); st.success("ลบแล้ว"); safe_rerun()

def page_reports(sh):
    st.subheader("📑 รายงาน / ประวัติ")
    txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    q = st.text_input("ค้นหา (ชื่อ/รหัส/สาขา)")
    today = datetime.now(TZ).date()
    colR = st.columns(5)
    with colR[0]: today_btn = st.button("วันนี้")
    with colR[1]: d7_btn = st.button("7 วันล่าสุด")
    with colR[2]: d30_btn = st.button("30 วันล่าสุด")
    with colR[3]: d90_btn = st.button("90 วันล่าสุด")
    with colR[4]: year_btn = st.button("ปีนี้")

    if "report_d1" not in st.session_state:
        st.session_state["report_d1"] = today - timedelta(days=30)
        st.session_state["report_d2"] = today
    if today_btn:
        st.session_state["report_d1"] = today
        st.session_state["report_d2"] = today
    if d7_btn:
        st.session_state["report_d1"] = today-timedelta(days=6)
        st.session_state["report_d2"] = today
    if d30_btn:
        st.session_state["report_d1"] = today-timedelta(days=29)
        st.session_state["report_d2"] = today
    if d90_btn:
        st.session_state["report_d1"] = today-timedelta(days=89)
        st.session_state["report_d2"] = today
    if year_btn:
        st.session_state["report_d1"] = date(today.year,1,1)
        st.session_state["report_d2"] = today

    d1 = st.date_input("ตั้งแต่", value=st.session_state["report_d1"])
    d2 = st.date_input("ถึง", value=st.session_state["report_d2"])
    st.session_state["report_d1"] = d1
    st.session_state["report_d2"] = d2

    if not txns.empty:
        df = txns.copy()
        df["วันเวลา"] = pd.to_datetime(df["วันเวลา"], errors="coerce")
        df = df.dropna(subset=["วันเวลา"])
        df = df[(df["วันเวลา"].dt.date >= d1) & (df["วันเวลา"].dt.date <= d2)]
        if q:
            mask = (
                df["ชื่ออุปกรณ์"].astype(str).str.contains(q, case=False, na=False) |
                df["รหัส"].astype(str).str.contains(q, case=False, na=False) |
                df["สาขา"].astype(str).str.contains(q, case=False, na=False)
            )
            df = df[mask]
    else:
        df = pd.DataFrame(columns=TXNS_HEADERS)
    st.dataframe(df.sort_values("วันเวลา", ascending=False), height=380, use_container_width=True)

def page_users(sh):
    st.subheader("👥 ผู้ใช้ (Admin)")
    users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    if hasattr(st, "data_editor"):
        users_disp = users.copy()
        users_disp.insert(0, "เลือก", False)
        edited = st.data_editor(users_disp, hide_index=True, num_rows="fixed", height=300, use_container_width=True,
                                column_config={"เลือก": st.column_config.CheckboxColumn()})
        picked = edited[edited["เลือก"]==True]
    else:
        st.dataframe(users, height=300, use_container_width=True)
        picked = pd.DataFrame()

    t_add, t_edit = st.tabs(["➕ เพิ่มผู้ใช้","✏️ แก้ไขผู้ใช้ (ติ๊กเลือก)"])

    with t_add:
        with st.form("add_user", clear_on_submit=True):
            c1,c2 = st.columns([2,1])
            with c1:
                un = st.text_input("Username*")
                dn = st.text_input("Display Name")
            with c2:
                role = st.selectbox("Role", ["admin","staff","viewer"], index=1)
                active = st.selectbox("Active", ["Y","N"], index=0)
            pwd = st.text_input("กำหนดรหัสผ่าน*", type="password")
            s = st.form_submit_button("บันทึกผู้ใช้ใหม่", use_container_width=True)
        if s:
            if not un.strip() or not pwd.strip():
                st.warning("กรอก Username/Password"); st.stop()
            if (users["Username"]==un).any():
                st.error("มี Username นี้แล้ว"); st.stop()
            ph = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
            new_row = pd.DataFrame([{
                "Username": un.strip(),
                "DisplayName": dn.strip(),
                "Role": role,
                "PasswordHash": ph,
                "Active": active,
            }])
            users2 = pd.concat([users, new_row], ignore_index=True)
            write_df(sh, SHEET_USERS, users2)
            st.success("เพิ่มผู้ใช้สำเร็จ"); safe_rerun()

    with t_edit:
        if picked.empty:
            st.info("ติ๊กเลือกรายการจากตารางด้านบนก่อน")
            return
        username = str(picked.iloc[0]["Username"])
        row = users[users["Username"]==username]
        if row.empty:
            st.warning("ไม่พบผู้ใช้ที่เลือก"); return
        data = row.iloc[0].to_dict()
        with st.form("edit_user", clear_on_submit=False):
            c1,c2 = st.columns([2,1])
            with c1:
                un = st.text_input("Username", value=data["Username"], disabled=True)
                dn = st.text_input("Display Name", value=data["DisplayName"])
            with c2:
                role = st.selectbox("Role", ["admin","staff","viewer"],
                                    index=["admin","staff","viewer"].index(data["Role"]) if data["Role"] in ["admin","staff","viewer"] else 1)
                active = st.selectbox("Active", ["Y","N"],
                                      index=["Y","N"].index(data["Active"]) if data["Active"] in ["Y","N"] else 0)
            pwd = st.text_input("ตั้ง/รีเซ็ตรหัสผ่าน (ปล่อยว่าง = ไม่เปลี่ยน)", type="password")
            b1,b2 = st.columns([2,1])
            save_btn = b1.form_submit_button("บันทึกการแก้ไข", use_container_width=True)
            del_btn  = b2.form_submit_button("ลบผู้ใช้นี้", use_container_width=True)
        if del_btn:
            if un.lower()=="admin":
                st.error("ห้ามลบผู้ใช้ admin")
            else:
                users2 = users[users["Username"]!=un]
                write_df(sh, SHEET_USERS, users2); st.success("ลบแล้ว"); safe_rerun()
        if save_btn:
            idx = users.index[users["Username"]==un][0]
            users.at[idx,"DisplayName"] = dn
            users.at[idx,"Role"] = role
            users.at[idx,"Active"] = active
            if pwd.strip():
                ph = bcrypt.hashpw(pwd.encode("utf-8"), bcrypt.gensalt(12)).decode("utf-8")
                users.at[idx,"PasswordHash"] = ph
            write_df(sh, SHEET_USERS, users)
            st.success("บันทึกแล้ว"); safe_rerun()

def page_import(sh):
    st.subheader("🗂️ นำเข้า/แก้ไข หมวดหมู่/สาขา แบบง่าย")
    t_cat, t_br = st.tabs(["หมวดหมู่","สาขา"])

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

    with t_cat:
        up = st.file_uploader("อัปโหลดไฟล์ หมวดหมู่ (CSV/Excel) : ต้องมีคอลัมน์ รหัสหมวด, ชื่อหมวด", type=["csv","xlsx"], key="up_cat1")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df, height=200, use_container_width=True)
                if {"รหัสหมวด","ชื่อหมวด"}.issubset(df.columns):
                    if st.button("นำเข้า/อัปเดต หมวดหมู่", use_container_width=True):
                        cur = read_df(sh, SHEET_CATS, CATS_HEADERS)
                        for _, r in df.iterrows():
                            code = str(r["รหัสหมวด"]).strip(); name = str(r["ชื่อหมวด"]).strip()
                            if not code: continue
                            if (cur["รหัสหมวด"]==code).any():
                                cur.loc[cur["รหัสหมวด"]==code, ["รหัสหมวด","ชื่อหมวด"]] = [code, name]
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code,name]], columns=CATS_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_CATS, cur); st.success("นำเข้าแล้ว"); safe_rerun()
                else:
                    st.warning("หัวตารางไม่ครบ")

    with t_br:
        up = st.file_uploader("อัปโหลดไฟล์ สาขา (CSV/Excel) : ต้องมีคอลัมน์ รหัสสาขา, ชื่อสาขา", type=["csv","xlsx"], key="up_br1")
        if up:
            df, err = _read_upload_df(up)
            if err: st.error(err)
            else:
                st.dataframe(df, height=200, use_container_width=True)
                if {"รหัสสาขา","ชื่อสาขา"}.issubset(df.columns):
                    if st.button("นำเข้า/อัปเดต สาขา", use_container_width=True):
                        cur = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
                        for _, r in df.iterrows():
                            code = str(r["รหัสสาขา"]).strip(); name = str(r["ชื่อสาขา"]).strip()
                            if not code: continue
                            if (cur["รหัสสาขา"]==code).any():
                                cur.loc[cur["รหัสสาขา"]==code, ["รหัสสาขา","ชื่อสาขา"]] = [code, name]
                            else:
                                cur = pd.concat([cur, pd.DataFrame([[code,name]], columns=BR_HEADERS)], ignore_index=True)
                        write_df(sh, SHEET_BRANCHES, cur); st.success("นำเข้าแล้ว"); safe_rerun()
                else:
                    st.warning("หัวตารางไม่ครบ")

def page_settings():
    st.subheader("⚙️ Settings")
    url = st.text_input("Google Sheet URL", value=st.session_state.get("sheet_url", DEFAULT_SHEET_URL))
    st.session_state["sheet_url"] = url
    if st.button("ทดสอบเชื่อมต่อ/ตรวจสอบชีตที่จำเป็น", use_container_width=True):
        try:
            sh = open_sheet_by_url(url)
            ensure_sheets_exist(sh)
            st.success("เชื่อมต่อสำเร็จ")
        except Exception as e:
            st.error(f"เชื่อมต่อไม่สำเร็จ: {e}")

# =========================
# Main
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    st.title(APP_TITLE); st.caption(APP_TAGLINE)

    if "sheet_url" not in st.session_state or not st.session_state.get("sheet_url"):
        st.session_state["sheet_url"] = DEFAULT_SHEET_URL

    with st.sidebar:
        st.markdown("---")
        page = st.radio("เมนู", ["📊 Dashboard","📦 คลังอุปกรณ์","🛠️ แจ้งปัญหา","🧾 เบิก/รับเข้า","📑 รายงาน","👥 ผู้ใช้","🗂️ นำเข้า/แก้ไข หมวดหมู่","⚙️ Settings"], index=0)

    # เปิดชีต
    try:
        sh = open_sheet_by_url(st.session_state["sheet_url"])
    except Exception as e:
        st.error(f"เปิดชีตไม่สำเร็จ: {e}")
        return

    ensure_sheets_exist(sh)

    # Auth (ยกเว้นหน้า Settings/Dashboard ให้เปิดดูได้, ที่เหลือต้อง Login)
    if not page.startswith("⚙️") and not page.startswith("📊"):
        if not auth_block(sh):
            st.info("เข้าสู่ระบบก่อนเพื่อใช้งานเมนูอื่น")
            return
    else:
        # ยังแสดงกรอบ login ได้ทาง sidebar
        auth_block(sh)

    if page.startswith("📊"): page_dashboard(sh)
    elif page.startswith("📦"): page_stock(sh)
    elif page.startswith("🛠️"): page_tickets(sh)
    elif page.startswith("🧾"): page_issue_receive(sh)
    elif page.startswith("📑"): page_reports(sh)
    elif page.startswith("👥"): page_users(sh)
    elif page.startswith("🗂️"): page_import(sh)
    elif page.startswith("⚙️"): page_settings()

    st.caption("© 2025 IT Stock · Streamlit + Google Sheets · iTao iT v11 (patched)")


if __name__ == "__main__":
    main()
