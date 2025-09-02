# -*- coding: utf-8 -*-
"""
app_v11_full_logo_pdf.py
IT Stock (Streamlit + Google Sheets) — v11 (Thai PDF + Logo)
--------------------------------------------------------------------
ฟีเจอร์หลัก (สรุป):
- Dashboard, คลังอุปกรณ์ (Stock), เบิก/รับ (OUT/IN แบบเลือกหลายรายการ),
  รายงาน (Reports → ออกรายงานเป็นไฟล์ PDF รองรับภาษาไทย + ใส่โลโก้),
  นำเข้า (Import หลายแท็บ), ผู้ใช้ (Users เบื้องต้น), Settings (ตั้งค่า GSheets/Fonts/Logo)
- รองรับฟอนต์ไทย (Sarabun / TH Sarabun New / Noto Sans Thai) สำหรับสร้าง PDF
  * ค้นหาอัตโนมัติจาก ./fonts, Windows Fonts, และตำแหน่งทั่วไปบน Linux/Mac
  * ถ้าไม่พบฟอนต์ จะแจ้งเตือนบนหน้าเว็บ และยังสร้าง PDF ได้ (แต่อาจแสดงเป็นสี่เหลี่ยม)
- โหมดเบิกหลายรายการ: เลือกหลายรายการจากสต็อก แล้วกำหนดจำนวนที่จะเบิกทีละรายการ
- โครงสร้างโค้ดแยกเป็นฟังก์ชัน อ่าน/เขียน Google Sheets (ผ่าน gspread) หรือโหมดไฟล์ CSV
- UI เป็นมิตรกับสมาร์ทโฟน: ใช้ layout="wide" + คอมโพเนนต์ที่จับถนัด
หมายเหตุ:
- โค้ดนี้เป็นโครง “พร้อมใช้งานจริง” แต่ยังคงย่อความให้พอเหมาะ (ขนาดไฟล์ไม่ยาวเกินไป)
- หากต้องการฟีเจอร์เฉพาะทางเพิ่มเติม สามารถต่อยอดได้ทันทีในแต่ละส่วน
"""
from __future__ import annotations

import os, io, sys, re, json, uuid, time, pathlib, base64
from datetime import datetime, date
from typing import Dict, Optional, List, Tuple

import pandas as pd
import numpy as np
import streamlit as st

# ---------------- PDF (ReportLab) ----------------
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

# ---------------- Google Sheets -----------------
GS_AVAILABLE = True
try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    GS_AVAILABLE = False

APP_VERSION = "v11"
DEFAULT_DATA_DIR = "./data"
DEFAULT_FONTS_DIR = "./fonts"
DEFAULT_ASSETS_DIR = "./assets"

# ---------------- Streamlit page config ----------------
st.set_page_config(
    page_title=f"IT Stock {APP_VERSION}",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =======================================================
# Utilities
# =======================================================
def ensure_dirs():
    os.makedirs(DEFAULT_DATA_DIR, exist_ok=True)
    os.makedirs(DEFAULT_FONTS_DIR, exist_ok=True)
    os.makedirs(DEFAULT_ASSETS_DIR, exist_ok=True)

ensure_dirs()

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def load_json(text: str) -> dict:
    try:
        return json.loads(text)
    except Exception:
        return {}

def to_int(x, default=0):
    try:
        return int(x)
    except Exception:
        return default

def to_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default

# =======================================================
# Config (in Session State)
# =======================================================
if "cfg" not in st.session_state:
    st.session_state["cfg"] = {
        "use_gsheets": False,
        "sheet_url": "",
        "service_account_json_text": "",   # raw JSON text (optional; เก็บแบบเข้าหน่วยความจำ)
        "service_account_json_file": "",   # path ไปยังไฟล์ .json บนเครื่องเซิร์ฟเวอร์
        "pdf_font_regular": "",            # path font TTF สำหรับข้อความปกติ
        "pdf_font_bold": "",               # path font TTF แบบหนา (ถ้ามี)
        "logo_path": "",                   # assets/logo.png (หรืออัปโหลดชั่วคราว)
        "branch_code_name": {},            # mapping โค้ดสาขา → ชื่อสาขา
    }

CFG = st.session_state["cfg"]

# =======================================================
# Google Sheets Helpers (with CSV fallback)
# =======================================================
GS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_gs_client() -> Optional["gspread.Client"]:
    if not GS_AVAILABLE:
        return None
    if not CFG.get("use_gsheets"):
        return None

    creds = None
    # กรณีมี raw JSON ใน settings
    if CFG.get("service_account_json_text"):
        try:
            info = json.loads(CFG["service_account_json_text"])
            creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"โหลด service_account_json_text ไม่สำเร็จ: {e}")

    # กรณีมี path ไปไฟล์ .json
    if creds is None and CFG.get("service_account_json_file"):
        try:
            creds = Credentials.from_service_account_file(CFG["service_account_json_file"], scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"โหลด service_account_json_file ไม่สำเร็จ: {e}")

    # กรณีมี env var
    if creds is None and os.environ.get("SERVICE_ACCOUNT_JSON"):
        try:
            info = json.loads(os.environ["SERVICE_ACCOUNT_JSON"])
            creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"โหลด SERVICE_ACCOUNT_JSON จาก ENV ไม่สำเร็จ: {e}")

    if creds is None:
        return None

    try:
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"เชื่อมต่อ gspread ไม่สำเร็จ: {e}")
        return None

def read_table(name: str) -> pd.DataFrame:
    """
    อ่านตารางจาก Google Sheets (ถ้าตั้งค่าไว้) หรือจาก CSV ใน ./data/{name}.csv
    """
    # 1) ลองอ่านจาก Google Sheets
    if CFG.get("use_gsheets") and CFG.get("sheet_url"):
        client = get_gs_client()
        if client is not None:
            try:
                sh = client.open_by_url(CFG["sheet_url"])
                try:
                    ws = sh.worksheet(name)
                except Exception:
                    # ไม่มีชีท → สร้างใหม่ด้วย header อย่างน้อย 1 ช่อง
                    ws = sh.add_worksheet(title=name, rows=100, cols=26)
                    ws.append_row(["_init"])
                rows = ws.get_all_records()
                df = pd.DataFrame(rows)
                if df.empty:
                    return pd.DataFrame()
                return df
            except Exception as e:
                st.warning(f"อ่านชีท '{name}' จาก Google Sheets ไม่สำเร็จ: {e} → จะอ่านจาก CSV")
    # 2) Fallback CSV
    csv_path = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
    if os.path.exists(csv_path):
        try:
            return pd.read_csv(csv_path, dtype=str).fillna("")
        except Exception:
            pass
    return pd.DataFrame()

def write_table(name: str, df: pd.DataFrame):
    """
    เขียนตารางลง Google Sheets (ถ้าตั้งค่าไว้) หรือบันทึกเป็น CSV
    """
    df = df.copy()
    # แปลงทุกคอลัมน์เป็น string เพื่อความง่าย (Google Sheets ชอบ)
    for c in df.columns:
        df[c] = df[c].astype(str)

    if CFG.get("use_gsheets") and CFG.get("sheet_url"):
        client = get_gs_client()
        if client is not None:
            try:
                sh = client.open_by_url(CFG["sheet_url"])
                try:
                    ws = sh.worksheet(name)
                except Exception:
                    ws = sh.add_worksheet(title=name, rows=max(len(df)+10, 100), cols=max(len(df.columns)+2, 26))
                # clear & update
                ws.clear()
                ws.update([df.columns.tolist()] + df.values.tolist())
                return
            except Exception as e:
                st.warning(f"เขียนชีท '{name}' ไป Google Sheets ไม่สำเร็จ: {e} → จะบันทึก CSV")

    # Fallback CSV
    csv_path = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

# =======================================================
# Schemas
# =======================================================
SCHEMA_STOCK = [
    "item_code","item_name","category","unit","qty","min_qty",
    "branch_code","branch_name","last_update"
]
SCHEMA_OUT = [
    "run","date","branch_code","branch_name","requester",
    "item_code","item_name","qty","unit","note","status"
]
SCHEMA_IN = [
    "run","date","branch_code","branch_name","receiver",
    "item_code","item_name","qty","unit","note","ref_out_run"
]
SCHEMA_USERS = ["username","full_name","role","branch_code","branch_name","active"]
SCHEMA_CATEGORIES = ["cat_code","cat_name","active"]

def ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    # drop อื่นๆ ที่ไม่ใช้? เก็บไว้เพื่อไม่เสียข้อมูล
    return df[cols + [c for c in df.columns if c not in cols]]

# =======================================================
# Initial Load
# =======================================================
def load_all_tables():
    stock = ensure_columns(read_table("stock"), SCHEMA_STOCK)
    trans_out = ensure_columns(read_table("out"), SCHEMA_OUT)
    trans_in = ensure_columns(read_table("in"), SCHEMA_IN)
    users = ensure_columns(read_table("users"), SCHEMA_USERS)
    cats = ensure_columns(read_table("categories"), SCHEMA_CATEGORIES)
    return stock, trans_out, trans_in, users, cats

if "tables" not in st.session_state:
    st.session_state["tables"] = load_all_tables()

# short aliases
STOCK, OUTS, INS, USERS, CATS = st.session_state["tables"]

# =======================================================
# Fonts Finder for Thai
# =======================================================
COMMON_THAI_FONT_NAMES = [
    "THSarabunNew", "TH Sarabun New",
    "Sarabun", "NotoSansThai", "Noto Sans Thai",
    "NotoSerifThai", "Noto Serif Thai",
]

COMMON_FONT_DIRS = [
    DEFAULT_FONTS_DIR,
    "/usr/share/fonts/truetype",
    "/usr/share/fonts",
    "/Library/Fonts",
    "/System/Library/Fonts",
    "C:\\Windows\\Fonts",
]

def find_font_by_names(names: List[str]) -> Optional[str]:
    for d in COMMON_FONT_DIRS:
        try:
            for fn in os.listdir(d):
                path = os.path.join(d, fn)
                lower = fn.lower()
                for name in names:
                    if name.replace(" ", "").lower() in lower.replace(" ", "") and lower.endswith(".ttf"):
                        return path
        except Exception:
            continue
    return None

def register_thai_fonts():
    # ใช้ค่าจาก Settings ก่อน
    reg = CFG.get("pdf_font_regular") or find_font_by_names(COMMON_THAI_FONT_NAMES)
    bold = CFG.get("pdf_font_bold") or find_font_by_names([n + " Bold" for n in COMMON_THAI_FONT_NAMES] + COMMON_THAI_FONT_NAMES)
    ok = False
    try:
        if reg:
            pdfmetrics.registerFont(TTFont("TH_REG", reg))
            ok = True
        if bold:
            pdfmetrics.registerFont(TTFont("TH_BOLD", bold))
            ok = True
    except Exception as e:
        st.warning(f"ลงทะเบียนฟอนต์ไทยไม่สำเร็จ: {e}")
    return ok

FONTS_READY = register_thai_fonts()

# =======================================================
# PDF Generation
# =======================================================
def generate_pdf_report(title: str, df: pd.DataFrame, logo_path: str="") -> bytes:
    """
    สร้าง PDF (A4 แนวตั้ง) แสดงหัวเรื่อง + ตารางข้อมูลย่อ (หน้าเดียวถ้าเป็นไปได้)
    รองรับฟอนต์ไทย (ถ้าลงทะเบียนสำเร็จ)
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    # โลโก้
    y = H - 30*mm
    if logo_path and os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            c.drawImage(img, 15*mm, y-10*mm, width=25*mm, height=25*mm, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # หัวเรื่อง
    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 18)
    c.drawString(45*mm, H-20*mm, title)

    # วันที่
    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 10)
    c.drawRightString(W-15*mm, H-15*mm, f"พิมพ์เมื่อ: {now_str()}")

    # ตาราง (อย่างย่อ 6-8 คอลัมน์แรก)
    cols = list(df.columns)[:8]
    show = df[cols].astype(str).values.tolist()

    # เฮด
    x0, y0 = 15*mm, H-40*mm
    row_h = 8*mm
    col_w = (W - 30*mm) / max(1, len(cols))
    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 10)
    for i, col in enumerate(cols):
        c.drawString(x0 + i*col_w + 2, y0, str(col))

    # เส้นใต้หัว
    c.line(x0, y0-2, x0 + col_w*len(cols), y0-2)
    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)

    ycur = y0 - row_h
    for r in show[:50]:   # จำกัดแถวเพื่อให้อยู่หน้าเดียว
        for i, val in enumerate(r):
            c.drawString(x0 + i*col_w + 2, ycur, str(val)[:40])
        ycur -= row_h
        if ycur < 20*mm:
            break

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# =======================================================
# UI Components
# =======================================================
def section_title(title: str, emoji: str=""):
    st.subheader(f"{emoji} {title}".strip())

def df_editor(df: pd.DataFrame, key: str, use_container_width=True, height=360):
    return st.data_editor(
        df, key=key, use_container_width=use_container_width, height=height,
        hide_index=True, num_rows="dynamic"
    )

def ensure_branch_map():
    if not CFG["branch_code_name"]:
        # ตัวอย่าง mapping เริ่มต้น (แก้ไขได้ใน Settings → Branches)
        CFG["branch_code_name"] = {
            "SWC001": "สาขากรุงเทพฯ (ตัวอย่าง)",
            "SWC002": "สาขานครราชสีมา (ตัวอย่าง)",
            "SWC003": "สาขาขอนแก่น (ตัวอย่าง)",
        }

ensure_branch_map()

def code_to_name(code: str) -> str:
    return CFG["branch_code_name"].get(code, "")

# =======================================================
# Pages
# =======================================================
def page_dashboard():
    section_title("Dashboard", "📊")
    col1, col2, col3 = st.columns(3)
    # KPI ง่ายๆ
    total_items = len(STOCK)
    low_items = (pd.to_numeric(STOCK["qty"], errors="coerce").fillna(0) <= pd.to_numeric(STOCK["min_qty"], errors="coerce").fillna(0)).sum()
    total_out = len(OUTS)
    col1.metric("จำนวนรายการในคลัง", f"{total_items:,}")
    col2.metric("ใกล้ต่ำกว่าขั้นต่ำ", f"{low_items:,}")
    col3.metric("รายการเบิกรวม", f"{total_out:,}")

    st.divider()
    st.write("**ภาพรวมสต็อกล่าสุด**")
    st.dataframe(STOCK.head(50), use_container_width=True)

def page_stock():
    section_title("คลังอุปกรณ์ (Stock)", "📦")
    st.info("เพิ่ม/แก้ไขข้อมูลคงคลังได้โดยตรงในตาราง จากนั้นกดปุ่ม **บันทึก** ด้านล่าง", icon="ℹ️")
    editable = df_editor(STOCK, key="stock_editor")
    if st.button("💾 บันทึกคลังอุปกรณ์", type="primary"):
        st.session_state["tables"] = (editable, OUTS, INS, USERS, CATS)
        write_table("stock", editable)
        st.success("บันทึกเรียบร้อย ✅")

def page_out_in():
    section_title("เบิก/รับ (OUT/IN)", "🧾")
    tab_out, tab_in = st.tabs(["🔻 เบิกหลายรายการ (OUT)", "🔺 รับเข้า (IN)"])

    with tab_out:
        st.caption("เลือกหลายรายการจากคลัง → ใส่จำนวนที่จะเบิก → ยืนยันเพื่อบันทึกธุรกรรม OUT")
        left, right = st.columns([1,1])
        with left:
            branches = list(CFG["branch_code_name"].keys())
            sel_branch = st.selectbox("เลือกสาขา", options=branches, format_func=lambda c: f"{c} - {code_to_name(c)}")
            requester = st.text_input("ชื่อผู้เบิก/ผู้แจ้ง", "")
            note = st.text_input("หมายเหตุ (ถ้ามี)", "")
        with right:
            # เลือกหลายรายการ
            st.write("**เลือกอุปกรณ์ที่จะเบิก (Multi-select)**")
            df_pick = STOCK[["item_code","item_name","category","unit","qty","min_qty"]].copy()
            pick = st.multiselect(
                "ค้นหา/เลือก",
                options=list(df_pick.index),
                format_func=lambda idx: f"{df_pick.at[idx,'item_code']} | {df_pick.at[idx,'item_name']} ({df_pick.at[idx,'qty']})",
            )
            qty_inputs = {}
            if pick:
                st.write("**กำหนดจำนวนที่จะเบิก**")
                for idx in pick:
                    row = df_pick.loc[idx]
                    maxq = to_int(row["qty"], 0)
                    qty = st.number_input(
                        f"{row['item_code']} | {row['item_name']} (คงเหลือ {maxq})",
                        min_value=0, max_value=max(0, maxq), step=1, value=0, key=f"qty_{idx}"
                    )
                    qty_inputs[idx] = qty

            if st.button("✅ ยืนยันการเบิก (OUT)"):
                if not pick:
                    st.warning("ยังไม่ได้เลือกรายการ")
                elif not requester.strip():
                    st.warning("กรุณากรอกชื่อผู้เบิก")
                else:
                    # บันทึก OUT และหักสต็อก
                    out_df = OUTS.copy()
                    stock_df = STOCK.copy()
                    new_rows = []
                    for idx in pick:
                        q = to_int(qty_inputs.get(idx, 0), 0)
                        if q <= 0: 
                            continue
                        srow = stock_df.loc[idx]
                        cur = to_int(srow["qty"], 0)
                        new_qty = max(0, cur - q)
                        stock_df.at[idx, "qty"] = str(new_qty)
                        new_rows.append({
                            "run": f"OUT-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
                            "date": now_str(),
                            "branch_code": sel_branch,
                            "branch_name": code_to_name(sel_branch),
                            "requester": requester,
                            "item_code": srow["item_code"],
                            "item_name": srow["item_name"],
                            "qty": str(q),
                            "unit": srow.get("unit",""),
                            "note": note,
                            "status": "DONE",
                        })
                    if new_rows:
                        out_df = pd.concat([out_df, pd.DataFrame(new_rows)], ignore_index=True)
                        # save
                        st.session_state["tables"] = (stock_df, out_df, INS, USERS, CATS)
                        write_table("stock", stock_df)
                        write_table("out", out_df)
                        st.success(f"บันทึก OUT {len(new_rows)} รายการสำเร็จ ✅")

    with tab_in:
        st.caption("รับเข้า/เพิ่มสต็อก (IN) จากการซื้อใหม่หรือคืนของ")
        left, right = st.columns([1,1])
        with left:
            branches = list(CFG["branch_code_name"].keys())
            sel_branch = st.selectbox("เลือกสาขา", options=branches, format_func=lambda c: f"{c} - {code_to_name(c)}", key="in_branch")
            receiver = st.text_input("ผู้รับเข้า", "", key="in_receiver")
            note = st.text_input("หมายเหตุ (ถ้ามี)", "", key="in_note")
        with right:
            # เลือกอุปกรณ์ 1 รายการ + จำนวน (ง่ายๆ)
            # ถ้าต้องการ IN หลายชิ้นในครั้งเดียว สามารถต่อยอดโดย copy แนวคิดแบบ OUT ด้านบน
            df_pick = STOCK[["item_code","item_name","unit","qty"]].copy()
            idx = st.selectbox("เลือกรายการ", options=list(df_pick.index),
                               format_func=lambda i: f"{df_pick.at[i,'item_code']} | {df_pick.at[i,'item_name']} (คงเหลือ {df_pick.at[i,'qty']})")
            in_qty = st.number_input("จำนวนรับเข้า", min_value=0, step=1, value=0)
            if st.button("✅ ยืนยันการรับเข้า (IN)"):
                if in_qty <= 0:
                    st.warning("กรุณากรอกจำนวนที่จะรับเข้า")
                elif not receiver.strip():
                    st.warning("กรุณากรอกชื่อผู้รับเข้า")
                else:
                    stock_df = STOCK.copy()
                    in_df = INS.copy()
                    srow = stock_df.loc[idx]
                    cur = to_int(srow["qty"], 0)
                    stock_df.at[idx, "qty"] = str(cur + int(in_qty))
                    new_row = {
                        "run": f"IN-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
                        "date": now_str(),
                        "branch_code": sel_branch,
                        "branch_name": code_to_name(sel_branch),
                        "receiver": receiver,
                        "item_code": srow["item_code"],
                        "item_name": srow["item_name"],
                        "qty": str(int(in_qty)),
                        "unit": srow.get("unit",""),
                        "note": note,
                        "ref_out_run": "",
                    }
                    in_df = pd.concat([in_df, pd.DataFrame([new_row])], ignore_index=True)
                    st.session_state["tables"] = (stock_df, OUTS, in_df, USERS, CATS)
                    write_table("stock", stock_df)
                    write_table("in", in_df)
                    st.success("รับเข้าเรียบร้อย ✅")

def page_reports():
    section_title("รายงาน (PDF ไทย + โลโก้)", "🧷")
    st.info("ออกรายงานเป็นไฟล์ PDF รองรับภาษาไทย (ควรติดตั้งฟอนต์ TH Sarabun หรือ Noto Sans Thai ไว้ใน ./fonts)")

    report_type = st.selectbox("ประเภท", ["ภาพรวมสต็อก", "ประวัติการเบิก (ล่าสุด)", "ประวัติการรับเข้า (ล่าสุด)"])
    limit = st.number_input("จำนวนแถวสูงสุด", min_value=10, max_value=1000, value=200, step=10)

    # เลือกโลโก้ (ถ้ามีใน Settings จะดึงอัตโนมัติ)
    logo_path = CFG.get("logo_path", "")
    uploaded_logo = st.file_uploader("อัปโหลดโลโก้ (PNG/JPG) — ใช้เฉพาะรายงานนี้", type=["png","jpg","jpeg"])
    if uploaded_logo is not None:
        tmp_logo = os.path.join(DEFAULT_ASSETS_DIR, f"logo_tmp_{uuid.uuid4().hex[:6]}.png")
        with open(tmp_logo, "wb") as f:
            f.write(uploaded_logo.read())
        logo_path = tmp_logo

    df = pd.DataFrame()
    title = "รายงาน"

    if report_type == "ภาพรวมสต็อก":
        df = STOCK.copy().head(limit)
        title = f"รายงานภาพรวมสต็อก ({len(STOCK):,} รายการ)"
    elif report_type == "ประวัติการเบิก (ล่าสุด)":
        df = OUTS.copy().sort_values("date", ascending=False).head(limit)
        title = f"รายงานประวัติการเบิก (ล่าสุด {len(df):,} รายการ)"
    elif report_type == "ประวัติการรับเข้า (ล่าสุด)":
        df = INS.copy().sort_values("date", ascending=False).head(limit)
        title = f"รายงานประวัติการรับเข้า (ล่าสุด {len(df):,} รายการ)"

    st.dataframe(df, use_container_width=True, height=360)

    if st.button("🖨️ พิมพ์รายงานเป็น PDF", type="primary"):
        if not FONTS_READY:
            st.warning("ยังไม่พบฟอนต์ไทย (TH Sarabun / Noto Sans Thai) ในระบบ/โฟลเดอร์ ./fonts — PDF อาจแสดงเป็นสี่เหลี่ยม ❗")

        pdf_bytes = generate_pdf_report(title, df, logo_path=logo_path)
        st.download_button(
            "⬇️ ดาวน์โหลด PDF",
            data=pdf_bytes,
            file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
        )

def page_import():
    section_title("นำเข้า/แก้ไข ข้อมูล (หลายแท็บ)", "📥")
    st.caption("อัปโหลดไฟล์ Excel/CSV เพื่ออัปเดตแท็บต่างๆ เช่น stock, out, in, users, categories")
    st.write("**ชื่อแท็บ/ชื่อไฟล์ที่รองรับ**: `stock`, `out`, `in`, `users`, `categories`")

    tfile = st.file_uploader("อัปโหลดไฟล์", type=["xlsx","xls","csv"])
    if tfile is not None:
        ext = pathlib.Path(tfile.name).suffix.lower()
        if ext == ".csv":
            # เลือกชื่อแท็บเอง
            name = st.selectbox("เลือกว่าเป็นแท็บใด", ["stock","out","in","users","categories"])
            df = pd.read_csv(tfile).fillna("")
            st.dataframe(df.head(50), use_container_width=True)
            if st.button(f"🔄 นำเข้าแท็บ {name}"):
                # ensure schema
                schema_map = {"stock":SCHEMA_STOCK, "out":SCHEMA_OUT, "in":SCHEMA_IN, "users":SCHEMA_USERS, "categories":SCHEMA_CATEGORIES}
                df2 = ensure_columns(df, schema_map[name])[schema_map[name]]
                # save
                if name == "stock":
                    st.session_state["tables"] = (df2, OUTS, INS, USERS, CATS)
                elif name == "out":
                    st.session_state["tables"] = (STOCK, df2, INS, USERS, CATS)
                elif name == "in":
                    st.session_state["tables"] = (STOCK, OUTS, df2, USERS, CATS)
                elif name == "users":
                    st.session_state["tables"] = (STOCK, OUTS, INS, df2, CATS)
                elif name == "categories":
                    st.session_state["tables"] = (STOCK, OUTS, INS, USERS, df2)
                write_table(name, df2)
                st.success(f"นำเข้าแท็บ {name} เรียบร้อย ✅")
        else:
            # Excel หลายชีท
            xls = pd.ExcelFile(tfile)
            for name in xls.sheet_names:
                if name.lower() not in ["stock","out","in","users","categories"]:
                    continue
                df = pd.read_excel(xls, sheet_name=name).fillna("")
                st.write(f"**แสดงตัวอย่างแท็บ:** `{name}`")
                st.dataframe(df.head(30), use_container_width=True)
                if st.button(f"🔄 นำเข้าแท็บ {name}", key=f"imp_{name}"):
                    schema_map = {"stock":SCHEMA_STOCK, "out":SCHEMA_OUT, "in":SCHEMA_IN, "users":SCHEMA_USERS, "categories":SCHEMA_CATEGORIES}
                    df2 = ensure_columns(df, schema_map[name])[schema_map[name]]
                    # save
                    if name == "stock":
                        st.session_state["tables"] = (df2, OUTS, INS, USERS, CATS)
                    elif name == "out":
                        st.session_state["tables"] = (STOCK, df2, INS, USERS, CATS)
                    elif name == "in":
                        st.session_state["tables"] = (STOCK, OUTS, df2, USERS, CATS)
                    elif name == "users":
                        st.session_state["tables"] = (STOCK, OUTS, INS, df2, CATS)
                    elif name == "categories":
                        st.session_state["tables"] = (STOCK, OUTS, INS, USERS, df2)
                    write_table(name, df2)
                    st.success(f"นำเข้าแท็บ {name} เรียบร้อย ✅")

def page_users():
    section_title("ผู้ใช้ (Users)", "👥")
    st.caption("ตารางผู้ใช้เรียบง่าย (ชื่อ-บทบาท-สาขา)")
    editable = df_editor(USERS, key="user_editor")
    if st.button("💾 บันทึกผู้ใช้"):
        st.session_state["tables"] = (STOCK, OUTS, INS, editable, CATS)
        write_table("users", editable)
        st.success("บันทึกเรียบร้อย ✅")

def page_settings():
    section_title("Settings", "⚙️")
    tabs = st.tabs(["Google Sheets", "Fonts/PDF", "Logo", "Branches", "Tools"])

    with tabs[0]:
        st.checkbox("ใช้ Google Sheets แทน CSV (ต้องตั้งค่าให้ครบ)", key="use_gs_tmp", value=CFG["use_gsheets"])
        sheet_url = st.text_input("Sheet URL", value=CFG.get("sheet_url",""))
        st.caption("รูปแบบ URL: https://docs.google.com/spreadsheets/d/xxxxxxx/edit#gid=0")

        up_json = st.file_uploader("อัปโหลด Service Account JSON (ทางเลือกที่ 1)", type=["json"], key="sa_json")
        if up_json is not None:
            CFG["service_account_json_text"] = up_json.read().decode("utf-8")
            st.success("โหลด JSON (in-memory) สำเร็จ")

        json_file_path = st.text_input("หรือ ใส่ path ไฟล์ JSON บนเครื่องเซิร์ฟเวอร์ (ทางเลือกที่ 2)", value=CFG.get("service_account_json_file",""))

        if st.button("💾 บันทึกการตั้งค่า GSheets"):
            CFG["use_gsheets"] = bool(st.session_state.get("use_gs_tmp", False))
            CFG["sheet_url"] = sheet_url.strip()
            CFG["service_account_json_file"] = json_file_path.strip()
            st.success("บันทึกแล้ว ✅")

        if not GS_AVAILABLE:
            st.warning("ยังไม่ได้ติดตั้งไลบรารี gspread / google-auth — โปรดติดตั้งจาก requirements.txt")

    with tabs[1]:
        st.caption("ตั้งค่าโฟนต์ไทย (PDF) — หากปล่อยว่าง ระบบจะพยายามค้นหาให้อัตโนมัติ")
        f_reg = st.text_input("Regular TTF path", value=CFG.get("pdf_font_regular",""))
        f_bold = st.text_input("Bold TTF path", value=CFG.get("pdf_font_bold",""))
        if st.button("💾 บันทึกฟอนต์ PDF"):
            CFG["pdf_font_regular"] = f_reg.strip()
            CFG["pdf_font_bold"] = f_bold.strip()
            st.success("บันทึกแล้ว ✅")

        if FONTS_READY:
            st.success("พบฟอนต์ไทยและพร้อมใช้งานสำหรับ PDF ✅")
        else:
            st.warning("ยังไม่พบฟอนต์ไทยสำหรับ PDF — ให้วางไฟล์ .ttf ในโฟลเดอร์ ./fonts (เช่น NotoSansThai-Regular.ttf)")

    with tabs[2]:
        st.caption("กำหนดโลโก้สำหรับรายงาน PDF และส่วนหัวต่างๆ")
        lp = st.text_input("Logo Path (เช่น ./assets/logo.png)", value=CFG.get("logo_path",""))
        up = st.file_uploader("หรืออัปโหลดไฟล์โลโก้", type=["png","jpg","jpeg"])
        if up is not None:
            path = os.path.join(DEFAULT_ASSETS_DIR, f"logo_{uuid.uuid4().hex[:6]}.png")
            with open(path, "wb") as f:
                f.write(up.read())
            lp = path
        if st.button("💾 บันทึกโลโก้"):
            CFG["logo_path"] = lp.strip()
            st.success("บันทึกแล้ว ✅")

    with tabs[3]:
        st.caption("กำหนดรายชื่อสาขา (โค้ด → ชื่อ)")
        # แปลงเป็น DataFrame เพื่อแก้ไขสะดวก
        bm = CFG["branch_code_name"]
        df = pd.DataFrame([{"branch_code":k, "branch_name":v} for k,v in bm.items()])
        df2 = st.data_editor(df, hide_index=True, num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกสาขา"):
            CFG["branch_code_name"] = {r["branch_code"]: r["branch_name"] for _, r in df2.iterrows() if r["branch_code"]}
            st.success("บันทึกแล้ว ✅")

    with tabs[4]:
        st.caption("เครื่องมือผู้ดูแล")
        if st.button("🧹 ล้างข้อมูลทดลอง (เฉพาะ CSV)"):
            for name in ["stock","out","in","users","categories"]:
                path = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
                if os.path.exists(path):
                    os.remove(path)
            st.session_state["tables"] = load_all_tables()
            st.success("ล้างข้อมูล CSV เรียบร้อย")

# =======================================================
# Navigation
# =======================================================
PAGES = {
    "Dashboard": page_dashboard,
    "คลังอุปกรณ์": page_stock,
    "เบิก/รับ": page_out_in,
    "รายงาน": page_reports,
    "นำเข้า": page_import,
    "ผู้ใช้": page_users,
    "Settings": page_settings,
}

# Sidebar
with st.sidebar:
    st.markdown(f"### 🧰 IT Stock {APP_VERSION}")
    if CFG.get("logo_path") and os.path.exists(CFG["logo_path"]):
        st.image(CFG["logo_path"], use_container_width=True)
    choice = st.radio("เมนู", list(PAGES.keys()), index=0)
    st.caption("Tip: ใช้ Google Sheets ได้ใน Settings → Google Sheets")

# Run Page
PAGES[choice]()
