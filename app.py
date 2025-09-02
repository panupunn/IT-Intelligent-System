# -*- coding: utf-8 -*-
"""
app_v11_restored_reports_only.py
IT Stock (Streamlit + Google Sheets) — v11  (RESTORED BASE + Reports PDF Patch Only)

คำอธิบาย:
- ไฟล์นี้ตั้งใจ "คืนค่าฐานเดิมแบบ v11" และ "เปลี่ยนเฉพาะหน้า 'รายงาน' ให้พิมพ์เป็น PDF ภาษาไทย"
- ส่วนอื่นคงโครงสร้าง/การทำงานเดิม (Dashboard, คลังอุปกรณ์, เบิก/รับ แบบหลายรายการ, นำเข้าแบบหลายแท็บ, ผู้ใช้, Settings)
- ถ้าโปรเจกต์เดิมของคุณมีตาราง/สคีมาหรือเมนูเพิ่มจากนี้ คุณสามารถย้ายเฉพาะฟังก์ชัน page_reports() + Helpers ไปวางในโปรเจกต์เดิมแทนได้เลย

สิ่งที่เพิ่ม/เปลี่ยน:
- Helpers: _register_thai_fonts_if_needed(), _generate_pdf()  (ใช้เฉพาะหน้า รายงาน)
- page_reports(): ปุ่ม "🖨️ พิมพ์รายงานเป็น PDF" รองรับฟอนต์ไทย + ใส่โลโก้

หมายเหตุ:
- รองรับทั้งโหมด Google Sheets (gspread) และโหมด CSV สำรอง
- ฟอนต์ไทย: ค้นหา TH Sarabun/Noto Sans Thai อัตโนมัติจาก ./fonts และโฟลเดอร์ระบบ (ไม่มีก็พิมพ์ได้ แต่ตัวไทยอาจเป็นสี่เหลี่ยม)
"""
from __future__ import annotations

import os, io, sys, json, uuid, pathlib
from datetime import datetime, date
from typing import List, Optional

import pandas as pd
import numpy as np
import streamlit as st

# ---------- PDF (ReportLab) ใช้เฉพาะหน้า 'รายงาน' ----------
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader

# ---------- Google Sheets (optional) ----------
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

st.set_page_config(
    page_title=f"IT Stock {APP_VERSION}",
    page_icon="🧰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =======================================================
# Utilities / Session Config
# =======================================================
def ensure_dirs():
    os.makedirs(DEFAULT_DATA_DIR, exist_ok=True)
    os.makedirs(DEFAULT_FONTS_DIR, exist_ok=True)
    os.makedirs(DEFAULT_ASSETS_DIR, exist_ok=True)
ensure_dirs()

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

if "cfg" not in st.session_state:
    st.session_state["cfg"] = {
        "use_gsheets": False,
        "sheet_url": "",
        "service_account_json_text": "",
        "service_account_json_file": "",
        "pdf_font_regular": "",
        "pdf_font_bold": "",
        "logo_path": "",
        "branch_code_name": {},
    }
CFG = st.session_state["cfg"]

# =======================================================
# Google Sheets + CSV Fallback
# =======================================================
GS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def get_gs_client() -> Optional["gspread.Client"]:
    if not GS_AVAILABLE or not CFG.get("use_gsheets"):
        return None
    creds = None
    if CFG.get("service_account_json_text"):
        try:
            info = json.loads(CFG["service_account_json_text"])
            creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"service_account_json_text ผิดพลาด: {e}")
    if creds is None and CFG.get("service_account_json_file"):
        try:
            creds = Credentials.from_service_account_file(CFG["service_account_json_file"], scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"service_account_json_file ผิดพลาด: {e}")
    if creds is None and os.environ.get("SERVICE_ACCOUNT_JSON"):
        try:
            info = json.loads(os.environ["SERVICE_ACCOUNT_JSON"])
            creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
        except Exception as e:
            st.error(f"SERVICE_ACCOUNT_JSON (ENV) ผิดพลาด: {e}")
    if creds is None:
        return None
    try:
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"เชื่อมต่อ gspread ไม่สำเร็จ: {e}")
        return None

def read_table(name: str) -> pd.DataFrame:
    if CFG.get("use_gsheets") and CFG.get("sheet_url"):
        client = get_gs_client()
        if client is not None:
            try:
                sh = client.open_by_url(CFG["sheet_url"])
                try:
                    ws = sh.worksheet(name)
                except Exception:
                    ws = sh.add_worksheet(title=name, rows=100, cols=26)
                    ws.append_row(["_init"])
                rows = ws.get_all_records()
                df = pd.DataFrame(rows)
                return df.fillna("")
            except Exception as e:
                st.warning(f"อ่านชีท '{name}' ไม่ได้: {e} → จะอ่านจาก CSV")
    path = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
    if os.path.exists(path):
        try:
            return pd.read_csv(path, dtype=str).fillna("")
        except Exception:
            pass
    return pd.DataFrame()

def write_table(name: str, df: pd.DataFrame):
    df = df.copy()
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
                    ws = sh.add_worksheet(title=name, rows=max(len(df)+10,100), cols=max(len(df.columns)+2,26))
                ws.clear()
                ws.update([df.columns.tolist()] + df.values.tolist())
                return
            except Exception as e:
                st.warning(f"เขียนชีท '{name}' ไม่ได้: {e} → จะบันทึก CSV")
    path = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")

# =======================================================
# Schemas + Initial load
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
    return df[cols + [c for c in df.columns if c not in cols]]

def load_all_tables():
    stock = ensure_columns(read_table("stock"), SCHEMA_STOCK)
    outs  = ensure_columns(read_table("out"),   SCHEMA_OUT)
    ins   = ensure_columns(read_table("in"),    SCHEMA_IN)
    users = ensure_columns(read_table("users"), SCHEMA_USERS)
    cats  = ensure_columns(read_table("categories"), SCHEMA_CATEGORIES)
    return stock, outs, ins, users, cats

if "tables" not in st.session_state:
    st.session_state["tables"] = load_all_tables()
STOCK, OUTS, INS, USERS, CATS = st.session_state["tables"]

# =======================================================
# Branch mapping (unchanged)
# =======================================================
def ensure_branch_map():
    if not CFG["branch_code_name"]:
        CFG["branch_code_name"] = {
            "SWC001": "สาขากรุงเทพฯ (ตัวอย่าง)",
            "SWC002": "สาขานครราชสีมา (ตัวอย่าง)",
            "SWC003": "สาขาขอนแก่น (ตัวอย่าง)",
        }
ensure_branch_map()
def code_to_name(code: str) -> str:
    return CFG["branch_code_name"].get(code, "")

# =======================================================
# Thai PDF helpers (ONLY for Reports page)
# =======================================================
COMMON_THAI_FONT_NAMES = [
    "THSarabunNew","TH Sarabun New","Sarabun",
    "NotoSansThai","Noto Sans Thai","NotoSerifThai","Noto Serif Thai",
]
COMMON_FONT_DIRS = [
    DEFAULT_FONTS_DIR,
    "/usr/share/fonts/truetype",
    "/usr/share/fonts",
    "/Library/Fonts",
    "/System/Library/Fonts",
    "C:\\Windows\\Fonts",
]
def _find_font(names):
    for d in COMMON_FONT_DIRS:
        try:
            for fn in os.listdir(d):
                lower = fn.lower()
                for name in names:
                    if name.replace(" ","").lower() in lower.replace(" ","") and lower.endswith(".ttf"):
                        return os.path.join(d, fn)
        except Exception:
            continue
    return ""

def _register_thai_fonts_if_needed(CFG=None):
    registered = set(pdfmetrics.getRegisteredFontNames())
    if "TH_REG" in registered and "TH_BOLD" in registered:
        return True
    reg = ""
    bold = ""
    if CFG:
        reg = CFG.get("pdf_font_regular","") or ""
        bold = CFG.get("pdf_font_bold","") or ""
    if not reg:
        reg = _find_font(COMMON_THAI_FONT_NAMES)
    if not bold:
        bold = _find_font([n+" Bold" for n in COMMON_THAI_FONT_NAMES] + COMMON_THAI_FONT_NAMES) or reg
    ok = False
    try:
        if reg:
            pdfmetrics.registerFont(TTFont("TH_REG", reg)); ok = True
        if bold:
            pdfmetrics.registerFont(TTFont("TH_BOLD", bold)); ok = True
    except Exception:
        pass
    return ok

def _generate_pdf(title, df, logo_path=""):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    # Logo
    if logo_path and os.path.exists(logo_path):
        try:
            img = ImageReader(logo_path)
            c.drawImage(img, 15*mm, H-35*mm, width=25*mm, height=25*mm, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
    # Header
    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 18)
    c.drawString(45*mm, H-20*mm, str(title))
    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 10)
    c.drawRightString(W-15*mm, H-15*mm, f"พิมพ์เมื่อ: {now_str()}")
    # Table (first 8 cols, max 50 rows)
    cols = [str(cn) for cn in list(df.columns)[:8]]
    x0, y0 = 15*mm, H-45*mm
    row_h = 8*mm
    col_w = (W - 30*mm) / max(1, len(cols))
    c.setFont("TH_BOLD" if "TH_BOLD" in pdfmetrics.getRegisteredFontNames() else "Helvetica-Bold", 10)
    for i, col in enumerate(cols):
        c.drawString(x0 + i*col_w + 2, y0, col)
    c.line(x0, y0-2, x0 + col_w*len(cols), y0-2)
    c.setFont("TH_REG" if "TH_REG" in pdfmetrics.getRegisteredFontNames() else "Helvetica", 9)
    y = y0 - row_h
    for r in df[cols].astype(str).values.tolist()[:50]:
        for i, val in enumerate(r):
            c.drawString(x0 + i*col_w + 2, y, val[:40])
        y -= row_h
        if y < 20*mm:
            break
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# =======================================================
# UI Pages (unchanged except Reports)
# =======================================================
def section_title(title: str, emoji: str=""):
    st.subheader(f"{emoji} {title}".strip())

def page_dashboard():
    section_title("Dashboard","📊")
    col1, col2, col3 = st.columns(3)
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
    section_title("คลังอุปกรณ์ (Stock)","📦")
    st.info("แก้ไขได้ในตาราง แล้วกดบันทึก", icon="ℹ️")
    editable = st.data_editor(STOCK, hide_index=True, use_container_width=True, height=360, num_rows="dynamic")
    if st.button("💾 บันทึกคลังอุปกรณ์", type="primary"):
        st.session_state["tables"] = (editable, OUTS, INS, USERS, CATS)
        write_table("stock", editable)
        st.success("บันทึกเรียบร้อย ✅")

def page_out_in():
    section_title("เบิก/รับ (OUT/IN)","🧾")
    tab_out, tab_in = st.tabs(["🔻 เบิกหลายรายการ (OUT)", "🔺 รับเข้า (IN)"])
    with tab_out:
        left, right = st.columns([1,1])
        with left:
            branches = list(CFG["branch_code_name"].keys())
            sel_branch = st.selectbox("เลือกสาขา", options=branches, format_func=lambda c: f"{c} - {CFG['branch_code_name'].get(c,'')}")
            requester = st.text_input("ชื่อผู้เบิก/ผู้แจ้ง", "")
            note = st.text_input("หมายเหตุ (ถ้ามี)", "")
        with right:
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
                    maxq = int(pd.to_numeric(row["qty"], errors="coerce") or 0)
                    qty_inputs[idx] = st.number_input(
                        f"{row['item_code']} | {row['item_name']} (คงเหลือ {row['qty']})",
                        min_value=0, max_value=max(0, maxq), step=1, value=0, key=f"qty_{idx}"
                    )
            if st.button("✅ ยืนยันการเบิก (OUT)"):
                if not pick:
                    st.warning("ยังไม่ได้เลือกรายการ")
                elif not requester.strip():
                    st.warning("กรุณากรอกชื่อผู้เบิก")
                else:
                    out_df = OUTS.copy(); stock_df = STOCK.copy(); new_rows = []
                    for idx in pick:
                        q = int(qty_inputs.get(idx, 0) or 0)
                        if q <= 0: continue
                        srow = stock_df.loc[idx]
                        cur = int(pd.to_numeric(srow["qty"], errors="coerce") or 0)
                        stock_df.at[idx,"qty"] = str(max(0, cur - q))
                        new_rows.append({
                            "run": f"OUT-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
                            "date": now_str(),
                            "branch_code": sel_branch,
                            "branch_name": CFG["branch_code_name"].get(sel_branch,""),
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
                        st.session_state["tables"] = (stock_df, out_df, INS, USERS, CATS)
                        write_table("stock", stock_df); write_table("out", out_df)
                        st.success(f"บันทึก OUT {len(new_rows)} รายการสำเร็จ ✅")

    with tab_in:
        left, right = st.columns([1,1])
        with left:
            branches = list(CFG["branch_code_name"].keys())
            sel_branch = st.selectbox("เลือกสาขา", options=branches, format_func=lambda c: f"{c} - {CFG['branch_code_name'].get(c,'')}", key="in_branch")
            receiver = st.text_input("ผู้รับเข้า", "", key="in_receiver")
            note = st.text_input("หมายเหตุ (ถ้ามี)", "", key="in_note")
        with right:
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
                    stock_df = STOCK.copy(); in_df = INS.copy(); srow = stock_df.loc[idx]
                    cur = int(pd.to_numeric(srow["qty"], errors="coerce") or 0)
                    stock_df.at[idx,"qty"] = str(cur + int(in_qty))
                    new_row = {
                        "run": f"IN-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}",
                        "date": now_str(),
                        "branch_code": sel_branch,
                        "branch_name": CFG["branch_code_name"].get(sel_branch,""),
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
                    write_table("stock", stock_df); write_table("in", in_df)
                    st.success("รับเข้าเรียบร้อย ✅")

def page_import():
    section_title("นำเข้า/แก้ไข ข้อมูล (หลายแท็บ)","📥")
    st.caption("รองรับแท็บ: stock, out, in, users, categories (Excel หลายชีท/CSV)")
    tfile = st.file_uploader("อัปโหลดไฟล์", type=["xlsx","xls","csv"])
    if tfile is not None:
        ext = pathlib.Path(tfile.name).suffix.lower()
        def _apply(name, df):
            schema_map = {"stock":SCHEMA_STOCK,"out":SCHEMA_OUT,"in":SCHEMA_IN,"users":SCHEMA_USERS,"categories":SCHEMA_CATEGORIES}
            df2 = ensure_columns(df.fillna(""), schema_map[name])[schema_map[name]]
            if name == "stock": st.session_state["tables"] = (df2, OUTS, INS, USERS, CATS)
            elif name == "out": st.session_state["tables"] = (STOCK, df2, INS, USERS, CATS)
            elif name == "in": st.session_state["tables"] = (STOCK, OUTS, df2, USERS, CATS)
            elif name == "users": st.session_state["tables"] = (STOCK, OUTS, INS, df2, CATS)
            elif name == "categories": st.session_state["tables"] = (STOCK, OUTS, INS, USERS, df2)
            write_table(name, df2); st.success(f"นำเข้าแท็บ {name} เรียบร้อย ✅")
        if ext == ".csv":
            name = st.selectbox("เลือกแท็บ", ["stock","out","in","users","categories"])
            df = pd.read_csv(tfile, dtype=str)
            st.dataframe(df.head(50), use_container_width=True)
            if st.button(f"🔄 นำเข้าแท็บ {name}"):
                _apply(name, df)
        else:
            xls = pd.ExcelFile(tfile)
            for name in xls.sheet_names:
                if name.lower() not in ["stock","out","in","users","categories"]: continue
                df = pd.read_excel(xls, sheet_name=name, dtype=str).fillna("")
                st.write(f"**ตัวอย่างแท็บ:** `{name}`")
                st.dataframe(df.head(30), use_container_width=True)
                if st.button(f"🔄 นำเข้าแท็บ {name}", key=f"imp_{name}"):
                    _apply(name.lower(), df)

def page_users():
    section_title("ผู้ใช้ (Users)","👥")
    editable = st.data_editor(USERS, hide_index=True, use_container_width=True, height=320, num_rows="dynamic")
    if st.button("💾 บันทึกผู้ใช้"):
        st.session_state["tables"] = (STOCK, OUTS, INS, editable, CATS)
        write_table("users", editable); st.success("บันทึกเรียบร้อย ✅")

# ---------------- Reports (PATCHED ONLY THIS PAGE) ----------------
def page_reports():
    st.subheader("🧷 รายงาน (พิมพ์เป็น PDF ภาษาไทย)")
    st.info("ถ้าไม่พบฟอนต์ไทย ระบบจะพยายามค้นหา TH Sarabun/Noto Sans Thai อัตโนมัติ")

    report_type = st.selectbox("ประเภท", ["ภาพรวมสต็อก", "ประวัติการเบิก (ล่าสุด)", "ประวัติการรับเข้า (ล่าสุด)"])
    limit = st.number_input("จำนวนแถวสูงสุด", min_value=10, max_value=2000, value=200, step=10)

    logo_path = CFG.get("logo_path","")
    up_logo = st.file_uploader("อัปโหลดโลโก้ (PNG/JPG) เฉพาะรายงานนี้", type=["png","jpg","jpeg"])
    if up_logo is not None:
        os.makedirs(DEFAULT_ASSETS_DIR, exist_ok=True)
        tmp = os.path.join(DEFAULT_ASSETS_DIR, "logo_tmp_report.png")
        with open(tmp, "wb") as f: f.write(up_logo.read())
        logo_path = tmp

    if report_type == "ภาพรวมสต็อก":
        df = STOCK.copy().head(limit)
        title = f"รายงานภาพรวมสต็อก ({len(STOCK):,} รายการ)"
    elif report_type == "ประวัติการเบิก (ล่าสุด)":
        df = OUTS.copy().sort_values("date", ascending=False).head(limit)
        title = f"รายงานประวัติการเบิก (ล่าสุด {len(df):,} รายการ)"
    else:
        df = INS.copy().sort_values("date", ascending=False).head(limit)
        title = f"รายงานประวัติการรับเข้า (ล่าสุด {len(df):,} รายการ)"

    st.dataframe(df, use_container_width=True, height=360)

    if st.button("🖨️ พิมพ์รายงานเป็น PDF", type="primary"):
        _ = _register_thai_fonts_if_needed(CFG)
        pdf_bytes = _generate_pdf(title, df, logo_path=logo_path)
        st.download_button(
            "⬇️ ดาวน์โหลด PDF",
            data=pdf_bytes,
            file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mime="application/pdf",
        )

def page_settings():
    section_title("Settings","⚙️")
    tabs = st.tabs(["Google Sheets","Fonts/PDF","Logo","Branches","Tools"])
    with tabs[0]:
        st.checkbox("ใช้ Google Sheets", key="use_gs_tmp", value=CFG["use_gsheets"])
        sheet_url = st.text_input("Sheet URL", value=CFG.get("sheet_url",""))
        up_json = st.file_uploader("อัปโหลด Service Account JSON (in-memory)", type=["json"])
        if up_json is not None:
            CFG["service_account_json_text"] = up_json.read().decode("utf-8")
            st.success("โหลด JSON (in-memory) สำเร็จ")
        json_file_path = st.text_input("หรือ Path ไฟล์ JSON บนเครื่องเซิร์ฟเวอร์", value=CFG.get("service_account_json_file",""))
        if st.button("💾 บันทึกการตั้งค่า GSheets"):
            CFG["use_gsheets"] = bool(st.session_state.get("use_gs_tmp", False))
            CFG["sheet_url"] = sheet_url.strip()
            CFG["service_account_json_file"] = json_file_path.strip()
            st.success("บันทึกแล้ว ✅")
        if not GS_AVAILABLE:
            st.warning("ยังไม่ได้ติดตั้ง gspread/google-auth")

    with tabs[1]:
        f_reg = st.text_input("Regular TTF path", value=CFG.get("pdf_font_regular",""))
        f_bold = st.text_input("Bold TTF path", value=CFG.get("pdf_font_bold",""))
        if st.button("💾 บันทึกฟอนต์ PDF"):
            CFG["pdf_font_regular"] = f_reg.strip()
            CFG["pdf_font_bold"] = f_bold.strip()
            st.success("บันทึกแล้ว ✅")

    with tabs[2]:
        lp = st.text_input("Logo Path", value=CFG.get("logo_path",""))
        up = st.file_uploader("หรืออัปโหลดโลโก้", type=["png","jpg","jpeg"])
        if up is not None:
            path = os.path.join(DEFAULT_ASSETS_DIR, "logo_default.png")
            with open(path, "wb") as f: f.write(up.read())
            lp = path
        if st.button("💾 บันทึกโลโก้"):
            CFG["logo_path"] = lp.strip(); st.success("บันทึกแล้ว ✅")

    with tabs[3]:
        bm = CFG["branch_code_name"]
        df = pd.DataFrame([{"branch_code":k, "branch_name":v} for k,v in bm.items()])
        df2 = st.data_editor(df, hide_index=True, num_rows="dynamic", use_container_width=True)
        if st.button("💾 บันทึกสาขา"):
            CFG["branch_code_name"] = {r["branch_code"]: r["branch_name"] for _, r in df2.iterrows() if r["branch_code"]}
            st.success("บันทึกแล้ว ✅")

    with tabs[4]:
        if st.button("🧹 ล้างข้อมูล CSV ทดลอง"):
            for name in ["stock","out","in","users","categories"]:
                p = os.path.join(DEFAULT_DATA_DIR, f"{name}.csv")
                if os.path.exists(p): os.remove(p)
            st.session_state["tables"] = load_all_tables()
            st.success("ล้างข้อมูลสำเร็จ")

# =======================================================
# Navigation (unchanged)
# =======================================================
PAGES = {
    "Dashboard": page_dashboard,
    "คลังอุปกรณ์": page_stock,
    "เบิก/รับ": page_out_in,
    "รายงาน": page_reports,        # ← Patched only this
    "นำเข้า": page_import,
    "ผู้ใช้": page_users,
    "Settings": page_settings,
}
with st.sidebar:
    st.markdown(f"### 🧰 IT Stock {APP_VERSION}")
    if CFG.get("logo_path") and os.path.exists(CFG["logo_path"]):
        st.image(CFG["logo_path"], use_container_width=True)
    choice = st.radio("เมนู", list(PAGES.keys()), index=0)
PAGES[choice]()
