# -*- coding: utf-8 -*-
"""
IT Intelligent System (Streamlit + Google Sheets)
v12 - Consolidated improvements
- Responsive UI for mobile, consistent icons in sidebar
- Robust Google Sheets handling (auto-create missing sheets, clearer errors)
- Unified Import UX with preview/confirm
- Categories management embedded into Stock page with delete-protection
- Issue/Receive with negative-stock guard and audit log
- Tickets: basic CRUD (create/update status), filters
- Reports: Low ROP, Transactions by date, CSV export
- Dashboard: small KPIs
- Settings: service account upload, test connection, sample PDF Thai font test
"""

import os, io, re, uuid, base64
from datetime import datetime, timedelta, date, time as dtime
import pytz, pandas as pd, streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import gspread
from google.oauth2.service_account import Credentials
import bcrypt

# ---------- Helpers ----------
def safe_rerun():
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()

APP_TITLE = "IT Intelligent System"
APP_TAGLINE = "Minimal, Modern, and Practical"
CREDENTIALS_FILE = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
CONFIG_FILE = os.environ.get("ITIS_CONFIG_FILE", "app_config.json")
TZ = pytz.timezone("Asia/Bangkok")

# Sheet names & headers
SHEET_ITEMS     = "Items"
SHEET_TXNS      = "Transactions"
SHEET_USERS     = "Users"
SHEET_CATS      = "Categories"
SHEET_BRANCHES  = "Branches"
SHEET_TICKETS   = "Tickets"
SHEET_TICKET_CATS = "TicketCategories"
SHEET_AUDIT = "AuditLog"

ITEMS_HEADERS   = ["รหัส","หมวดหมู่","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
TXNS_HEADERS    = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","ผู้ดำเนินการ","หมายเหตุ"]
USERS_HEADERS   = ["Username","DisplayName","Role","PasswordHash","Active"]
CATS_HEADERS    = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS      = ["รหัสสาขา","ชื่อสาขา"]
TICKETS_HEADERS = ["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]
TICKET_CAT_HEADERS = ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]
AUDIT_HEADERS   = ["เมื่อ","ผู้ใช้","การทำงาน","รายละเอียด"]

# ---------- Minimal CSS + Responsive ----------
MINIMAL_CSS = """
<style>
:root { --radius: 16px; }
section.main > div { padding-top: 8px; }
.block-card { background: #fff; border:1px solid #eee; border-radius:16px; padding:16px; }
.kpi { display:grid; grid-template-columns: repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
.danger { color:#b00020; }
</style>"""
RESPONSIVE_CSS = """
<style>
html, body, [data-testid="stAppViewContainer"] { font-size: 15px; }
@media (max-width: 768px){
  html, body, [data-testid="stAppViewContainer"]{ font-size: 14px; }
  h1{font-size:1.6rem;} h2{font-size:1.35rem;} h3{font-size:1.15rem;}
}
@media (max-width: 480px){
  html, body, [data-testid="stAppViewContainer"]{ font-size: 13px; }
  h1{font-size:1.45rem;} h2{font-size:1.25rem;} h3{font-size:1.1rem;}
}
</style>
"""

# ---------- Google Sheets ----------
@st.cache_resource(show_spinner=False)
def _get_client():
    if not os.path.exists(CREDENTIALS_FILE):
        return None, "ไม่พบไฟล์ service_account.json (อัปโหลดใน Settings)"
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, f"โหลด service account ไม่สำเร็จ: {e}"

@st.cache_resource(show_spinner=False)
def open_sheet_by_url(sheet_url: str):
    client, err = _get_client()
    if err: raise RuntimeError(err)
    return client.open_by_url(sheet_url)

def ensure_sheets_exist(sh):
    """สร้างชีตที่จำเป็นถ้ายังไม่มี + header"""
    titles = [ws.title for ws in sh.worksheets()]
    def _make(name, rows, cols, headers):
        ws = sh.add_worksheet(name, rows, cols)
        ws.append_row(headers)
    if SHEET_ITEMS not in titles: _make(SHEET_ITEMS, 1000, len(ITEMS_HEADERS)+5, ITEMS_HEADERS)
    if SHEET_TXNS not in titles: _make(SHEET_TXNS, 2000, len(TXNS_HEADERS)+5, TXNS_HEADERS)
    if SHEET_USERS not in titles:
        _make(SHEET_USERS, 100, len(USERS_HEADERS)+2, USERS_HEADERS)
        default_pwd = bcrypt.hashpw("admin123".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        sh.worksheet(SHEET_USERS).append_row(["admin","Administrator","admin",default_pwd,"Y"])
    if SHEET_CATS not in titles: _make(SHEET_CATS, 200, len(CATS_HEADERS)+2, CATS_HEADERS)
    if SHEET_BRANCHES not in titles: _make(SHEET_BRANCHES, 200, len(BR_HEADERS)+2, BR_HEADERS)
    if SHEET_TICKETS not in titles: _make(SHEET_TICKETS, 1000, len(TICKETS_HEADERS)+5, TICKETS_HEADERS)
    if SHEET_TICKET_CATS not in titles: _make(SHEET_TICKET_CATS, 200, len(TICKET_CAT_HEADERS)+2, TICKET_CAT_HEADERS)
    if SHEET_AUDIT not in titles: _make(SHEET_AUDIT, 2000, len(AUDIT_HEADERS)+2, AUDIT_HEADERS)

def read_df(sh, title, headers):
    """อ่านข้อมูลจากชีตแบบทนทาน"""
    try:
        ws = sh.worksheet(title)
    except Exception:
        try:
            ensure_sheets_exist(sh)
            ws = sh.worksheet(title)
        except Exception as e2:
            try:
                titles = [w.title for w in sh.worksheets()]
            except Exception:
                titles = []
            st.error("""ไม่สามารถเปิดชีตชื่อ **{}** ได้

- ตรวจสอบว่า URL ของ Google Sheet ถูกต้องและแชร์ให้ service account แล้ว
- ตรวจสอบว่ามีแท็บชื่อ **{}** อยู่จริง (ปัจจุบันพบ: {})
- ถ้าเพิ่งเปลี่ยนสิทธิ์การเข้าถึง ให้กดปุ่มรีเฟรช/ลองใหม่อีกครั้ง

รายละเอียดระบบ: {}""".format(title, title, ", ".join(titles) if titles else "ไม่สามารถอ่านรายชื่อชีตได้", str(e2)), icon="⚠️")
            raise
    vals = ws.get_all_values()
    if not vals: return pd.DataFrame(columns=headers)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    return df if not df.empty else pd.DataFrame(columns=headers)

def write_df(sh, title, df):
    ws = sh.worksheet(title)
    ws.clear()
    ws.append_row(df.columns.tolist())
    if not df.empty:
        ws.append_rows(df.astype(str).values.tolist())

def log_event(sh, user, action, detail):
    df = read_df(sh, SHEET_AUDIT, AUDIT_HEADERS)
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    df = pd.concat([df, pd.DataFrame([[now, user, action, detail]], columns=AUDIT_HEADERS)], ignore_index=True)
    write_df(sh, SHEET_AUDIT, df)

# ---------- Utility ----------
# ---------- Auth & Connection ----------
def require_login():
    if not st.session_state.get("logged_in", False):
        with st.form("login_form"):
            u = st.text_input("Username", value=st.session_state.get("username","admin"))
            p = st.text_input("Password", type="password", value="")
            s = st.form_submit_button("เข้าสู่ระบบ")
        if s:
            # demo: allow anything; set role admin by default
            st.session_state["logged_in"] = True
            st.session_state["username"] = u or "admin"
            st.session_state["role"] = "admin"
            safe_rerun()
        st.stop()

def connect_sheet_if_ready():
    """Connect once and cache in session_state['sh'] after successful test/save in Settings."""
    if st.session_state.get("sh"):  # already connected
        return st.session_state["sh"]
    url = st.session_state.get("sheet_url","")
    if not url:
        return None
    try:
        sh = open_sheet_by_url(url)
        ensure_sheets_exist(sh)
        st.session_state["sh"] = sh
        st.session_state["connected"] = True
        return sh
    except Exception:
        return None


def load_config_into_session():
    """โหลดค่าคอนฟิก (เช่น sheet_url) ใส่ session_state ถ้ายังไม่มี พร้อมเชื่อมต่ออัตโนมัติถ้า URL พร้อม"""
    try:
        if "sheet_url" not in st.session_state and os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            url = cfg.get("sheet_url", "")
            if url:
                st.session_state["sheet_url"] = url
                if "sh" not in st.session_state:
                    try:
                        sh = open_sheet_by_url(url)
                        ensure_sheets_exist(sh)
                        st.session_state["sh"] = sh
                        st.session_state["connected"] = True
                    except Exception:
                        pass
    except Exception:
        pass

def save_config_from_session():
    """บันทึกค่า sheet_url ลงไฟล์คอนฟิก เพื่อให้คงอยู่ข้ามการ rerun/menu"""
    try:
        url = st.session_state.get("sheet_url", "")
        if url:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"sheet_url": url, "connected": True}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def save_config_from_session():
    """บันทึกค่า sheet_url ลงไฟล์คอนฟิก เพื่อให้คงอยู่ข้ามการ rerun/menu"""
    try:
        url = st.session_state.get("sheet_url", "")
        if url:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"sheet_url": url, "connected": True}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def find_thai_font():
    candidates = [
        ("ThaiFont", "./fonts/Sarabun-Regular.ttf", "./fonts/Sarabun-Bold.ttf"),
        ("ThaiFont", "./fonts/THSarabunNew.ttf", "./fonts/THSarabunNew Bold.ttf"),
        ("ThaiFont", "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf"),
        ("ThaiFont", "/usr/share/fonts/truetype/sarabun/Sarabun-Regular.ttf", "/usr/share/fonts/truetype/sarabun/Sarabun-Bold.ttf"),
        ("ThaiFont", "/Library/Fonts/NotoSansThai-Regular.ttf", "/Library/Fonts/NotoSansThai-Bold.ttf"),
    ]
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
    return None

def sample_pdf(use_thai=True):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=14, rightMargin=14, topMargin=14, bottomMargin=14)
    styles = getSampleStyleSheet()
    if use_thai:
        f = find_thai_font()
        if f is None:
            st.warning("ไม่พบฟอนต์ไทยสำหรับ PDF (Sarabun / TH Sarabun New / Noto Sans Thai). โปรดวางไฟล์ .ttf ไว้ในโฟลเดอร์ ./fonts แล้วลองใหม่อีกครั้ง.", icon="⚠️")
        else:
            styles["Normal"].fontName = f["normal"]; styles["Normal"].leading = 14
            styles["Heading1"].fontName = f["normal"]
    story = []
    story.append(Paragraph("ตัวอย่างรายงาน (PDF) — ภาษาไทยทดสอบการแสดงผล", styles["Heading1"]))
    story.append(Spacer(0, 8))
    story.append(Paragraph("ระบบ IT Intelligent System", styles["Normal"]))
    doc.build(story)
    return buf.getvalue()

def get_username():
    return st.session_state.get("username", "admin")

# ---------- Import UX (shared) ----------
def render_import_box(df_upload, required_cols, rename_map=None):
    if rename_map:
        df_upload.columns = [rename_map.get(c.strip(), c.strip()) for c in df_upload.columns]
    df_upload = df_upload.fillna("").applymap(lambda x: str(x).strip())

    missing = [c for c in required_cols if c not in df_upload.columns]
    if missing:
        st.error("ไฟล์ขาดคอลัมน์บังคับ: " + ", ".join(missing), icon="⚠️")
        return None

    st.success(f"พรีวิว {len(df_upload):,} แถว", icon="✅")
    st.dataframe(df_upload.head(100), use_container_width=True, height=260)
    return df_upload

# ---------- Pages ----------

def page_dashboard(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("📊 Dashboard")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    total_items = len(items)
    low_rop = 0
    if not items.empty:
        try:
            low_rop = int((items["คงเหลือ"].astype(float) <= items["จุดสั่งซื้อ"].astype(float)).sum())
        except Exception:
            low_rop = 0

    st.markdown("<div class='kpi'>", unsafe_allow_html=True)
    st.metric("จำนวนอุปกรณ์", f"{total_items:,}")
    st.metric("ต่ำกว่า ROP", f"{low_rop:,}")
    st.metric("Tickets ทั้งหมด", f"{len(tickets):,}")
    # Top branches by OUT (30 days)
    with c3:
        st.markdown("**TOP 5 สาขาที่เบิกมากสุด (30 วัน)**")
        if txns.empty:
            st.info("ยังไม่มีธุรกรรม", icon="ℹ️")
        else:
            try:
                df = txns.copy()
                df["วันเวลา"] = pd.to_datetime(df["วันเวลา"], errors="coerce")
                df = df.dropna(subset=["วันเวลา"])
                cutoff = pd.Timestamp.utcnow().tz_localize("UTC").tz_convert(TZ) - pd.Timedelta(days=30)
                df = df[(df["วันเวลา"] >= cutoff) & (df["ประเภท"]=="OUT")]
                df["จำนวน"] = pd.to_numeric(df["จำนวน"], errors="coerce").fillna(0)
                top = df.groupby("สาขา")["จำนวน"].sum().sort_values(ascending=False).head(5)
                st.bar_chart(top)
            except Exception:
                st.info("ไม่สามารถสร้างกราฟได้จากข้อมูลปัจจุบัน", icon="ℹ️")

    # Top branches by Tickets (30 days)
    with c4:
        st.markdown("**TOP 5 สาขาที่แจ้งซ่อมมากสุด (30 วัน)**")
        if tickets.empty:
            st.info("ยังไม่มี Tickets", icon="ℹ️")
        else:
            try:
                tv = tickets.copy()
                tv["วันที่แจ้ง"] = pd.to_datetime(tv["วันที่แจ้ง"], errors="coerce")
                tv = tv.dropna(subset=["วันที่แจ้ง"])
                cutoff = pd.Timestamp.utcnow().tz_localize("UTC").tz_convert(TZ) - pd.Timedelta(days=30)
                tv = tv[tv["วันที่แจ้ง"] >= cutoff]
                cnt = tv.groupby("สาขา")["TicketID"].count().sort_values(ascending=False).head(5)
                st.bar_chart(cnt)
            except Exception:
                st.info("ไม่สามารถสร้างกราฟได้จากข้อมูลปัจจุบัน", icon="ℹ️")

    st.markdown("</div>", unsafe_allow_html=True)

    # ---------- Charts ----------
    c1, c2 = st.columns(2)

    # Chart 1: ยอดคงเหลือต่อหมวดหมู่
    with c1:
        st.markdown("**คงเหลือรวมต่อหมวดหมู่**")
        chart_type = st.radio("รูปแบบกราฟ", ["แท่ง (Bar)","วงกลม (Pie)"], horizontal=True, key="chart_cat_type")
        if items.empty:
            st.info("ยังไม่มีข้อมูลคงเหลือ", icon="ℹ️")
        else:
            try:
                grp = items.copy()
                grp["คงเหลือ"] = pd.to_numeric(grp["คงเหลือ"], errors="coerce").fillna(0)
                chart_df = grp.groupby("หมวดหมู่")["คงเหลือ"].sum().sort_values(ascending=False).head(10)
                if chart_type.startswith("แท่ง"):
                    st.bar_chart(chart_df)
                else:
                    # Render pie via matplotlib
                    import matplotlib.pyplot as plt
                    fig, ax = plt.subplots()
                    ax.pie(chart_df.values, labels=chart_df.index, autopct='%1.1f%%')
                    ax.axis('equal')
                    st.pyplot(fig)
            except Exception:
                st.info("ไม่สามารถสร้างกราฟได้จากข้อมูลปัจจุบัน", icon="ℹ️")
    # Chart 2: IN/OUT ตามวัน 30 วันล่าสุด
    with c2:
        st.markdown("**ธุรกรรม IN/OUT ย้อนหลัง 30 วัน**")
        if txns.empty:
            st.info("ยังไม่มีธุรกรรม", icon="ℹ️")
        else:
            try:
                df = txns.copy()
                df["วันเวลา"] = pd.to_datetime(df["วันเวลา"], errors="coerce")
                df = df.dropna(subset=["วันเวลา"])
                cutoff = pd.Timestamp.utcnow().tz_localize("UTC").tz_convert(TZ) - pd.Timedelta(days=30)
                df = df[df["วันเวลา"] >= cutoff]
                df["count"] = 1
                pv = df.pivot_table(index=df["วันเวลา"].dt.date, columns="ประเภท", values="count", aggfunc="sum").fillna(0)
                st.line_chart(pv)
            except Exception:
                st.info("ไม่สามารถสร้างกราฟได้จากข้อมูลปัจจุบัน", icon="ℹ️")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

def get_unit_options(items_df): return ["พิมพ์เอง","ชิ้น","กล่อง","ชุด","เครื่อง"]
def get_loc_options(items_df): return ["พิมพ์เอง","คลังกลาง","สาขา1","สาขา2"]

def render_categories_admin(sh):
    st.markdown("#### 🏷️ จัดการหมวดหมู่")
    cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
    tab1, tab2, tab3 = st.tabs(["✏️ เพิ่ม/แก้ไข 1 รายการ", "📥 นำเข้าไฟล์ (หลายรายการ)", "🔎 ค้นหา/แก้ไข/ลบ (ตาราง)"])

    with tab1:
        c1, c2 = st.columns([1,2])
        code_in = c1.text_input("รหัสหมวด", placeholder="เช่น PRT, KBD").upper().strip()
        name_in = c2.text_input("ชื่อหมวด", placeholder="เช่น หมึกพิมพ์, คีย์บอร์ด").strip()
        if st.button("บันทึก/แก้ไข 1 รายการ", use_container_width=True):
            if not code_in or not name_in:
                st.warning("กรุณากรอกรหัสและชื่อหมวดให้ครบ", icon="⚠️")
            else:
                df = read_df(sh, SHEET_CATS, CATS_HEADERS)
                if (df["รหัสหมวด"] == code_in).any():
                    df.loc[df["รหัสหมวด"] == code_in, "ชื่อหมวด"] = name_in; msg="อัปเดต"
                else:
                    df = pd.concat([df, pd.DataFrame([[code_in, name_in]], columns=CATS_HEADERS)], ignore_index=True); msg="เพิ่มใหม่"
                write_df(sh, SHEET_CATS, df); log_event(sh, get_username(), "CAT_SAVE", f"{msg}: {code_in} -> {name_in}")
                st.success(f"{msg}เรียบร้อย", icon="✅"); safe_rerun()

    with tab2:
        with st.expander("วิธีใช้งาน/เทมเพลต", expanded=False):
            st.markdown("""- รองรับไฟล์ .csv หรือ .xlsx (คอลัมน์: **รหัสหมวด, ชื่อหมวด**)
- ถ้ารหัสซ้ำ จะอัปเดตชื่อหมวด
- ถ้าเปิดโหมด 'แทนที่ทั้งชีต' จะตรวจสอบว่าหมวดที่กำลังใช้งานอยู่ใน Items ยังอยู่ในไฟล์ใหม่""")
            tpl = "รหัสหมวด,ชื่อหมวด\nPRT,หมึกพิมพ์\nKBD,คีย์บอร์ด\n"
            st.download_button("ดาวน์โหลดเทมเพลต (CSV)", data=tpl.encode("utf-8-sig"), file_name="template_categories.csv", mime="text/csv")
        cA, cB = st.columns([2,1])
        up = cA.file_uploader("เลือกไฟล์ (.csv, .xlsx)", type=["csv","xlsx"])
        replace_all = cB.checkbox("แทนที่ทั้งชีต (ล้างและใส่ใหม่)", value=False)
        if up is not None:
            try:
                df_up = pd.read_csv(up, dtype=str) if up.name.lower().endswith(".csv") else pd.read_excel(up, dtype=str)
            except Exception as e:
                st.error(f"อ่านไฟล์ไม่สำเร็จ: {e}", icon="❌")
                df_up = None
            if df_up is not None:
                df_up = render_import_box(df_up, ["รหัสหมวด","ชื่อหมวด"],
                    rename_map={"รหัสหมวดหมู่":"รหัสหมวด","ชื่อหมวดหมู่":"ชื่อหมวด","code":"รหัสหมวด","name":"ชื่อหมวด"})
                if df_up is not None and st.button("ดำเนินการนำเข้า/อัปเดต", use_container_width=True):
                    base = read_df(sh, SHEET_CATS, CATS_HEADERS)
                    if replace_all:
                        items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                        used = set(items["หมวดหมู่"].tolist()) if not items.empty else set()
                        newcats = set(df_up["รหัสหมวด"].str.upper().tolist())
                        if used - newcats:
                            st.error("ไม่สามารถแทนที่ทั้งชีตได้: พบหมวดที่ถูกใช้งานใน Items แต่ไม่อยู่ในไฟล์ใหม่นี้", icon="⚠️")
                        else:
                            write_df(sh, SHEET_CATS, df_up[CATS_HEADERS]); log_event(sh, get_username(), "CAT_REPLACE_ALL", f"{len(df_up)} rows")
                            st.success("แทนที่ทั้งชีตสำเร็จ", icon="✅"); safe_rerun()
                    else:
                        added, updated = 0, 0
                        for _, r in df_up.iterrows():
                            cd = str(r["รหัสหมวด"]).strip().upper(); nm = str(r["ชื่อหมวด"]).strip()
                            if not cd or not nm: continue
                            if (base["รหัสหมวด"] == cd).any():
                                base.loc[base["รหัสหมวด"] == cd, "ชื่อหมวด"] = nm; updated += 1
                            else:
                                base = pd.concat([base, pd.DataFrame([[cd, nm]], columns=CATS_HEADERS)], ignore_index=True); added += 1
                        write_df(sh, SHEET_CATS, base); log_event(sh, get_username(), "CAT_IMPORT", f"add={added}, upd={updated}")
                        st.success(f"สำเร็จ • เพิ่ม {added} • อัปเดต {updated}", icon="✅"); safe_rerun()

    with tab3:
        q = st.text_input("ค้นหา (รหัส/ชื่อ)")
        view = cats if not q else cats[cats.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
        edited = st.data_editor(view.sort_values("รหัสหมวด"), use_container_width=True, height=360, disabled=["รหัสหมวด"])
        cL, cM, cR = st.columns(3)
        if cL.button("บันทึกการแก้ไข"):
            base = read_df(sh, SHEET_CATS, CATS_HEADERS)
            for _, r in edited.iterrows():
                base.loc[base["รหัสหมวด"] == str(r["รหัสหมวด"]).strip().upper(), "ชื่อหมวด"] = str(r["ชื่อหมวด"]).strip()
            write_df(sh, SHEET_CATS, base); log_event(sh, get_username(), "CAT_EDIT_TABLE", f"{len(edited)} rows")
            st.success("บันทึกแล้ว", icon="✅"); safe_rerun()
        with cR:
            base = read_df(sh, SHEET_CATS, CATS_HEADERS)
            opts = (base["รหัสหมวด"]+" | "+base["ชื่อหมวด"]).tolist() if not base.empty else []
            picks = st.multiselect("เลือกลบ (เฉพาะหมวดที่ไม่ถูกใช้งาน)", options=opts)
            if st.button("ลบที่เลือก"):
                items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                used = set(items["หมวดหมู่"].tolist()) if not items.empty else set()
                to_del = {x.split(" | ")[0] for x in picks}
                blocked = sorted(list(used.intersection(to_del)))
                if blocked:
                    st.error("ไม่สามารถลบได้: หมวดถูกใช้งานใน Items: " + ", ".join(blocked), icon="⚠️")
                else:
                    base = base[~base["รหัสหมวด"].isin(list(to_del))]
                    write_df(sh, SHEET_CATS, base); log_event(sh, get_username(), "CAT_DELETE", f"{len(to_del)} rows")
                    st.success("ลบแล้ว", icon="✅"); safe_rerun()

def generate_item_code(items_df):
    prefix = "IT"
    if items_df.empty:
        return f"{prefix}0001"
    nums = [int(re.sub(r"\D","", str(x))[0:6] or 0) for x in items_df["รหัส"].tolist()]
    n = max(nums) if nums else 0
    return f"{prefix}{n+1:04d}"

def page_stock(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("📦 คลังอุปกรณ์")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS); cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)
    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view_df = items if not q else items[items.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
    st.dataframe(view_df, use_container_width=True, height=320)

    if st.session_state.get("role","admin") in ("admin","staff"):
        t_add, t_edit, t_cat = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ (เลือกรายการเดิม)","🏷️ หมวดหมู่"])

        with t_add:
            with st.form("item_add", clear_on_submit=True):
                c1,c2,c3 = st.columns(3)
                with c1:
                    if cats.empty: st.info("ยังไม่มีหมวดหมู่ (ไปที่แท็บ '🏷️ หมวดหมู่' เพื่อเพิ่ม)", icon="ℹ️"); cat_opt=""
                    else:
                        opts = (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist(); selected = st.selectbox("หมวดหมู่", options=opts)
                        cat_opt = selected.split(" | ")[0]
                    name = st.text_input("ชื่ออุปกรณ์")
                with c2:
                    unit = st.text_input("หน่วย", value="ชิ้น")
                    qty = st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                    rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
                with c3:
                    loc = st.text_input("ที่เก็บ", value="คลังกลาง")
                    active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                    code = st.text_input("รหัส (เว้นว่างให้ระบบรันอัตโนมัติ)", value="")
                s = st.form_submit_button("บันทึก")
            if s:
                df = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
                code_final = code.strip().upper() or generate_item_code(df)
                new_row = [code_final, cat_opt, name.strip(), unit.strip(), str(qty), str(rop), loc.strip(), active]
                # update if exists else append
                if (df["รหัส"] == code_final).any():
                    df.loc[df["รหัส"] == code_final, ITEMS_HEADERS[1]:] = new_row[1:]
                    msg="อัปเดต"
                else:
                    df = pd.concat([df, pd.DataFrame([new_row], columns=ITEMS_HEADERS)], ignore_index=True); msg="เพิ่มใหม่"
                write_df(sh, SHEET_ITEMS, df); log_event(sh, get_username(), "ITEM_SAVE", f"{msg}: {code_final}")
                st.success(f"{msg}เรียบร้อย", icon="✅"); safe_rerun()

        with t_edit:
            if items.empty: st.info("ยังไม่มีรายการอุปกรณ์ในคลัง", icon="ℹ️")
            else:
                pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
                code_sel = pick.split(" | ")[0]
                row = items[items["รหัส"]==code_sel].iloc[0]
                with st.form("item_edit"):
                    name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
                    unit = st.text_input("หน่วย", value=row["หน่วย"])
                    qty = st.number_input("คงเหลือ", min_value=0, value=int(float(row["คงเหลือ"] or 0)))
                    rop = st.number_input("จุดสั่งซื้อ", min_value=0, value=int(float(row["จุดสั่งซื้อ"] or 0)))
                    loc = st.text_input("ที่เก็บ", value=row["ที่เก็บ"])
                    active = st.selectbox("ใช้งาน", ["Y","N"], index=0 if row["ใช้งาน"]=="Y" else 1)
                    save = st.form_submit_button("บันทึกการแก้ไข")
                if save:
                    items.loc[items["รหัส"]==code_sel, ["ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]] = [name, unit, str(qty), str(rop), loc, "Y" if active=="Y" else "N"]
                    write_df(sh, SHEET_ITEMS, items); log_event(sh, get_username(), "ITEM_UPDATE", code_sel)
                    st.success("บันทึกแล้ว", icon="✅"); safe_rerun()

        with t_cat:
            render_categories_admin(sh)

    st.markdown("</div>", unsafe_allow_html=True)


def page_issue_receive(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("📥 เบิก/รับเข้า")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty: 
        st.info("ยังไม่มีรายการอุปกรณ์ในคลัง", icon="ℹ️"); 
        st.markdown("</div>", unsafe_allow_html=True); 
        return
    if branches.empty:
        st.warning("ยังไม่มีข้อมูลสาขา (ไปที่ชีต Branches เพื่อเพิ่ม)", icon="⚠️")

    t1,t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])

    with t1:
        with st.form("issue", clear_on_submit=True):
            pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
            bopt = st.selectbox("เลือกสาขาที่เบิก", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
            qty = st.number_input("จำนวนที่เบิก", min_value=1, value=1, step=1)
            by = st.text_input("ผู้ดำเนินการ", value=get_username())
            note = st.text_input("หมายเหตุ", value="")
            s = st.form_submit_button("บันทึกการเบิก")
        if s:
            code_sel = pick.split(" | ")[0]
            row = items[items["รหัส"]==code_sel].iloc[0]
            cur = int(float(row["คงเหลือ"] or 0))
            if qty > cur:
                st.error("สต็อกไม่พอสำหรับการเบิก", icon="⚠️")
            else:
                items.loc[items["รหัส"]==code_sel, "คงเหลือ"] = str(cur - qty)
                write_df(sh, SHEET_ITEMS, items)
                txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
                branch_code = bopt.split(" | ")[0] if bopt else ""
                txns = pd.concat([txns, pd.DataFrame([[
                    str(uuid.uuid4())[:8],
                    datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
                    "OUT", code_sel, row["ชื่ออุปกรณ์"], branch_code, str(qty), by, note
                ]], columns=TXNS_HEADERS)], ignore_index=True)
                write_df(sh, SHEET_TXNS, txns); log_event(sh, get_username(), "ISSUE", f"{code_sel} x {qty} @ {branch_code}")
                st.success("บันทึกแล้ว", icon="✅")

    with t2:
        with st.form("receive", clear_on_submit=True):
            pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist(), key="recvpick")
            bopt = st.selectbox("เลือกสาขาที่รับเข้า", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [], key="recvbranch")
            qty = st.number_input("จำนวนที่รับเข้า", min_value=1, value=1, step=1, key="recvqty")
            by = st.text_input("ผู้ดำเนินการ", value=get_username(), key="recvby")
            note = st.text_input("หมายเหตุ", value="", key="recvnote")
            s = st.form_submit_button("บันทึกรับเข้า")
        if s:
            code_sel = pick.split(" | ")[0]
            row = items[items["รหัส"]==code_sel].iloc[0]
            cur = int(float(row["คงเหลือ"] or 0))
            items.loc[items["รหัส"]==code_sel, "คงเหลือ"] = str(cur + qty)
            write_df(sh, SHEET_ITEMS, items)
            txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
            branch_code = bopt.split(" | ")[0] if bopt else ""
            txns = pd.concat([txns, pd.DataFrame([[
                str(uuid.uuid4())[:8],
                datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "IN", code_sel, row["ชื่ออุปกรณ์"], branch_code, str(qty), by, note
            ]], columns=TXNS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_TXNS, txns); log_event(sh, get_username(), "RECEIVE", f"{code_sel} x {qty} @ {branch_code}")
            st.success("บันทึกแล้ว", icon="✅")

    st.markdown("</div>", unsafe_allow_html=True)


def page_tickets(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")
    cats = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    # tabs: สร้าง/รายการ/หมวดหมู่ปัญหา (สำหรับ admin/staff)
    if st.session_state.get("role","admin") in ("admin","staff"):
        tab1, tab2, tab3 = st.tabs(["สร้างคำขอ", "รายการทั้งหมด", "หมวดหมู่ปัญหา"])
    else:
        tab1, tab2 = st.tabs(["สร้างคำขอ", "รายการทั้งหมด"])
        tab3 = None

    with tab1:
        with st.form("tick_new", clear_on_submit=True):
            bopt = st.selectbox("เลือกสาขาที่แจ้ง", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
            cat = st.selectbox("หมวดหมู่ปัญหา", options=(cats["รหัสหมวดปัญหา"]+" | "+cats["ชื่อหมวดปัญหา"]).tolist() if not cats.empty else [])
            who = st.text_input("ผู้แจ้ง", value=get_username())
            detail = st.text_area("รายละเอียด")
            s = st.form_submit_button("สร้าง Ticket")
        if s:
            df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
            tid = "T" + datetime.now(TZ).strftime("%y%m%d%H%M%S")
            now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
            catname = cat.split(" | ")[1] if cat else ""
            row = [tid, now, (bopt.split(" | ")[0] if bopt else ""), who, catname, detail, "เปิด", "", now, ""]
            df = pd.concat([df, pd.DataFrame([row], columns=TICKETS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_TICKETS, df); log_event(sh, get_username(), "TICKET_NEW", tid)
            st.success("สร้าง Ticket แล้ว", icon="✅")

    with tab2:
        st.caption("กรองข้อมูล")
        c1,c2,c3 = st.columns(3)
        status = c1.selectbox("สถานะ", options=["ทั้งหมด","เปิด","กำลังทำ","รออะไหล่","เสร็จ"], index=0)
        who = c2.text_input("ผู้แจ้ง (ค้นหา)")
        q = c3.text_input("คำค้น (รายละเอียด/หมวด)")
        view = tickets.copy()
        if status!="ทั้งหมด": view = view[view["สถานะ"]==status]
        if who: view = view[view["ผู้แจ้ง"].str.contains(who, case=False, na=False)]
        if q: view = view[view.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
        st.dataframe(view, use_container_width=True, height=360)
        if not view.empty:
            with st.expander("อัปเดตสถานะ"):
                sel = st.selectbox("เลือกรายการ", options=(view["TicketID"]+" | "+view["รายละเอียด"].str.slice(0,30)).tolist())
                tid = sel.split(" | ")[0]
                st_new = st.selectbox("สถานะใหม่", options=["เปิด","กำลังทำ","รออะไหล่","เสร็จ"], index=0)
                assignee = st.text_input("ผู้รับผิดชอบ", value="")
                note = st.text_input("หมายเหตุเพิ่มเติม", value="")
                if st.button("บันทึกการเปลี่ยนแปลง"):
                    tickets.loc[tickets["TicketID"]==tid, ["สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]] = [st_new, assignee, datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"), note]
                    write_df(sh, SHEET_TICKETS, tickets); log_event(sh, get_username(), "TICKET_UPDATE", f"{tid} -> {st_new}")
                    st.success("อัปเดตแล้ว", icon="✅"); safe_rerun()

    if tab3 is not None:
        with tab3:
            st.markdown("#### 🗂️ จัดการหมวดหมู่ปัญหา")
            t1, t2 = st.tabs(["เพิ่ม/แก้ไข", "นำเข้า/แก้ไขแบบตาราง"])

            with t1:
                c1, c2 = st.columns([1,2])
                code_in = c1.text_input("รหัสหมวดปัญหา").upper().strip()
                name_in = c2.text_input("ชื่อหมวดปัญหา").strip()
                if st.button("บันทึก/แก้ไข"):
                    if not code_in or not name_in:
                        st.warning("กรุณากรอกให้ครบ", icon="⚠️")
                    else:
                        base = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
                        if (base["รหัสหมวดปัญหา"] == code_in).any():
                            base.loc[base["รหัสหมวดปัญหา"] == code_in, "ชื่อหมวดปัญหา"] = name_in; msg="อัปเดต"
                        else:
                            base = pd.concat([base, pd.DataFrame([[code_in, name_in]], columns=TICKET_CAT_HEADERS)], ignore_index=True); msg="เพิ่มใหม่"
                        write_df(sh, SHEET_TICKET_CATS, base); log_event(sh, get_username(), "TICKET_CAT_SAVE", f"{msg}: {code_in}")
                        st.success(f"{msg}เรียบร้อย", icon="✅"); safe_rerun()

            with t2:
                q = st.text_input("ค้นหา (รหัส/ชื่อ)", key="tkcat_search")
                view = cats if not q else cats[cats.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
                edited = st.data_editor(view.sort_values("รหัสหมวดปัญหา"), use_container_width=True, height=360, disabled=["รหัสหมวดปัญหา"])
                if st.button("บันทึกการแก้ไข (หมวดปัญหา)"):
                    base = read_df(sh, SHEET_TICKET_CATS, TICKET_CAT_HEADERS)
                    for _, r in edited.iterrows():
                        base.loc[base["รหัสหมวดปัญหา"] == str(r["รหัสหมวดปัญหา"]).strip().upper(), "ชื่อหมวดปัญหา"] = str(r["ชื่อหมวดปัญหา"]).strip()
                    write_df(sh, SHEET_TICKET_CATS, base); log_event(sh, get_username(), "TICKET_CAT_EDIT_TABLE", f"{len(edited)} rows")
                    st.success("บันทึกแล้ว", icon="✅"); safe_rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def page_reports(sh):
    st.markdown("<div class='block-card'>", unsafe_allow_html=True); st.subheader("📑 รายงาน")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    st.markdown("### รายงานสินค้าต่ำกว่า ROP")
    low = pd.DataFrame(columns=ITEMS_HEADERS)
    if not items.empty:
        try:
            mask = items["คงเหลือ"].astype(float) <= items["จุดสั่งซื้อ"].astype(float)
            low = items[mask]
        except Exception:
            low = pd.DataFrame(columns=ITEMS_HEADERS)
    st.dataframe(low, use_container_width=True, height=240)
    if not low.empty:
        st.download_button("ดาวน์โหลด CSV (ต่ำกว่า ROP)", data=low.to_csv(index=False).encode("utf-8-sig"), file_name="low_rop.csv", mime="text/csv")

    st.markdown("### รายงานธุรกรรมตามช่วงเวลา")
    c1,c2 = st.columns(2)
    since = c1.date_input("ตั้งแต่", value=date.today()-timedelta(days=30))
    until = c2.date_input("ถึง", value=date.today())

    view = txns.copy()
    if not view.empty:
        try:
            _dtc = pd.to_datetime(view["วันเวลา"])
            view = view[( _dtc.dt.date >= since ) & ( _dtc.dt.date <= until )]
        except Exception:
            pass
    st.dataframe(view, use_container_width=True, height=260)
    if not view.empty:
        st.download_button("ดาวน์โหลด CSV (ธุรกรรม)", data=view.to_csv(index=False).encode("utf-8-sig"), file_name="transactions.csv", mime="text/csv")

    # --- New: Pivot การเบิกตามสาขา/อุปกรณ์ ---
    st.markdown("### สรุปการเบิกตามสาขาและอุปกรณ์ (ช่วงเวลาที่เลือก)")
    out = view[view["ประเภท"]=="OUT"].copy() if not view.empty else pd.DataFrame(columns=TXNS_HEADERS)
    if not out.empty:
        out["จำนวน"] = pd.to_numeric(out["จำนวน"], errors="coerce").fillna(0)
        pvt = out.pivot_table(index="สาขา", columns="ชื่ออุปกรณ์", values="จำนวน", aggfunc="sum", fill_value=0)
        st.dataframe(pvt, use_container_width=True, height=240)
        st.markdown("**กราฟรวมต่อสาขา (OUT)**")
        try:
            st.bar_chart(pvt.sum(axis=1))
        except Exception:
            pass
    else:
        st.info("ยังไม่มีข้อมูล OUT ในช่วงเวลาที่เลือก", icon="ℹ️")

    # --- New: Tickets by Branch/Category ---
    st.markdown("### สรุป Tickets แยกตามสาขาและหมวดหมู่ปัญหา (ช่วงเวลาที่เลือก)")
    tv = tickets.copy()
    if not tv.empty:
        try:
            tv["วันที่แจ้ง"] = pd.to_datetime(tv["วันที่แจ้ง"], errors="coerce")
            tv = tv.dropna(subset=["วันที่แจ้ง"])
            tv = tv[(tv["วันที่แจ้ง"].dt.date >= since) & (tv["วันที่แจ้ง"].dt.date <= until)]
        except Exception:
            tv = pd.DataFrame(columns=TICKETS_HEADERS)
    if not tv.empty:
        pvt2 = tv.pivot_table(index="สาขา", columns="หมวดหมู่", values="TicketID", aggfunc="count", fill_value=0)
        st.dataframe(pvt2, use_container_width=True, height=240)
        st.markdown("**กราฟสัดส่วน Tickets ต่อสาขา**")
        try:
            import matplotlib.pyplot as plt
            totals = pvt2.sum(axis=1).sort_values(ascending=False)
            fig, ax = plt.subplots()
            ax.pie(totals.values, labels=totals.index, autopct='%1.1f%%')
            ax.axis('equal')
            st.pyplot(fig)
        except Exception:
            pass
    else:
        st.info("ยังไม่มี Tickets ในช่วงเวลาที่เลือก", icon="ℹ️")

    st.markdown("</div>", unsafe_allow_html=True)

def ensure_credentials_ui():
    if os.path.exists(CREDENTIALS_FILE): return True
    st.warning("ยังไม่พบไฟล์ service_account.json", icon="⚠️")
    up = st.file_uploader("อัปโหลดไฟล์ service_account.json", type=["json"])
    if up is not None:
        with open(CREDENTIALS_FILE, "wb") as f: f.write(up.getbuffer())
        st.success("บันทึกไฟล์แล้ว", icon="✅"); safe_rerun()
    return False

def test_sheet_connection(url):
    try:
        sh = open_sheet_by_url(url); ensure_sheets_exist(sh)
        titles = [ws.title for ws in sh.worksheets()]
        return True, titles
    except Exception as e:
        return False, str(e)

def page_settings():
    st.subheader("⚙️ Settings")
    ok = ensure_credentials_ui()
    st.text_input("Google Sheet URL", key="sheet_url", value=st.session_state.get("sheet_url",""))
    if st.button("บันทึก URL"):
        st.success("บันทึก URL แล้ว", icon="✅")
    c1,c2,c3 = st.columns(3)
    if c1.button("ทดสอบการเชื่อมต่อ"):
        url = st.session_state.get("sheet_url","")
        if not url:
            st.error("กรุณาใส่ Google Sheet URL ก่อน", icon="⚠️")
        else:
            ok, info = test_sheet_connection(url)
            if ok:
                st.success("เชื่อมต่อได้ และตรวจสอบ/สร้างชีตที่จำเป็นแล้ว: " + ", ".join(info), icon="✅")
                st.session_state["connected"]=True
                try:
                    st.session_state["sh"]=open_sheet_by_url(st.session_state.get("sheet_url",""))
                except Exception:
                    pass
                save_config_from_session()
            else:
                st.error("เชื่อมต่อไม่สำเร็จ: " + str(info), icon="❌")
    if c2.button("สร้าง PDF ทดสอบฟอนต์ไทย"):
        data = sample_pdf(True)
        st.download_button("ดาวน์โหลด PDF", data=data, file_name="sample_thai.pdf", mime="application/pdf")
    if c3.button("ล้างแคชการเชื่อมต่อ"):
        _get_client.clear(); open_sheet_by_url.clear(); st.success("ล้างแคชแล้ว", icon="✅")

def page_users_admin(sh):
    st.subheader("👥 ผู้ใช้")
    users = read_df(sh, SHEET_USERS, USERS_HEADERS)
    st.dataframe(users, use_container_width=True, height=260)
    with st.form("user_add", clear_on_submit=True):
        u = st.text_input("Username"); d = st.text_input("Display Name"); r = st.selectbox("Role", ["admin","staff","viewer"]); p = st.text_input("รหัสผ่าน (จะถูกแฮช)")
        s = st.form_submit_button("เพิ่มผู้ใช้")
    if s:
        if not u or not p:
            st.warning("กรอก Username/Password ให้ครบ", icon="⚠️")
        else:
            pwd = bcrypt.hashpw(p.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            row = [u, d, r, pwd, "Y"]
            users = pd.concat([users, pd.DataFrame([row], columns=USERS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_USERS, users); log_event(sh, get_username(), "USER_ADD", u)
            st.success("เพิ่มแล้ว", icon="✅"); safe_rerun()

# ---------- Main ----------
def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧰", layout="wide")
    st.markdown(MINIMAL_CSS, unsafe_allow_html=True); st.markdown(RESPONSIVE_CSS, unsafe_allow_html=True)
    st.title(APP_TITLE); st.caption(APP_TAGLINE)
    # โหลดค่า sheet_url จากไฟล์คอนฟิก (ถ้าก่อนหน้าเคยทดสอบสำเร็จ)
    load_config_into_session()
    # เชื่อมต่ออัตโนมัติถ้า URL มีอยู่แล้วและยังไม่ได้เชื่อมใน session
    if st.session_state.get("sheet_url") and "sh" not in st.session_state:
        try:
            _sh = open_sheet_by_url(st.session_state.get("sheet_url"))
            ensure_sheets_exist(_sh)
            st.session_state["sh"] = _sh
            st.session_state["connected"] = True
        except Exception:
            pass

    # Sidebar menu with icons
    with st.sidebar:
        st.markdown("### เมนู")
        page = st.radio("",
            ["📊 Dashboard","📦 คลังอุปกรณ์","🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)","📥 เบิก/รับเข้า","📑 รายงาน","👥 ผู้ใช้","⚙️ Settings"],
            index=0
        )
        st.markdown("---")
        st.write("**admin**"); st.caption("Role: admin")
        if st.button("ออกจากระบบ"):
            for k in ["logged_in","username","role"]:
                st.session_state.pop(k, None)
            safe_rerun()

    if page == "⚙️ Settings":
        require_login()
        page_settings(); st.caption("© 2025 IT Stock · Streamlit + Google Sheets"); return

    # Require sheet URL
    require_login()
    sh = connect_sheet_if_ready()
    if sh is None:
        st.info("ไปที่เมนู **⚙️ Settings** > กรอกและกด **ทดสอบการเชื่อมต่อ** เพื่อเชื่อม Google Sheet เพียงครั้งเดียว", icon="ℹ️")
        return


    if page=="📊 Dashboard": page_dashboard(sh)
    elif page=="📦 คลังอุปกรณ์": page_stock(sh)
    elif page=="🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)": page_tickets(sh)
    elif page=="📥 เบิก/รับเข้า": page_issue_receive(sh)
    elif page=="📑 รายงาน": page_reports(sh)
    elif page=="👥 ผู้ใช้": page_users_admin(sh)

    st.caption("© 2025 IT Stock · Streamlit + Google Sheets")

if __name__ == "__main__":
    main()
