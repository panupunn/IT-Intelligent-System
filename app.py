
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta, timezone
import json, os, io, time, uuid

# =============================
# Config & Constants
# =============================
CONFIG_FILE = "app_config.json"
TZ = timezone(timedelta(hours=7))  # Asia/Bangkok

SHEET_ITEMS   = "Items"
SHEET_CATS    = "ItemCategories"
SHEET_BRANCHES= "Branches"
SHEET_TXNS    = "Transactions"
SHEET_TICKETS = "Tickets"
SHEET_TICKET_CATS = "TicketCategories"
SHEET_AUDIT   = "AuditLog"

ITEMS_HEADERS = ["รหัส","รหัสหมวด","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]
CATS_HEADERS  = ["รหัสหมวด","ชื่อหมวด"]
BR_HEADERS    = ["รหัสสาขา","ชื่อสาขา"]
TXNS_HEADERS  = ["TxnID","วันเวลา","ประเภท","รหัส","ชื่ออุปกรณ์","สาขา","จำนวน","โดย","หมายเหตุ"]
TICKETS_HEADERS=["TicketID","วันที่แจ้ง","สาขา","ผู้แจ้ง","หมวดหมู่","รายละเอียด","สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]

# =============================
# Utilities & Persistence
# =============================
def load_config_into_session():
    try:
        if os.path.exists(CONFIG_FILE):
            cfg = json.load(open(CONFIG_FILE,"r",encoding="utf-8"))
        else:
            cfg = {}
        st.session_state.setdefault("sheet_url", cfg.get("sheet_url",""))
        st.session_state.setdefault("cache_ttl", int(cfg.get("cache_ttl", 120)))
        if cfg.get("connected") and st.session_state.get("sheet_url"):
            _connect_if_needed()
    except Exception:
        pass

def save_config_from_session():
    cfg = {
        "sheet_url": st.session_state.get("sheet_url",""),
        "cache_ttl": int(st.session_state.get("cache_ttl",120)),
        "connected": bool(st.session_state.get("sh"))
    }
    json.dump(cfg, open(CONFIG_FILE,"w",encoding="utf-8"), ensure_ascii=False, indent=2)

def get_username(): return st.session_state.get("username","admin")

# ---------- Google Sheets ----------
def _connect_if_needed():
    if st.session_state.get("sh"):  # already
        return st.session_state["sh"]
    url = st.session_state.get("sheet_url","")
    if not url:
        return None
    sh = open_sheet_by_url(url)
    ensure_sheets_exist(sh)
    st.session_state["sh"] = sh
    st.session_state["connected"] = True
    return sh

def open_sheet_by_url(url: str):
    import gspread
    # try service_account.json from working dir or from st.secrets if available
    creds_path = "service_account.json"
    if not os.path.exists(creds_path) and "service_account" in st.secrets:
        json.dump(dict(st.secrets["service_account"]), open(creds_path,"w"))
    gc = gspread.service_account(filename=creds_path)
    return gc.open_by_url(url)

def ensure_sheets_exist(sh):
    needed = {
        SHEET_ITEMS: ITEMS_HEADERS,
        SHEET_CATS: CATS_HEADERS,
        SHEET_BRANCHES: BR_HEADERS,
        SHEET_TXNS: TXNS_HEADERS,
        SHEET_TICKETS: TICKETS_HEADERS,
        SHEET_TICKET_CATS: ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"],
        SHEET_AUDIT: ["เวลา","ผู้ใช้","เหตุการณ์","รายละเอียด"]
    }
    titles = [ws.title for ws in sh.worksheets()]
    for t, hdr in needed.items():
        if t not in titles:
            ws = sh.add_worksheet(t, rows=1000, cols=max(10,len(hdr)))
            ws.update([hdr])
    return True

def log_event(sh, user, event, detail=""):
    try:
        ws = sh.worksheet(SHEET_AUDIT)
        now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([now, user, event, detail], value_input_option="USER_ENTERED")
    except Exception:
        pass

def read_df(sh, title, headers):
    if sh is None: return pd.DataFrame(columns=headers)
    tries=0
    while True:
        try:
            ws = sh.worksheet(title)
            vals = ws.get_all_values()
            break
        except Exception as e:
            msg=str(e)
            if ("429" in msg or "quota" in msg.lower()) and tries<2:
                time.sleep(0.8*(tries+1)); tries+=1; continue
            st.warning(f"อ่านชีต '{title}' ไม่สำเร็จ: {msg}", icon="⚠️")
            return pd.DataFrame(columns=headers)
    if not vals: return pd.DataFrame(columns=headers)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    # ensure all headers exist
    for h in headers:
        if h not in df.columns: df[h]=""
    return df[headers]

def write_df(sh, title, df):
    ws = sh.worksheet(title)
    values = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")

# ---------- UI Helpers ----------
def add_reload_button():
    st.button("🔁 รีโหลดข้อมูล", on_click=lambda: (st.cache_data.clear(), st.toast("รีโหลดแล้ว", icon="🔁")))

def record_recent(key: str, row: list, headers: list):
    df = st.session_state.get(f"recent_{key}")
    new = pd.DataFrame([row], columns=headers)
    st.session_state[f"recent_{key}"] = new if df is None else pd.concat([df, new], ignore_index=True).tail(10)

def require_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"]=True
        st.session_state["username"]="admin"
        st.session_state["role"]="admin"
    return True

# =============================
# Pages
# =============================
def page_dashboard(sh):
    add_reload_button()
    st.subheader("📊 Dashboard")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)

    total_items = len(items)
    low_rop=0
    if not items.empty:
        try:
            low_rop = (pd.to_numeric(items["คงเหลือ"],errors="coerce") <= pd.to_numeric(items["จุดสั่งซื้อ"],errors="coerce")).sum()
        except Exception: pass
    st.metric("จำนวนอุปกรณ์", f"{total_items:,}")
    st.metric("ต่ำกว่า ROP", f"{low_rop:,}")
    st.metric("Tickets ทั้งหมด", f"{len(tickets):,}")

    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**คงเหลือรวมต่อหมวดหมู่ (Top 10)**")
        if items.empty:
            st.info("ยังไม่มีข้อมูล", icon="ℹ️")
        else:
            grp = items.copy()
            grp["คงเหลือ"] = pd.to_numeric(grp["คงเหลือ"], errors="coerce").fillna(0)
            top = grp.groupby("รหัสหมวด")["คงเหลือ"].sum().sort_values(ascending=False).head(10)
            st.bar_chart(top)
    with c2:
        st.markdown("**ธุรกรรม IN/OUT 30 วัน**")
        if txns.empty:
            st.info("ยังไม่มีธุรกรรม", icon="ℹ️")
        else:
            df = txns.copy()
            df["วันเวลา"]=pd.to_datetime(df["วันเวลา"], errors="coerce")
            df = df.dropna(subset=["วันเวลา"])
            cutoff = datetime.now(TZ) - timedelta(days=30)
            df = df[df["วันเวลา"]>=cutoff]
            df["count"]=1
            pv=df.pivot_table(index=df["วันเวลา"].dt.date, columns="ประเภท", values="count", aggfunc="sum").fillna(0)
            st.line_chart(pv)

def render_categories_admin(sh):
    st.markdown("### 🏷️ หมวดหมู่ (อุปกรณ์)")
    cats = read_df(sh, SHEET_CATS, CATS_HEADERS)
    c1,c2 = st.columns([1,2])
    code = c1.text_input("รหัสหมวด")
    name = c2.text_input("ชื่อหมวด")
    if st.button("บันทึก/แก้ไข"):
        if not code or not name:
            st.warning("กรอกให้ครบ", icon="⚠️")
        else:
            base = read_df(sh, SHEET_CATS, CATS_HEADERS)
            if (base["รหัสหมวด"]==code).any():
                base.loc[base["รหัสหมวด"]==code,"ชื่อหมวด"]=name; msg="อัปเดต"
            else:
                base = pd.concat([base, pd.DataFrame([[code,name]], columns=CATS_HEADERS)], ignore_index=True); msg="เพิ่มใหม่"
            write_df(sh, SHEET_CATS, base); log_event(sh, get_username(), "CAT_SAVE", f"{msg}:{code}")
            st.success(f"{msg}แล้ว", icon="✅")
    st.dataframe(cats, use_container_width=True, height=240)

def page_stock(sh):
    add_reload_button()
    st.subheader("📦 คลังอุปกรณ์")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    cats  = read_df(sh, SHEET_CATS, CATS_HEADERS)

    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)")
    view = items if not q else items[items.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
    st.dataframe(view, use_container_width=True, height=300)

    t1,t2,t3 = st.tabs(["➕ เพิ่ม/อัปเดต (รหัสใหม่)","✏️ แก้ไข/ลบ","🏷️ หมวดหมู่"])

    with t1:
        with st.form("item_add", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            with c1:
                opt = (cats["รหัสหมวด"]+" | "+cats["ชื่อหมวด"]).tolist() if not cats.empty else []
                cat_sel = st.selectbox("หมวดหมู่", options=opt)
                cat_code = cat_sel.split(" | ")[0] if cat_sel else ""
                name = st.text_input("ชื่ออุปกรณ์")
            with c2:
                unit = st.text_input("หน่วย", value="ชิ้น")
                qty  = st.number_input("คงเหลือ", min_value=0, value=0, step=1)
                rop  = st.number_input("จุดสั่งซื้อ", min_value=0, value=0, step=1)
            with c3:
                loc  = st.text_input("ที่เก็บ", value="คลังกลาง")
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0)
                code = st.text_input("รหัส (เว้นว่างให้ระบบรัน)", value="")
            s = st.form_submit_button("บันทึก")
        if s:
            df = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
            code_final = (code or f"ITM{len(df)+1:04d}").upper()
            new_row = [code_final, cat_code, name, unit, str(qty), str(rop), loc, active]
            if (df["รหัส"]==code_final).any():
                df.loc[df["รหัส"]==code_final, ITEMS_HEADERS[1]:] = new_row[1:]; msg="อัปเดต"
            else:
                df = pd.concat([df, pd.DataFrame([new_row], columns=ITEMS_HEADERS)], ignore_index=True); msg="เพิ่มใหม่"
            write_df(sh, SHEET_ITEMS, df); log_event(sh, get_username(), "ITEM_SAVE", f"{msg}:{code_final}")
            st.success(f"{msg}แล้ว", icon="✅")
            record_recent("items", new_row, ITEMS_HEADERS)
            st.markdown("#### รายการที่เพิ่ม/แก้ไขล่าสุด")
            st.dataframe(st.session_state.get("recent_items"), use_container_width=True, height=160)

    with t2:
        if items.empty:
            st.info("ยังไม่มีอุปกรณ์", icon="ℹ️")
        else:
            pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
            code_sel = pick.split(" | ")[0]
            row = items[items["รหัส"]==code_sel].iloc[0]
            with st.form("item_edit"):
                name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
                unit = st.text_input("หน่วย", value=row["หน่วย"])
                qty  = st.number_input("คงเหลือ", min_value=0, value=int(float(row["คงเหลือ"] or 0)))
                rop  = st.number_input("จุดสั่งซื้อ", min_value=0, value=int(float(row["จุดสั่งซื้อ"] or 0)))
                loc  = st.text_input("ที่เก็บ", value=row["ที่เก็บ"])
                active = st.selectbox("ใช้งาน", ["Y","N"], index=0 if row["ใช้งาน"]=="Y" else 1)
                save = st.form_submit_button("บันทึกการแก้ไข")
            if save:
                items.loc[items["รหัส"]==code_sel, ["ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]] = [name, unit, str(qty), str(rop), loc, "Y" if active=="Y" else "N"]
                write_df(sh, SHEET_ITEMS, items); log_event(sh, get_username(), "ITEM_UPDATE", code_sel)
                st.success("บันทึกแล้ว", icon="✅")

    with t3:
        render_categories_admin(sh)

def page_issue_receive(sh):
    add_reload_button()
    st.subheader("📥 เบิก/รับเข้า")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)
    if items.empty:
        st.info("ยังไม่มีรายการอุปกรณ์", icon="ℹ️"); return
    t1,t2 = st.tabs(["เบิก (OUT)","รับเข้า (IN)"])

    with t1:
        with st.form("issue", clear_on_submit=True):
            pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist())
            bopt = st.selectbox("เลือกสาขาที่เบิก", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
            qty  = st.number_input("จำนวนที่เบิก", min_value=1, value=1, step=1)
            by   = st.text_input("ผู้ดำเนินการ", value=get_username())
            note = st.text_input("หมายเหตุ", value="")
            s = st.form_submit_button("บันทึกการเบิก")
        if s:
            code_sel = pick.split(" | ")[0]
            row = items[items["รหัส"]==code_sel].iloc[0]
            cur = int(float(row["คงเหลือ"] or 0))
            if qty>cur: st.error("สต็อกไม่พอ", icon="⚠️")
            else:
                items.loc[items["รหัส"]==code_sel,"คงเหลือ"]=str(cur-qty); write_df(sh, SHEET_ITEMS, items)
                txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
                branch_code = bopt.split(" | ")[0] if bopt else ""
                new_txn = [str(uuid.uuid4())[:8], datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"), "OUT", code_sel, row["ชื่ออุปกรณ์"], branch_code, str(qty), by, note]
                txns = pd.concat([txns, pd.DataFrame([new_txn], columns=TXNS_HEADERS)], ignore_index=True)
                write_df(sh, SHEET_TXNS, txns); log_event(sh, get_username(), "ISSUE", f"{code_sel} x {qty} @ {branch_code}")
                st.success("บันทึกแล้ว", icon="✅")
                record_recent("txns", new_txn, TXNS_HEADERS)
                st.markdown("#### รายการที่บันทึกล่าสุด")
                st.dataframe(st.session_state.get("recent_txns"), use_container_width=True, height=160)

    with t2:
        with st.form("receive", clear_on_submit=True):
            pick = st.selectbox("เลือกรายการ", options=(items["รหัส"]+" | "+items["ชื่ออุปกรณ์"]).tolist(), key="recvpick")
            bopt = st.selectbox("เลือกสาขาที่รับเข้า", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [], key="recvbranch")
            qty  = st.number_input("จำนวนที่รับเข้า", min_value=1, value=1, step=1, key="recvqty")
            by   = st.text_input("ผู้ดำเนินการ", value=get_username(), key="recvby")
            note = st.text_input("หมายเหตุ", value="", key="recvnote")
            s = st.form_submit_button("บันทึกรับเข้า")
        if s:
            code_sel = pick.split(" | ")[0]
            row = items[items["รหัส"]==code_sel].iloc[0]
            cur = int(float(row["คงเหลือ"] or 0))
            items.loc[items["รหัส"]==code_sel,"คงเหลือ"]=str(cur+qty); write_df(sh, SHEET_ITEMS, items)
            txns = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
            branch_code = bopt.split(" | ")[0] if bopt else ""
            new_txn = [str(uuid.uuid4())[:8], datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"), "IN", code_sel, row["ชื่ออุปกรณ์"], branch_code, str(qty), by, note]
            txns = pd.concat([txns, pd.DataFrame([new_txn], columns=TXNS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_TXNS, txns); log_event(sh, get_username(), "RECEIVE", f"{code_sel} x {qty} @ {branch_code}")
            st.success("บันทึกรับเข้าแล้ว", icon="✅")
            record_recent("txns", new_txn, TXNS_HEADERS)
            st.markdown("#### รายการที่บันทึกล่าสุด")
            st.dataframe(st.session_state.get("recent_txns"), use_container_width=True, height=160)

def page_tickets(sh):
    add_reload_button()
    st.subheader("🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)")
    cats = read_df(sh, SHEET_TICKET_CATS, ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"])
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    branches = read_df(sh, SHEET_BRANCHES, BR_HEADERS)

    if st.session_state.get("role","admin") in ("admin","staff"):
        tab1, tab2, tab3 = st.tabs(["สร้างคำขอ","รายการทั้งหมด","หมวดหมู่ปัญหา"])
    else:
        tab1, tab2 = st.tabs(["สร้างคำขอ","รายการทั้งหมด"]); tab3=None

    with tab1:
        with st.form("tick_new", clear_on_submit=True):
            bopt = st.selectbox("เลือกสาขาที่แจ้ง", options=(branches["รหัสสาขา"]+" | "+branches["ชื่อสาขา"]).tolist() if not branches.empty else [])
            cat  = st.selectbox("หมวดหมู่ปัญหา", options=(cats["รหัสหมวดปัญหา"]+" | "+cats["ชื่อหมวดปัญหา"]).tolist() if not cats.empty else [])
            who  = st.text_input("ผู้แจ้ง", value=get_username())
            detail = st.text_area("รายละเอียด")
            s = st.form_submit_button("สร้าง Ticket")
        if s:
            df = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
            tid = "T"+datetime.now(TZ).strftime("%y%m%d%H%M%S")
            now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
            catname = cat.split(" | ")[1] if cat else ""
            branch_code = bopt.split(" | ")[0] if bopt else ""
            row=[tid, now, branch_code, who, catname, detail, "เปิดงาน", "", now, ""]
            df = pd.concat([df, pd.DataFrame([row], columns=TICKETS_HEADERS)], ignore_index=True)
            write_df(sh, SHEET_TICKETS, df); log_event(sh, get_username(), "TICKET_NEW", tid)
            st.success("สร้าง Ticket แล้ว", icon="✅")
            record_recent("tickets", row, TICKETS_HEADERS)
            st.markdown("#### รายการ Ticket ที่สร้างล่าสุด")
            st.dataframe(st.session_state.get("recent_tickets"), use_container_width=True, height=160)

    with tab2:
        st.caption("กรองข้อมูล")
        c1,c2,c3 = st.columns(3)
        status = c1.selectbox("สถานะ", ["ทั้งหมด","เปิดงาน","กำลังดำเนินการ","รออะไหล่","จบงาน"])
        whof   = c2.text_input("ผู้แจ้ง (ค้นหา)")
        q      = c3.text_input("คำค้น (รายละเอียด/หมวด)")
        view = tickets.copy()
        if status!="ทั้งหมด": view = view[view["สถานะ"]==status]
        if whof: view = view[view["ผู้แจ้ง"].str.contains(whof, case=False, na=False)]
        if q: view = view[view.apply(lambda r: r.astype(str).str.contains(q, case=False).any(), axis=1)]
        if view.empty:
            st.info("ยังไม่มีรายการ", icon="ℹ️")
        else:
            view = view.copy(); view.insert(0,"เลือก", False)
            ed = st.data_editor(view, use_container_width=True, height=360,
                                column_config={"เลือก": st.column_config.CheckboxColumn()},
                                disabled=[c for c in view.columns if c!="เลือก"], hide_index=True, key="tickets_table")
            sel = ed[ed["เลือก"]==True]
            selected_tid = sel["TicketID"].iloc[0] if len(sel)==1 else None
            with st.expander("อัปเดตสถานะ/ผู้รับผิดชอบ/หมายเหตุ", expanded=bool(selected_tid)):
                if not selected_tid:
                    st.info("เลือก 1 แถวจากตารางด้านบนก่อน", icon="ℹ️")
                else:
                    target = tickets[tickets["TicketID"]==selected_tid].iloc[0]
                    st.write(f"Ticket **{selected_tid}** · สาขา: **{target['สาขา']}** · หมวด: **{target['หมวดหมู่']}**")
                    c1,c2,c3 = st.columns(3)
                    st_new = c1.selectbox("สถานะใหม่", ["เปิดงาน","กำลังดำเนินการ","รออะไหล่","จบงาน"],
                                          index=["เปิดงาน","กำลังดำเนินการ","รออะไหล่","จบงาน"].index(target["สถานะ"]))
                    assignee = c2.text_input("ผู้รับผิดชอบ", value=str(target.get("ผู้รับผิดชอบ","") or ""))
                    note = c3.text_input("หมายเหตุเพิ่มเติม", value=str(target.get("หมายเหตุ","") or ""))
                    if st.button("บันทึกการเปลี่ยนแปลง", type="primary"):
                        tickets.loc[tickets["TicketID"]==selected_tid, ["สถานะ","ผู้รับผิดชอบ","อัปเดตล่าสุด","หมายเหตุ"]] = \
                            [st_new, assignee, datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"), note]
                        write_df(sh, SHEET_TICKETS, tickets); log_event(sh, get_username(), "TICKET_UPDATE", f"{selected_tid}->{st_new}")
                        st.success("อัปเดตแล้ว", icon="✅")
                        record_recent("tickets", tickets[tickets["TicketID"]==selected_tid].iloc[0].values.tolist(), TICKETS_HEADERS)
                        st.markdown("#### รายการ Ticket ที่อัปเดตล่าสุด")
                        st.dataframe(st.session_state.get("recent_tickets"), use_container_width=True, height=160)

    if tab3 is not None:
        with tab3:
            st.markdown("### 🗂️ หมวดหมู่ปัญหา")
            c1,c2 = st.columns([1,2])
            code = c1.text_input("รหัสหมวดปัญหา")
            name = c2.text_input("ชื่อหมวดปัญหา")
            if st.button("บันทึก/แก้ไข"):
                base = read_df(sh, SHEET_TICKET_CATS, ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"])
                if (base["รหัสหมวดปัญหา"]==code).any():
                    base.loc[base["รหัสหมวดปัญหา"]==code,"ชื่อหมวดปัญหา"]=name; msg="อัปเดต"
                else:
                    base = pd.concat([base, pd.DataFrame([[code,name]], columns=["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"])], ignore_index=True); msg="เพิ่มใหม่"
                write_df(sh, SHEET_TICKET_CATS, base); log_event(sh, get_username(), "TICKET_CAT_SAVE", f"{msg}:{code}")
                st.success(f"{msg}แล้ว", icon="✅")
            st.dataframe(read_df(sh, SHEET_TICKET_CATS, ["รหัสหมวดปัญหา","ชื่อหมวดปัญหา"]), use_container_width=True, height=240)

def page_reports(sh):
    add_reload_button()
    st.subheader("📑 รายงาน")
    items = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)
    txns  = read_df(sh, SHEET_TXNS, TXNS_HEADERS)
    tickets = read_df(sh, SHEET_TICKETS, TICKETS_HEADERS)
    c1,c2 = st.columns(2)
    since = c1.date_input("ตั้งแต่", value=date.today()-timedelta(days=30))
    until = c2.date_input("ถึง", value=date.today())

    st.markdown("### รายงานสินค้าต่ำกว่า ROP")
    low = pd.DataFrame(columns=ITEMS_HEADERS)
    if not items.empty:
        try:
            m = pd.to_numeric(items["คงเหลือ"],errors="coerce") <= pd.to_numeric(items["จุดสั่งซื้อ"],errors="coerce")
            low = items[m]
        except Exception: pass
    st.dataframe(low, use_container_width=True, height=200)

    st.markdown("### ธุรกรรมตามช่วงเวลา")
    view = txns.copy()
    if not view.empty:
        view["วันเวลา"]=pd.to_datetime(view["วันเวลา"], errors="coerce")
        view=view.dropna(subset=["วันเวลา"])
        view = view[(view["วันเวลา"].dt.date>=since) & (view["วันเวลา"].dt.date<=until)]
    st.dataframe(view, use_container_width=True, height=260)

    st.markdown("### สรุปการเบิกตามสาขาและอุปกรณ์ (ช่วงเวลาที่เลือก)")
    out = view[view["ประเภท"]=="OUT"].copy() if not view.empty else pd.DataFrame(columns=TXNS_HEADERS)
    if not out.empty:
        out["จำนวน"]=pd.to_numeric(out["จำนวน"], errors="coerce").fillna(0)
        pvt = out.pivot_table(index="สาขา", columns="ชื่ออุปกรณ์", values="จำนวน", aggfunc="sum", fill_value=0)
        st.dataframe(pvt, use_container_width=True, height=200)
        st.bar_chart(pvt.sum(axis=1))

    st.markdown("### สรุป Tickets แยกตามสาขาและหมวดหมู่ (ช่วงเวลาที่เลือก)")
    tv = tickets.copy()
    if not tv.empty:
        tv["วันที่แจ้ง"]=pd.to_datetime(tv["วันที่แจ้ง"], errors="coerce")
        tv=tv.dropna(subset=["วันที่แจ้ง"])
        tv = tv[(tv["วันที่แจ้ง"].dt.date>=since) & (tv["วันที่แจ้ง"].dt.date<=until)]
    if not tv.empty:
        pvt2 = tv.pivot_table(index="สาขา", columns="หมวดหมู่", values="TicketID", aggfunc="count", fill_value=0)
        st.dataframe(pvt2, use_container_width=True, height=200)

def page_users_admin(sh):
    add_reload_button()
    st.subheader("👥 ผู้ใช้")
    st.info("เวอร์ชันนี้มีระบบล็อกอินอย่างง่าย (จำลอง). ถ้าต้องการเชื่อมต่อชีต Users จริง แจ้งได้ครับ", icon="ℹ️")

def page_settings(sh):
    add_reload_button()
    st.subheader("⚙️ Settings")
    st.text_input("Google Sheet URL", key="sheet_url", value=st.session_state.get("sheet_url",""))
    up = st.file_uploader("อัปโหลด service_account.json", type=["json"])
    if up is not None:
        open("service_account.json","wb").write(up.read())
        st.success("อัปโหลด service_account.json แล้ว", icon="✅")
    c1,c2,c3 = st.columns(3)
    if c1.button("บันทึก URL/TTL"):
        save_config_from_session()
        sh = _connect_if_needed()
        if sh: st.success("เชื่อมต่อสำเร็จ", icon="✅")
    if c2.button("ทดสอบการเชื่อมต่อ"):
        sh = _connect_if_needed()
        if sh: st.success("เชื่อมต่อสำเร็จ", icon="✅")
        else: st.error("ยังไม่ได้ตั้งค่า", icon="❌")
    st.slider("TTL แคช (วินาที)", 10, 600, key="cache_ttl")
    st.write("สถานะการเชื่อมต่อ:", "✅ พร้อม" if st.session_state.get("sh") else "❌ ยังไม่ได้เชื่อม")

# =============================
# App main
# =============================
def main():
    st.set_page_config(page_title="IT Intelligent System", layout="wide")
    require_login()
    load_config_into_session()
    sh = _connect_if_needed()
    st.sidebar.title("เมนู")
    page = st.sidebar.radio("", ["📊 Dashboard","📦 คลังอุปกรณ์","🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)","📥 เบิก/รับเข้า","📑 รายงาน","👥 ผู้ใช้","⚙️ Settings"])
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Role: {st.session_state.get('role','admin')}")
    if page=="📊 Dashboard": page_dashboard(sh)
    elif page=="📦 คลังอุปกรณ์": page_stock(sh)
    elif page=="🛠️ แจ้งซ่อม / แจ้งปัญหา (Tickets)": page_tickets(sh)
    elif page=="📥 เบิก/รับเข้า": page_issue_receive(sh)
    elif page=="📑 รายงาน": page_reports(sh)
    elif page=="👥 ผู้ใช้": page_users_admin(sh)
    elif page=="⚙️ Settings": page_settings(sh)

if __name__ == "__main__":
    main()
