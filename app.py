#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_stock_optionB_patch.py
---------------------------------
Patch your current app file to enable "เลือก 1 แถวจากตาราง แล้วแก้ไขในแบบฟอร์มด้านล่าง" (Option B)
for the Stock page.

Usage:
    1) Put this file next to your app file (e.g., app.py)
    2) Run:  python apply_stock_optionB_patch.py app.py
    3) It will create: app_patched_stockB.py (safe copy), leaving your original app unchanged.
"""
import sys, re, io, os

TEMPLATE_FUNC = r"""
def page_stock(sh):
    add_reload_button()
    st.subheader("📦 คลังอุปกรณ์")

    import pandas as pd
    df = read_df(sh, SHEET_ITEMS, ITEMS_HEADERS)

    q = st.text_input("ค้นหา (รหัส/ชื่อ/หมวด)", "")
    if q.strip():
        ql = q.strip().lower()
        df = df[df.apply(lambda r: ql in str(r["รหัส"]).lower()
                                   or ql in str(r["ชื่ออุปกรณ์"]).lower()
                                   or ql in str(r["รหัสหมวด"]).lower(), axis=1)]

    # เพิ่มคอลัมน์เลือก
    df_show = df.copy()
    if "เลือก" not in df_show.columns:
        df_show.insert(0, "เลือก", False)

    st.caption("ติ๊กเลือก 1 แถวจากตารางเพื่อแก้ไขรายละเอียดด้านล่าง")
    edited_table = st.data_editor(
        df_show,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "เลือก": st.column_config.CheckboxColumn(help="เลือก 1 แถวเพื่อแก้ไขด้านล่าง"),
            "รหัส": st.column_config.TextColumn(disabled=True),
        },
        disabled=[c for c in df_show.columns if c not in ["เลือก"]],
        key="items_picker",
    )

    selected = edited_table[edited_table["เลือก"] == True]
    if len(selected) != 1:
        st.info("เลือก 1 แถวจากตารางด้านบนเพื่อแก้ไขรายละเอียด", icon="ℹ️")
        return

    row = selected.iloc[0]
    st.markdown("### ✏️ แก้ไขอุปกรณ์")

    col1, col2 = st.columns(2)
    code_id = col1.text_input("รหัส", value=row["รหัส"], disabled=True)
    cat  = col2.text_input("รหัสหมวด", value=row["รหัสหมวด"])
    name = st.text_input("ชื่ออุปกรณ์", value=row["ชื่ออุปกรณ์"])
    unit = st.text_input("หน่วย", value=row["หน่วย"])
    bal  = st.number_input("คงเหลือ", min_value=0, step=1, value=int(pd.to_numeric(row["คงเหลือ"], errors="coerce") or 0))
    rop  = st.number_input("จุดสั่งซื้อ", min_value=0, step=1, value=int(pd.to_numeric(row["จุดสั่งซื้อ"], errors="coerce") or 0))
    loc  = st.text_input("ที่เก็บ", value=row["ที่เก็บ"])
    use  = st.selectbox("ใช้งาน", ["Y","N"], index=0 if str(row["ใช้งาน"]).upper()=="Y" else 1)

    if st.button("บันทึกการแก้ไข", type="primary"):
        df2 = df.copy()
        df2.loc[df2["รหัส"] == code_id, ["รหัสหมวด","ชื่ออุปกรณ์","หน่วย","คงเหลือ","จุดสั่งซื้อ","ที่เก็บ","ใช้งาน"]] =             [cat, name, unit, str(bal), str(rop), loc, use]
        write_df(sh, SHEET_ITEMS, df2.astype(str))
        st.success("บันทึกการแก้ไขเรียบร้อย", icon="✅")
        st.rerun()
"""

def main():
    if len(sys.argv) < 2:
        print("Usage: python apply_stock_optionB_patch.py <your_app_file.py>")
        sys.exit(2)

    path = sys.argv[1]
    with io.open(path, "r", encoding="utf-8") as f:
        src = f.read()

    # Try replace existing page_stock
    new_src, n = re.subn(r'def\s+page_stock\([^\)]*\):[\s\S]*?(?=\ndef\s+\w+\()', TEMPLATE_FUNC + "\n", src, count=1)
    if n == 0:
        # append at end (if not found)
        new_src = src + "\n\n" + TEMPLATE_FUNC + "\n"

    out_path = os.path.splitext(path)[0] + "_patched_stockB.py"
    with io.open(out_path, "w", encoding="utf-8") as f:
        f.write(new_src)

    print("Patched:", out_path)

if __name__ == "__main__":
    main()
