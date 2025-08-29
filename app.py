
# -*- coding: utf-8 -*-
"""
app_fixed_v12k3_dashboard_pdf_full_thaifont_finalfix.py
Dashboard demo with Thai-font-safe PDF export.

- Upload a Thai TTF once (e.g., NotoSansThai-Regular.ttf or Sarabun-Regular.ttf)
- Charts on screen and in PDF will use that font
- PDF embeds TrueType so Thai shows correctly

You can copy only the functions you need (ensure_thai_font, make_pie/make_bar,
export_charts_to_pdf, the PDF expander block) into your existing app.
"""

import os
import io
import tempfile
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st

import matplotlib
from matplotlib import pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.backends.backend_pdf import PdfPages


# ---------------------------- Thai Font Helper ---------------------------- #

def ensure_thai_font(font_path: str = None):
    """
    Load Thai TTF and return a FontProperties; also set rcParams so PDF shows Thai.
    If font_path is None, try common Thai families; fallback to DejaVu Sans.
    """
    try:
        if font_path and os.path.exists(font_path):
            fm.fontManager.addfont(font_path)
            prop = fm.FontProperties(fname=font_path)
            matplotlib.rcParams["font.family"] = prop.get_name()
        else:
            preferred = [
                "Noto Sans Thai", "Sarabun", "TH Sarabun New",
                "Kanit", "Prompt", "Leelawadee UI", "Tahoma"
            ]
            # Build name->path map of installed fonts
            available = {f.name: f.fname for f in fm.fontManager.ttflist}
            chosen = None
            for name in preferred:
                if name in available:
                    chosen = name
                    break
            if chosen:
                fm.fontManager.addfont(available[chosen])
                prop = fm.FontProperties(fname=available[chosen])
                matplotlib.rcParams["font.family"] = chosen
            else:
                prop = fm.FontProperties(family="DejaVu Sans")
                matplotlib.rcParams["font.family"] = "DejaVu Sans"

        # Critical for embedding TrueType into PDF
        matplotlib.rcParams["axes.unicode_minus"] = False
        matplotlib.rcParams["pdf.fonttype"] = 42
        matplotlib.rcParams["ps.fonttype"] = 42
        return prop
    except Exception:
        # very defensive fallback
        matplotlib.rcParams["font.family"] = "DejaVu Sans"
        matplotlib.rcParams["axes.unicode_minus"] = False
        matplotlib.rcParams["pdf.fonttype"] = 42
        matplotlib.rcParams["ps.fonttype"] = 42
        return fm.FontProperties(family="DejaVu Sans")


# ---------------------------- Chart helpers ---------------------------- #

def _topn_series(df: pd.DataFrame, label_col: str, value_col: str, top_n: int):
    s = df.groupby(label_col)[value_col].sum().sort_values(ascending=False)
    if top_n and len(s) > top_n:
        top = s.iloc[:top_n]
        other = s.iloc[top_n:].sum()
        s = pd.concat([top, pd.Series({"อื่น ๆ": float(other)})])
    return s


def make_pie(df: pd.DataFrame, label_col: str, value_col: str, top_n: int, title: str, font_prop=None):
    fig, ax = plt.subplots(figsize=(6, 6))
    s = _topn_series(df, label_col, value_col, top_n)
    wedges, texts, autotexts = ax.pie(
        s.values, labels=s.index, autopct="%1.1f%%", startangle=90, counterclock=False
    )
    ax.axis("equal")
    if font_prop is not None:
        for t in texts + autotexts:
            t.set_fontproperties(font_prop)
        ax.set_title(title, fontproperties=font_prop)
    else:
        ax.set_title(title)
    fig.tight_layout()
    return fig


def make_bar(df: pd.DataFrame, label_col: str, value_col: str, top_n: int, title: str, font_prop=None):
    fig, ax = plt.subplots(figsize=(8, 5))
    s = _topn_series(df, label_col, value_col, top_n)
    ax.bar(s.index, s.values)
    if font_prop is not None:
        ax.set_title(title, fontproperties=font_prop)
        ax.set_ylabel(value_col, fontproperties=font_prop)
        for lab in ax.get_xticklabels() + ax.get_yticklabels():
            lab.set_fontproperties(font_prop)
    else:
        ax.set_title(title)
        ax.set_ylabel(value_col)
    fig.tight_layout()
    return fig


# ---------------------------- PDF Export ---------------------------- #

def export_charts_to_pdf(charts, selected_titles, chart_kind: str):
    """
    charts: list of (title, df, label_col, value_col)
    selected_titles: list[str]
    chart_kind: "Pie" or "Bar"
    """
    font_path = st.session_state.get("thai_font_path") if "thai_font_path" in st.session_state else None
    prop = ensure_thai_font(font_path)

    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        for title, df, label_col, value_col in charts:
            if title not in selected_titles:
                continue
            if chart_kind == "Bar":
                fig = make_bar(df, label_col, value_col, top_n=10, title=title, font_prop=prop)
            else:
                fig = make_pie(df, label_col, value_col, top_n=10, title=title, font_prop=prop)
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------- Demo Dashboard ---------------------------- #

def _demo_data():
    rng = np.random.default_rng(42)
    equip = ["เมาส์", "คีย์บอร์ด", "หมึกพิมพ์", "ดรัม", "คอมพิวเตอร์", "อุปกรณ์ อื่นๆ"]
    branch = ["HQ02 | ฝ่ายA", "HQ03 | GR", "HQ14 | ฝ่ายB", "SWC001 | สำนักงานใหญ่",
              "SWC015 | เขตฝั่งน", "SWC019 | เขตฝั่งตะวันตก"]
    rows = []
    today = date.today()
    for i in range(400):
        rows.append({
            "วันที่": today - timedelta(days=int(rng.integers(0, 60))),
            "อุปกรณ์": rng.choice(equip),
            "จำนวน": int(rng.integers(1, 6)),
            "สาขา": rng.choice(branch)
        })
    df = pd.DataFrame(rows)
    return df


def page_dashboard():
    st.title("ช่วงเวลา (ใช้กับกราฟประเภท 'เบิก ... (OUT)' เท่านั้น)")

    # --- filters ---
    kind = st.radio("ชนิดกราฟ", ["Pie", "Bar"], horizontal=True)
    col1, col2 = st.columns(2)
    start = col1.date_input("วันที่เริ่ม", value=date.today() - timedelta(days=30))
    end = col2.date_input("วันที่สิ้นสุด", value=date.today())
    df = _demo_data()
    df = df[(df["วันที่"] >= start) & (df["วันที่"] <= end)].copy()

    # --- Charts to show ---
    charts = [
        ("เบิกตามหมวดหมู่ (OUT)", df.rename(columns={"อุปกรณ์": "หมวดหมู่"}), "หมวดหมู่", "จำนวน"),
        ("เบิกตามสาขา (OUT)", df.rename(columns={"สาขา": "สาขา"}), "สาขา", "จำนวน"),
    ]

    # Thai font for on-screen charts
    font_path = st.session_state.get("thai_font_path") if "thai_font_path" in st.session_state else None
    prop = ensure_thai_font(font_path)

    # --- render ---
    for title, data, label_col, value_col in charts:
        if kind == "Bar":
            fig = make_bar(data, label_col, value_col, top_n=10, title=title, font_prop=prop)
        else:
            fig = make_pie(data, label_col, value_col, top_n=10, title=title, font_prop=prop)
        st.pyplot(fig, use_container_width=True)

    # --- PDF Panel ---
    with st.expander("พิมพ์/ดาวน์โหลดกราฟเป็น PDF", expanded=False):
        up = st.file_uploader("อัปโหลดฟอนต์ไทย (.ttf) เพื่อให้ PDF แสดงไทยถูกต้อง", type=["ttf"])
        if up is not None:
            save_dir = os.path.join(tempfile.gettempdir(), "thai_fonts")
            os.makedirs(save_dir, exist_ok=True)
            save_path = os.path.join(save_dir, up.name or "thai_font.ttf")
            with open(save_path, "wb") as f:
                f.write(up.read())
            st.session_state["thai_font_path"] = save_path
            st.success("บันทึกฟอนต์ไทยแล้ว: จะใช้ในการสร้าง PDF")

        if "thai_font_path" in st.session_state:
            st.caption("ใช้ฟอนต์ไทยจาก: " + st.session_state["thai_font_path"])

        titles_all = [t for (t, *_rest) in charts]
        sel = st.multiselect("เลือกกราฟที่จะพิมพ์เป็น PDF", options=titles_all, default=titles_all)
        if sel:
            pdf_bytes = export_charts_to_pdf(charts, sel, kind)
            st.download_button("ดาวน์โหลด PDF กราฟที่เลือก", data=pdf_bytes,
                               file_name="dashboard_charts.pdf", mime="application/pdf")


def main():
    st.set_page_config(page_title="IT Intelligent System (Demo PDF Thai)", layout="wide")
    page_dashboard()


if __name__ == "__main__":
    main()
