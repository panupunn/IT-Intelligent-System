# IT Intelligent System (Streamlit + Google Sheets)

ระบบสต็อก IT + เบิก/รับเข้า + Dashboard + รายงาน + แจ้งปัญหา  
ดีไซน์มินิมอล ใช้งานง่าย รองรับการนำไป Deploy บน **GitHub + Streamlit Community Cloud**

## โครงสร้างโปรเจกต์
```
it-intelligent-system/
├─ app.py                        # ไฟล์หลักของแอป (เปลี่ยนชื่อจากเวอร์ชันล่าสุด)
├─ requirements.txt              # ไลบรารีที่ต้องใช้
├─ .gitignore
├─ .streamlit/
│  ├─ config.toml
│  └─ secrets.toml.example       # ตัวอย่างไฟล์ secrets (ห้าม commit ไฟล์จริง)
├─ templates/                    # เทมเพลต CSV นำเข้า
│  ├─ template_categories.csv
│  ├─ template_branches.csv
│  ├─ template_items.csv
│  └─ template_ticket_categories.csv
└─ fonts/
   ├─ Sarabun-Regular.ttf        # ฟอนต์ไทยสำหรับ PDF
   └─ Sarabun-Bold.ttf
```

## การเตรียมค่า (Secrets)
แนะนำให้เก็บค่าใน **`.streamlit/secrets.toml`** (บน Streamlit Cloud ก็ใช้เมนู Secrets ได้)
ตัวอย่างไฟล์ดูที่ `.streamlit/secrets.toml.example`

ค่าที่ต้องใส่:
- `SHEET_URL` : ลิงก์ Google Sheet ของคุณ
- `SERVICE_ACCOUNT_JSON` : เนื้อหา JSON ของ Service Account (คัดลอกทั้งไฟล์มาแปะ)

> อย่าลืมแชร์ Google Sheet ให้ **อีเมลของ Service Account** ด้วย

## การรันแบบ Local
```bash
# Python 3.10+ แนะนำให้สร้าง virtual env
pip install -r requirements.txt
streamlit run app.py
```

## การ Deploy บน Streamlit Community Cloud
1. สร้าง GitHub repo แล้วอัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้
2. ไปที่ https://share.streamlit.io/ → Connect to GitHub → เลือก repo
3. ตั้งค่า **Main file path** เป็น `app.py`
4. เปิดเมนู **Settings → Secrets** แล้ววางค่าตาม `.streamlit/secrets.toml.example`
5. กด Deploy ได้เลย

## หมายเหตุเรื่องฟอนต์ไทยใน PDF
ระบบมีฟอนต์ **Sarabun** ใน `fonts/` แล้ว หากโค้ดออกรายงาน PDF แนะให้ชี้ไปที่ฟอนต์นี้
(ตัวอย่างการลงทะเบียนฟอนต์ด้วย reportlab)
```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont("Sarabun", "fonts/Sarabun-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Sarabun-Bold", "fonts/Sarabun-Bold.ttf"))
```

## License
MIT (แก้ไขได้ตามสะดวก)
