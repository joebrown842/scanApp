# -*- coding: utf-8 -*-

import streamlit as st
import json, os, re, shutil, tempfile, datetime
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageOps
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ---------- constants & helpers ----------
PRESET_FILE = Path("presets.json")
BORDER = Border(
    top=Side(style="thin", color="000000"),
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)

def default_presets():
    return {
        "buildings": {
            "BLDG-1": {
                "Process Equip": {
                    "project": "Chase Bank Renovation",
                    "location": "32 Chase Way · Cedar Park TX 78613",
                    "site_contact": "Trevor Cantor",
                    "phone": "512-915-3075",
                }
            }
        }
    }

def load_presets() -> dict:
    if PRESET_FILE.exists():
        return json.loads(PRESET_FILE.read_text())
    PRESET_FILE.write_text(json.dumps(default_presets(), indent=2))
    return default_presets()

def save_presets(d: dict):
    PRESET_FILE.write_text(json.dumps(d, indent=2))

def ocr_crop(img, box):
    gray = img.crop(box).convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def clean_line(txt: str):
    txt = txt.replace("lJ", "U").replace("l", "I").replace("LOT: STORAGE", "")
    txt = txt.replace(":", "").strip()
    return re.sub(r"\s{2,}", " ", txt)

def extract_items(lines):
    items, i = [], 0
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            if not re.match(r"^\d+\s", nxt):
                desc += " " + nxt
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                items.append((clean_line(desc), qty))
        i += 1
    return items

def fill_excel(template, out, items, meta):
    wb = load_workbook(template)
    ws = wb.active
    header = {
        "B5": meta["project"],
        "B6": meta["location"],
        "B7": meta["delivery_date"],
        "E6": meta["site_contact"],
        "E7": meta["phone"],
    }
    for cell, val in header.items():
        if not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    row = ws.max_row + 1
    for desc, qty in items:
        ws.cell(row, 1, desc).border = BORDER
        ws.cell(row, 2, qty).border = BORDER
        ws.cell(row, 3, meta["building"]).border = BORDER
        ws.cell(row, 4, meta["category"]).border = BORDER
        row += 1
    wb.save(out)

# ---------- Streamlit UI ----------
st.set_page_config(page_title="PDF-to-Excel Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet → Excel Loader")

tabs = st.tabs(["🚚 Process PDF", "🔧 Preset Manager"])
presets = load_presets()

# --------------- TAB 0 : PROCESS PDF ---------------
with tabs[0]:
    st.subheader("1️⃣ Upload scanned PDF + Excel template")
    pdf_file  = st.file_uploader("PDF file", type=["pdf"])
    xlsx_file = st.file_uploader("Excel template (.xlsx)", type=["xlsx"])

    st.subheader("2️⃣ Select Building / Category")
    bldg = st.selectbox("Building", sorted(presets["buildings"].keys()))
    cat_options = sorted(presets["buildings"][bldg].keys())
    category = st.selectbox("Category", cat_options)

    if st.button("🚀 Run & Download") and pdf_file and xlsx_file:
        with st.spinner("Running OCR …"):
            # --- OCR pipeline ---
            tmp_pdf  = Path(tempfile.mktemp(suffix=".pdf"))
            tmp_pdf.write_bytes(pdf_file.read())
            pages = convert_from_path(tmp_pdf)
            lines = []
            for pg in pages:
                w, h = pg.size
                lines += [
                    ln.strip()
                    for ln in ocr_crop(
                        pg, (150, int(h * 0.25), w, int(h * 0.90))
                    ).split("\n")
                    if ln.strip()
                ]
            items = extract_items(lines)
            if not items:
                st.error("❗ No LOT/TYPE lines detected."); st.stop()

            tmp_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
            tmp_xlsx.write_bytes(xlsx_file.read())

            meta = presets["buildings"][bldg][category].copy()
            meta.update({
                "delivery_date": str(datetime.date.today()),
                "building": bldg,
                "category": category,
            })
            fill_excel(tmp_xlsx, tmp_xlsx, items, meta)

            st.success("Finished!")
            st.download_button(
                "⬇️ Download populated workbook",
                data=tmp_xlsx.read_bytes(),
                file_name="filled_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# --------------- TAB 1 : PRESET MANAGER ---------------
with tabs[1]:
    st.subheader("Existing presets")
    table_rows = []
    for b, cats in presets["buildings"].items():
        for c, data in cats.items():
            table_rows.append(
                [b, c, data["site_contact"], data["project"], data["location"], data["phone"]]
            )
    st.dataframe(
        table_rows,
        column_config={
            0: "Building",
            1: "Category",
            2: "Site Contact",
            3: "Project",
            4: "Location",
            5: "Phone",
        },
        hide_index=True,
        use_container_width=True,
    )

    st.divider()
    st.subheader("Add new preset")
    with st.form("new_preset"):
        col1, col2 = st.columns(2)
        with col1:
            nbldg    = st.text_input("Building")
            ncat     = st.text_input("Category")
            nproject = st.text_input("Project")
        with col2:
            nloc  = st.text_input("Location")
            nname = st.text_input("Site Contact")
            nphone = st.text_input("Phone")
        submitted = st.form_submit_button("➕ Save Preset")
        if submitted:
            if not (nbldg and ncat and nproject and nloc and nname and nphone):
                st.warning("Please fill in all fields.")
            else:
                presets["buildings"].setdefault(nbldg, {})[ncat] = {
                    "project": nproject,
                    "location": nloc,
                    "site_contact": nname,
                    "phone": nphone,
                }
                save_presets(presets)
                st.success("Preset saved - refresh Process tab to use it.")
