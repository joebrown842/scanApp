# -*- coding: utf-8 -*-

import streamlit as st
import json, os, re, shutil, tempfile, datetime
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageOps
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# -------------------------  CONSTANTS & HELPERS  ------------------------ #

PRESET_FILE = Path("presets.json")
DEFAULT_PRESETS = {"names": ["Joe Brown"],
                   "buildings": ["BLDG-1"],
                   "categories": ["Lighting"],
                   # project meta that seldom changes
                   "project":  "Temu Data Center",
                   "location": "32 Chase Way · Cedar Park TX 78613",
                   "delivery_date": str(datetime.date.today()),
                   "phone": "(999) 888-7777"
                   }

BORDER = Border(top=Side(style="thin", color="000000"),
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"))

def load_presets() -> dict:
    if PRESET_FILE.exists():
        with PRESET_FILE.open() as f:
            return json.load(f)
    PRESET_FILE.write_text(json.dumps(DEFAULT_PRESETS, indent=2))
    return DEFAULT_PRESETS.copy()

def save_presets(data: dict):
    PRESET_FILE.write_text(json.dumps(data, indent=2))

def ocr_pillow_crop(img: Image.Image, box):
    """basic-contrast OCR of a crop (box = (l,t,r,b))"""
    crop = img.crop(box)
    gray = crop.convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def extract_items(lines):
    """Return list[(desc, qty)] where line starts with qty and has LOT & TYPE"""
    def clean(txt):
        txt = txt.replace("lJ", "U").replace("l", "I")  # common mis-reads
        txt = txt.replace("LOT: STORAGE", "")
        txt = re.sub(r"\s{2,}", " ", txt)
        return txt.strip(" :")
    out, i = [], 0
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            # join continuation line
            nxt = lines[i+1] if i+1 < len(lines) else ""
            if not re.match(r"^\d+\s", nxt):  # continuation?
                desc += " " + nxt
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                out.append((clean(desc), qty))
        i += 1
    return out

def fill_workbook(template_path, output_path, items,
                  presets, name_sel, building_sel, category_sel):
    wb = load_workbook(template_path)
    ws = wb.active
    # header cells (adjust to match your template)
    header_map = {"B5": presets["project"],
                  "B6": presets["location"],
                  "B7": presets["delivery_date"],
                  "E6": name_sel,
                  "E7": presets["phone"]}
    for cell, val in header_map.items():
        if not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    # start after current data
    start_row = ws.max_row + 1
    for desc, qty in items:
        ws.cell(start_row, 1, desc).border = BORDER
        ws.cell(start_row, 2, qty).border  = BORDER
        ws.cell(start_row, 3, building_sel).border  = BORDER
        ws.cell(start_row, 4, category_sel).border  = BORDER
        start_row += 1
    wb.save(output_path)

# ---------------------------  STREAMLIT UI  ----------------------------- #

st.set_page_config(page_title="PDF → Excel Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet to Excel Loader")

tabs = st.tabs(["🔧 Manage Presets", "🚚 Process PDF"])

# ---------------------  TAB 1 : PRESET EDITOR  -------------------------- #
with tabs[0]:
    st.subheader("Add / Remove preset names, buildings, categories")
    presets = load_presets()

    col1, col2, col3 = st.columns(3)
    # ---- Names ----
    with col1:
        st.markdown("**Names**")
        name_to_add = st.text_input("Add new name", key="new_name")
        if st.button("➕ Add Name"):
            if name_to_add and name_to_add not in presets["names"]:
                presets["names"].append(name_to_add)
                save_presets(presets)
                st.experimental_rerun()
        if st.button("❌ Clear Names"):
            presets["names"].clear(); save_presets(presets); st.experimental_rerun()
        st.write(presets["names"])

    # ---- Buildings ----
    with col2:
        st.markdown("**Buildings**")
        bldg_new = st.text_input("Add new building", key="new_bldg")
        if st.button("➕ Add Building"):
            if bldg_new and bldg_new not in presets["buildings"]:
                presets["buildings"].append(bldg_new)
                save_presets(presets)
                st.experimental_rerun()
        if st.button("❌ Clear Buildings"):
            presets["buildings"].clear(); save_presets(presets); st.experimental_rerun()
        st.write(presets["buildings"])

    # ---- Categories ----
    with col3:
        st.markdown("**Categories**")
        cat_new = st.text_input("Add new category", key="new_cat")
        if st.button("➕ Add Category"):
            if cat_new and cat_new not in presets["categories"]:
                presets["categories"].append(cat_new)
                save_presets(presets)
                st.experimental_rerun()
        if st.button("❌ Clear Categories"):
            presets["categories"].clear(); save_presets(presets); st.experimental_rerun()
        st.write(presets["categories"])

# ---------------------  TAB 2 : PROCESS PDF  ---------------------------- #
with tabs[1]:
    st.subheader("1️⃣ Upload scanned PDF")
    pdf_file = st.file_uploader("PDF file", type=["pdf"])

    st.subheader("2️⃣ Upload Excel template (.xlsx)")
    excel_template = st.file_uploader("Excel template", type=["xlsx"])

    st.subheader("3️⃣ Select metadata")
    presets = load_presets()   # reload in case tab 1 changed it
    name_sel = st.selectbox("Name (Site Contact)", options=presets["names"])
    building_sel = st.selectbox("Building", options=presets["buildings"])
    category_sel = st.selectbox("Category", options=presets["categories"])

    if st.button("🚀 Run OCR & Populate Excel") and pdf_file and excel_template:
        with st.spinner("Running OCR and filling spreadsheet…"):
            # ---- OCR pipeline ----
            tmp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
            tmp_pdf.write_bytes(pdf_file.read())
            images = convert_from_path(tmp_pdf)
            all_lines = []
            for img in images:
                w, h = img.size
                crop = img.crop((150, int(h*0.25), w, int(h*0.90)))
                lines = [ln.strip() for ln in ocr_pillow_crop(img, (150, int(h*0.25), w, int(h*0.90))).split("\n") if ln.strip()]
                all_lines.extend(lines)
            items = extract_items(all_lines)
            if not items:
                st.error("No valid LOT/TYPE lines found.")
                st.stop()

            # ---- Save Excel ----
            tmp_xlsx = Path(tempfile.mktemp(suffix=".xlsx"))
            tmp_xlsx.write_bytes(excel_template.read())

            fill_workbook(tmp_xlsx, tmp_xlsx, items, presets,
                          name_sel, building_sel, category_sel)

            st.success("Done! Click to download.")
            st.download_button("⬇️ Download populated Excel",
                               data=tmp_xlsx.read_bytes(),
                               file_name="filled_template.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
