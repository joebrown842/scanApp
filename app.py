# -*- coding: utf-8 -*-

# -------------------- app.py --------------------
import streamlit as st
import json, re, tempfile, datetime
from pathlib import Path
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ---------- constants ----------
PRESET_FILE = Path("presets.json")
BORDER = Border(
    top=Side(style="thin"), left=Side(style="thin"),
    right=Side(style="thin"), bottom=Side(style="thin")
)

def default_presets():
    return {"buildings": {}}  # empty structure

# ---------- helpers ----------
def load_presets():
    if PRESET_FILE.exists():
        return json.loads(PRESET_FILE.read_text())
    PRESET_FILE.write_text(json.dumps(default_presets(), indent=2))
    return default_presets()

def save_presets(d: dict):
    PRESET_FILE.write_text(json.dumps(d, indent=2))

def ocr_crop(img: Image.Image, box):
    gray = img.crop(box).convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def clean_text(txt: str):
    txt = txt.replace("lJ", "U").replace("l", "I")
    txt = txt.replace("LOT: STORAGE", "")
    txt = re.sub(r"\s{2,}", " ", txt)
    return txt.strip(" :")

def extract_items(lines):
    items, i = [], 0
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            # join continuation
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i + 1]):
                desc += " " + lines[i + 1]
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                items.append((clean_text(desc), qty))
        i += 1
    return items

def fill_excel(template, output, items, meta):
    wb = load_workbook(template)
    ws = wb.active
    header_map = {
        "B5": meta["project"],
        "B6": meta["location"],
        "B7": meta["delivery_date"],
        "E6": meta["site_contact"],
        "E7": meta["phone"],
    }
    for cell, val in header_map.items():
        if not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    row = ws.max_row + 1
    for desc, qty in items:
        ws.cell(row, 1, desc).border = BORDER
        ws.cell(row, 2, qty).border = BORDER
        ws.cell(row, 3, meta["building"]).border = BORDER
        ws.cell(row, 4, meta["category"]).border = BORDER
        row += 1
    wb.save(output)

# -------------------- Streamlit UI --------------------
st.set_page_config(page_title="PDF → Excel Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet → Excel Loader")

tabs = st.tabs(["🚚 Process PDF", "🔧 Preset Manager"])
presets = load_presets()

# ============ TAB 0  (Process PDF) ============
with tabs[0]:
    st.subheader("1️⃣ Upload files")
    pdf_file  = st.file_uploader("Scanned PDF", type=["pdf"])
    xlsx_file = st.file_uploader("Excel template (.xlsx)", type=["xlsx"])

    if not presets["buildings"]:
        st.info("No presets yet – add one in *Preset Manager* tab first.")

    bldg = st.selectbox("Building", sorted(presets["buildings"].keys()) or [" "])
    cat_list = sorted(presets["buildings"].get(bldg, {}).keys())
    category = st.selectbox("Category", cat_list or [" "])

    if st.button("🚀 Run OCR & Populate") and pdf_file and xlsx_file and cat_list:
        with st.spinner("Running OCR…"):
            tmp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
            tmp_pdf.write_bytes(pdf_file.read())
            pages = convert_from_path(tmp_pdf)

            lines = []
            for pg in pages:
                w, h = pg.size
                lines += [
                    ln.strip() for ln in ocr_crop(
                        pg, (150, int(h * 0.25), w, int(h * 0.90))
                    ).split("\n") if ln.strip()
                ]

            items = extract_items(lines)
            if not items:
                st.error("No valid LOT/TYPE lines detected."); st.stop()

            tmp_xls = Path(tempfile.mktemp(suffix=".xlsx"))
            tmp_xls.write_bytes(xlsx_file.read())

            meta = presets["buildings"][bldg][category].copy()
            meta.update({
                "delivery_date": str(datetime.date.today()),
                "building": bldg,
                "category": category,
            })
            fill_excel(tmp_xls, tmp_xls, items, meta)

            st.success("Finished!")
            st.download_button(
                "⬇️ Download workbook",
                data=tmp_xls.read_bytes(),
                file_name="filled_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ============ TAB 1  (Preset Manager) ============
with tabs[1]:
    st.subheader("Existing presets")
    rows = []
    for b, cats in presets["buildings"].items():
        for c, d in cats.items():
            rows.append([b, c, d["project"], d["location"], d["site_contact"], d["phone"]])
    st.dataframe(rows, hide_index=True, use_container_width=True,
                 column_config={0: "Building", 1: "Category",
                                2: "Project", 3: "Location",
                                4: "Site Contact", 5: "Phone"})

    st.divider()
    st.subheader("Add new preset")
    with st.form("add_preset"):
        col1, col2 = st.columns(2)
        with col1:
            nbldg     = st.text_input("Building")
            ncat      = st.text_input("Category")
            nproj     = st.text_input("Project Name")
            nloc      = st.text_input("Site Location")
        with col2:
            ncontact  = st.text_input("Site Contact Name")
            nphone    = st.text_input("Phone Number")
        submitted = st.form_submit_button("💾 Save preset")
        if submitted:
            if not all([nbldg, ncat, nproj, nloc, ncontact, nphone]):
                st.warning("Fill in every field")
            else:
                presets["buildings"].setdefault(nbldg, {})[ncat] = {
                    "project": nproj,
                    "location": nloc,
                    "site_contact": ncontact,
                    "phone": nphone,
                }
                save_presets(presets)
                st.success("Preset saved – return to *Process PDF* tab.")
