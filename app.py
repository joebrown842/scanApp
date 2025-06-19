# -*- coding: utf-8 -*-
# app.py – Streamlit App for PDF-to-Excel Shipping Sheet with Project & Preset Management
import json, re, datetime, tempfile
from pathlib import Path
import streamlit as st
from PIL import Image
import pytesseract
import pdf2image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

PRESET_PATH = Path("presets.json")
BORDER = Border(*(Side(style="thin") for _ in range(4)))

def _default_struct():
    return {"projects": {}}

def load_presets():
    if not PRESET_PATH.exists():
        PRESET_PATH.write_text(json.dumps(_default_struct(), indent=2))
    try:
        return json.loads(PRESET_PATH.read_text())
    except:
        PRESET_PATH.write_text(json.dumps(_default_struct(), indent=2))
        return _default_struct()

def save_presets(data): 
    PRESET_PATH.write_text(json.dumps(data, indent=2))

presets = load_presets()

def ocr_text(img: Image.Image, box):
    gray = img.crop(box).convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def clean_line(t: str):
    t = (
        t.replace("LOT: STORAGE", "")
        .replace("lJ", "U")
        .replace("l", "I")
        .strip(" :")
    )
    return re.sub(r"\s{2,}", " ", t)

def extract_items(lines):
    out, i = 0, 0
    items = []
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i + 1]):
                desc += " " + lines[i + 1]
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                items.append((clean_line(desc), qty))
        i += 1
    return items

def fill_workbook(tmpl, out, items, meta):
    wb = load_workbook(tmpl)
    ws = wb.active
    header_map = {
        "B5": meta["project"],
        "B6": meta["location"],
        "B7": str(datetime.date.today()),
        "E6": meta["site_contact"],
        "E7": meta["phone"],
    }
    for cell, val in header_map.items():
        if not isinstance(ws[cell], type(ws["A1"]).__class__) or not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    r = ws.max_row + 1
    for desc, qty in items:
        for c, val in enumerate((desc, qty, meta["building"], meta["category"]), 1):
            ws.cell(r, c, val).border = BORDER
        r += 1
    wb.save(out)

st.set_page_config(page_title="Shipping Sheet Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet ➜ Excel Loader")
tab_proc, tab_mgr = st.tabs(["🚚 Process PDF", "🛠️ Preset Manager"])

# Process Tab
with tab_proc:
    if not presets["projects"]:
        st.warning("Create a project first in the Preset Manager.")
    else:
        proj = st.selectbox("Project", list(presets["projects"]))
        ppl = presets["projects"][proj]["personnel"]
        if not ppl:
            st.warning("Add personnel to this project.")
        else:
            person = st.selectbox("Prepared By", ppl)
            pres_tree = presets["projects"][proj]["presets"]
            if not pres_tree:
                st.warning("Add a preset for this project.")
            else:
                bldg = st.selectbox("Building", list(pres_tree))
                cats = pres_tree[bldg]
                cat = st.selectbox("Category", list(cats))
                pdf = st.file_uploader("Scanned PDF", ["pdf"])
                xlsx = st.file_uploader("Excel Template", ["xlsx"])
                if st.button("🚀 Generate Excel") and pdf and xlsx:
                    with st.spinner("Running OCR…"):
                        tmp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
                        tmp_pdf.write_bytes(pdf.read())
                        pages = pdf2image.convert_from_path(tmp_pdf)
                        lines = []
                        for pg in pages:
                            w, h = pg.size
                            crop_box = (150, int(h * 0.25), w, int(h * 0.9))
                            lines += [
                                ln.strip()
                                for ln in ocr_text(pg, crop_box).split("\n")
                                if ln.strip()
                            ]
                        items = extract_items(lines)
                        if not items:
                            st.error("No LOT/TYPE rows detected.")
                        else:
                            tmp_xls = Path(tempfile.mktemp(suffix=".xlsx"))
                            tmp_xls.write_bytes(xlsx.read())
                            preset = cats[cat]
                            meta = {
                                "project": proj,
                                "location": preset["location"],
                                "phone": preset["phone"],
                                "site_contact": preset["contact"],
                                "building": bldg,
                                "category": cat,
                            }
                            fill_workbook(tmp_xls, tmp_xls, items, meta)
                            st.success("Excel ready!")
                            st.download_button("⬇ Download Excel", tmp_xls.read_bytes(), "filled_template.xlsx")

# Manager Tab
with tab_mgr:
    st.subheader("📁 Projects")

    if "new_proj_name" not in st.session_state:
        st.session_state["new_proj_name"] = ""
    st.session_state["new_proj_name"] = st.text_input("New project name", key="new_project_input", value=st.session_state["new_proj_name"])
    if st.button("Add Project"):
        name = st.session_state["new_proj_name"].strip()
        if name:
            if name in presets["projects"]:
                st.warning("Project already exists.")
            else:
                presets["projects"][name] = {"personnel": [], "presets": {}}
                save_presets(presets)
                st.success("Project added.")
                st.session_state["new_proj_name"] = ""
                st.rerun()

    if not presets["projects"]:
        st.stop()

    proj = st.selectbox("Manage Project", list(presets["projects"]))
    proj_data = presets["projects"][proj]
    st.markdown('### 👥 Personnel')
    for i, person in enumerate(proj_data['personnel']):
        col1, col2 = st.columns([5, 1])
        col1.markdown(f'- 👤 {person}')
        if col2.button("🗑️", key=f'del_pers_{i}'):
            proj_data['personnel'].remove(person)
            save_presets(presets)
            st.success('Person deleted.')
            st.rerun()

    st.markdown("### 👥 Personnel")
    if "new_person" not in st.session_state:
        st.session_state["new_person_input"] = ""
pers_input = st.text_input("Add Person", key="new_person_input")
if st.button("Add Person") and pers_input.strip():
    new_p = pers_input.strip()
    if new_p not in proj_data["personnel"]:
        proj_data["personnel"].append(new_p)
        save_presets(presets)
        st.success("Person added.")
        st.session_state["new_person_input"] = ""
        st.rerun()
        st.rerun()

    st.markdown("---\n### 🏗 Presets")
    with st.form("add_preset", clear_on_submit=True):
        b = st.text_input("Building")
        c = st.text_input("Category")
        loc = st.text_input("Location")
        ct = st.text_input("Site Contact")
        ph = st.text_input("Phone")
        if st.form_submit_button("💾 Save Preset"):
            if all([b, c, loc, ph, ct]):
                proj_data["presets"].setdefault(b, {})[c] = {
                    "location": loc,
                    "phone": ph,
                    "contact": ct,
                }
                save_presets(presets)
                st.success("Preset added.")
                st.rerun()

    for bldg, cats in proj_data["presets"].items():
        st.markdown(f"#### 🏢 {bldg}")
        for cat, val in list(cats.items()):
            cols = st.columns([2, 2, 3, 1, 1])
            cols[0].markdown(f"📦 **{cat}**")
            cols[1].markdown(f"📍 {val['location']}")
            cols[2].markdown(f"👤 {val['contact']} | 📞 {val['phone']}")
