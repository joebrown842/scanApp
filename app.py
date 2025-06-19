# -*- coding: utf-8 -*-
import streamlit as st, json, re, datetime, tempfile
from pathlib import Path
import pytesseract, pdf2image
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

PRESET_PATH = Path("presets.json")
BORDER = Border(top=Side(style="thin"), bottom=Side(style="thin"),
                left=Side(style="thin"), right=Side(style="thin"))

def _default():
    return {"projects": {}}

def presets_load():
    if PRESET_PATH.exists():
        try:
            data = json.loads(PRESET_PATH.read_text())
            if "projects" not in data or not isinstance(data["projects"], dict):
                raise ValueError
            return data
        except Exception:
            PRESET_PATH.write_text(json.dumps(_default(), indent=2))
            return _default()
    PRESET_PATH.write_text(json.dumps(_default(), indent=2))
    return _default()

def presets_save(d: dict):
    PRESET_PATH.write_text(json.dumps(d, indent=2))

presets = presets_load()

def ocr_crop(pg: Image.Image, box):
    gray = pg.crop(box).convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def clean(t: str):
    t = t.replace("lJ", "U").replace("l", "I").replace("LOT: STORAGE", "")
    return re.sub(r"\s{2,}", " ", t).strip(" :")

def extract(lines):
    out, i = 0, []
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i+1]):
                desc += " " + lines[i+1]; i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                out.append((clean(desc), qty))
        i += 1
    return out

def fill_wb(template, out, items, meta):
    wb = load_workbook(template); ws = wb.active
    hdr = {"B5": meta["project"],  "B6": meta["location"],
           "B7": str(datetime.date.today()),
           "E6": meta["site_contact"], "E7": meta["phone"]}
    for c, v in hdr.items():
        if not isinstance(ws[c], type(ws["A1"]).MergedCell):
            ws[c].value = v
    row = ws.max_row + 1
    for desc, qty in items:
        for col, val in enumerate((desc, qty, meta["building"], meta["category"]), 1):
            ws.cell(row, col, val).border = BORDER
        row += 1
    wb.save(out)

st.set_page_config(page_title="PDF → Excel Loader", layout="wide")
st.title("📁 PDF Shipping-Sheet ➔ Excel Loader")

tab_proc, tab_preset = st.tabs(["🚚 Process PDF", "🔧 Preset Manager"])

with tab_proc:
    if not presets["projects"]:
        st.info("No projects yet ➔ add one in *Preset Manager*")
    else:
        proj = st.selectbox("Project", sorted(presets["projects"]))
        people = presets["projects"][proj]["personnel"]
        if not people:
            st.warning("Add personnel first in Preset Manager")
        else:
            user = st.selectbox("Report Prepared By", people)
            bldgs = sorted(presets["projects"][proj]["presets"])
            if not bldgs:
                st.warning("Add a building preset first")
            else:
                bldg = st.selectbox("Building", bldgs)
                cats = sorted(presets["projects"][proj]["presets"][bldg])
                if not cats:
                    st.warning("Add a category under this building")
                else:
                    cat = st.selectbox("Category", cats)
                    pdf_upl = st.file_uploader("Scanned PDF", ["pdf"])
                    xls_upl = st.file_uploader("Excel template (.xlsx)", ["xlsx"])

                    if st.button("🚀 Run OCR & Populate") and pdf_upl and xls_upl:
                        with st.spinner("OCR in progress..."):
                            tmp_pdf = Path(tempfile.mktemp(suffix=".pdf"))
                            tmp_pdf.write_bytes(pdf_upl.read())
                            pages = pdf2image.convert_from_path(tmp_pdf)

                            lines = []
                            for pg in pages:
                                w, h = pg.size
                                lines += [ln.strip() for ln in
                                          ocr_crop(pg, (150, int(h*0.25), w, int(h*0.90))).split("\n")
                                          if ln.strip()]
                            items = extract(lines)
                            if not items:
                                st.error("No LOT/TYPE rows detected.")
                            else:
                                tmp_xls = Path(tempfile.mktemp(suffix=".xlsx"))
                                tmp_xls.write_bytes(xls_upl.read())

                                preset = presets["projects"][proj]["presets"][bldg][cat]
                                meta = {"project": proj, "location": preset["location"],
                                        "phone": preset["phone"], "site_contact": preset["contact"],
                                        "building": bldg, "category": cat}
                                fill_wb(tmp_xls, tmp_xls, items, meta)

                                st.success("Workbook ready")
                                st.download_button("⬇️ Download file", tmp_xls.read_bytes(),
                                                   "filled_template.xlsx",
                type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_preset:
    st.subheader("📁 Projects")
    st.markdown("### ➕ Create New Project")
    new_proj = st.text_input("Project name", key="create_proj")
    if st.button("Add project"):
        if not new_proj:
            st.warning("Please enter a project name.")
        elif new_proj in presets["projects"]:
            st.warning("Project already exists.")
        else:
            presets["projects"][new_proj] = {"personnel": [], "presets": {}}
            presets_save(presets)
            st.success("✅ Project added.")
            st.rerun()

    if presets["projects"]:
        st.divider()
        proj = st.selectbox("Manage project", sorted(presets["projects"]))
        if st.button("🗑️ Delete Project"):
            if st.button("Confirm Delete", key="del_confirm"):
                presets["projects"].pop(proj)
                presets_save(presets)
                st.rerun()

        proj_data = presets["projects"][proj]

        st.markdown("### 👤 Project Personnel")
        col1, col2 = st.columns([2, 1])
        col1.write(proj_data["personnel"] or "*None yet*")
        person_to_add = col2.text_input("Add person", key="add_person")
        if col2.button("Add", key="btn_add_person") and person_to_add:
            if person_to_add not in proj_data["personnel"]:
                proj_data["personnel"].append(person_to_add)
                presets_save(presets); st.rerun()

        st.divider()
        st.markdown("### 🗑️ Delete a Preset")
        if proj_data["presets"]:
            b_sel = st.selectbox("Building", sorted(proj_data["presets"]), key="del_bldg")
            c_sel = st.selectbox("Category", sorted(proj_data["presets"][b_sel]), key="del_cat")
            if st.button("Delete Preset"):
                proj_data["presets"][b_sel].pop(c_sel, None)
                if not proj_data["presets"][b_sel]:
                    proj_data["presets"].pop(b_sel)
                presets_save(presets); st.rerun()
        else:
            st.info("No presets yet.")

        st.divider()
        st.markdown("### 🏧 Existing Building / Category Presets")
        rows = [[b, c, d["location"], d["phone"], d["contact"]]
                for b, cats in proj_data["presets"].items()
                for c, d in cats.items()]
        st.dataframe(rows, hide_index=True,
                     column_config={0:"Building",1:"Category",2:"Location",3:"Phone",4:"Site Contact"},
                     use_container_width=True)

        st.divider()
        st.markdown("### ➕ Add or Update a Preset")
        with st.form("preset_form", clear_on_submit=True):
            b = st.text_input("Building")
            c = st.text_input("Category")
            loc = st.text_input("Site Location")
            ph = st.text_input("Phone")
            ct = st.text_input("Site Contact Name")
            if st.form_submit_button("Save Preset"):
                if not all([b, c, loc, ph, ct]):
                    st.warning("Fill all fields.")
                else:
                    proj_data["presets"].setdefault(b, {})[c] = {
                        "location": loc, "phone": ph, "contact": ct
                    }
                    presets_save(presets)
                    st.success("Preset saved.")
                    st.rerun()
