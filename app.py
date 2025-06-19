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
    out, i = [], 0
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
st.title("📁 PDF Shipping-Sheet → Excel Loader")

preset_data = presets_load()

with st.sidebar:
    st.header("Project Selection")
    all_projects = sorted(preset_data["projects"].keys())
    selected_project = st.selectbox("Select a Project", all_projects if all_projects else ["No projects"],
                                     key="active_project")
    if selected_project != "No projects":
        proj_info = preset_data["projects"][selected_project]
        selected_person = st.selectbox("Personnel", proj_info.get("personnel", []), key="active_person")
        selected_building = st.selectbox("Building", sorted(proj_info.get("presets", {}).keys()), key="active_building")
        selected_category = st.selectbox("Category", sorted(proj_info["presets"].get(selected_building, {}).keys()),
                                         key="active_category")


tab_proc, tab_preset = st.tabs(["📄 Process PDF", "🛠️ Preset Manager"])

with tab_proc:
    if not all_projects:
        st.info("No projects yet → add one in Preset Manager")
    else:
        preset = preset_data["projects"][selected_project]["presets"][selected_building][selected_category]
        st.subheader("Upload Files")
        pdf_upl = st.file_uploader("Upload scanned PDF", ["pdf"])
        xls_upl = st.file_uploader("Upload Excel template (.xlsx)", ["xlsx"])

        if st.button("🚀 Run OCR & Populate") and pdf_upl and xls_upl:
            with st.spinner("Running OCR and populating Excel..."):
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
                    meta = {
                        "project": selected_project,
                        "location": preset["location"],
                        "phone": preset["phone"],
                        "site_contact": preset["contact"],
                        "building": selected_building,
                        "category": selected_category
                    }
                    fill_wb(tmp_xls, tmp_xls, items, meta)
                    st.success("Workbook populated.")
                    st.download_button("⬇️ Download file", tmp_xls.read_bytes(),
                                       "filled_template.xlsx",
                                       type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_preset:
    st.subheader("📁 Manage Projects & Presets")

    new_proj = st.text_input("Create New Project")
    if st.button("Add Project"):
        if new_proj and new_proj not in preset_data["projects"]:
            preset_data["projects"][new_proj] = {"personnel": [], "presets": {}}
            presets_save(preset_data)
            st.success("Project added.")
            st.rerun()

    st.divider()
    if selected_project != "No projects":
        st.markdown(f"### 👤 Personnel for `{selected_project}`")
        col1, col2 = st.columns([2, 1])
        col1.write(preset_data["projects"][selected_project]["personnel"] or "*None yet*")
        person_to_add = col2.text_input("Add person")
        if col2.button("Add Person") and person_to_add:
            if person_to_add not in preset_data["projects"][selected_project]["personnel"]:
                preset_data["projects"][selected_project]["personnel"].append(person_to_add)
                presets_save(preset_data); st.rerun()

        person_to_remove = col2.selectbox("Remove person", preset_data["projects"][selected_project]["personnel"])
        if col2.button("Remove Person"):
            preset_data["projects"][selected_project]["personnel"].remove(person_to_remove)
            presets_save(preset_data); st.rerun()

        st.markdown(f"### 🏢 Add/Edit Preset for `{selected_project}`")
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
                    preset_data["projects"][selected_project]["presets"].setdefault(b, {})[c] = {
                        "location": loc, "phone": ph, "contact": ct
                    }
                    presets_save(preset_data)
                    st.success("Preset saved.")
                    st.rerun()

        st.markdown("### 🗑️ Delete Project")
        if st.button("Delete Selected Project"):
            preset_data["projects"].pop(selected_project)
            presets_save(preset_data)
            st.success("Project deleted.")
            st.rerun()

        st.markdown("### 🗑️ Delete Preset")
        if st.button("Delete Selected Preset"):
            preset_data["projects"][selected_project]["presets"][selected_building].pop(selected_category)
            if not preset_data["projects"][selected_project]["presets"][selected_building]:
                preset_data["projects"][selected_project]["presets"].pop(selected_building)
            presets_save(preset_data)
            st.success("Preset deleted.")
            st.rerun()

        st.markdown("### ✏️ Edit Preset")
        curr = preset_data["projects"][selected_project]["presets"][selected_building][selected_category]
        with st.form("edit_form"):
            loc = st.text_input("Site Location", value=curr["location"])
            ph = st.text_input("Phone", value=curr["phone"])
            ct = st.text_input("Site Contact Name", value=curr["contact"])
            if st.form_submit_button("Update Preset"):
                curr.update({"location": loc, "phone": ph, "contact": ct})
                presets_save(preset_data)
                st.success("Preset updated.")
                st.rerun()
