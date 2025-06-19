# -*- coding: utf-8 -*-
import streamlit as st, json, re, datetime, tempfile
from pathlib import Path
from PIL import Image
import pytesseract
import pdf2image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ──────────────── File & Style Setup ────────────────
PRESET_PATH = Path("presets.json")
BORDER = Border(top=Side(style="thin"), bottom=Side(style="thin"),
                left=Side(style="thin"), right=Side(style="thin"))

def _default(): return {"projects": {}}

def presets_load():
    if not PRESET_PATH.exists():
        PRESET_PATH.write_text(json.dumps(_default(), indent=2))
    try:
        data = json.loads(PRESET_PATH.read_text())
        return data if "projects" in data else _default()
    except:
        PRESET_PATH.write_text(json.dumps(_default(), indent=2))
        return _default()

def presets_save(data): PRESET_PATH.write_text(json.dumps(data, indent=2))
presets = presets_load()

# ──────────────── OCR / Excel Functions ────────────────
def ocr_crop(pg: Image.Image, box):
    gray = pg.crop(box).convert("L")
    bw = gray.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(bw)

def clean(text):
    return re.sub(r"\s{2,}", " ", text.replace("LOT: STORAGE", "").replace("lJ", "U").replace("l", "I")).strip(" :")

def extract(lines):
    out, i = 0, 0
    results = []
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i + 1]):
                desc += " " + lines[i + 1]
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                results.append((clean(desc), qty))
        i += 1
    return results

def fill_wb(template_path, out_path, items, meta):
    wb = load_workbook(template_path)
    ws = wb.active
    headers = {
        "B5": meta["project"],
        "B6": meta["location"],
        "B7": str(datetime.date.today()),
        "E6": meta["site_contact"],
        "E7": meta["phone"],
    }
    for cell, val in headers.items():
        if not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    r = ws.max_row + 1
    for desc, qty in items:
        for col, val in enumerate((desc, qty, meta["building"], meta["category"]), 1):
            ws.cell(row=r, column=col, value=val).border = BORDER
        r += 1
    wb.save(out_path)

# ──────────────── Streamlit UI ────────────────
st.set_page_config(page_title="Shipping Sheet Loader", layout="wide")
st.title("📦 PDF Shipping-Sheet ➜ Excel Loader")
tab_proc, tab_preset = st.tabs(["🚚 Process PDF", "🛠️ Preset Manager"])

# ──────────────── Tab 1: Process PDF ────────────────
with tab_proc:
    if not presets["projects"]:
        st.warning("No projects yet. Add one in the Preset Manager tab.")
    else:
        project = st.selectbox("Select Project", list(presets["projects"]))
        person_list = presets["projects"][project]["personnel"]
        person = st.selectbox("Prepared By", person_list) if person_list else st.warning("Add personnel first.")

        preset_tree = presets["projects"][project]["presets"]
        buildings = list(preset_tree)
        if buildings:
            building = st.selectbox("Building", buildings)
            categories = list(preset_tree[building])
            if categories:
                category = st.selectbox("Category", categories)
                pdf_file = st.file_uploader("Upload scanned PDF", type="pdf")
                excel_template = st.file_uploader("Upload Excel Template", type="xlsx")
                if st.button("🚀 Run OCR & Generate Excel") and pdf_file and excel_template:
                    with st.spinner("Processing..."):
                        pdf_tmp = Path(tempfile.mktemp(suffix=".pdf"))
                        pdf_tmp.write_bytes(pdf_file.read())
                        pages = pdf2image.convert_from_path(pdf_tmp)

                        all_lines = []
                        for pg in pages:
                            w, h = pg.size
                            lines = ocr_crop(pg, (150, int(h * 0.25), w, int(h * 0.9))).split("\n")
                            all_lines.extend([ln.strip() for ln in lines if ln.strip()])

                        items = extract(all_lines)
                        if not items:
                            st.error("No valid LOT/TYPE entries found.")
                        else:
                            excel_tmp = Path(tempfile.mktemp(suffix=".xlsx"))
                            excel_tmp.write_bytes(excel_template.read())

                            preset = preset_tree[building][category]
                            metadata = {
                                "project": project,
                                "location": preset["location"],
                                "phone": preset["phone"],
                                "site_contact": preset["contact"],
                                "building": building,
                                "category": category,
                            }
                            fill_wb(excel_tmp, excel_tmp, items, metadata)

                            st.success("Excel file ready!")
                            st.download_button("⬇ Download", excel_tmp.read_bytes(), file_name="filled_template.xlsx")
        else:
            st.warning("No presets yet for this project.")

# ──────────────── Tab 2: Preset Manager ────────────────
with tab_preset:
    st.subheader("📁 Project Setup")

    # Create project
    new_project = st.text_input("New Project Name")
    if st.button("Add Project") and new_project:
        if new_project in presets["projects"]:
            st.warning("Project already exists.")
        else:
            presets["projects"][new_project] = {"personnel": [], "presets": {}}
            presets_save(presets)
            st.success(f"Project '{new_project}' added.")
            st.rerun()

    # Select project to manage
    if presets["projects"]:
        project = st.selectbox("Select Project", list(presets["projects"]), key="proj_mgr")
        if st.button("🗑 Delete Project"):
            presets["projects"].pop(project)
            presets_save(presets)
            st.success("Project deleted.")
            st.rerun()

        # Edit project name
        new_name = st.text_input("Rename Project", value=project, key="rename_proj")
        if st.button("✏ Rename"):
            if new_name and new_name != project:
                presets["projects"][new_name] = presets["projects"].pop(project)
                presets_save(presets)
                st.success(f"Renamed to '{new_name}'")
                st.rerun()

        st.markdown("### 👤 Manage Personnel")
        person_add = st.text_input("Add Person")
        if st.button("Add Person") and person_add:
            if person_add not in presets["projects"][project]["personnel"]:
                presets["projects"][project]["personnel"].append(person_add)
                presets_save(presets)
                st.success("Person added.")
                st.rerun()

        # Edit/delete personnel
        for i, p in enumerate(presets["projects"][project]["personnel"]):
            col1, col2, col3 = st.columns([3, 1, 1])
            new_val = col1.text_input(f"Person {i+1}", value=p, key=f"person_{i}")
            if col2.button("✏", key=f"edit_person_{i}") and new_val != p:
                presets["projects"][project]["personnel"][i] = new_val
                presets_save(presets)
                st.rerun()
            if col3.button("🗑", key=f"del_person_{i}"):
                presets["projects"][project]["personnel"].pop(i)
                presets_save(presets)
                st.rerun()

        st.markdown("### 🏗 Manage Presets")

        # Add/edit preset
        with st.form("add_preset_form", clear_on_submit=True):
            b = st.text_input("Building")
            c = st.text_input("Category")
            loc = st.text_input("Site Location")
            phone = st.text_input("Phone")
            contact = st.text_input("Site Contact Name")
            if st.form_submit_button("💾 Save Preset"):
                if not all([b, c, loc, phone, contact]):
                    st.warning("Please fill out all fields.")
                else:
                    presets["projects"][project]["presets"].setdefault(b, {})[c] = {
                        "location": loc,
                        "phone": phone,
                        "contact": contact,
                    }
                    presets_save(presets)
                    st.success("Preset saved.")
                    st.rerun()

        # View/edit existing presets
        for bldg, cats in presets["projects"][project]["presets"].items():
            st.markdown(f"**🏢 {bldg}**")
            for cat, val in cats.items():
                col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 2, 1])
                with col1: val["location"] = st.text_input(f"{bldg}_{cat}_loc", val["location"], key=f"{bldg}_{cat}_loc")
                with col2: val["phone"] = st.text_input(f"{bldg}_{cat}_ph", val["phone"], key=f"{bldg}_{cat}_ph")
                with col3: val["contact"] = st.text_input(f"{bldg}_{cat}_ct", val["contact"], key=f"{bldg}_{cat}_ct")
                if col4.button("💾", key=f"save_{bldg}_{cat}"):
                    presets_save(presets)
                    st.success("Preset updated.")
                if col5.button("🗑", key=f"del_{bldg}_{cat}"):
                    del cats[cat]
                    if not cats: del presets["projects"][project]["presets"][bldg]
                    presets_save(presets)
                    st.rerun()

