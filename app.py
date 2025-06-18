# -*- coding: utf-8 -*-
import streamlit as st, json, re, datetime, tempfile
from pathlib import Path
import pytesseract, pdf2image
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ──────────────── Preset storage ────────────────
PRESET_PATH = Path("presets.json")
BORDER = Border(
    top=Side(style="thin"), bottom=Side(style="thin"),
    left=Side(style="thin"), right=Side(style="thin")
)

def _default():
    return {"projects": {}}

def presets_load():
    if PRESET_PATH.exists():
        try:
            data = json.loads(PRESET_PATH.read_text())
            if "projects" not in data or not isinstance(data["projects"], dict):
                raise ValueError("Invalid format")
            return data
        except Exception:
            st.warning("⚠️ presets.json was missing or corrupt. Recreated clean file.")
            PRESET_PATH.write_text(json.dumps(_default(), indent=2))
            return _default()
    else:
        PRESET_PATH.write_text(json.dumps(_default(), indent=2))
        return _default()

def presets_save(data): PRESET_PATH.write_text(json.dumps(data, indent=2))

presets = presets_load()

# ──────────────── OCR helpers ────────────────
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
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i + 1]):
                desc += " " + lines[i + 1]
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                out.append((clean(desc), qty))
        i += 1
    return out

# ──────────────── Excel Writer ────────────────
def fill_wb(template, out, items, meta):
    wb = load_workbook(template); ws = wb.active
    hdr = {
        "B5": meta["project"],
        "B6": meta["location"],
        "B7": str(datetime.date.today()),
        "E6": meta["site_contact"],
        "E7": meta["phone"],
    }
    for c, v in hdr.items():
        if not isinstance(ws[c], type(ws["A1"]).MergedCell): ws[c].value = v
    r = ws.max_row + 1
    for desc, qty in items:
        for col, val in enumerate((desc, qty, meta["building"], meta["category"]), 1):
            ws.cell(r, col, val).border = BORDER
        r += 1
    wb.save(out)

# ──────────────── Streamlit UI ────────────────
st.set_page_config(page_title="PDF → Excel Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet → Excel Loader")

tab_proc, tab_preset = st.tabs(["🚚 Process PDF", "🔧 Preset Manager"])

# ============ TAB 1: PROCESS PDF ============ #
with tab_proc:
    if not presets["projects"]:
        st.info("No projects found. Please add one in the *Preset Manager* tab.")
    else:
        proj = st.selectbox("Project", sorted(presets["projects"].keys()))
        ppl  = presets["projects"][proj]["personnel"]
        if not ppl:
            st.warning("Add personnel to this project in Preset Manager.")
        else:
            person = st.selectbox("Report Prepared By", ppl)

            bldg_opts = sorted(presets["projects"][proj]["presets"].keys())
            if not bldg_opts:
                st.warning("No building presets in project. Add in Preset Manager.")
            else:
                bldg = st.selectbox("Building", bldg_opts)

                cat_opts = sorted(presets["projects"][proj]["presets"][bldg].keys())
                if not cat_opts:
                    st.warning("No categories under selected building.")
                else:
                    cat = st.selectbox("Category", cat_opts)

                    pdf_upl = st.file_uploader("Scanned PDF", type=["pdf"])
                    xls_upl = st.file_uploader("Excel Template (.xlsx)", type=["xlsx"])

                    if st.button("🚀 Run OCR & Populate") and pdf_upl and xls_upl:
                        with st.spinner("Running OCR…"):
                            pdf_tmp = Path(tempfile.mktemp(suffix=".pdf"))
                            pdf_tmp.write_bytes(pdf_upl.read())
                            pages = pdf2image.convert_from_path(pdf_tmp)

                            all_lines = []
                            for pg in pages:
                                w, h = pg.size
                                lines = ocr_crop(pg, (150, int(h*0.25), w, int(h*0.90))).split("\n")
                                all_lines += [ln.strip() for ln in lines if ln.strip()]
                            items = extract(all_lines)
                            if not items:
                                st.error("No LOT/TYPE lines detected.")
                            else:
                                xls_tmp = Path(tempfile.mktemp(suffix=".xlsx"))
                                xls_tmp.write_bytes(xls_upl.read())

                                preset = presets["projects"][proj]["presets"][bldg][cat]
                                meta = {
                                    "project": proj,
                                    "location": preset["location"],
                                    "phone": preset["phone"],
                                    "site_contact": person,
                                    "building": bldg,
                                    "category": cat,
                                }
                                fill_wb(xls_tmp, xls_tmp, items, meta)

                                st.success("Workbook ready!")
                                st.download_button("⬇️ Download workbook",
                                                   xls_tmp.read_bytes(),
                                                   file_name="filled_template.xlsx",
                                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============ TAB 2: PRESET MANAGER ============ #
with tab_preset:
    st.subheader("📁 Projects")

    st.markdown("### ➕ Create New Project")
    new_proj = st.text_input("Project Name", key="create_project")
    if st.button("➕ Add Project"):
        if not new_proj:
            st.warning("Please enter a project name.")
        elif new_proj in presets["projects"]:
            st.warning("Project already exists.")
        else:
            presets["projects"][new_proj] = {"personnel": [], "presets": {}}
            presets_save(presets)
            st.success(f"✅ Project '{new_proj}' created.")
            st.experimental_rerun()

    if presets["projects"]:
        st.divider()
        proj = st.selectbox("Manage project", sorted(presets["projects"].keys()))
        proj_data = presets["projects"][proj]

        # ─── Manage Personnel ───
        st.markdown("### 👤 Project Personnel")
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write(proj_data["personnel"] or "*None yet*")
        with col2:
            p_new = st.text_input("Add person")
            if st.button("Add Person"):
                if p_new and p_new not in proj_data["personnel"]:
                    proj_data["personnel"].append(p_new)
                    presets_save(presets)
                    st.experimental_rerun()

        st.divider()

        # ─── Existing presets ───
        st.markdown("### 🏗️ Existing Building/Category Presets")
        rows = []
        for b, cats in proj_data["presets"].items():
            for c, d in cats.items():
                rows.append([b, c, d["location"], d["phone"]])
        st.dataframe(rows, hide_index=True,
                     column_config={0:"Building",1:"Category",2:"Location",3:"Phone"},
                     use_container_width=True)

        st.divider()

        # ─── Add new preset ───
        st.markdown("### ➕ Add or Update a Preset")
        with st.form("add_preset"):
            b = st.text_input("Building")
            c = st.text_input("Category")
            loc = st.text_input("Site Location")
            ph = st.text_input("Phone Number")
            if st.form_submit_button("💾 Save Preset"):
                if not all([b, c, loc, ph]):
                    st.warning("Please fill out all fields.")
                else:
                    proj_data["presets"].setdefault(b, {})[c] = {
                        "location": loc, "phone": ph
                    }
                    presets_save(presets)
                    st.success("Preset saved.")
                    st.experimental_rerun()
