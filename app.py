# -*- coding: utf-8 -*-
# PDF-to-Excel Shipping Sheet App
import json, re, datetime, tempfile
from pathlib import Path
import streamlit as st
from PIL import Image
import pytesseract
import pdf2image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ───── Data Setup ─────
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

def save_presets(data): PRESET_PATH.write_text(json.dumps(data, indent=2))
presets = load_presets()

# ───── OCR & Excel Helpers ─────
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
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            if i + 1 < len(lines) and not re.match(r"^\d+\s", lines[i + 1]):
                desc += " " + lines[i + 1]
                i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                out.append((clean_line(desc), qty))
        i += 1
    return out

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
        if not isinstance(ws[cell], type(ws["A1"]).MergedCell):
            ws[cell].value = val
    r = ws.max_row + 1
    for desc, qty in items:
        for c, val in enumerate((desc, qty, meta["building"], meta["category"]), 1):
            ws.cell(r, c, val).border = BORDER
        r += 1
    wb.save(out)

# ───── Streamlit Setup ─────
st.set_page_config(page_title="Shipping Sheet Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet ➜ Excel Loader")
tab_proc, tab_mgr = st.tabs(["🚚 Process PDF", "🛠️ Preset Manager"])

# ════════════════════════ Process Tab ════════════════════════
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
                            preset = pres_tree[bldg][cat]
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

# ════════════════════════ Preset Manager Tab ════════════════════════
with tab_mgr:
    st.subheader("📁 Projects")
    new_proj = st.text_input("New project name")
    if st.button("Add Project") and new_proj:
        if new_proj in presets["projects"]:
            st.warning("Project exists.")
        else:
            presets["projects"][new_proj] = {"personnel": [], "presets": {}}
            save_presets(presets)
            st.success("Project added.")
            st.rerun()

    if not presets["projects"]:
        st.stop()

    proj = st.selectbox("Manage Project", list(presets["projects"]))
    proj_data = presets["projects"][proj]

    # Project rename/delete
    proj_flags = st.session_state.setdefault("proj_edit_flags", {})
    key = f"proj_{proj}"
    col1, col2 = st.columns([4, 1])
    if proj_flags.get(key, False):
        new_name = col1.text_input("Rename Project", value=proj, key=f"rename_input_{proj}")
        if col2.button("💾 Save", key=f"proj_save_{proj}"):
            if new_name and new_name != proj:
                presets["projects"][new_name] = presets["projects"].pop(proj)
                save_presets(presets)
                proj_flags.pop(key)
                st.success("Renamed.")
                st.rerun()
    else:
        col1.markdown(f"### 📁 Project: `{proj}`")
        if col2.button("✏ Rename", key=f"proj_edit_{proj}"):
            proj_flags[key] = True
            st.rerun()
    if st.button("🗑 Delete Project"):
        presets["projects"].pop(proj)
        save_presets(presets)
        st.success("Project deleted.")
        st.rerun()

    # Personnel
    st.markdown("### 👥 Personnel")
    new_p = st.text_input("Add Person")
    if st.button("Add Person") and new_p:
        if new_p not in proj_data["personnel"]:
            proj_data["personnel"].append(new_p)
            save_presets(presets)
            st.rerun()

    person_flags = st.session_state.setdefault("person_edit_flags", {})
    for i, name in enumerate(proj_data["personnel"]):
        pkey = f"{proj}_{i}"
        col1, col2, col3 = st.columns([4, 1, 1])
        if person_flags.get(pkey, False):
            new_val = col1.text_input(f"Person {i+1}", value=name, key=f"pinput_{pkey}")
            if col2.button("💾", key=f"psave_{pkey}"):
                proj_data["personnel"][i] = new_val
                person_flags[pkey] = False
                save_presets(presets)
                st.rerun()
        else:
            col1.markdown(f"**👤 Person {i+1}:** {name}")
            if col2.button("✏", key=f"pedit_{pkey}"):
                person_flags[pkey] = True
                st.rerun()
        if col3.button("🗑", key=f"pdel_{pkey}"):
            proj_data["personnel"].pop(i)
            save_presets(presets)
            st.rerun()

    # Preset form
    st.markdown("---\n### 🏗 Presets")
    with st.form("add_preset", clear_on_submit=True):
        b = st.text_input("Building")
        c = st.text_input("Category")
        loc = st.text_input("Location")
        ph = st.text_input("Phone")
        ct = st.text_input("Site Contact")
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

    # Preset display & editing
    edit_flags = st.session_state.setdefault("preset_edit", {})
    for bldg, cats in proj_data["presets"].items():
        st.markdown(f"#### 🏢 {bldg}")
        for cat, val in cats.items():
            pkey = f"{bldg}_{cat}"
            cols = st.columns([3, 3, 3, 1, 1])
            if edit_flags.get(pkey, False):
                val["location"] = cols[0].text_input("Location", val["location"], key=f"{pkey}_loc")
                val["phone"]    = cols[1].text_input("Phone", val["phone"], key=f"{pkey}_ph")
                val["contact"]  = cols[2].text_input("Contact", val["contact"], key=f"{pkey}_ct")
                if cols[3].button("💾", key=f"save_{pkey}"):
                    save_presets(presets)
                    edit_flags[pkey] = False
                    st.success("Saved.")
                    st.rerun()
            else:
                cols[0].markdown(f"📦 **{cat}**")
                cols[1].markdown(f"📍 {val['location']}")
                cols[2].markdown(f"📞 {val['phone']} | 👤 {val['contact']}")
                if cols[3].button("✏", key=f"edit_{pkey}"):
                    edit_flags[pkey] = True
                    st.rerun()
            if cols[4].button("🗑", key=f"del_{pkey}"):
                cats.pop(cat)
                if not cats:
                    proj_data["presets"].pop(bldg)
                save_presets(presets)
                st.rerun()
