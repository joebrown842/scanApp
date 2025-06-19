# -*- coding: utf-8 -*-
# app.py – Streamlit App for PDF-to-Excel Shipping Sheet with Project & Preset Management
import json, re, datetime, tempfile
from pathlib import Path
import streamlit as st
from PIL import Image
import pytesseract, pdf2image
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# ─────────────── File helpers ───────────────
PRESET_PATH = Path("presets.json")
def _default(): return {"projects": {}}

def load_presets():
    if not PRESET_PATH.exists():
        PRESET_PATH.write_text(json.dumps(_default(), indent=2))
    try:
        data = json.loads(PRESET_PATH.read_text())
        return data if "projects" in data else _default()
    except Exception:
        PRESET_PATH.write_text(json.dumps(_default(), indent=2))
        return _default()

def save_presets(d): PRESET_PATH.write_text(json.dumps(d, indent=2))
presets = load_presets()

# ─────────────── OCR / Excel helpers ───────────────
BORDER = Border(*(Side(style="thin") for _ in range(4)))

def ocr(img: Image.Image, box):
    g = img.crop(box).convert("L")
    b = g.point(lambda x: 0 if x < 180 else 255, "1")
    return pytesseract.image_to_string(b)

def clean(t): return re.sub(r"\s{2,}", " ", t.replace("LOT: STORAGE", "").strip(" :"))

def extract(lines):
    items = []; i = 0
    while i < len(lines):
        m = re.match(r"^(\d+)\s+(.*)", lines[i])
        if m:
            qty, desc = m.group(1), m.group(2)
            if i+1 < len(lines) and not re.match(r"^\d+\s", lines[i+1]):
                desc += " " + lines[i+1]; i += 1
            if "LOT" in desc.upper() and "TYPE" in desc.upper():
                items.append((clean(desc), qty))
        i += 1
    return items

def fill_xlsx(template, out, rows, meta):
    wb = load_workbook(template); ws = wb.active
    hdr = {"B5":meta["project"],"B6":meta["location"],
           "B7":str(datetime.date.today()),
           "E6":meta["contact"],"E7":meta["phone"]}
    for c,v in hdr.items():
        if not isinstance(ws[c], type(ws["A1"]).MergedCell): ws[c].value=v
    r=ws.max_row+1
    for d,q in rows:
        for col,val in enumerate((d,q,meta["building"],meta["category"]),1):
            ws.cell(r,col,val).border=BORDER
        r+=1
    wb.save(out)

# ─────────────── Streamlit App ───────────────
st.set_page_config("PDF → Excel Loader", layout="wide")
st.title("📑 PDF Shipping-Sheet ➜ Excel Loader")
tab_proc, tab_mgr = st.tabs(["🚚 Process PDF", "🛠️ Preset Manager"])

# ══════════ Process PDF ══════════
with tab_proc:
    if not presets["projects"]:
        st.info("Add a project in Preset Manager first.")
    else:
        proj = st.selectbox("Project", list(presets["projects"]))
        people = presets["projects"][proj]["personnel"]
        if not people:
            st.warning("Add personnel in Preset Manager.")
            st.stop()
        person = st.selectbox("Prepared By", people)
        tree = presets["projects"][proj]["presets"]
        if not tree:
            st.warning("Add presets in Preset Manager.")
            st.stop()
        bldg = st.selectbox("Building", list(tree))
        cat  = st.selectbox("Category", list(tree[bldg]))
        pdf  = st.file_uploader("Scanned PDF",["pdf"])
        xls  = st.file_uploader("Excel Template",["xlsx"])
        if st.button("🚀 Generate Excel") and pdf and xls:
            with st.spinner("OCR in progress…"):
                tmp_pdf=Path(tempfile.mktemp(suffix=".pdf")); tmp_pdf.write_bytes(pdf.read())
                pages=pdf2image.convert_from_path(tmp_pdf)
                lines=[]
                for p in pages:
                    w,h=p.size; lines+= [ln.strip() for ln in ocr(p,(150,int(h*0.25),w,int(h*0.9))).split("\n") if ln.strip()]
                items=extract(lines)
                if not items: st.error("No LOT/TYPE rows found."); st.stop()
                tmp_xls=Path(tempfile.mktemp(suffix=".xlsx")); tmp_xls.write_bytes(xls.read())
                preset=tree[bldg][cat]
                meta=dict(project=proj,location=preset["location"],phone=preset["phone"],
                          contact=preset["contact"],building=bldg,category=cat)
                fill_xlsx(tmp_xls,tmp_xls,items,meta)
                st.success("Workbook ready!")
                st.download_button("⬇ Download Excel", tmp_xls.read_bytes(),"filled_template.xlsx")

# ══════════ Preset Manager ══════════
with tab_mgr:
    st.subheader("📁 Projects")

    proj_input = st.text_input("New Project")
    if st.button("Add Project"):
        name = proj_input.strip()
        if name and name not in presets["projects"]:
            presets["projects"][name]={"personnel":[],"presets":{}}
            save_presets(presets)
            st.success("Project added.")
            st.rerun()

    if not presets["projects"]: st.stop()
    proj = st.selectbox("Select Project", list(presets["projects"]))
    pdata=presets["projects"][proj]

    pers_input = st.text_input("Add Person")
    if st.button("Add Person") and pers_input.strip():
        if pers_input not in pdata["personnel"]:
            pdata["personnel"].append(pers_input.strip()); save_presets(presets); st.success("Person added."); st.rerun()

    st.markdown("### Personnel")
    for i,p in enumerate(pdata["personnel"]):
        st.markdown(f"- {p}")

    st.markdown("---\n### Add Preset")
    with st.form("add_preset", clear_on_submit=True):
        b = st.text_input("Building")
        c = st.text_input("Category")
        loc = st.text_input("Location")
        ct  = st.text_input("Site Contact")
        ph  = st.text_input("Phone")
        if st.form_submit_button("💾 Save"):
            if all([b,c,loc,ph,ct]):
                pdata["presets"].setdefault(b,{})[c]={"location":loc,"phone":ph,"contact":ct}
                save_presets(presets); st.success("Preset saved."); st.rerun()

    st.markdown("---\n### Existing Presets")
    for b,cats in pdata["presets"].items():
        st.markdown(f"#### 🏢 {b}")
        for c,val in cats.items():
            st.markdown(f"- **{c}** — 📍 {val['location']} — 👤 {val['contact']} | 📞 {val['phone']}")
