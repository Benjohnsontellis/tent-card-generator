import streamlit as st
import pandas as pd
import os

from generate_cards import build_pdf
from generate_cards_docx import build_docx

st.set_page_config(page_title="Tent Card Generator", layout="centered")

st.title("🪧 Tent Card Generator")
st.divider()

# ── 1. EXCEL FILE ────────────────────────────────────────────
st.subheader("📄 Excel File")
excel_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

st.divider()

# ── 2. LOGO ──────────────────────────────────────────────────
st.subheader("🖼️ Logo")
logo_file = st.file_uploader("Upload logo image (optional)", type=["png", "jpg", "jpeg"])
if logo_file:
    st.image(logo_file, width=160)

st.divider()

# ── 3. EXPORT FORMAT ─────────────────────────────────────────
st.subheader("📁 Export Format")
export_format = st.radio(
    "Choose output format:",
    ["PDF", "Word Document (.docx)"],
    horizontal=True
)

st.divider()

# ── 4. PASSWORD OVERRIDE ─────────────────────────────────────
st.subheader("🔐 Password Override")
st.caption("Leave blank to use passwords from the Excel file. Enter a value to override all passwords.")
default_password = st.text_input("Override Password", type="password", placeholder="Leave blank to use Excel passwords")

st.divider()

# ── 5. PATIENT LABEL OVERRIDES ───────────────────────────────
st.subheader("🏷️ Patient Label Overrides")
st.caption("Set the label shown before each patient name. Leave blank to use the default 'PATIENT'.")
p1_label = st.text_input("Patient 1 Label", placeholder="e.g. Outpatient  (default: PATIENT)")
p2_label = st.text_input("Patient 2 Label", placeholder="e.g. Inpatient   (default: PATIENT)")
p3_label = st.text_input("Patient 3 Label", placeholder="e.g. Patient 3   (default: PATIENT)")
patient_labels = [p1_label, p2_label, p3_label]

st.divider()

# ── 6. WRISTBAND LABEL ───────────────────────────────────────
st.subheader("📎 Wristband Label")
st.caption("Prefix shown above each wristband QR code. A number is added automatically — e.g. 'Patient Wristband 1'.")
wristband_label = st.text_input("Wristband Label Prefix", value="Patient Wristband", placeholder="e.g. Patient Wristband")

st.divider()

# ── GENERATE ─────────────────────────────────────────────────
if st.button("⚙️ Generate Tent Cards", type="primary", use_container_width=True):

    if excel_file is None:
        st.error("❌ Please upload an Excel file first.")

    else:
        logo_path = None

        if logo_file:
            os.makedirs("logos", exist_ok=True)
            logo_path = f"logos/{logo_file.name}"
            with open(logo_path, "wb") as f:
                f.write(logo_file.read())

        xls       = pd.ExcelFile(excel_file)
        generated = []
        is_docx   = export_format == "Word Document (.docx)"

        with st.spinner("Generating..."):
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str)

                kwargs = dict(
                    logo_path=logo_path,
                    default_password=default_password.strip() if default_password.strip() else None,
                    patient_labels=patient_labels,
                    wristband_label=wristband_label.strip() if wristband_label.strip() else None
                )

                if is_docx:
                    out_path = build_docx(df, sheet, **kwargs)
                    mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ext      = "docx"
                else:
                    out_path = build_pdf(df, sheet, **kwargs)
                    mime     = "application/pdf"
                    ext      = "pdf"

                generated.append((sheet, out_path, mime, ext))

        st.success(f"✅ {len(generated)} file(s) generated!")

        for sheet_name, path, mime, ext in generated:
            with open(path, "rb") as f:
                st.download_button(
                    label=f"⬇️ Download — {sheet_name}.{ext}",
                    data=f,
                    file_name=f"tent_cards_{sheet_name}.{ext}",
                    mime=mime,
                    use_container_width=True
                )
