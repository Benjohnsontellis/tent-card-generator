import pandas as pd
import os
import datetime
import re
import qrcode

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image
)

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Page width available inside default margins
PAGE_W = 7.0 * inch


# -----------------------------
# CLEAN VALUE
# -----------------------------
def clean(value):

    if pd.isna(value):
        return None

    value = str(value).strip()

    if value == "" or value.lower() == "nan":
        return None

    return value


# -----------------------------
# CREATE QR CODE
# -----------------------------
def create_qr(data):

    img = qrcode.make(data)

    path = "temp_qr.png"

    img.save(path)

    return path


# -----------------------------
# DETECT PATIENT NUMBERS
# -----------------------------
def detect_patients(columns):

    nums = []

    for c in columns:

        match = re.match(r"patient (\d+) mrn", c.lower())

        if match:
            nums.append(int(match.group(1)))

    return sorted(set(nums))


# -----------------------------
# GET LOCATION FIELDS
# -----------------------------
def get_location_fields(row, patient_number):

    suffix = "" if patient_number == 1 else str(patient_number)

    location_fields = [
        "Site",
        "Building",
        "NurseStation",
        "Room",
        "Bed"
    ]

    result = []

    for field in location_fields:

        column = f"{field}{suffix}"

        value = clean(row.get(column))

        if value:
            label = field.replace("NurseStation", "Ward")
            result.append((label, value))

    return result


# -----------------------------
# BUILD PDF
# -----------------------------
def build_pdf(
    df,
    sheet,
    logo_path=None,
    default_password=None,
    patient_labels=None,
    wristband_label=None
):

    if not os.path.exists("output"):
        os.makedirs("output")

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    filename = f"output/tent_cards_{timestamp}.pdf"

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch
    )

    elements = []
    styles   = getSampleStyleSheet()

    # ── Styles ───────────────────────────────────────────────
    small = ParagraphStyle(
        'small',
        parent=styles['Normal'],
        fontSize=9,
        leading=12          # Fix 3: tighter line spacing in patient blocks
    )

    small_bold = ParagraphStyle(
        'small_bold',
        parent=styles['Normal'],
        fontSize=9,
        leading=15,
        fontName='Helvetica-Bold'
    )

    bold = ParagraphStyle(
        'bold',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica-Bold',
        leading=14
    )

    section_head = ParagraphStyle(
        'section_head',
        parent=styles['Normal'],
        fontSize=11,
        fontName='Helvetica-Bold',
        leading=14,
        leftIndent=0,       # flush left, no indent
        spaceBefore=0
    )

    bold_center = ParagraphStyle(
        'bold_center',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        alignment=TA_CENTER,
        leading=13
    )

    title_style = ParagraphStyle(
        'title_style',
        parent=styles['Heading1'],
        fontSize=16,
        fontName='Helvetica-Bold',
        alignment=TA_LEFT,
        leading=20,
        textColor=colors.HexColor('#1a1a1a')
    )

    # Location heading — bold, slightly larger
    location_label_style = ParagraphStyle(
        'location_label',
        parent=styles['Normal'],
        fontSize=10,
        fontName='Helvetica-Bold',
        leading=13,          # Fix 3: tighter
        spaceBefore=0,
        spaceAfter=1
    )

    # Location detail rows — normal weight, compact
    location_text_style = ParagraphStyle(
        'location_text',
        parent=styles['Normal'],
        fontSize=9,
        leading=12          # Fix 3: match small style leading
    )

    # ── Per-row loop ─────────────────────────────────────────
    for _, row in df.iterrows():

        columns = df.columns

        # ── TITLE HEADER ─────────────────────────────────────
        title = clean(row.get("CLASS")) or sheet

        if logo_path and os.path.exists(logo_path):
            from PIL import Image as PILImage
            with PILImage.open(logo_path) as pil_img:
                orig_w, orig_h = pil_img.size
            # Max box: 1.2 x 0.45 inch — small and neat, fit inside, no stretch, no upscale
            max_w = 1.2 * inch
            max_h = 0.45 * inch
            ratio = min(max_w / orig_w, max_h / orig_h, 1.0)  # never upscale small images
            logo_w = orig_w * ratio
            logo_h = orig_h * ratio
            logo_cell = Image(logo_path, logo_w, logo_h)
        else:
            logo_cell = Paragraph("", small)

        header = Table(
            [[Paragraph(title.upper(), title_style), logo_cell]],
            colWidths=[PAGE_W - 2 * inch, 2 * inch]
        )

        header.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), colors.HexColor('#e8e8e8')),
            ('ALIGN',         (0, 0), (0,  0),  'LEFT'),
            ('ALIGN',         (1, 0), (1,  0),  'RIGHT'),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',   (0, 0), (-1, -1), 12),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 12),
            ('TOPPADDING',    (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            # thin bottom border for definition
            ('LINEBELOW',     (0, 0), (-1, -1), 1, colors.HexColor('#999999')),
        ]))

        elements.append(header)
        elements.append(Spacer(1, 14))

        # ── LOGIN BOXES ───────────────────────────────────────
        # Collect login data first, then render as a flat table
        # so we can control per-cell borders precisely (no nested tables)
        login_data = []

        for i in range(1, 4):
            user      = clean(row.get(f"Log On {i}"))
            excel_pwd = clean(row.get(f"Password {i}"))
            pwd       = default_password if default_password else excel_pwd
            if user:
                login_data.append((user, pwd or ''))

        if login_data:

            n_boxes = len(login_data)
            GAP     = 12    # visual gap between boxes (points)
            BOX_W   = 2.1 * inch   # fixed comfortable width per box

            if n_boxes == 1:
                # Single box — just its own width, left-aligned
                u, p   = login_data[0]
                row1   = [Paragraph(f"<b>User Name:</b> {u}", small)]
                row2   = [Paragraph(f"<b>Password:</b> {p}", small)]
                cw     = [BOX_W]
                style  = [
                    ('BOX',           (0, 0), (0, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BACKGROUND',    (0, 0), (0, -1), colors.HexColor('#fafafa')),
                    ('LEFTPADDING',   (0, 0), (-1, -1), 8),
                    ('RIGHTPADDING',  (0, 0), (-1, -1), 8),
                    ('TOPPADDING',    (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]

                login_table = Table([row1, row2], colWidths=cw, hAlign='LEFT')
                login_table.setStyle(TableStyle(style))

            elif n_boxes == 2:
                # Two boxes side by side, left-aligned, small gap between them
                u0, p0 = login_data[0]
                u1, p1 = login_data[1]
                GAP_COL = 10   # small gap between the two boxes
                cw     = [BOX_W, GAP_COL, BOX_W]
                row1   = [Paragraph(f"<b>User Name:</b> {u0}", small), Paragraph("", small), Paragraph(f"<b>User Name:</b> {u1}", small)]
                row2   = [Paragraph(f"<b>Password:</b> {p0}", small), Paragraph("", small), Paragraph(f"<b>Password:</b> {p1}", small)]

                login_table = Table([row1, row2], colWidths=cw, hAlign='LEFT')
                login_table.setStyle(TableStyle([
                    ('BOX',           (0, 0), (0, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BOX',           (2, 0), (2, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BACKGROUND',    (0, 0), (0, -1), colors.HexColor('#fafafa')),
                    ('BACKGROUND',    (2, 0), (2, -1), colors.HexColor('#fafafa')),
                    ('LEFTPADDING',   (0, 0), (0, -1), 8),
                    ('RIGHTPADDING',  (0, 0), (0, -1), 8),
                    ('LEFTPADDING',   (2, 0), (2, -1), 8),
                    ('RIGHTPADDING',  (2, 0), (2, -1), 8),
                    ('LEFTPADDING',   (1, 0), (1, -1), 0),
                    ('RIGHTPADDING',  (1, 0), (1, -1), 0),
                    ('TOPPADDING',    (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]))

            else:
                # Three boxes — equal width with a gap column between each
                u0, p0 = login_data[0]
                u1, p1 = login_data[1]
                u2, p2 = login_data[2]
                spacer_w = (PAGE_W - 3 * BOX_W) / 2
                cw     = [BOX_W, spacer_w, BOX_W, spacer_w, BOX_W]
                row1   = [
                    Paragraph(f"<b>User Name:</b> {u0}", small), Paragraph("", small),
                    Paragraph(f"<b>User Name:</b> {u1}", small), Paragraph("", small),
                    Paragraph(f"<b>User Name:</b> {u2}", small),
                ]
                row2   = [
                    Paragraph(f"<b>Password:</b> {p0}", small), Paragraph("", small),
                    Paragraph(f"<b>Password:</b> {p1}", small), Paragraph("", small),
                    Paragraph(f"<b>Password:</b> {p2}", small),
                ]

                login_table = Table([row1, row2], colWidths=cw, hAlign='LEFT')
                login_table.setStyle(TableStyle([
                    # Box only around cols 0, 2, 4 — cols 1 and 3 are spacers
                    ('BOX',           (0, 0), (0, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BOX',           (2, 0), (2, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BOX',           (4, 0), (4, -1), 0.75, colors.HexColor('#aaaaaa')),
                    ('BACKGROUND',    (0, 0), (0, -1), colors.HexColor('#fafafa')),
                    ('BACKGROUND',    (2, 0), (2, -1), colors.HexColor('#fafafa')),
                    ('BACKGROUND',    (4, 0), (4, -1), colors.HexColor('#fafafa')),
                    ('LEFTPADDING',   (0, 0), (0, -1), 8),
                    ('RIGHTPADDING',  (0, 0), (0, -1), 8),
                    ('LEFTPADDING',   (2, 0), (2, -1), 8),
                    ('RIGHTPADDING',  (2, 0), (2, -1), 8),
                    ('LEFTPADDING',   (4, 0), (4, -1), 8),
                    ('RIGHTPADDING',  (4, 0), (4, -1), 8),
                    ('LEFTPADDING',   (1, 0), (1, -1), 0),
                    ('RIGHTPADDING',  (1, 0), (1, -1), 0),
                    ('LEFTPADDING',   (3, 0), (3, -1), 0),
                    ('RIGHTPADDING',  (3, 0), (3, -1), 0),
                    ('TOPPADDING',    (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]))

            elements.append(login_table)
            elements.append(Spacer(1, 16))

        # ── PATIENT BLOCKS ────────────────────────────────────
        patients = detect_patients(columns)

        col_half = PAGE_W / 2   # each column = half page width

        for n in patients:

            mrn = clean(row.get(f"Patient {n} MRN"))

            if not mrn:
                continue

            fin   = clean(row.get(f"Patient {n} FIN"))
            last  = clean(row.get(f"Patient {n} Last Name"))
            first = clean(row.get(f"Patient {n} First Name"))

            name = f"{last or ''}, {first or ''}".strip().strip(',')

            # Patient label override
            if patient_labels and len(patient_labels) >= n and patient_labels[n - 1].strip():
                pat_label = patient_labels[n - 1].strip()
            else:
                pat_label = "PATIENT"

            location_fields = get_location_fields(row, n)

            # ── Left column: patient identity ─────────────────
            patient_info = Table([
                [Paragraph(f"<b>{pat_label.upper()}:</b> {name.upper()}", small)],
                [Paragraph(f"<b>MRN:</b> {mrn}", small)],
                [Paragraph(f"<b>FIN:</b> {fin or ''}", small)],
            ], colWidths=[col_half - 0.3 * inch])

            patient_info.setStyle(TableStyle([
                ('TOPPADDING',    (0, 0), (-1, -1), 2),   # Fix 3: tight rows
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('LEFTPADDING',   (0, 0), (-1, -1), 0),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ]))

            # ── Right column: location ─────────────────────────
            # Pair location fields two per line for compactness
            combined_lines = []

            for i in range(0, len(location_fields), 2):
                pair = location_fields[i:i + 2]
                line = ",  ".join([f"<b>{lbl}:</b> {val}" for lbl, val in pair])
                combined_lines.append([Paragraph(line, location_text_style)])

            location_rows = [[Paragraph("<b>LOCATION</b>", location_label_style)]]
            location_rows.extend(combined_lines)

            location_info = Table(location_rows, colWidths=[col_half - 0.3 * inch])

            location_info.setStyle(TableStyle([
                ('TOPPADDING',    (0, 0), (-1, -1), 2),   # Fix 3: tight rows
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('LEFTPADDING',   (0, 0), (-1, -1), 0),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ]))

            # ── Outer patient table ────────────────────────────
            patient_table = Table(
                [[patient_info, location_info]],
                colWidths=[col_half, col_half]
            )

            patient_table.setStyle(TableStyle([
                ('BOX',           (0, 0), (-1, -1), 1,   colors.black),
                # vertical divider between the two columns
                ('LINEBEFORE',    (1, 0), (1, -1),  0.5, colors.HexColor('#cccccc')),
                ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING',   (0, 0), (-1, -1), 12),
                ('RIGHTPADDING',  (0, 0), (-1, -1), 12),
                ('TOPPADDING',    (0, 0), (-1, -1), 8),   # Fix 3: was 10
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))

            elements.append(patient_table)
            elements.append(Spacer(1, 6))

        # ── WRISTBAND QR ──────────────────────────────────────
        wrist_headers = []
        wrist_qr      = []

        for n in patients:

            visit = clean(row.get(f"Visit ID{n}"))

            if visit:

                # Clean wristband label — strip trailing digits/spaces to avoid "Barcode1 1"
                wb_label = (wristband_label or "Patient Wristband").strip()

                wrist_headers.append(
                    Paragraph(f"{wb_label} {n}", bold_center)
                )

                qr = create_qr(visit)

                wrist_qr.append(Image(qr, 1.2 * inch, 1.2 * inch))

        if wrist_headers:

            elements.append(Spacer(1, 16))

            n_wrist   = len(wrist_headers)
            wrist_col = PAGE_W / n_wrist

            wrist_table = Table(
                [wrist_headers, wrist_qr],
                colWidths=[wrist_col] * n_wrist
            )

            wrist_table.setStyle(TableStyle([
                ('GRID',          (0, 0), (-1, -1), 0.5, colors.HexColor('#aaaaaa')),
                ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
                ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor('#f5f5f5')),
                ('TOPPADDING',    (0, 0), (-1, 0),  6),
                ('BOTTOMPADDING', (0, 0), (-1, 0),  6),
                ('TOPPADDING',    (0, 1), (-1, 1),  10),
                ('BOTTOMPADDING', (0, 1), (-1, 1),  10),
            ]))

            elements.append(wrist_table)

        # ── MEDICATION QR ─────────────────────────────────────
        meds = []

        for i in range(1, 15):

            med_name = clean(row.get(f"Product{i}"))
            ndc      = clean(row.get(f"NDC{i}"))

            if med_name and ndc:

                qr     = create_qr(ndc)
                qr_img = Image(qr, 0.85 * inch, 0.85 * inch)   # Fix 5: smaller QR

                cell = Table([
                    [Paragraph(f"{med_name}", small)],
                    [qr_img]
                ], colWidths=[PAGE_W / 2 - 0.2 * inch], rowHeights=[0.45 * inch, 1.0 * inch])

                cell.setStyle(TableStyle([
                    ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
                    # Fix 4: zero left/right padding so separator line touches borders
                    ('LINEBELOW',     (0, 0), (-1,  0), 0.75, colors.HexColor('#555555')),
                    ('LEFTPADDING',   (0, 0), (-1, -1), 0),
                    ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
                    # Fix 5: tighter vertical padding around name and QR
                    ('TOPPADDING',    (0, 0), (-1,  0), 5),
                    ('BOTTOMPADDING', (0, 0), (-1,  0), 3),
                    ('TOPPADDING',    (0, 1), (-1,  1), 4),
                    ('BOTTOMPADDING', (0, 1), (-1,  1), 4),
                ]))

                meds.append(cell)

        if meds:

            elements.append(Spacer(1, 16))
            elements.append(Paragraph("<b>Medications</b>", bold_center))
            elements.append(Spacer(1, 6))

            grid  = []
            row_g = []

            for m in meds:
                row_g.append(m)
                if len(row_g) == 2:
                    grid.append(row_g)
                    row_g = []

            if row_g:
                while len(row_g) < 2:
                    row_g.append(Paragraph("", small))
                grid.append(row_g)

            med_table = Table(grid, colWidths=[PAGE_W / 2, PAGE_W / 2])

            # Only draw borders around cells that have content
            n_rows = len(grid)
            n_cols = 2
            med_style = [
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN',(0, 0), (-1, -1), 'MIDDLE'),
            ]
            for r, grow in enumerate(grid):
                for c, cell in enumerate(grow):
                    if hasattr(cell, 'text') or (hasattr(cell, '__class__') and cell.__class__.__name__ == 'Table'):
                        med_style.append(('BOX', (c, r), (c, r), 0.5, colors.HexColor('#aaaaaa')))
                    elif isinstance(cell, str) and cell == "":
                        pass  # no border on empty string padding cells
                    else:
                        med_style.append(('BOX', (c, r), (c, r), 0.5, colors.HexColor('#aaaaaa')))
            med_table.setStyle(TableStyle(med_style))

            elements.append(med_table)

        # ── SPECIMEN QR ───────────────────────────────────────
        specimens = []

        for i in range(1, 10):

            specimen = clean(row.get(f"Specimen{i}"))
            barcode  = clean(row.get(f"Barcode{i}"))

            if specimen and barcode:

                qr     = create_qr(barcode)
                qr_img = Image(qr, 0.85 * inch, 0.85 * inch)   # Fix 5: smaller QR

                cell = Table([
                    [Paragraph(f"{specimen}", small)],
                    [qr_img]
                ], colWidths=[PAGE_W / 2 - 0.2 * inch], rowHeights=[0.45 * inch, 1.0 * inch])

                cell.setStyle(TableStyle([
                    ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
                    # Fix 4: zero left/right padding so separator line touches borders
                    ('LINEBELOW',     (0, 0), (-1,  0), 0.75, colors.HexColor('#555555')),
                    ('LEFTPADDING',   (0, 0), (-1, -1), 0),
                    ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
                    # Fix 5: tighter vertical padding around name and QR
                    ('TOPPADDING',    (0, 0), (-1,  0), 5),
                    ('BOTTOMPADDING', (0, 0), (-1,  0), 3),
                    ('TOPPADDING',    (0, 1), (-1,  1), 4),
                    ('BOTTOMPADDING', (0, 1), (-1,  1), 4),
                ]))

                specimens.append(cell)

        if specimens:

            elements.append(Spacer(1, 16))
            elements.append(Paragraph("<b>Specimens</b>", bold_center))
            elements.append(Spacer(1, 6))

            grid  = []
            row_g = []

            for s in specimens:
                row_g.append(s)
                if len(row_g) == 2:
                    grid.append(row_g)
                    row_g = []

            if row_g:
                while len(row_g) < 2:
                    row_g.append(Paragraph("", small))
                grid.append(row_g)

            spec_table = Table(grid, colWidths=[PAGE_W / 2, PAGE_W / 2])

            spec_style = [
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN',(0, 0), (-1, -1), 'MIDDLE'),
            ]
            for r, grow in enumerate(grid):
                for c, cell in enumerate(grow):
                    if hasattr(cell, '__class__') and cell.__class__.__name__ == 'Table':
                        spec_style.append(('BOX', (c, r), (c, r), 0.5, colors.HexColor('#aaaaaa')))
                    elif isinstance(cell, str) and cell == "":
                        pass
                    else:
                        spec_style.append(('BOX', (c, r), (c, r), 0.5, colors.HexColor('#aaaaaa')))
            spec_table.setStyle(TableStyle(spec_style))

            elements.append(spec_table)

        elements.append(PageBreak())

    doc.build(elements)

    print("PDF Generated:", filename)

    return filename


# -----------------------------
# MAIN
# -----------------------------
if __name__ == "__main__":

    excel_file = "input.xlsx"

    xls = pd.ExcelFile(excel_file)

    for sheet in xls.sheet_names:

        df = pd.read_excel(xls, sheet_name=sheet)

        build_pdf(df, sheet)
