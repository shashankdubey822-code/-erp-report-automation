import logging
from io import BytesIO
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from .utilities import safe_str, clean_subject_label, format_attendance_value

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# -------------------------
# PDF generation (reportlab)
# -------------------------
def create_pdf_file(df: pd.DataFrame, subject_details: dict, metadata: dict, chart_image: BytesIO = None, subjects_with_zero_attendance: list = None) -> BytesIO:
    """
    Create a well-formatted PDF using reportlab.
    - Single header row (subject names only)
    - Dynamic column width calculation to avoid overlap
    - Header wrapping and auto-shrink by using modest font sizes
    - Summary table appended
    - Includes a message about subjects with 0% attendance if provided.
    Returns a BytesIO buffer containing PDF bytes.
    """
    # defensive imports
    from reportlab.lib.pagesizes import landscape
    styles = getSampleStyleSheet()

    # Prepare doc
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter),
                            leftMargin=0.35 * inch, rightMargin=0.35 * inch,
                            topMargin=0.35 * inch, bottomMargin=0.35 * inch)
    elements = []

    # Title block
    title_style = ParagraphStyle("title", parent=styles["Title"], alignment=1, fontSize=14, leading=16)
    title_text = (
        f"{safe_str(metadata.get('department_name', 'DEPARTMENT')).upper()}<br/>"
        f"{safe_str(metadata.get('date_range', '')).upper()}<br/>"
        f"{safe_str(metadata.get('report_title', 'ATTENDANCE REPORT')).upper()}"
    )
    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 0.12 * inch))

    # Metadata
    meta_style = ParagraphStyle("meta", parent=styles["Normal"], fontSize=9, leading=11)
    meta_block = (
        f"<b>Branch:</b> {safe_str(metadata.get('branch', 'N/A'))} &nbsp; "
        f"<b>Department:</b> {safe_str(metadata.get('department_specialization', 'N/A'))}<br/>"
        f"<b>Class:</b> {safe_str(metadata.get('class_name_division', 'N/A'))} &nbsp; "
        f"<b>Division:</b> {safe_str(metadata.get('division', 'N/A'))} &nbsp; "
        f"<b>Date:</b> {safe_str(metadata.get('date_range', 'N/A'))} &nbsp; "
        f"<b>Coordinator:</b> {safe_str(metadata.get('coordinator', ''))}"
    )
    elements.append(Paragraph(meta_block, meta_style))
    elements.append(Spacer(1, 0.12 * inch))

    # Add message for subjects with 0% attendance
    if subjects_with_zero_attendance:
        zero_att_message = "The following subjects have 0% attendance for all students: " + \
                           ", ".join([f"<b>{s}</b>" for s in subjects_with_zero_attendance]) + \
                           ". These subjects are not included in the main table."
        warning_style = ParagraphStyle("warning", parent=styles["Normal"], textColor=colors.red, fontSize=9, leading=11, spaceAfter=6)
        elements.append(Paragraph(zero_att_message, warning_style))
        elements.append(Spacer(1, 0.12 * inch))

    # Build headers
    raw_headers = list(df.columns)
    cleaned_headers = [clean_subject_label(h) for h in raw_headers]
    hdr_style = ParagraphStyle("hdr", fontSize=8, leading=9, alignment=1)
    wrapped_headers = [Paragraph(h or "", hdr_style) for h in cleaned_headers]
    
    num_cols = len(cleaned_headers)

    # 1. Define the "invisible boundary" trigger
    COLUMN_THRESHOLD = 15
    apply_aggressive_wrapping = num_cols > COLUMN_THRESHOLD

    # 2. Dynamically set the width of the name column
    name_col_idx = 2
    try:
        header_texts = [h.text.strip() for h in wrapped_headers]
        name_col_idx = header_texts.index("Student Name")
    except ValueError:
        pass

    if apply_aggressive_wrapping:
        name_col_width = 0.8 * inch
    else:
        name_col_width = 1.65 * inch
    
    left_fixed = [0.45 * inch, 1.15 * inch, 1.65 * inch]  # Default widths
    left_fixed[name_col_idx] = name_col_width  # Overwrite with dynamic width

    # 3. Normalize data rows with selective wrapping
    left_align_style = ParagraphStyle("data_cell_left", parent=styles["Normal"], fontSize=8, leading=9, alignment=0)
    normalized_rows = []

    for row in df.values.tolist():
        row_list = []
        for i, cell_text in enumerate(list(row)):
            if i == name_col_idx:
                formatted_text = safe_str(cell_text)
                if apply_aggressive_wrapping:
                    formatted_text = formatted_text.replace(' ', '<br/>')
                row_list.append(Paragraph(formatted_text, left_align_style))
            elif i > name_col_idx: # Apply formatting to columns after Student Name (likely attendance)
                row_list.append(safe_str(format_attendance_value(cell_text)))
            else:
                row_list.append(safe_str(cell_text))
        
        if len(row_list) < num_cols:
            row_list += [""] * (num_cols - len(row_list))
        elif len(row_list) > num_cols:
            row_list = row_list[:num_cols]
        normalized_rows.append(row_list)

    table_data = [wrapped_headers] + normalized_rows

    # Column width engine
    page_w = landscape(letter)[0]
    available = page_w - doc.leftMargin - doc.rightMargin

    right_fixed_count = 4 if num_cols >= (len(left_fixed) + 4) else max(0, num_cols - len(left_fixed) - len(subject_details))
    right_fixed_width = 0.8 * inch

    num_subject_cols = max(1, num_cols - (len(left_fixed) + right_fixed_count))
    remaining_width = available - sum(left_fixed) - (right_fixed_count * right_fixed_width)
    subject_w = max(0.42 * inch, remaining_width / max(1, num_subject_cols))
    
    col_widths = list(left_fixed)
    col_widths += [subject_w] * num_subject_cols
    col_widths += [right_fixed_width] * right_fixed_count

    if len(col_widths) < num_cols:
        col_widths += [subject_w] * (num_cols - len(col_widths))
    col_widths = col_widths[:num_cols]

    # Build table
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(1, 1, 0, alpha=0.2)),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
    ])

    # 4. Update highlighting logic
    min_att = metadata.get("min_attendance", 75)
    for subj, col_name in subject_details.items():
        possible_cols = [c for c in raw_headers if str(c).startswith(subj) or clean_subject_label(c) == subj]
        for pc in possible_cols:
            if pc not in raw_headers: continue
            col_idx = raw_headers.index(pc)
            
            for ridx, row in enumerate(normalized_rows, start=1):
                try:
                    cell_content = row[col_idx]
                    val_str = cell_content.text if isinstance(cell_content, Paragraph) else cell_content
                    val = float(safe_str(val_str))
                    if val < min_att:
                        tbl_style.add("BACKGROUND", (col_idx, ridx), (col_idx, ridx), colors.lightgrey)
                except (ValueError, IndexError):
                    pass

    table.setStyle(tbl_style)
    elements.append(table)
    elements.append(Spacer(1, 0.12 * inch))

    # Page break after main table
    elements.append(PageBreak())

    # Summary table
    subjects = list(subject_details.keys())
    valid_subjects = [s for s in subjects if s in df.columns or any(str(c).startswith(s) for c in df.columns)]
    summary_headers = ["Subject", "Students < 75%", "Students < 70%", "Students < 65%", "Students < 60%"]
    summary_rows = [summary_headers]
    for s in valid_subjects:
        col_key = None
        for c in df.columns:
            if str(c).startswith(s) or clean_subject_label(c) == s:
                col_key = c
                break
        if col_key is None:
            continue
        summary_rows.append([
            s,
            int((pd.to_numeric(df[col_key], errors="coerce") < 75).sum()),
            int((pd.to_numeric(df[col_key], errors="coerce") < 70).sum()),
            int((pd.to_numeric(df[col_key], errors="coerce") < 65).sum()),
            int((pd.to_numeric(df[col_key], errors="coerce") < 60).sum()),
        ])

    sum_hdr = ParagraphStyle("sumhdr", fontSize=8, alignment=1)
    sum_sub = ParagraphStyle("sumsub", fontSize=8, alignment=0)
    wrapped_summary = []
    for i, row in enumerate(summary_rows):
        new_row = []
        for j, cell in enumerate(row):
            if j == 0:
                if i == 0:
                    new_row.append(Paragraph(str(cell), sum_hdr))
                else:
                    new_row.append(Paragraph(str(cell), sum_sub))
            else:
                new_row.append(str(cell))
        wrapped_summary.append(new_row)

    summary_col_widths = [
        4.0 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch
    ]

    summary_table = Table(wrapped_summary, colWidths=summary_col_widths, hAlign="LEFT")
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.12 * inch))

    # Page break after summary table
    elements.append(PageBreak())

    # Optional chart image
    if chart_image:
        try:
            rl_img = RLImage(chart_image)
            
            img_w, img_h = rl_img.imageWidth, rl_img.imageHeight
            if img_w <= 0 or img_h <= 0:
                raise ValueError("Invalid image dimensions")

            aspect = img_h / float(img_w)
            new_w = doc.width
            new_h = new_w * aspect

            if new_h > doc.height:
                new_h = doc.height
                new_w = new_h / aspect
            
            rl_img.drawWidth = new_w
            rl_img.drawHeight = new_h
            rl_img.hAlign = 'CENTER'
            elements.append(rl_img)
        except Exception:
            logger.exception("Failed to attach chart image to PDF")

    # Build and return buffer
    doc.build(elements)
    buffer.seek(0)
    return buffer