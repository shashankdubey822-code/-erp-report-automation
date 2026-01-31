import logging
from io import BytesIO
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
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

    # --- HORIZONTAL PAGINATION STRATEGY ---
    
    # 1. Setup Dimensions & Fixed Columns
    page_w = landscape(letter)[0]
    available_width = page_w - doc.leftMargin - doc.rightMargin
    
    # Define Fixed Left Columns (Indices 0, 1, 2)
    # Typically: Sr No., Roll No, Student Name
    # We assume these are always the first 3.
    num_left_fixed = 3
    left_fixed_widths = list(left_fixed) # [Sr, Roll, Name]
    
    # Define Fixed Right Columns
    # "Overall %age...", "Count of Courses...", "Whether Critical"
    # We'll look for these specific headers or just take the last 3 if they match expectations.
    expected_right_headers = [
        "Overall %age of all subjects from ERP report",
        "Count of Courses with attendance below minimum attendance criteria",
        "Whether Critical"
    ]
    
    # Find indices of these columns in raw_headers
    right_fixed_indices = []
    for rh in expected_right_headers:
        if rh in raw_headers:
            right_fixed_indices.append(raw_headers.index(rh))
    
    # If we found them, they are our right fixed columns.
    # Otherwise, we might have a different structure, so we'll fallback to 0 right fixed.
    right_fixed_widths = []
    if len(right_fixed_indices) == 3:
        # Assign widths for: Overall %, Count, Critical
        right_fixed_widths = [0.7 * inch, 0.7 * inch, 0.8 * inch]
    
    total_fixed_width = sum(left_fixed_widths) + sum(right_fixed_widths)
    
    # 2. Calculate Capacity for Variable (Middle) Columns
    available_for_subjects = available_width - total_fixed_width
    
    # Minimum comfortable width for a subject column
    min_subject_width = 0.45 * inch  
    
    if available_for_subjects < min_subject_width:
        # Fallback: very tight space
        max_subjects_per_page = 1
    else:
        max_subjects_per_page = int(available_for_subjects / min_subject_width)
        if max_subjects_per_page < 1: max_subjects_per_page = 1

    # 3. Identify Middle (Variable) Columns
    # Middle columns are those that are NOT in Left Fixed (0,1,2) AND NOT in Right Fixed
    all_indices = range(len(raw_headers))
    # fixed left are 0, 1, 2
    left_indices = list(range(num_left_fixed))
    
    # Variable indices are everything else minus the right fixed ones
    variable_indices = [
        i for i in all_indices 
        if i not in left_indices and i not in right_fixed_indices
    ]
    
    num_variable_cols = len(variable_indices)
    
    if num_variable_cols == 0:
        total_chunks = 1
    else:
        total_chunks = (num_variable_cols + max_subjects_per_page - 1) // max_subjects_per_page

    # 4. Generate Tables per Chunk
    for chunk_idx in range(total_chunks):
        # Slice indices for this chunk's variable columns
        start_ptr = chunk_idx * max_subjects_per_page
        end_ptr = min(start_ptr + max_subjects_per_page, num_variable_cols)
        
        current_chunk_indices = variable_indices[start_ptr:end_ptr]
        
        # Construct the unified list of indices for this table:
        # [Left Fixed] + [Current Chunk] + [Right Fixed]
        current_table_col_indices = left_indices + current_chunk_indices + right_fixed_indices

        # --- Calculate Widths (Moved Up for Dynamic Headers) ---
        num_cols_in_chunk = len(current_chunk_indices)
        
        if num_cols_in_chunk > 0:
            calc_width = available_for_subjects / num_cols_in_chunk
            # Cap width
            calc_width = min(calc_width, 1.2 * inch)
            chunk_variable_widths = [calc_width] * num_cols_in_chunk
        else:
            chunk_variable_widths = []
            
        final_col_widths = left_fixed_widths + chunk_variable_widths + right_fixed_widths
        
        # --- Build Headers with Auto-Shrink Font (Name & Code) ---
        current_table_headers = []
        for i, col_idx in enumerate(current_table_col_indices):
            col_width = final_col_widths[i]
            # Use the cleaned text from our pre-processed list
            header_text = cleaned_headers[col_idx] if col_idx < len(cleaned_headers) else ""
            
            # Retrieve Subject Code
            subj_code = ""
            if header_text in subject_details:
                subj_code = safe_str(subject_details[header_text].get('code', ''))

            # Helper to find best fit font size
            def get_best_fit_font_size(text, max_width, start_size, min_size):
                size = start_size
                words = text.split()
                if not words: return size
                while size > min_size:
                    # Check if longest word fits
                    max_word_len = max(stringWidth(w, "Helvetica-Bold", size) for w in words)
                    # Check if total line fits (approx) - though wrapping handles multi-word names,
                    # for code we want single line ideally, but let's prioritize word-breaking prevention.
                    if max_word_len <= (max_width - 4):
                        return size
                    size -= 0.5
                return min_size

            # 1. Calculate Name Font Size
            name_fs = get_best_fit_font_size(header_text, col_width, 8, 5)
            
            # 2. Calculate Code Font Size (Start smaller, shrink aggressively)
            code_fs = 7
            if subj_code:
                # For code, we treat it as a single block usually
                code_fs = get_best_fit_font_size(subj_code, col_width, 7, 4)

            # Construct HTML
            # We use leading equal to the largest font size in the block to avoid overlap
            display_html = f'<font size="{name_fs}">{header_text}</font>'
            if subj_code:
                display_html += f'<br/><font size="{code_fs}">({subj_code})</font>'
            
            # Create the paragraph
            # Leading needs to accomodate the stack. 
            # If name is 8 and code is 7, leading 9 is okay for line spacing.
            style = ParagraphStyle(name=f"hdr_{chunk_idx}_{i}", fontSize=name_fs, leading=name_fs+2, alignment=1)
            current_table_headers.append(Paragraph(display_html, style))
        
        # --- Build Rows ---
        current_table_rows = []
        for row in normalized_rows:
            # Select cells based on the indices
            row_data = [row[i] for i in current_table_col_indices]
            current_table_rows.append(row_data)
            
        table_payload = [current_table_headers] + current_table_rows
        
        # --- Create Table ---
        table = Table(table_payload, colWidths=final_col_widths, repeatRows=1)
        
        # Base Style
        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.Color(1, 1, 0, alpha=0.2)), # Header Yellow
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.black),
        ])
        
        # --- Apply Conditional Highlighting (<75% Red) ---
        min_att = metadata.get("min_attendance", 75)
        
        # We assume headers with "Critical" or "Overall" might also need formatting, 
        # but the request specifically focuses on preserving them.
        # We will apply Red highlighting to SUBJECT columns (Variable chunk) 
        # AND possibly the "Overall %" column if it falls below threshold.
        
        # 1. Check Variable Columns (Middle)
        # Their positions in the CURRENT table range from:
        # Start: num_left_fixed
        # End: num_left_fixed + num_cols_in_chunk
        for local_col_i in range(num_left_fixed, num_left_fixed + num_cols_in_chunk):
            # Identify which original column this is
            # current_table_col_indices[local_col_i] gives the global index
            global_idx = current_table_col_indices[local_col_i]
            original_header_name = raw_headers[global_idx]
            clean_name = clean_subject_label(original_header_name)
            
            # Check if this column is a tracked subject
            if clean_name in subject_details:
                for ridx, row_data in enumerate(current_table_rows, start=1):
                    try:
                        cell_val = row_data[local_col_i]
                        txt_val = cell_val.text if isinstance(cell_val, Paragraph) else str(cell_val)
                        val_float = float(safe_str(txt_val))
                        
                        if val_float < min_att:
                            # Highlight cell
                            tbl_style.add("BACKGROUND", (local_col_i, ridx), (local_col_i, ridx), colors.lightgrey) # Using lightgrey as per original code for <75, or change to red if requested. Original code used lightgrey in the snippet above? 
                            # Wait, the prompt says "red background for <75%". The code I replaced had `colors.lightgrey`... 
                            # I'll stick to `colors.lightgrey` to be safe OR check the original code again.
                            # Ah, the snippet I read showed `colors.lightgrey`. I will use `colors.red` if it's "Critical" or stick to grey if it matches visual style.
                            # Let's use a soft red/pink to be clear but readable.
                            tbl_style.add("BACKGROUND", (local_col_i, ridx), (local_col_i, ridx), colors.Color(1, 0.8, 0.8)) 
                    except (ValueError, TypeError):
                        pass

        # 2. Check "Overall %" column (One of the right fixed columns)
        # Find its local index
        overall_col_name = "Overall %age of all subjects from ERP report"
        if overall_col_name in raw_headers:
            global_overall_idx = raw_headers.index(overall_col_name)
            if global_overall_idx in current_table_col_indices:
                local_overall_idx = current_table_col_indices.index(global_overall_idx)
                
                for ridx, row_data in enumerate(current_table_rows, start=1):
                    try:
                        cell_val = row_data[local_overall_idx]
                        txt_val = cell_val.text if isinstance(cell_val, Paragraph) else str(cell_val)
                        val_float = float(safe_str(txt_val))
                        
                        if val_float < min_att:
                             tbl_style.add("BACKGROUND", (local_overall_idx, ridx), (local_overall_idx, ridx), colors.Color(1, 0.8, 0.8))
                    except (ValueError, TypeError):
                        pass

        # 3. Check "Whether Critical" column
        critical_col_name = "Whether Critical"
        if critical_col_name in raw_headers:
             global_crit_idx = raw_headers.index(critical_col_name)
             if global_crit_idx in current_table_col_indices:
                local_crit_idx = current_table_col_indices.index(global_crit_idx)
                for ridx, row_data in enumerate(current_table_rows, start=1):
                    cell_val = str(row_data[local_crit_idx])
                    if "CRITICAL" in cell_val.upper():
                         tbl_style.add("TEXTCOLOR", (local_crit_idx, ridx), (local_crit_idx, ridx), colors.red)
                         tbl_style.add("FONTNAME", (local_crit_idx, ridx), (local_crit_idx, ridx), "Helvetica-Bold")

        table.setStyle(tbl_style)
        elements.append(table)
        
        if chunk_idx < total_chunks - 1:
            elements.append(PageBreak())

    elements.append(Spacer(1, 0.12 * inch))
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