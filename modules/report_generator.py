import pandas as pd
from io import BytesIO, StringIO
import re
import csv
import logging
from typing import Dict, Tuple, Any

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Advanced imports for styling, charts, and graph label formatting
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.label import DataLabelList

# Professional styling constants
PROFESSIONAL_STYLES = {
    'thin_border': Border(
        left=Side(border_style='thin', color='000000'),
        right=Side(border_style='thin', color='000000'),
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000')
    ),
    'medium_border': Border(
        left=Side(border_style='medium', color='000000'),
        right=Side(border_style='medium', color='000000'),
        top=Side(border_style='medium', color='000000'),
        bottom=Side(border_style='medium', color='000000')
    ),
    'header_fill': PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
    'category_fill': PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid"),
    'data_fill': PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
    'critical_fill': PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid"),
    'header_font': Font(color="FFFFFF", bold=True, size=10),
    'category_font': Font(color="000000", bold=True, size=10),
    'data_font': Font(color="000000", size=9),
    'center_align': Alignment(horizontal='center', vertical='center', wrap_text=True)
}

# Subject categorization for better organization
SUBJECT_CATEGORIES = {
    'CORE_SUBJECTS': ['COMPUTER', 'PROGRAMMING', 'DATA', 'SYSTEM', 'SOFTWARE'],
    'MATHEMATICS': ['MATHEMATICS', 'STATISTICS', 'PROBABILITY', 'CALCULUS'],
    'ENGINEERING': ['ENGINEERING', 'MECHANICS', 'PHYSICS', 'QUANTUM'],
    'COMMUNICATION': ['COMMUNICATION', 'ENGLISH', 'TECHNICAL WRITING'],
    'MANAGEMENT': ['MANAGEMENT', 'ECONOMICS', 'BUSINESS'],
    'SCIENCE': ['ENVIRONMENTAL', 'CHEMISTRY', 'BIOLOGY'],
    'RESEARCH': ['RESEARCH', 'INNOVATION', 'PROJECT']
}

def categorize_subject(subject_name):
    """Categorize subjects based on name patterns"""
    subject_upper = subject_name.upper()
    
    for category, keywords in SUBJECT_CATEGORIES.items():
        for keyword in keywords:
            if keyword in subject_upper:
                return category.replace('_', ' & ')
    
    return 'OTHER SUBJECTS'

def create_report_dataframe(erp_file, min_attendance_criteria, filename):
    """
    This is the definitive, unified function with a robust parser
    that correctly handles all file types and special characters like '&'.
    """
    erp_file.seek(0)
    file_content = erp_file.read().decode('utf-8', errors='ignore')
    
    # Use StringIO to treat the string content like a file, which is memory efficient
    string_io_file = StringIO(file_content)
    
    # --- UPGRADED PARSER: Use Python's built-in CSV reader for robust parsing ---
    # This correctly handles any commas or special characters inside quoted subject names
    reader = csv.reader(string_io_file)
    lines_as_lists = list(reader)
    # Re-join the correctly parsed rows into strings for the existing logic
    lines = [",".join(row) for row in lines_as_lists]
    
    # --- From this point on, the logic remains the same, but it's now working with clean data ---
    header_start_index = -1
    
    # Try multiple possible header patterns for flexibility
    header_patterns = [
        'Sr.,Division/Section,Unique id',
        'Sr,Division/Section,Unique id', 
        'Sr.,Division,Unique id',
        'Unique id',
        'Roll',
        'Student Name'
    ]
    
    for i, line in enumerate(lines):
        line_clean = line.replace('"', '').replace(' ', '')
        for pattern in header_patterns:
            pattern_clean = pattern.replace(' ', '')
            if pattern_clean in line_clean:
                header_start_index = i
                logger.info(f"Found header at line {i} with pattern: {pattern}")
                break
        if header_start_index != -1:
            break
    
    if header_start_index == -1:
        logger.error(f"Could not find data table header. Searched for patterns: {header_patterns}")
        logger.error(f"First 10 lines of file: {lines[:10]}")
        raise ValueError("Could not find the data table header in the ERP file. Please check the file format.")
    
    # This logic now works correctly because 'lines' was parsed by the csv reader
    try:
        h1_subjects = [h.strip() for h in lines_as_lists[header_start_index]]
        
        # Safely get subsequent header rows with bounds checking
        max_lines = len(lines_as_lists)
        h2_codes = [h.strip() for h in lines_as_lists[header_start_index + 2]] if header_start_index + 2 < max_lines else ['']
        h3_types = [h.strip() for h in lines_as_lists[header_start_index + 3]] if header_start_index + 3 < max_lines else ['']
        h_metrics = [h.strip() for h in lines_as_lists[header_start_index + 4]] if header_start_index + 4 < max_lines else ['']
        
        logger.info(f"Parsed headers - Subjects: {len(h1_subjects)}, Codes: {len(h2_codes)}, Types: {len(h3_types)}, Metrics: {len(h_metrics)}")
        
    except IndexError as e:
        logger.error(f"Error parsing headers at index {header_start_index}: {e}")
        raise ValueError(f"Invalid file format: Unable to parse headers starting at line {header_start_index + 1}")

    last_subject, last_code, last_type = "", "", ""
    for i, subject in enumerate(h1_subjects):
        if subject: last_subject = subject
        else: h1_subjects[i] = last_subject
        if h2_codes[i]: last_code = h2_codes[i]
        else: h2_codes[i] = last_code
        if h3_types[i]: last_type = h3_types[i]
        else: h3_types[i] = last_type

    final_headers, subject_details = [], {}
    for i, metric in enumerate(h_metrics):
        subject, code, type = h1_subjects[i], h2_codes[i], h3_types[i]
        if subject in ['Sr.', 'Division/Section', 'Unique id', 'Rollno', 'Student Name', 'PRN / Enroll']:
            final_headers.append(subject)
        elif 'Total' in subject:
            final_headers.append(f"Grand Total - {metric}")
        else:
            final_headers.append(f"{subject} - {metric}")
            if subject not in subject_details: subject_details[subject] = {'code': code, 'type': type}
    
    # We must re-encode the data to pass it to pandas
    data_as_string_list = [",".join(row) for row in lines_as_lists[header_start_index + 6:]]
    csv_data_string = "\n".join(data_as_string_list)
    
    df = pd.read_csv(StringIO(csv_data_string), header=None, names=final_headers, on_bad_lines='skip')
    df.dropna(subset=['Rollno'], inplace=True)
    
    output_df = pd.DataFrame({'Sr No.': range(1, len(df) + 1), 'Roll No': df['Rollno'], 'Student Name': df['Student Name']})
    
    subject_percent_cols = {}
    logger.info(f"Processing {len(subject_details)} subjects: {list(subject_details.keys())}")
    logger.info(f"Available columns: {list(df.columns)}")
    
    for subject in subject_details.keys():
        found_column = None
        # Try multiple possible suffix patterns including TUT type
        possible_suffixes = [' - Total %', ' - % (PP)', ' - % (PR)', ' - % (TUT)', ' - %', ' - Total', '- Total %', '- % (PP)', '- % (PR)', '- % (TUT)']
        
        for suffix in possible_suffixes:
            potential_col = f"{subject}{suffix}"
            if potential_col in df.columns:
                found_column = potential_col
                break
        
        if found_column:
            subject_percent_cols[subject] = found_column
            logger.info(f"Found column for {subject}: {found_column}")
        else:
            logger.warning(f"No matching column found for subject: {subject}")
            logger.warning(f"Looked for: {[f'{subject}{s}' for s in possible_suffixes]}")
            # For missing subjects, add them with 0 values so they still appear in the report
            output_df[subject] = 0
            logger.info(f"Added {subject} with default 0 values")
    
    # Only process subjects that have matching columns
    for subject, col_name in subject_percent_cols.items():
        try:
            output_df[subject] = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
        except KeyError as e:
            logger.error(f"KeyError processing subject {subject} with column {col_name}: {e}")
            # Create a default column with 0 values
            output_df[subject] = 0
    
    output_df['Overall %age of all subjects from ERP report'] = pd.to_numeric(df.get('Grand Total - %', df.get('Total - %', 0)), errors='coerce').fillna(0)
    output_df['Roll No_duplicate'] = df['Rollno'] 
    
    # Only use subjects that actually have data columns
    subject_keys = list(subject_percent_cols.keys())
    logger.info(f"Using {len(subject_keys)} subjects for attendance calculation: {subject_keys}")
    
    if subject_keys:
        # Safely calculate attendance with only existing columns
        try:
            output_df['Count of Courses with attendance below minimum attendance criteria'] = output_df[subject_keys].apply(lambda row: (row < min_attendance_criteria).sum(), axis=1)
        except KeyError as e:
            logger.error(f"KeyError in attendance calculation: {e}")
            logger.error(f"Subject keys: {subject_keys}")
            logger.error(f"Available columns: {list(output_df.columns)}")
            # Fallback: set count to 0 for all rows
            output_df['Count of Courses with attendance below minimum attendance criteria'] = 0
    else:
        logger.warning("No subject columns found for attendance calculation")
        output_df['Count of Courses with attendance below minimum attendance criteria'] = 0
    
    output_df['Whether Critical'] = output_df['Count of Courses with attendance below minimum attendance criteria'].apply(lambda count: 'CRITICAL' if count > 4 else 'Not Critical')
    
    return output_df, subject_details

def create_excel_file(df, subject_details, metadata):
    """
    Creates Excel report matching the exact ERP format from the uploaded image.
    """
    df.rename(columns={'Roll No_duplicate': 'Roll No'}, inplace=True)
    output_buffer = BytesIO()
    
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        sheet_name = metadata['monitoring_stage']
        worksheet = writer.book.create_sheet(title=sheet_name)
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])
        
        # Create exact ERP format matching the uploaded image
        _create_erp_format_report(worksheet, df, subject_details, metadata)

    output_buffer.seek(0)
    return output_buffer

def _create_erp_format_report(worksheet, df, subject_details, metadata):
    """Create exact ERP format matching the uploaded image"""
    
    # Define exact colors and styles from the image
    YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    RED_FILL = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    THIN_BORDER = Border(
        left=Side(border_style='thin', color='000000'),
        right=Side(border_style='thin', color='000000'),
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000')
    )
    
    BLACK_FONT = Font(color="000000", size=9, bold=False)
    BOLD_FONT = Font(color="000000", size=9, bold=True)
    CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Header section (rows 1-10)
    _create_erp_header_section(worksheet, metadata, YELLOW_FILL, RED_FILL, BLACK_FONT, BOLD_FONT, CENTER_ALIGN, THIN_BORDER)
    
    # Data headers (rows 11-15)
    header_start_row = 11
    data_start_row = _create_erp_data_headers(worksheet, df, subject_details, header_start_row, YELLOW_FILL, BLACK_FONT, BOLD_FONT, CENTER_ALIGN, THIN_BORDER)
    
    # Data rows
    _add_erp_data_rows(worksheet, df, data_start_row, YELLOW_FILL, WHITE_FILL, BLACK_FONT, CENTER_ALIGN, THIN_BORDER)
    
    # Set column widths
    _set_erp_column_widths(worksheet)

def _create_erp_header_section(worksheet, metadata, yellow_fill, red_fill, black_font, bold_font, center_align, thin_border):
    """Create the header section matching the ERP format"""
    
    # Row 1: Branch
    cell = worksheet.cell(row=1, column=1, value="Branch: MRU-School of Engineering")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A1:AK1')
    
    # Row 2: Department
    cell = worksheet.cell(row=2, column=1, value="Department: Bachelor of Technology in Computer Science and Engineering")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A2:AK2')
    
    # Row 3: Class Name
    cell = worksheet.cell(row=3, column=1, value=f"Class Name: {metadata['class_name']}")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A3:T3')
    
    # Row 4: Division
    cell = worksheet.cell(row=4, column=1, value=f"Division: {metadata['division']}")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A4:T4')
    
    # Row 5: Date
    cell = worksheet.cell(row=5, column=1, value=f"Date: {metadata['date_range']}")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A5:T5')
    
    # Row 6: Program Coordinator
    cell = worksheet.cell(row=6, column=1, value=f"Program Coordinator: {metadata['coordinator']}")
    cell.font = bold_font
    cell.fill = yellow_fill
    cell.border = thin_border
    worksheet.merge_cells('A6:T6')
    
    # Red box for "Please update the Minimum Attendance Criteria"
    cell = worksheet.cell(row=3, column=21, value="Please update the Minimum Attendance Criteria (Engineering, Management & Applied Sciences)(75%), Law (70%), Education (60%)")
    cell.font = black_font
    cell.fill = red_fill
    cell.border = thin_border
    cell.alignment = center_align
    worksheet.merge_cells('U3:AK6')

def _create_erp_data_headers(worksheet, df, subject_details, start_row, yellow_fill, black_font, bold_font, center_align, thin_border):
    """Create multi-level data headers matching ERP format"""
    
    # Row 11: Subject names (Level 1)
    basic_headers = ['Sr No.', 'Roll No', 'Student Name', 'Faculty Member Name']
    current_col = 1
    
    # Basic columns
    for header in basic_headers:
        cell = worksheet.cell(row=start_row, column=current_col, value=header)
        cell.font = bold_font
        cell.fill = yellow_fill
        cell.border = thin_border
        cell.alignment = center_align
        current_col += 1
    
    # Subject columns - create proper multi-level structure
    subjects = list(subject_details.keys())
    
    for subject in subjects:
        if subject in df.columns:
            # Subject name
            cell = worksheet.cell(row=start_row, column=current_col, value=subject[:20])  # Truncate long names
            cell.font = bold_font
            cell.fill = yellow_fill
            cell.border = thin_border
            cell.alignment = center_align
            
            # Subject code (Row 12)
            code = subject_details.get(subject, {}).get('code', '')
            cell = worksheet.cell(row=start_row + 1, column=current_col, value=code)
            cell.font = black_font
            cell.fill = yellow_fill
            cell.border = thin_border
            cell.alignment = center_align
            
            # Subject type (Row 13)
            subject_type = subject_details.get(subject, {}).get('type', '')
            cell = worksheet.cell(row=start_row + 2, column=current_col, value=subject_type)
            cell.font = black_font
            cell.fill = yellow_fill
            cell.border = thin_border
            cell.alignment = center_align
            
            current_col += 1
    
    # Summary columns
    summary_headers = ['Overall %age of all subjects from ERP report', 'Count of Courses with attendance below minimum', 'Whether Critical']
    for header in summary_headers:
        if header in df.columns or any(h in df.columns for h in summary_headers):
            cell = worksheet.cell(row=start_row, column=current_col, value=header)
            cell.font = bold_font
            cell.fill = yellow_fill
            cell.border = thin_border
            cell.alignment = center_align
            
            # Empty cells for rows 12-13
            for row_offset in [1, 2]:
                cell = worksheet.cell(row=start_row + row_offset, column=current_col, value="")
                cell.fill = yellow_fill
                cell.border = thin_border
            
            current_col += 1
    
    return start_row + 3

def _add_erp_data_rows(worksheet, df, start_row, yellow_fill, white_fill, black_font, center_align, thin_border):
    """Add data rows with yellow background matching ERP format"""
    
    for row_idx, row_data in enumerate(df.values.tolist()):
        excel_row = start_row + row_idx
        
        for col_idx, value in enumerate(row_data):
            excel_col = col_idx + 1
            cell = worksheet.cell(row=excel_row, column=excel_col, value=value)
            
            # Apply yellow background to all data cells (matching the image)
            cell.fill = yellow_fill
            cell.font = black_font
            cell.border = thin_border
            cell.alignment = center_align
            
            # Special formatting for critical students
            if col_idx < 4:  # Basic info columns
                cell.alignment = Alignment(horizontal='left', vertical='center')

def _set_erp_column_widths(worksheet):
    """Set column widths matching ERP format"""
    # Set specific widths for different column types
    widths = {
        'A': 8,   # Sr No
        'B': 12,  # Roll No
        'C': 20,  # Student Name
        'D': 15,  # Faculty Member Name
    }
    
    # Apply basic column widths
    for col, width in widths.items():
        worksheet.column_dimensions[col].width = width
    
    # Set subject column widths
    for col_num in range(5, 30):  # Subject columns
        col_letter = get_column_letter(col_num)
        worksheet.column_dimensions[col_letter].width = 12

# Old helper functions removed - using new ERP format functions instead
    output_buffer.seek(0)
    return output_buffer

