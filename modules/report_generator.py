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
        # Try multiple possible suffix patterns
        possible_suffixes = [' - Total %', ' - % (PP)', ' - % (PR)', ' - %', ' - Total', '- Total %', '- % (PP)', '- % (PR)']
        
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
    Professional Excel report generator with multi-level headers, borders, and styling like Image 2.
    """
    df.rename(columns={'Roll No_duplicate': 'Roll No'}, inplace=True)
    output_buffer = BytesIO()
    
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        sheet_name = metadata['monitoring_stage']
        worksheet = writer.book.create_sheet(title=sheet_name)
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])
        
        # Get styles
        styles = PROFESSIONAL_STYLES
        
        # Create header section with professional styling
        _create_professional_header(worksheet, metadata, df.shape[1])
        
        # Create multi-level data headers with categories
        header_start_row = 11
        data_start_row = _create_multi_level_headers(worksheet, df, subject_details, header_start_row, styles)
        
        # Add data with borders and styling
        _add_data_with_styling(worksheet, df, data_start_row, styles)
        
        # Add professional summary section
        _create_professional_summary(worksheet, df, subject_details, len(df) + data_start_row + 2, styles)
        
        # Apply column widths
        _apply_column_widths(worksheet)
        
        # Add conditional formatting for critical attendance
        _apply_conditional_formatting(worksheet, df, subject_details, data_start_row, metadata['min_attendance'])

    output_buffer.seek(0)
    return output_buffer

def _create_professional_header(worksheet, metadata, total_cols):
    """Create professional header section like Image 2"""
    styles = PROFESSIONAL_STYLES
    last_col = get_column_letter(total_cols)
    
    # Main title
    cell = worksheet.cell(row=1, column=1, value='DEPARTMENT OF CST')
    cell.font = Font(bold=True, size=14)
    cell.alignment = styles['center_align']
    worksheet.merge_cells(f'A1:{last_col}1')
    
    # Subtitle
    cell = worksheet.cell(row=2, column=1, value='LOW ATTENDANCE REVIEW REPORT (1 WEEK PRIOR TO LAST DAY OF CLASSES)')
    cell.font = Font(bold=True, size=12)
    cell.alignment = styles['center_align']
    worksheet.merge_cells(f'A2:{last_col}2')
    
    # Details section
    details = [
        f"Branch: MRU-School of Engineering",
        f"Department: Bachelor of Technology in Computer Science and Engineering",
        f"Class Name: {metadata['class_name']}",
        f"Division: {metadata['division']}",
        f"Date: {metadata['date_range']}",
        f"Program Coordinator: {metadata['coordinator']}"
    ]
    
    for i, detail in enumerate(details, 3):
        cell = worksheet.cell(row=i, column=1, value=detail)
        cell.font = Font(bold=True, size=10)
        worksheet.merge_cells(f'A{i}:{last_col}{i}')

def _create_multi_level_headers(worksheet, df, subject_details, start_row, styles):
    """Create multi-level headers with subject categories"""
    # Categorize subjects
    subject_categories = {}
    basic_cols = ['Sr No.', 'Roll No', 'Student Name']
    
    for subject in subject_details.keys():
        category = categorize_subject(subject)
        if category not in subject_categories:
            subject_categories[category] = []
        subject_categories[category].append(subject)
    
    # Create category headers (Level 1)
    current_col = 1
    
    # Basic columns
    for col_name in basic_cols:
        cell = worksheet.cell(row=start_row, column=current_col, value=col_name)
        cell.font = styles['header_font']
        cell.fill = styles['header_fill']
        cell.alignment = styles['center_align']
        cell.border = styles['medium_border']
        worksheet.merge_cells(f'{get_column_letter(current_col)}{start_row}:{get_column_letter(current_col)}{start_row + 1}')
        current_col += 1
    
    # Subject categories
    for category, subjects in subject_categories.items():
        if subjects:  # Only if subjects exist in the dataframe
            valid_subjects = [s for s in subjects if s in df.columns]
            if valid_subjects:
                # Category header
                start_col = current_col
                end_col = current_col + len(valid_subjects) - 1
                
                cell = worksheet.cell(row=start_row, column=start_col, value=category)
                cell.font = styles['category_font']
                cell.fill = styles['category_fill']
                cell.alignment = styles['center_align']
                cell.border = styles['medium_border']
                
                if end_col > start_col:
                    worksheet.merge_cells(f'{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{start_row}')
                
                # Subject headers (Level 2)
                for subject in valid_subjects:
                    cell = worksheet.cell(row=start_row + 1, column=current_col, value=subject)
                    cell.font = styles['header_font']
                    cell.fill = styles['header_fill']
                    cell.alignment = styles['center_align']
                    cell.border = styles['thin_border']
                    current_col += 1
    
    # Add summary columns
    summary_cols = ['Overall %age of all subjects from ERP report', 'Count of Courses with attendance below minimum attendance criteria', 'Whether Critical']
    for col_name in summary_cols:
        if col_name in df.columns:
            cell = worksheet.cell(row=start_row, column=current_col, value=col_name)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.alignment = styles['center_align']
            cell.border = styles['medium_border']
            worksheet.merge_cells(f'{get_column_letter(current_col)}{start_row}:{get_column_letter(current_col)}{start_row + 1}')
            current_col += 1
    
    return start_row + 2

def _add_data_with_styling(worksheet, df, start_row, styles):
    """Add data with professional styling and borders"""
    for row_idx, row_data in enumerate(df.values.tolist()):
        excel_row = start_row + row_idx
        for col_idx, value in enumerate(row_data):
            excel_col = col_idx + 1
            cell = worksheet.cell(row=excel_row, column=excel_col, value=value)
            
            # Apply styling
            cell.font = styles['data_font']
            cell.fill = styles['data_fill']
            cell.border = styles['thin_border']
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Special styling for critical rows
            if 'Whether Critical' in df.columns and row_data[df.columns.get_loc('Whether Critical')] == 'CRITICAL':
                if col_idx >= 3:  # Subject columns
                    cell.fill = styles['critical_fill']

def _create_professional_summary(worksheet, df, subject_details, start_row, styles):
    """Create professional summary section"""
    valid_subjects = [s for s in subject_details.keys() if s in df.columns]
    thresholds = [75, 70, 65, 60]
    
    # Summary title
    cell = worksheet.cell(row=start_row, column=1, value="ATTENDANCE SUMMARY")
    cell.font = Font(bold=True, size=12)
    cell.border = styles['medium_border']
    
    start_row += 2
    
    # Headers
    cell = worksheet.cell(row=start_row, column=1, value="Metrics")
    cell.font = styles['header_font']
    cell.fill = styles['header_fill']
    cell.border = styles['medium_border']
    
    for i, subject in enumerate(valid_subjects[:10], 2):  # Limit to avoid overflow
        cell = worksheet.cell(row=start_row, column=i, value=subject[:15] + '...' if len(subject) > 15 else subject)
        cell.font = styles['header_font']
        cell.fill = styles['header_fill']
        cell.border = styles['thin_border']
        cell.alignment = styles['center_align']
    
    # Summary data
    metrics = [("Students in course", lambda s: (df[s] > 0).sum())] + \
              [(f"Students below {th}%", lambda s, t=th: (df[s] < t).sum()) for th in thresholds]
    
    for row_offset, (metric_name, calc_func) in enumerate(metrics, 1):
        excel_row = start_row + row_offset
        
        cell = worksheet.cell(row=excel_row, column=1, value=metric_name)
        cell.font = styles['data_font']
        cell.border = styles['thin_border']
        
        for col_offset, subject in enumerate(valid_subjects[:10], 2):
            try:
                value = calc_func(subject) if subject in df.columns else 0
                cell = worksheet.cell(row=excel_row, column=col_offset, value=value)
                cell.font = styles['data_font']
                cell.border = styles['thin_border']
                cell.alignment = styles['center_align']
            except Exception as e:
                logger.error(f"Error calculating summary for {subject}: {e}")

def _apply_column_widths(worksheet):
    """Apply professional column widths"""
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value or '')) for cell in column_cells)
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max(length + 2, 10), 25)

def _apply_conditional_formatting(worksheet, df, subject_details, data_start_row, min_attendance):
    """Apply conditional formatting for low attendance"""
    styles = PROFESSIONAL_STYLES
    
    # Find subject columns
    subject_cols = []
    for col_idx, col_name in enumerate(df.columns):
        if col_name in subject_details:
            subject_cols.append(get_column_letter(col_idx + 1))
    
    if subject_cols:
        for col_letter in subject_cols:
            range_str = f"{col_letter}{data_start_row}:{col_letter}{data_start_row + len(df) - 1}"
            rule = CellIsRule(operator='lessThan', formula=[min_attendance], 
                            stopIfTrue=True, 
                            fill=PatternFill(start_color="FFB3B3", end_color="FFB3B3", fill_type="solid"))
            worksheet.conditional_formatting.add(range_str, rule)
    output_buffer.seek(0)
    return output_buffer

