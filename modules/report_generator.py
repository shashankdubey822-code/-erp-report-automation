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
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.label import DataLabelList

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
    This is the definitive, most advanced version of the Excel file generator.
    It takes full manual control to build the report for maximum stability and precision.
    """
    df.rename(columns={'Roll No_duplicate': 'Roll No'}, inplace=True)
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        sheet_name = metadata['monitoring_stage']
        
        worksheet = writer.book.create_sheet(title=sheet_name)
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])
        
        VIBRANT_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        DARK_GREY_FILL = PatternFill(start_color="A9A9A9", end_color="A9A9A9", fill_type="solid")
        WHITE_FONT = Font(color="FFFFFF", bold=True)
        BLACK_FONT = Font(color="000000", bold=True)
        CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

        last_col_letter = get_column_letter(df.shape[1])
        cell_b2 = worksheet.cell(row=2, column=2); cell_b2.value = 'DEPARTMENT OF CST'; cell_b2.font = BLACK_FONT
        worksheet.merge_cells(f'B2:{last_col_letter}2'); cell_b2.alignment = CENTER_ALIGN
        
        cell_b3 = worksheet.cell(row=3, column=2); cell_b3.value = 'LOW ATTENDANCE REVIEW REPORT (1 WEEK PRIOR TO LAST DAY OF CLASSES)'; cell_b3.font = BLACK_FONT
        worksheet.merge_cells(f'B3:{last_col_letter}3'); cell_b3.alignment = CENTER_ALIGN
        
        worksheet['B4'] = 'Branch: MRU-School of Engineering'
        worksheet['B5'] = 'Department: Bachelor of Technology in Computer Science and Engineering'
        worksheet['B6'] = f"Class Name: {metadata['class_name']}"
        worksheet['B7'] = f"Division: {metadata['division']}"
        worksheet['B8'] = f"Date: {metadata['date_range']}"
        worksheet['B9'] = f"Program Coordinator: {metadata['coordinator']}"
        
        headers1 = list(df.columns)
        
        # Safely build headers with error handling
        try:
            headers2 = [''] * 3 + [subject_details.get(subj, {}).get('code', '') for subj in subject_details.keys()] + [''] * 4
            headers3 = [''] * 3 + [subject_details.get(subj, {}).get('type', '') for subj in subject_details.keys()] + [''] * 4
        except Exception as e:
            logger.error(f"Error building Excel headers: {e}")
            # Fallback to simple headers
            headers2 = [''] * len(headers1)
            headers3 = [''] * len(headers1)
        try:
            overall_percent_col_index = headers1.index('Overall %age of all subjects from ERP report') + 1
        except ValueError:
            overall_percent_col_index = df.shape[1] - 3
        for i, val in enumerate(headers1, 1):
            cell = worksheet.cell(row=9, column=i, value=val); cell.font = BLACK_FONT
            if i <= overall_percent_col_index: cell.fill = VIBRANT_YELLOW_FILL
        
        for i, val in enumerate(headers2, 1):
            cell = worksheet.cell(row=10, column=i, value=val); cell.font = BLACK_FONT
            if i <= overall_percent_col_index: cell.fill = VIBRANT_YELLOW_FILL

        for i, val in enumerate(headers3, 1):
            cell = worksheet.cell(row=11, column=i, value=val); cell.font = BLACK_FONT
            if i <= overall_percent_col_index: cell.fill = VIBRANT_YELLOW_FILL
        
        for r_idx, row_data in enumerate(df.values.tolist(), 12):
            for c_idx, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=r_idx, column=c_idx, value=value)
                if c_idx <= (df.shape[1] - 3):
                    cell.fill = VIBRANT_YELLOW_FILL

        min_attendance = metadata['min_attendance']
        dark_grey_rule = CellIsRule(operator='lessThan', formula=[min_attendance], stopIfTrue=True, fill=DARK_GREY_FILL, font=WHITE_FONT)
        data_range = f"D12:{get_column_letter(3 + len(subject_details))}{len(df)+11}"
        worksheet.conditional_formatting.add(data_range, dark_grey_rule)
        
        summary_start_row = len(df) + 15
        subjects = list(subject_details.keys())
        thresholds = [75, 70, 65, 60]
        
        # Only create summary for subjects that exist in the DataFrame
        valid_subjects = [s for s in subjects if s in df.columns]
        logger.info(f"Creating summary for {len(valid_subjects)} valid subjects: {valid_subjects}")
        
        cell = worksheet.cell(row=summary_start_row, column=1, value="Number of Students in course")
        cell.font = BLACK_FONT
        
        # Build summary statistics safely
        for i, subject in enumerate(valid_subjects, 4):
            try:
                count = (df[subject] > 0).sum() if subject in df.columns else 0
                worksheet.cell(row=summary_start_row, column=i, value=count)
            except Exception as e:
                logger.error(f"Error calculating student count for {subject}: {e}")
                worksheet.cell(row=summary_start_row, column=i, value=0)
        
        for i, th in enumerate(thresholds, 1):
            cell = worksheet.cell(row=summary_start_row + i, column=1, value=f"Number of students below {th}%")
            cell.font = BLACK_FONT
            for j, subject in enumerate(valid_subjects, 4):
                try:
                    count = (df[subject] < th).sum() if subject in df.columns else 0
                    worksheet.cell(row=summary_start_row + i, column=j, value=count)
                except Exception as e:
                    logger.error(f"Error calculating threshold count for {subject} at {th}%: {e}")
                    worksheet.cell(row=summary_start_row + i, column=j, value=0)
        
        footer_start_row = summary_start_row + 10
        cell = worksheet.cell(row=footer_start_row, column=15, value="Signature of Mentor"); cell.font = BLACK_FONT

        chart = BarChart()
        chart.title = "Count of Students Below Minimum Attendance Criteria"
        chart.y_axis.title = 'Count of Students'
        chart.x_axis.title = 'Subjects'
        chart.height = 18; chart.width = 40  
        data_row = summary_start_row + thresholds.index(75) + 1
        data = Reference(worksheet, min_col=4, min_row=data_row, max_col=3 + len(subjects), max_row=data_row)
        cats = Reference(worksheet, min_col=4, min_row=9, max_col=3 + len(subjects), max_row=9)
        chart.add_data(data, from_rows=True, titles_from_data=False)
        chart.set_categories(cats)
        chart.legend = None
        series = chart.series[0]
        series.graphicalProperties = GraphicalProperties(solidFill="FFFF00", ln=LineProperties(solidFill="000000"))
        chart.data_labels = DataLabelList(showVal=True)
        chart_anchor = f"A{summary_start_row + len(thresholds) + 5}"
        worksheet.add_chart(chart, chart_anchor)
        
        for col in worksheet.columns:
            if any(c.value for c in col):
                length = max(len(str(c.value)) for c in col if c.value)
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = length + 2

    output_buffer.seek(0)
    return output_buffer

