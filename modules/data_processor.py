import csv
import logging
from io import StringIO
from typing import Dict, Tuple
from datetime import datetime

import pandas as pd

from .utilities import safe_str
from .database import get_clean_name, init_db

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ensure DB is initialized on module load or first use
init_db()

def determine_academic_period(metadata: Dict) -> str:
    """
    Determines the academic period string (e.g., '2024 Odd', '2025 Even')
    based on 'date_range' or 'academic_year'.
    """
    date_range = metadata.get('date_range', '')
    
    # Try parsing date range first (Format: DD/MM/YYYY to ...)
    if date_range:
        try:
            # Handle "30/07/2024 to 14/11/2024" or similar
            start_date_str = date_range.split('to')[0].strip()
            # Try a few common formats
            for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
                try:
                    dt = datetime.strptime(start_date_str, fmt)
                    year = dt.year
                    month = dt.month
                    
                    if 7 <= month <= 12:
                        return f"{year} Odd"
                    else:
                        return f"{year} Even"
                except ValueError:
                    continue
        except Exception:
            pass
            
    # Fallback to academic_year + semester
    acad_year = metadata.get('academic_year', '').strip()
    sem = metadata.get('semester', '').strip()
    
    if acad_year:
        # If acad_year is "2024", check semester
        if sem:
            # Check for Roman Numerals or digits
            is_odd = False
            if sem.isdigit():
                is_odd = int(sem) % 2 != 0
            else:
                # Roman numeral check (I, III, V, VII are Odd)
                is_odd = sem.upper() in ['I', 'III', 'V', 'VII']
            
            suffix = "Odd" if is_odd else "Even"
            return f"{acad_year} {suffix}"
            
        return acad_year

    return "Unknown"

# -------------------------
# Dataframe creation
# -------------------------
def create_report_dataframe(erp_file, min_attendance_criteria: int = 75) -> Tuple[pd.DataFrame, Dict, Dict]:
    """
    Parse ERP CSV-like file object and return (output_df, subject_details, extracted_metadata).
    output_df: DataFrame with 'Sr No.', 'Roll No', 'Student Name', subject percentage columns, etc.
    subject_details: dict mapping subject_name -> {'code': code, 'type': type}
    extracted_metadata: dict with dynamically extracted header fields
    """
    erp_file.seek(0)
    raw = erp_file.read()
    if isinstance(raw, bytes):
        content = raw.decode("utf-8", errors="ignore")
    else:
        content = str(raw)

    sio = StringIO(content)
    reader = csv.reader(sio)
    rows = list(reader)
    
    # Metadata extraction
    extracted_metadata = {}
    for row in rows[:20]:  # Scan top 20 rows for metadata
        line = ",".join(row)
        
        # Helper to safely extract value after a keyword (handles "Key:", "Key :", etc.)
        def extract_val(keyword, text):
            # patterns to try: "Keyword:", "Keyword :"
            for k in [f"{keyword}:", f"{keyword} :"]:
                if k in text:
                    parts = text.split(k)
                    if len(parts) > 1:
                        # Take the part after the keyword, split by comma to get just the cell value
                        return parts[1].split(",")[0].strip()
            return None

        # Extract fields using the helper
        if val := extract_val("Branch", line): extracted_metadata['branch'] = val
        if val := extract_val("Department", line): 
            # Heuristic: If it's long, it's specialization. If short/Dept, it's name.
            if "Bachelor" in val or "Master" in val or "Engineering" in val:
                extracted_metadata['department_specialization'] = val
            else:
                extracted_metadata['department_name'] = val
                
        if val := extract_val("Class Name", line): extracted_metadata['class_name_division'] = val
        elif val := extract_val("Class", line): extracted_metadata['class_name_division'] = val # Fallback
        
        if val := extract_val("Division", line): extracted_metadata['division'] = val
        
        if val := extract_val("Date", line): extracted_metadata['date_range'] = val
        if val := extract_val("Program Coordinator", line): extracted_metadata['coordinator'] = val
        if val := extract_val("Academic Year", line): extracted_metadata['academic_year'] = val.split("-")[0].strip()
        if val := extract_val("Semester", line): extracted_metadata['semester'] = val
        
        # Look for a line that might be the report title
        if "ATTENDANCE" in line.upper() and 'report_title' not in extracted_metadata:
            extracted_metadata['report_title'] = line.split(',')[0].strip()
    
    # Basic heuristics for header start
    header_start_index = -1
    header_patterns = [
        "Sr.,Division/Section,Unique id",
        "Sr,Division/Section,Unique id",
        "Sr.,Division,Unique id",
        "Unique id",
        "Roll",
        "Student Name",
    ]

    # join each row as comma-joined string
    joined = [",".join(r) for r in rows]
    for i, line in enumerate(joined[:60]):
        cleaned = line.replace('"', "").replace(" ", "")
        for pat in header_patterns:
            if pat.replace(" ", "") in cleaned:
                header_start_index = i
                logger.info("Header detected at line %d using pattern '%s'", i, pat)
                break
        if header_start_index != -1:
            break

    if header_start_index == -1:
        # fallback: look for 'Roll' or 'Student Name' exact tokens
        for i, r in enumerate(rows[:60]):
            for cell in r:
                if str(cell).strip().lower() in ("roll", "student name"):
                    header_start_index = i
                    break
            if header_start_index != -1:
                break

    if header_start_index == -1:
        logger.error("Failed to detect header in ERP. First 10 lines: %s", joined[:10])
        raise ValueError("Could not find the data table header in the ERP file. Please check the file format.")

    # Extract header rows (we expect multiple rows describing subject names / codes / types / metrics)
    # We'll be defensive: if rows missing, fill with empty strings.
    def row_at(idx):
        return rows[idx] if 0 <= idx < len(rows) else []

    h1 = [safe_str(x).strip() for x in row_at(header_start_index)]
    h2 = [safe_str(x).strip() for x in row_at(header_start_index + 2)]
    h3 = [safe_str(x).strip() for x in row_at(header_start_index + 3)]
    h4 = [safe_str(x).strip() for x in row_at(header_start_index + 4)]

    # Expand to same length
    maxlen = max(len(h1), len(h2), len(h3), len(h4))
    def pad(lst):
        return lst + [""] * (maxlen - len(lst))
    h1, h2, h3, h4 = pad(h1), pad(h2), pad(h3), pad(h4)

    # Fill empty subject names in h1 with last seen (ERP often uses merged header cells)
    last = ""
    for i, val in enumerate(h1):
        if val:
            last = val
        else:
            h1[i] = last

    def is_valid_code(val: str) -> bool:
        """Returns True if val starts with a letter and contains at least one digit."""
        s = val.strip()
        if not s: return False
        # Must start with a letter (e.g., "KCS-101")
        if not s[0].isalpha():
            return False
        # Must contain at least one digit to distinguish from "Core", "Theory"
        return any(char.isdigit() for char in s)

    # Construct final_headers - pair subject + metric info where appropriate
    final_headers = []
    subject_details = {}
    for i, metric in enumerate(h4):
        subj = h1[i]
        
        # Smart Code Detection: Check h2 and h3 for a valid code
        raw_h2 = h2[i] if i < len(h2) else ""
        raw_h3 = h3[i] if i < len(h3) else ""
        
        code = ""
        # Prioritize h2, then h3
        if is_valid_code(raw_h2):
            code = raw_h2
        elif is_valid_code(raw_h3):
            code = raw_h3
            
        typ = raw_h3 # Default to h3 for type, or maybe h2 if swapped? 
        # For now, we trust the existing logic for type but override code if needed.

        # Identify special columns
        if subj.strip() in ("Sr.", "Division/Section", "Unique id", "Rollno", "Student Name", "PRN / Enroll"):
            final_headers.append(subj.strip())
        elif "Total" in subj or "Grand Total" in subj or "Total" in metric:
            # unify grand total naming
            final_headers.append(f"Grand Total - {metric}".strip())
        else:
            # typical subject column: "SUBJ - metric"
            # Use database lookup to get formatted name
            clean_subj = get_clean_name(subj.strip())
            label = f"{clean_subj} - {metric}".strip()
            final_headers.append(label)
            if clean_subj not in subject_details:
                subject_details[clean_subj] = {"code": code, "type": typ}
                # Code saving removed as per request


    # Deduplicate headers to prevent pandas error
    new_headers = []
    counts = {}
    for header in final_headers:
        if header in counts:
            counts[header] += 1
            new_headers.append(f"{header}_duplicate_{counts[header]}")
        else:
            counts[header] = 1
            new_headers.append(header)
    final_headers = new_headers

    # Data rows start: ERP often has 6 header lines; we'll use header_start_index + 6 as before
    data_start = header_start_index + 6
    data_rows = rows[data_start:]
    data_joined = "\n".join([",".join(r) for r in data_rows])
    df = pd.read_csv(StringIO(data_joined), header=None, names=final_headers, on_bad_lines="skip")
    # Ensure Rollno column exists
    roll_col = next((c for c in df.columns if "Roll" in c or "roll" in c or "Rollno" in c), None)
    if roll_col is None:
        # look for Unique id
        roll_col = next((c for c in df.columns if "Unique" in c or "unique" in c), None)
    if roll_col is None:
        raise ValueError("Could not find Roll/Rollno column in parsed data")
    df.dropna(subset=[roll_col], inplace=True)

    # Build output_df with basic info
    output_df = pd.DataFrame({
        "Sr No.": range(1, len(df) + 1),
        "Roll No": df[roll_col].astype(str)
    })

    # For "Student Name" column detection
    name_col = next((c for c in df.columns if "Student Name" in c or "Student" in c), None)
    if name_col:
        output_df["Student Name"] = df[name_col].astype(str)
    else:
        # fallback: attempt to find a sensible column
        output_df["Student Name"] = df.iloc[:, 1].astype(str) if df.shape[1] > 1 else ""

    # Map subjects -> column names that contain percentage metrics
    subject_percent_cols = {}
    for subj in subject_details:
        # try many suffixes to find correct column name
        found = None
        suffixes = [
            " - Total %", " - % (PP)", " - % (PR)", " - % (TUT)",
            " - %", " - Total", "- Total %", "- % (PP)", "- % (PR)", "- % (TUT)",
            f"{subj} - %", f"{subj} - Total %"
        ]
        for s in suffixes:
            candidate = f"{subj}{s}"
            if candidate in df.columns:
                found = candidate
                break
        # fallback: find any column whose name starts with subject and contains '%'
        if not found:
            for col in df.columns:
                if col.startswith(subj) and "%" in col:
                    found = col
                    break
        if found:
            subject_percent_cols[subj] = found
        else:
            # warn but continue
            logger.warning("No percentage column found for subject '%s'", subj)

    # add numeric subject percentage columns to output_df
    for subj, col_name in subject_percent_cols.items():
        output_df[subj] = pd.to_numeric(df.get(col_name, 0), errors="coerce").fillna(0)

    # overall percentage
    overall_col = next((c for c in df.columns if "Grand Total - %" in c or "Total - %" in c or "Overall" in c), None)
    if overall_col:
        output_df["Overall %age of all subjects from ERP report"] = pd.to_numeric(df.get(overall_col), errors="coerce").fillna(0)
    else:
        # safe fallback: create zeros
        output_df["Overall %age of all subjects from ERP report"] = 0

    

    # count of courses below threshold
    subject_keys = list(subject_percent_cols.keys())
    if subject_keys:
        output_df["Count of Courses with attendance below minimum attendance criteria"] = output_df[subject_keys].apply(
            lambda row: (row < min_attendance_criteria).sum(), axis=1
        )
    else:
        output_df["Count of Courses with attendance below minimum attendance criteria"] = 0

    output_df["Whether Critical"] = output_df["Count of Courses with attendance below minimum attendance criteria"].apply(
        lambda c: "CRITICAL" if c >= 3 else ""
    )

    subjects_with_zero_attendance = []
    subjects_to_drop_from_df = []

    # Iterate through subject percentage columns to find those with 0% attendance for all students
    for subj_name in subject_percent_cols.keys():
        if not output_df[subj_name].empty and (output_df[subj_name] == 0).all():
            subjects_with_zero_attendance.append(subj_name)
            subjects_to_drop_from_df.append(subj_name)
            logger.info("Subject '%s' identified with 0%% attendance for all students.", subj_name)
            
    # Drop these subjects from the output_df
    if subjects_to_drop_from_df:
        output_df = output_df.drop(columns=subjects_to_drop_from_df)
        # Also remove them from subject_details
        for subj_name in subjects_to_drop_from_df:
            subject_details.pop(subj_name, None)

    return output_df, subject_details, extracted_metadata, subjects_with_zero_attendance
