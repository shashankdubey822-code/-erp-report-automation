import csv
import logging
from io import StringIO
from typing import Dict, Tuple

import pandas as pd

from .utilities import safe_str

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
        if "Branch:" in line:
            extracted_metadata['branch'] = line.split("Branch:")[1].split(",")[0].strip()
        if "Department:" in line:
            extracted_metadata['department_specialization'] = line.split("Department:")[1].split(",")[0].strip()
        if "Class Name:" in line:
            extracted_metadata['class_name_division'] = line.split("Class Name:")[1].split(",")[0].strip()
        if "Date:" in line:
            extracted_metadata['date_range'] = line.split("Date:")[1].split(",")[0].strip()
        if "Program Coordinator:" in line:
            extracted_metadata['coordinator'] = line.split("Program Coordinator:")[1].split(",")[0].strip()
        if "Academic Year:" in line:
            extracted_metadata['academic_year'] = line.split("Academic Year:")[1].split("-")[0].strip()
        if "Semester:" in line:
            extracted_metadata['semester'] = line.split("Semester:")[1].split(",")[0].strip()
        # Look for a line that might be the department name (all caps, single entry in row)
        if len(row) == 1 and row[0].isupper() and 'department_name' not in extracted_metadata:
             extracted_metadata['department_name'] = row[0].strip()
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

    # Construct final_headers - pair subject + metric info where appropriate
    final_headers = []
    subject_details = {}
    for i, metric in enumerate(h4):
        subj = h1[i]
        code = h2[i] if i < len(h2) else ""
        typ = h3[i] if i < len(h3) else ""
        # Identify special columns
        if subj.strip() in ("Sr.", "Division/Section", "Unique id", "Rollno", "Student Name", "PRN / Enroll"):
            final_headers.append(subj.strip())
        elif "Total" in subj or "Grand Total" in subj or "Total" in metric:
            # unify grand total naming
            final_headers.append(f"Grand Total - {metric}".strip())
        else:
            # typical subject column: "SUBJ - metric"
            label = f"{subj} - {metric}".strip()
            final_headers.append(label)
            if subj not in subject_details:
                subject_details[subj] = {"code": code, "type": typ}

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
