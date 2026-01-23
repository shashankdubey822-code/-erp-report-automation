import pandas as pd

# -------------------------
# HTML summary helper
# -------------------------
def generate_summary_table_html(df: pd.DataFrame, min_attendance: int = 75, subjects_with_zero_attendance: list = None) -> str:
    """
    Return an HTML snippet representing summary counts for each subject.
    Also prepends a message if there are subjects with 0% attendance for all students.
    """
    html_output = ""
    if subjects_with_zero_attendance:
        zero_att_message = "The following subjects have 0% attendance for all students: " + \
                           ", ".join([f"<strong>{s}</strong>" for s in subjects_with_zero_attendance]) + \
                           ". These subjects are not included in the main table."
        html_output += f"<div class='alert alert-warning mb-3' role='alert'>{zero_att_message}</div>"

    subject_columns = [col for col in df.columns if col not in [
        "Sr No.", "Roll No", "Student Name", "Overall %age of all subjects from ERP report",
        "Roll No_duplicate", "Count of Courses with attendance below minimum attendance criteria",
        "Whether Critical"
    ]]

    summary_data = []
    for subject in subject_columns:
        summary_data.append({
            "Subject": subject,
            f"Below {min_attendance}%": int((pd.to_numeric(df[subject], errors="coerce") < min_attendance).sum()),
            "Below 70%": int((pd.to_numeric(df[subject], errors="coerce") < 70).sum()),
            "Below 65%": int((pd.to_numeric(df[subject], errors="coerce") < 65).sum()),
            "Below 60%": int((pd.to_numeric(df[subject], errors="coerce") < 60).sum()),
        })
    summary_df = pd.DataFrame(summary_data)
    
    html_output += summary_df.to_html(classes="summary-table", index=False)
    return html_output
