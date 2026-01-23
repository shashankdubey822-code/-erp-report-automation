import logging
from io import BytesIO
from typing import Dict, Any

from .data_processor import create_report_dataframe
from .pdf_generator import create_pdf_file
from .excel_generator import create_excel_file
from .html_summary_generator import generate_summary_table_html
from .chart_image_generator import generate_chart_image

logger = logging.getLogger(__name__)

def generate_all_reports(erp_file_content: BytesIO, min_attendance_criteria: int = 75, user_metadata: Dict = None) -> Dict[str, Any]:
    """
    Orchestrates the generation of all report types from an ERP file.
    
    Args:
        erp_file_content: The content of the ERP CSV file as a BytesIO object.
        min_attendance_criteria: The minimum attendance percentage for flagging students.
        user_metadata: Optional. A dictionary of metadata from the user form to override extracted values.

    Returns:
        A dictionary containing the generated report buffers and metadata.
    """
    logger.info("Starting report generation process.")

    # 1. Process data and extract metadata from the file
    df, subject_details, extracted_metadata, subjects_with_zero_attendance = create_report_dataframe(erp_file_content, min_attendance_criteria)
    
    # Create the final metadata object
    final_metadata = extracted_metadata
    if user_metadata:
        # The user's metadata from the form takes precedence
        final_metadata.update(user_metadata)
    
    final_metadata['min_attendance'] = min_attendance_criteria

    # 2. Generate chart image (used by both PDF and Excel)
    chart_original_buffer = generate_chart_image(df)
    chart_bytes = chart_original_buffer.getvalue()

    # 3. Generate PDF report using the final, merged metadata
    pdf_buffer = create_pdf_file(df, subject_details, final_metadata, chart_image=BytesIO(chart_bytes), subjects_with_zero_attendance=subjects_with_zero_attendance)

    # 4. Generate Excel report using the final, merged metadata
    excel_buffer = create_excel_file(df, subject_details, final_metadata, chart_image=BytesIO(chart_bytes), subjects_with_zero_attendance=subjects_with_zero_attendance)

    # 5. Generate HTML summary, passing the subjects with zero attendance
    html_summary = generate_summary_table_html(df, min_attendance_criteria, subjects_with_zero_attendance)

    logger.info("Report generation process completed.")

    return {
        "dataframe": df,
        "pdf_buffer": pdf_buffer,
        "excel_buffer": excel_buffer,
        "html_summary": html_summary,
        "metadata": final_metadata, # Return the final, merged metadata
        "chart_buffer": BytesIO(chart_bytes),
        "subjects_with_zero_attendance": subjects_with_zero_attendance, # Also return this list
        "subject_details": subject_details # ADD THIS LINE
    }

