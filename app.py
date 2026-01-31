import pyquotegen
"""This is the main application file for the ERP Report Automation tool."""


from pickle import APPEND
import uuid  # For creating unique filenames
import logging
import json
import base64
import time
from io import BytesIO
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
import logging
from werkzeug.utils import secure_filename
from matplotlib.figure import Figure
import pandas as pd

from modules.data_processor import create_report_dataframe
from modules.report_manager import generate_all_reports
from modules.database import update_mapping



# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

@app.route('/download_pdf/<filename>', methods=['POST'])
def download_pdf(filename):
    """Generates and downloads the final PDF report."""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        flash('File not found. Please upload the file again.')
        return redirect(url_for('index'))

    try:
        metadata = request.form.to_dict()
        metadata['min_attendance'] = float(metadata.get('min_attendance', 75))
        original_filename = metadata.get('original_filename', filename)

        with open(filepath, 'rb') as f:
            file_content_bytesio = BytesIO(f.read())
            # Pass the form metadata to the report generator
            report_output = generate_all_reports(
                file_content_bytesio,
                metadata['min_attendance'],
                user_metadata=metadata
            )
        
        report_df = report_output["dataframe"]
        # Capture subjects_with_zero_attendance for consistency, though not directly used here
        subjects_with_zero_attendance = report_output["subjects_with_zero_attendance"]
        
        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))
        
        # The final metadata is now in report_output["metadata"]
        final_metadata = report_output["metadata"]
        
        subject_details_count = len(report_output["subject_details"]) # Use report_output["subject_details"] which is filtered
        logger.info("Generated report with %d records and %d subjects", len(report_df), subject_details_count)

        pdf_buffer = report_output["pdf_buffer"]
        download_filename = f"{final_metadata.get('monitoring_stage', 'Report').replace(' ', '_')}.pdf"

        logger.info("PDF file generated successfully: %s", download_filename)

        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/pdf'
        )

    except KeyError as e:
        logger.error("Missing data error for %s: %s", filename, e, exc_info=True)
        flash(f"Data processing error: Missing expected data '{str(e)}'. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except ValueError as e:
        logger.error("Data validation error for %s: %s", filename, e, exc_info=True)
        flash(f"Data validation error: {str(e)}. Please check your input values.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except Exception as e:
        logger.error("Unexpected error during PDF generation for %s: %s", filename, e, exc_info=True)
        flash(f"An unexpected error occurred while generating the report: {str(e)}")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
# Use environment variable for secret key, with fallback for development
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-key-change-in-production')

# File upload security settings
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
ALLOWED_EXTENSIONS = {'csv', 'xls', 'xlsx'}
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def validate_file_content(file):
    """Basic validation of file content to prevent malicious uploads."""
    try:
        # Read first few bytes to check file signature
        file.seek(0)
        first_bytes = file.read(1024)
        file.seek(0)  # Reset file pointer

        # Basic checks for CSV files
        if file.filename.lower().endswith('.csv'):
            # Check if it looks like text content
            try:
                first_bytes.decode('utf-8')
                return True
            except UnicodeDecodeError:
                try:
                    first_bytes.decode('latin-1')
                    return True
                except UnicodeDecodeError:
                    return False

        # For Excel files, check basic file signatures
        if file.filename.lower().endswith(('.xls', '.xlsx')):
            # Basic Excel file signature checks
            excel_signatures = [b'PK\x03\x04', b'\xd0\xcf\x11\xe0']
            return any(first_bytes.startswith(sig) for sig in excel_signatures)

        return False
    except IOError as e:
        logger.error("Error validating file content: %s", e)
        return False


@app.route('/')
def index():
    """Renders the main landing page."""
    try:
        motivational_quote_text = pyquotegen.get_quote(category="motivational")
        display_quote = f'"{motivational_quote_text}"'
    except Exception as e:
        logger.error(f"Error fetching motivational quote: {e}")
        display_quote = '"The secret of getting ahead is getting started." (Fallback Quote)'
    return render_template('index.html', motivational_quote=display_quote)


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles the initial file upload with improved security and validation."""
    try:
        if 'erp_file' not in request.files:
            flash('No file part in the request.')
            return redirect(url_for('index'))

        file = request.files['erp_file']
        if file.filename == '':
            flash('No file selected.')
            return redirect(url_for('index'))

        # Secure the filename
        original_filename = secure_filename(file.filename)

        # Validate file extension
        if not allowed_file(original_filename):
            flash('Invalid file type. Please upload a CSV or Excel file.')
            return redirect(url_for('index'))

        # Validate file content
        if not validate_file_content(file):
            flash('Invalid file content. Please ensure the file is not corrupted.')
            return redirect(url_for('index'))

        # Generate a unique, secure filename and save the file
        _, extension = os.path.splitext(original_filename)
        unique_filename = f"{uuid.uuid4().hex}{extension.lower()}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)

        file.save(filepath)
        logger.info("File uploaded successfully: %s -> %s", original_filename, unique_filename)

        return redirect(url_for('view_file', filename=unique_filename, original_filename=original_filename))

    except (IOError, OSError) as e:
        logger.error("Error during file upload: %s", e)
        flash('An error occurred during file upload. Please try again.')
        return redirect(url_for('index'))


        return False


def detect_monitoring_stage(report_title):
    """
    Detects the monitoring stage from the report title.
    """
    search_text = (report_title or "").lower()
    
    if "first" in search_text or "1st" in search_text:
        return "First Att Monitoring"
    elif "second" in search_text or "2nd" in search_text:
        return "Second Att Monitoring"
    elif "third" in search_text or "3rd" in search_text:
        return "Third Att Monitoring"
    elif "low" in search_text or "review" in search_text:
        return "Low Attendance Review"
    elif "final" in search_text or "end" in search_text:
        return "Final Att Monitoring"
    
    return "First Att Monitoring" # Default


@app.route('/view/<filename>')
def view_file(filename):
    """Shows the user the options for their uploaded file."""
    original_filename = request.args.get('original_filename')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    if not os.path.exists(filepath):
        flash('File not found. Please upload the file again.')
        return redirect(url_for('index'))

    default_metadata = {
        'department_name': '',
        'report_title': 'ATTENDANCE MONITORING REPORT',
        'monitoring_stage': 'First Att Monitoring',
        'class_name_division': '',
        'division': '',
        'date_range': '',
        'coordinator': '',
        'min_attendance': 75,
        'report_color': '#FFFF00'
    }

    try:
        with open(filepath, 'rb') as f:
            # We parse the file here just to get the metadata for the form
            file_content_bytesio = BytesIO(f.read())
            _, _, extracted_metadata, _ = create_report_dataframe(file_content_bytesio, 75) # Unpack 4, discard 2
            # Merge default metadata with extracted metadata
            default_metadata.update(extracted_metadata)
            
            # Auto-detect monitoring stage
            detected_stage = detect_monitoring_stage(default_metadata.get('report_title'))
            default_metadata['monitoring_stage'] = detected_stage
            
    except Exception as e:
        logger.error("Failed to pre-parse file %s: %s", filename, e, exc_info=True)
        flash("Could not read metadata from file. Please check the file format or fill in the details manually.")
    
    return render_template('view_file.html',
                           filename=filename,
                           original_filename=original_filename,
                           metadata=default_metadata)


@app.route('/preview/<filename>', methods=['POST'])
def preview_file(filename):
    """Generates and displays the HTML preview table with improved error handling."""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        flash('File not found. Please upload the file again.')
        return redirect(url_for('index'))

    try:
        min_attendance = float(request.form.get('min_attendance', 75))
        original_filename = request.form.get('original_filename', filename)
        report_color = request.form.get('report_color', '#FFFF00')
        
        # Get metadata from the form
        form_metadata = request.form.to_dict()

        logger.info("Generating preview for file: %s (original: %s)", filename, original_filename)

        with open(filepath, 'rb') as f:
            file_content_bytesio = BytesIO(f.read())
            # Pass the form metadata to the report generator
            report_output = generate_all_reports(
                file_content_bytesio,
                min_attendance,
                user_metadata=form_metadata
            )
        
        report_df = report_output["dataframe"]
        
        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))
        
        # The final, merged metadata is in report_output["metadata"]
        final_metadata = report_output["metadata"]

        summary_html = report_output["html_summary"]
        chart_image_buf = report_output["chart_buffer"]
        subjects_with_zero_attendance = report_output["subjects_with_zero_attendance"] # Capture the new output
        chart_image_buf.seek(0)
        chart_image_base64 = base64.b64encode(chart_image_buf.read()).decode('utf-8')
        chart_image = f"data:image/png;base64,{chart_image_base64}"

        data_json = report_df.to_json(orient='split')
        subject_details_json = json.dumps(final_metadata.get("subject_details", {}))

        logger.info("Preview generated successfully for %d records", len(report_df))

        return render_template('preview.html',
                               data_json=data_json,
                               filename=filename,
                               metadata=final_metadata,
                               subject_details=final_metadata.get("subject_details", {}),
                               subject_details_json=subject_details_json,
                               summary_table=summary_html,
                               chart_image=chart_image,
                               report_color=report_color,
                               subjects_with_zero_attendance=subjects_with_zero_attendance) # Pass to template

    except ValueError as e:
        logger.error("Data processing error for %s: %s", filename, e, exc_info=True)
        flash(f"Data processing error: {str(e)}. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=original_filename))
    except (IOError, OSError) as e:
        logger.error("Unexpected error during preview generation for %s: %s", filename, e, exc_info=True)
        flash('An unexpected error occurred during preview generation. Please try again.')
        return redirect(url_for('view_file', filename=filename, original_filename=original_filename))


@app.route('/download/<filename>', methods=['POST'])
def download_file(filename):
    """Generates and downloads the final Excel report with improved error handling."""
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        flash('File not found. Please upload the file again.')
        return redirect(url_for('index'))

    try:
        metadata = request.form.to_dict()
        metadata['min_attendance'] = float(metadata.get('min_attendance', 75))
        original_filename = metadata.get('original_filename', filename)

        with open(filepath, 'rb') as f:
            file_content_bytesio = BytesIO(f.read())
            # Pass the form metadata to the report generator
            report_output = generate_all_reports(
                file_content_bytesio,
                metadata['min_attendance'],
                user_metadata=metadata
            )
        
        report_df = report_output["dataframe"]
        # Capture subjects_with_zero_attendance for consistency, though not directly used here
        subjects_with_zero_attendance = report_output["subjects_with_zero_attendance"]
        
        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))

        # The final metadata is now in report_output["metadata"]
        final_metadata = report_output["metadata"]
        
        subject_details_count = len(report_output["subject_details"]) # Use report_output["subject_details"] which is filtered
        logger.info("Generated report with %d records and %d subjects", len(report_df), subject_details_count)

        excel_buffer = report_output["excel_buffer"]
        download_filename = f"{final_metadata.get('monitoring_stage', 'Report').replace(' ', '_')}.xlsx"

        logger.info("Excel file generated successfully: %s", download_filename)

        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except KeyError as e:
        logger.error("Missing data error for %s: %s", filename, e, exc_info=True)
        flash(f"Data processing error: Missing expected data '{str(e)}'. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except ValueError as e:
        logger.error("Data validation error for %s: %s", filename, e, exc_info=True)
        flash(f"Data validation error: {str(e)}. Please check your input values.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except (IOError, OSError) as e:
        logger.error("Unexpected error during Excel generation for %s: %s", filename, e, exc_info=True)
        flash(f"An unexpected error occurred while generating the report: {str(e)}")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
