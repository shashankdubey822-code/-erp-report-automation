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

from modules.report_generator import (
    create_report_dataframe,
    create_excel_file,
    create_pdf_file,
    generate_summary_table_html,
    generate_chart_image,
)



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
            report_df, subject_details, extracted_metadata = create_report_dataframe(f, metadata['min_attendance'])
        
        metadata.update(extracted_metadata)

        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))

        logger.info("Generated report with %d records and %d subjects", len(report_df), len(subject_details))

        # Generate chart
        chart_image = generate_chart_image(report_df)

        pdf_buffer = create_pdf_file(report_df, subject_details, metadata, chart_image=chart_image)
        download_filename = f"{metadata.get('monitoring_stage', 'Report').replace(' ', '_')}.pdf"

        logger.info("PDF file generated successfully: %s", download_filename)

        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/pdf'
        )

    except KeyError as e:
        logger.error("Missing data error for %s: %s", filename, e)
        flash(f"Data processing error: Missing expected data '{str(e)}'. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except ValueError as e:
        logger.error("Data validation error for %s: %s", filename, e)
        flash(f"Data validation error: {str(e)}. Please check your input values.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except Exception as e:
        logger.error("Unexpected error during PDF generation for %s: %s", filename, e)
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
                return False

        # For Excel files, check basic file signatures
        if file.filename.lower().endswith(('.xls', '.xlsx')):
            # Basic Excel file signature checks
            excel_signatures = [b'PK\x03\x04', b'\xd0\xcf\x11\xe0']
            return any(first_bytes.startswith(sig) for sig in excel_signatures)

        return True
    except IOError as e:
        logger.error("Error validating file content: %s", e)
        return False


@app.route('/')
def index():
    """Renders the main landing page."""
    return render_template('index.html')


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


@app.route('/view/<filename>')
def view_file(filename):
    """Shows the user the options for their uploaded file."""
    # EDIT: Correctly get the original filename passed from the upload step
    original_filename = request.args.get('original_filename')
    return render_template('view_file.html',
                           filename=filename,
                           original_filename=original_filename)


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

        logger.info("Generating preview for file: %s (original: %s)", filename, original_filename)

        with open(filepath, 'rb') as f:
            report_df, subject_details, extracted_metadata = create_report_dataframe(f, min_attendance)
        
        metadata = request.form.to_dict()
        metadata.update(extracted_metadata)

        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))

        # Generate summary and chart on the server-side
        summary_html = generate_summary_table_html(report_df, min_attendance)
        chart_image_buf = generate_chart_image(report_df)
        chart_image_base64 = base64.b64encode(chart_image_buf.read()).decode('utf-8')
        chart_image = f"data:image/png;base64,{chart_image_base64}"

        data_json = report_df.to_json(orient='split')
        subject_details_json = json.dumps(subject_details)

        logger.info("Preview generated successfully for %d records", len(report_df))

        return render_template('preview.html',
                               data_json=data_json,
                               filename=filename,
                               metadata=metadata,
                               subject_details=subject_details,
                               subject_details_json=subject_details_json,
                               summary_table=summary_html,
                               chart_image=chart_image)

    except ValueError as e:
        logger.error("Data processing error for %s: %s", filename, e)
        flash(f"Data processing error: {str(e)}. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=original_filename))
    except (IOError, OSError) as e:
        logger.error("Unexpected error during preview generation for %s: %s", filename, e)
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
            report_df, subject_details, extracted_metadata = create_report_dataframe(f, metadata['min_attendance'])
        
        metadata.update(extracted_metadata)

        if report_df.empty:
            flash('No data found in the uploaded file. Please check the file format.')
            return redirect(url_for('view_file', filename=filename, original_filename=original_filename))

        logger.info("Generated report with %d records and %d subjects", len(report_df), len(subject_details))

        # Generate chart
        chart_image = generate_chart_image(report_df)

        excel_buffer = create_excel_file(report_df, subject_details, metadata, chart_image=chart_image)
        download_filename = f"{metadata.get('monitoring_stage', 'Report').replace(' ', '_')}.xlsx"

        logger.info("Excel file generated successfully: %s", download_filename)

        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except KeyError as e:
        logger.error("Missing data error for %s: %s", filename, e)
        flash(f"Data processing error: Missing expected data '{str(e)}'. Please check your file format.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except ValueError as e:
        logger.error("Data validation error for %s: %s", filename, e)
        flash(f"Data validation error: {str(e)}. Please check your input values.")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))
    except (IOError, OSError) as e:
        logger.error("Unexpected error during Excel generation for %s: %s", filename, e)
        flash(f"An unexpected error occurred while generating the report: {str(e)}")
        return redirect(url_for('view_file', filename=filename, original_filename=metadata.get('original_filename', filename)))



if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
