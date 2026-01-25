# 📊 ERP Report Automator - Deep Analysis

## 1. Project Overview
**ERP Report Automator** is a specialized web application designed to streamline the processing of raw student attendance data from ERP systems. It transforms unformatted, messy CSV/Excel exports into professional, actionable reports (PDF, Excel, and HTML summaries).

**Core Mission:** To automate the identification of "At-Risk" students (low attendance) and provide faculty with instant, visual insights without manual Excel manipulation.

**Tech Stack:**
*   **Backend:** Python (Flask)
*   **Data Engine:** Pandas, NumPy
*   **Reporting:** ReportLab (PDF), OpenPyXL (Excel), Matplotlib (Charts)
*   **Frontend:** HTML5, CSS3 (Bootstrap 5), JavaScript (Vanilla + Alpine.js concepts)

---

## 2. Directory Structure & File Analysis

The project follows a modular "Controller-Service" architecture where `app.py` acts as the controller and the `modules/` package contains the business logic services.

```text
-erp-report-automation/
├── app.py                      # [Controller] Main Flask application entry point. Handles routing, uploads, and session management.
├── .env.example                # [Config] Template for environment variables.
├── requirements.txt            # [Config] Python dependencies list.
├── render.yaml                 # [Config] Deployment configuration for Render.com.
├── modules/                    # [Service Layer] Core business logic.
│   ├── __init__.py             # Makes 'modules' a Python package.
│   ├── data_processor.py       # [The Brain] Parses raw ERP files, cleans data, and calculates attendance stats.
│   ├── report_manager.py       # [Orchestrator] Coordinates the generation of PDF, Excel, and HTML outputs.
│   ├── pdf_generator.py        # [Output] Generates pixel-perfect PDF reports using ReportLab.
│   ├── excel_generator.py      # [Output] Creates styled Excel sheets with conditional formatting.
│   ├── html_summary_generator.py # [Output] Generates HTML snippets for the preview page.
│   ├── chart_image_generator.py # [Visualization] Creates bar/pie charts using Matplotlib.
│   └── utilities.py            # [Helper] Shared utility functions (e.g., file cleanup).
├── static/                     # [Assets] Frontend resources.
│   ├── style.css               # [Styling] Custom CSS including the detailed theming system.
│   ├── theme.js                # [Logic] Handles time-based themes (Morning/Afternoon/Evening) and Dark Mode.
│   ├── owl.js                  # [Interactive] Logic for the "Night Mode" interactive owl animation.
│   └── upload.js               # [Logic] Drag-and-drop file upload handling and progress bars.
├── templates/                  # [Views] Jinja2 HTML templates.
│   ├── index.html              # [View] Landing page with upload zone and dashboard.
│   ├── preview.html            # [View] Results page showing analysis summary and download options.
│   └── view_file.html          # [View] Simple file viewer.
└── uploads/                    # [Storage] Temporary storage for uploaded and generated files.
```

---

## 3. Deep Dive: Core Modules

### 🧠 `modules/data_processor.py` (The Parser)
This is the most critical module. It doesn't just read a CSV; it "understands" the specific, often messy format of ERP exports.
*   **Intelligent Header Search:** It scans the first 20 rows of a file to locate the actual header row (looking for keywords like "Sr. No." or "Roll No."), skipping metadata at the top.
*   **Metadata Extraction:** It scrapes context from the top rows (Department, Semester, Batch, Subject) before the actual table starts.
*   **Data Cleaning:** Handles missing values, converts string percentages (e.g., "75%") to floats, and normalizes student names.
*   **Business Logic:**
    *   Calculates `Total Students`.
    *   Identifies `Zero Attendance` cases.
    *   Flags `Below 75%` (At Risk) students.
    *   Segments data into buckets (Below 50%, 50-75%, Above 75%).

### 📑 `modules/report_manager.py` (The Manager)
Acts as the bridge between the Flask app and the specialized generators.
*   **Workflow:**
    1.  Receives a file path.
    2.  Calls `data_processor` to get a clean DataFrame.
    3.  Calls `chart_image_generator` to create visual assets.
    4.  Calls `pdf_generator`, `excel_generator`, and `html_summary_generator` in parallel (conceptually) to create all outputs.
    5.  Returns a dictionary of paths to the generated files.

### 📉 `modules/chart_image_generator.py` (The Artist)
Uses `matplotlib` to generate static images for the PDF report.
*   **Charts Created:**
    *   **Attendance Distribution (Pie Chart):** Visualizes the proportion of students in varying attendance brackets.
    *   **Category Breakdown (Bar Chart):** Compares "Safe" vs. "At Risk" counts.
*   **Optimization:** Uses the `Agg` backend to generate images without needing a display server (crucial for server deployments).

---

## 4. Frontend Architecture & Theming

The frontend is not just a form; it's a rich, interactive experience built with **Bootstrap 5** and custom JavaScript.

### 🎨 Theming Engine (`static/theme.js` & `static/style.css`)
The application features a robust **Time-Based Theming System**:
1.  **Morning (6 AM - 12 PM):** Bright, fresh colors (Teal/Green).
2.  **Afternoon (12 PM - 5 PM):** Warm, energetic tones (Orange/Yellow).
3.  **Evening (5 PM - 9 PM):** Calm, sophisticated gradients (Deep Purple/Blue).
4.  **Night Mode (9 PM - 6 AM or Manual Toggle):** High-contrast Dark Mode with specific overrides for all Bootstrap components.

**Key Mechanics:**
*   **CSS Variables (`:root`):** All colors are defined as variables (e.g., `--theme-bg-primary`). Changing the theme class on `<body>` (e.g., `.theme-morning`) instantly repaints the entire app by redefining these variables.
*   **Interactive Owl (`static/owl.js`):** A complex CSS/JS animation that appears only in Night Mode. It tracks mouse movement with its eyes and has "Sleep" and "Angry" states based on user interaction.

---

## 5. Data Flow Lifecycle

1.  **Upload:**
    *   User drags a CSV file to `index.html`.
    *   `upload.js` sends it via POST to `/upload`.
    *   `app.py` saves it to `uploads/`.

2.  **Processing:**
    *   `app.py` triggers `ReportManager`.
    *   `DataProcessor` reads the CSV, cleans it, and extracts stats.
    *   `ChartGenerator` creates `.png` graphs of the stats.

3.  **Generation:**
    *   `PDFGenerator` compiles text, tables, and graphs into a PDF.
    *   `ExcelGenerator` writes the raw data into an `.xlsx` with conditional formatting (Red for <75%).
    *   `HTMLSummaryGenerator` creates a quick preview snippet.

4.  **Presentation:**
    *   User is redirected to `preview.html`.
    *   They see the HTML summary and "Download" buttons for the generated PDF and Excel files.

---

## 6. Setup & Installation

**Prerequisites:** Python 3.9+

1.  **Clone the Repository:**
    ```bash
    git clone <repo_url>
    cd -erp-report-automation
    ```

2.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the Application:**
    ```bash
    python app.py
    ```

4.  **Access:**
    Open `http://127.0.0.1:5000` in your web browser.

---

**Developer Credits:**
*   **Developers:** Shashank Dubey, Vineet Kumar
*   **Mentor:** Dr. Mamta Arora
*   **Institution:** Manav Rachna University
