# 📊 ERP Report Automator

A smart web application that transforms raw ERP attendance data into professional, insightful reports with just a few clicks.

## ✨ Features

-   **Intelligent Data Processing:** Automatically parses and processes complex CSV and Excel files from your ERP system.
-   **Dynamic Web Interface:** A modern, responsive interface with a multi-step workflow for a smooth user experience.
-   **Customizable Reports:** Easily configure report metadata like title, department, date ranges, and minimum attendance criteria.
-   **Live Preview:** Instantly generate an HTML preview of the report, including data tables, summaries, and charts, before downloading.
-   **Multiple Download Formats:**
    -   **Excel (`.xlsx`):** Generates a professionally formatted, single-sheet Excel dashboard with conditional formatting and analytical charts.
    -   **PDF:** Creates a clean and printable PDF version of the report.
-   **Dynamic Theming:**
    -   The UI automatically switches between **Morning**, **Afternoon**, and **Evening** themes based on the time of day.
    -   Includes an automatic **Dark Mode** for nighttime hours (9 PM - 6 AM).
    -   Users can manually override the theme at any time.

## 🛠️ Tech Stack

-   **Backend:** Python, Flask
-   **Data Processing:** Pandas
-   **Excel Generation:** Openpyxl
-   **PDF Generation:** (Depends on the library used, e.g., FPDF, WeasyPrint)
-   **Charting:** Matplotlib
-   **Frontend:** HTML, CSS, JavaScript, Bootstrap 5

## 🚀 Getting Started

### Prerequisites

-   Python 3.8+
-   Git

### Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/shashankdubey822-code/-erp-report-automation.git
    cd -erp-report-automation
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # For Windows
    python -m venv venv
    .\venv\Scripts\activate

    # For macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1.  **Run the application:**
    ```bash
    flask run
    ```

2.  Open your web browser and navigate to `http://127.0.0.1:5000`.

3.  Upload your ERP data file, configure the report options, and generate your report!

## 🔧 Configuration

The application can be configured using a `.env` file. Create a `.env` file in the root of the project (you can copy `.env.example`) to set the following variables:

-   `SECRET_KEY`: A secret key for Flask session management.
-   `MAX_CONTENT_LENGTH`: The maximum file size for uploads (e.g., `16 * 1024 * 1024` for 16MB).

## 🤝 Contributing

Contributions are welcome! If you have suggestions for improvements, please feel free to fork the repository and submit a pull request.

## 📄 License

This project is for educational purposes.
