import logging
import textwrap
from io import BytesIO
import pandas as pd
from matplotlib.figure import Figure

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# -------------------------
# Chart generation
# -------------------------
def generate_chart_image(df: pd.DataFrame) -> BytesIO:
    """
    Generate a PNG Bar chart as BytesIO where x=courses and y=students below 75%.
    """
    subject_columns = [col for col in df.columns if col not in [
        "Sr No.", "Roll No", "Student Name", "Overall %age of all subjects from ERP report"
        ,"Count of Courses with attendance below minimum attendance criteria",
        "Whether Critical"
    ]]

    courses = subject_columns
    students_below_75 = [(pd.to_numeric(df[c], errors="coerce") < 75).sum() for c in courses]

    # Wrap long course names
    wrapped_courses = [textwrap.fill(course, 15) for course in courses]

    fig = Figure(figsize=(12, 6))
    ax = fig.subplots()
    bars = ax.bar(wrapped_courses, students_below_75)
    ax.set_title("Number of Students with Attendance Below 75% per Course")
    ax.set_xlabel("Courses")
    ax.set_ylabel("Number of Students below 75%")
    ax.tick_params(axis="x", rotation=45, labelsize=8)
    for b in bars:
        yval = b.get_height()
        ax.text(b.get_x() + b.get_width() / 2.0, yval, str(int(yval)), va="bottom", ha="center", fontsize=8)
    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png")
    buf.seek(0)
    return buf
