import re

# -------------------------
# Utility helpers
# -------------------------
def safe_str(x) -> str:
    """Safely converts a value to a string, returning empty for None."""
    return "" if x is None else str(x)


def clean_subject_label(label: str) -> str:
    """Strips trailing metrics like '- %', '- Total', '(TUT)' etc. from a subject name."""
    if label is None:
        return ""
    s = str(label).strip()
    # Remove patterns like " - % (PP)", " - Total - something", "(PR)"
    s = re.sub(r"\s*-\s*%.*$", "", s)  # "- % ..." and suffixes
    s = re.sub(r"\s*-\s*Total.*$", "", s)  # "- Total ..."
    s = re.sub(r"\s*\(.*\)$", "", s)  # trailing parentheses
    s = re.sub(r"\s*-\s*$", "", s)  # trailing hyphen
    return s.strip()

def format_attendance_value(value):
    """Formats a numeric value to remove .0 if it's an integer, otherwise keeps float."""
    try:
        f_val = float(value)
        if f_val.is_integer():
            return int(f_val)
        return f_val
    except (ValueError, TypeError):
        return value
