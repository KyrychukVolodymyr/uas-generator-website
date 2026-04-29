from pathlib import Path
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject
from commercial_profile import get_assessor_display_name, get_assessor_plain_name

NEGATIVE_VALUES = {
    "no selection",
    "not in last 3 days",
    "person could not (would not) respond",
    "person could not respond",
    "person would not respond",
}

POSITIVE_VALUES = {
    "not in last 3 days, but often feels that way",
    "in 1-2 days of last 3 days",
    "daily in last 3 days",
}

def _safe_cell(rows, row_idx_1_based, col_idx_0_based):
    try:
        return str(rows[row_idx_1_based - 1][col_idx_0_based]).strip()
    except Exception:
        return ""


def should_generate_phq9_from_rows(rows):
    try:
        blob = " ".join(
            " ".join(map(str, r)) if isinstance(r, (list, tuple)) else str(r)
            for r in rows
        ).lower()

        return "phq-9 score is" in blob
    except Exception as e:
        print("PHQ9 trigger error:", e)
        return False


def generate_phq9_pdf(template_path, output_path, member_name, member_id, assessment_date):
    template_path = Path(template_path)
    output_path = Path(output_path)

    reader = PdfReader(str(template_path))
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    if "/AcroForm" in reader.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
        })
        try:
            writer._root_object["/AcroForm"].update({
                NameObject("/NeedAppearances"): BooleanObject(True)
            })
        except Exception:
            pass

    fields = {
        "Member Name": member_name,
        "Member ID": str(member_id),
        "Date of Assessment": assessment_date,
        "Signature Date_af_date": assessment_date,
        "Date4_af_date": assessment_date,
        "Assessor's Name": get_assessor_plain_name(),
        "Assessor's Signature": get_assessor_display_name(),
        "Signature and Title": get_assessor_display_name(),
    }

    for page in writer.pages:
        writer.update_page_form_field_values(page, fields)

    with open(output_path, "wb") as f:
        writer.write(f)
