import re
from datetime import datetime
import pandas as pd
from uas_automation import parse_header

def normalize_date_for_folder(mmddyyyy: str):
    try:
        dt = datetime.strptime(mmddyyyy, "%m/%d/%Y")
        return dt.strftime("%m-%d-%Y")
    except Exception:
        return mmddyyyy.replace("/", "-")

def extract_assessment_date_from_rows(rows):
    patterns = [
        re.compile(r'Assessment Date:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE),
        re.compile(r'Assessment Reference Date:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE),
    ]
    for row in rows:
        joined = " | ".join(cell.strip() for cell in row if cell and cell.strip())
        for pat in patterns:
            m = pat.search(joined)
            if m:
                return m.group(1)
    for row in rows:
        for i, cell in enumerate(row):
            value = str(cell).strip().lower()
            if value in {"assessment date:", "assessment reference date:"}:
                for j in range(i + 1, len(row)):
                    val = str(row[j]).strip()
                    if re.fullmatch(r'\d{2}/\d{2}/\d{4}', val):
                        return val
    raise RuntimeError("Could not extract Assessment Date from CSV.")

def extract_mode_of_assessment_from_rows(rows):
    for idx, row in enumerate(rows):
        cleaned = [str(c).strip() for c in row if str(c).strip()]
        joined = " | ".join(cleaned).lower()
        if "mode of assessment" in joined:
            window_parts = [joined]
            for extra in range(1, 5):
                if idx + extra < len(rows):
                    next_joined = " | ".join(str(c).strip() for c in rows[idx + extra] if str(c).strip()).lower()
                    if next_joined:
                        window_parts.append(next_joined)
            window = " | ".join(window_parts)
            if "in-person only" in window:
                return "in_person"
            if "interactive video teleconference only" in window:
                return "telehealth"
    full_text = "\n".join(" | ".join(str(c).strip() for c in row if str(c).strip()).lower() for row in rows)
    if re.search(r'mode of assessment.*in-person only', full_text, re.IGNORECASE):
        return "in_person"
    if re.search(r'mode of assessment.*interactive video teleconference only', full_text, re.IGNORECASE):
        return "telehealth"
    return "telehealth"

def extract_medicaid_from_rows(rows):
    code_pat = re.compile(r'\b[A-Z]{2}\d{5,6}[A-Z]\b')
    priority_labels = [
        "member cin", "member cin no", "cin", "cin no",
        "medicaid", "medicaid number", "medicaid no",
    ]
    for row in rows:
        for i, cell in enumerate(row):
            text = str(cell).strip()
            low = text.lower()
            if any(label in low for label in priority_labels):
                scan = row[i:i+6]
                joined = " | ".join(str(x).strip() for x in scan if str(x).strip())
                m = code_pat.search(joined.upper())
                if m:
                    return m.group(0)
    for row in rows:
        joined = " | ".join(str(x).strip() for x in row if str(x).strip())
        m = code_pat.search(joined.upper())
        if m:
            return m.group(0)
    return None

def extract_dob_from_rows(rows):
    patterns = [
        re.compile(r"Date of Birth:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE),
        re.compile(r"DOB:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE),
    ]
    for row in rows:
        joined = " | ".join(str(cell).strip() for cell in row if str(cell).strip())
        for pat in patterns:
            m = pat.search(joined)
            if m:
                return m.group(1)
    for row in rows:
        for i, cell in enumerate(row):
            value = str(cell).strip().lower()
            if value in {"date of birth:", "dob:"}:
                for j in range(i + 1, len(row)):
                    val = str(row[j]).strip()
                    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", val):
                        return val
    return ""

def extract_member_name_file_from_rows(rows):
    try:
        df = pd.DataFrame(rows)
        header = parse_header(df)
        return (header.get("patient_name_file") or "").strip().upper()
    except Exception:
        return ""

def extract_member_name_parts_from_rows(rows):
    name_file = extract_member_name_file_from_rows(rows)
    if not name_file:
        return "", "", ""
    parts = [p.strip().upper() for p in name_file.split("_") if p.strip()]
    if len(parts) >= 2:
        last = parts[0]
        first = "_".join(parts[1:])
        return first, last, name_file
    if len(parts) == 1:
        return "", parts[0], name_file
    return "", "", ""

def extract_member_id_from_rows(rows):
    label_patterns = [
        re.compile(r'^member id:?$', re.IGNORECASE),
        re.compile(r'^member number:?$', re.IGNORECASE),
        re.compile(r'^member no:?$', re.IGNORECASE),
        re.compile(r'^member identifier:?$', re.IGNORECASE),
    ]

    for row in rows:
        vals = [str(c).strip() for c in row if str(c).strip()]
        if not vals:
            continue
        for i, val in enumerate(vals):
            if any(p.match(val) for p in label_patterns) or "member id" in val.lower():
                scan = vals[i:i+6]
                joined = " | ".join(scan)
                m = re.search(r'\b\d{4,12}\b', joined)
                if m:
                    return m.group(0)

    joined_all = "\n".join(" | ".join(str(c).strip() for c in row if str(c).strip()) for row in rows)
    m = re.search(r'Member ID\s*:?\s*(\d{4,12})', joined_all, re.IGNORECASE)
    if m:
        return m.group(1)

    return None
