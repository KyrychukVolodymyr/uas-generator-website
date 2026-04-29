import pandas as pd
import argparse
import re
from datetime import datetime
from collections import OrderedDict

TT_ORDER = [
    "cognition",
    "meal preparation",
    "ordinary housework",
    "managing medications",
    "shopping",
    "personal hygiene",
    "dressing upper body",
    "dressing lower body",
    "bathing",
    "locomotion",
    "transfer toilet",
    "toilet use",
    "eating",
]

SCORE_MAP = OrderedDict([
    ("total dependence", 6),
    ("maximal assistance", 5),
    ("extensive assistance", 4),
    ("limited assistance", 3),
])

TASK_PATTERNS = OrderedDict([
    ("meal preparation", ["meal preparation - performance:", "meal preparation"]),
    ("ordinary housework", ["ordinary housework - performance:", "ordinary housework"]),
    ("managing medications", ["managing medications - performance:", "managing medications"]),
    ("shopping", ["shopping - performance:", "shopping"]),
    ("personal hygiene", ["personal hygiene - performance:", "personal hygiene/grooming", "personal hygiene"]),
    ("dressing upper body", ["dressing upper body - performance:", "upper body dressing - performance:", "dressing upper body"]),
    ("dressing lower body", ["dressing lower body - performance:", "lower body dressing - performance:", "dressing lower body"]),
    ("bathing", ["bathing - performance:", "bathing - how takes bath or shower", "bathing"]),
    ("locomotion", ["locomotion - performance:", "locomotion"]),
    ("transfer toilet", ["transfer toilet - performance:", "transfer toilet", "transfer"]),
    ("toilet use", ["toilet use - performance:", "toilet use"]),
    ("eating", ["eating - performance:", "eating - how eats and drinks", "eating/feeding"]),
])

NEGATIVE_PHRASES = [
    "not present",
    "not ordered and did not occur",
    "not ordered",
    "did not occur",
    "none of the above",
]

COGNITIVE_DIAG_TERMS = [
    "alzheimer's disease",
    "dementia other than alzheimer's disease",
    "unspecified dementia",
    "cognitive decline",
    "schizophrenia",
    "bipolar",
]

POSITIVE_DIAG_STATUS_TERMS = [
    "diagnosis present",
    "present, receiving active treatment",
    "present receiving active treatment",
    "present, no treatment",
    "present no treatment",
    "monitored",
]

def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def row_values(row):
    return [clean_text(x) for x in row]

def row_join(row):
    vals = [v for v in row_values(row) if v]
    return " | ".join(vals)

def extract_after_label(text, label):
    m = re.search(re.escape(label) + r"\s*(.*)", text, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def contains_any(text, phrases):
    low = text.lower()
    return any(p in low for p in phrases)

def normalize_score(text):
    low = text.lower()
    for phrase, score in SCORE_MAP.items():
        if phrase in low:
            return score
    return None

def parse_header(df):
    patient_name = ""
    dob = ""
    cin = ""
    assessment_date = ""

    for _, row in df.head(15).iterrows():
        vals = row_values(row)
        for v in vals:
            v_clean = v.strip()
            if not patient_name and re.match(r"^[A-Za-z' -]+,\s*[A-Za-z' -]+$", v_clean):
                patient_name = v_clean.replace(",", " ").replace("  ", " ").strip()
            if not dob and "Date of Birth:" in v_clean:
                dob = extract_after_label(v_clean, "Date of Birth:")
            if not cin and "Medicaid ID:" in v_clean:
                cin = extract_after_label(v_clean, "Medicaid ID:")
            if not assessment_date and "Assessment Date:" in v_clean:
                assessment_date = extract_after_label(v_clean, "Assessment Date:")

    patient_name_file = patient_name.replace(" ", "_").upper()

    return {
        "patient_name_raw": patient_name,
        "patient_name_file": patient_name_file,
        "dob": dob,
        "cin": cin,
        "assessment_date": assessment_date,
        "today_date": datetime.now().strftime("%m/%d/%Y"),
    }

def parse_tasks(df):
    tasks = {}

    for _, row in df.iterrows():
        joined = row_join(row)
        low = joined.lower()
        for task_name, patterns in TASK_PATTERNS.items():
            if any(p in low for p in patterns):
                score = normalize_score(low)
                if score is not None:
                    tasks[task_name] = score

    cognition_score = None
    for _, row in df.iterrows():
        low = " ".join(str(x).lower() for x in row if str(x) != "nan")

        if "decision" in low and ("cognitive" in low or "making" in low):
            if "no discernable consciousness" in low or "no discernible consciousness" in low:
                cognition_score = 6
            elif "severely impaired" in low:
                cognition_score = 6
            elif "moderately impaired" in low:
                cognition_score = 5
            elif "minimally impaired" in low:
                cognition_score = 4
            elif "modified independent" in low or "independent" in low:
                cognition_score = 3

    for _, row in df.iterrows():
        low = row_join(row).lower()
        if (
            "making decisions regarding tasks of daily life" in low
            or "cognitive skills for daily decision making" in low
            or "decision making" in low
        ):
            if "no discernable consciousness" in low or "no discernible consciousness" in low:
                cognition_score = 6
            elif "severely impaired" in low:
                cognition_score = 6
            elif "moderately impaired" in low:
                cognition_score = 5
            elif "minimally impaired" in low:
                cognition_score = 4
            elif "modified independent" in low or "independent" in low:
                cognition_score = 3

    if cognition_score is not None:
        tasks["cognition"] = cognition_score

    return tasks

def parse_incontinence(df):
    bladder = None
    bowel = None

    for _, row in df.iterrows():
        low = row_join(row).lower()
        if "bladder continence:" in low:
            bladder = low
        if "bowel continence:" in low:
            bowel = low

    def is_incontinent_line(line):
        if not line:
            return False
        return "incontinent" in line

    incontinent = is_incontinent_line(bladder) or is_incontinent_line(bowel)
    return incontinent, bladder, bowel

def collect_positive_diagnoses(df):
    positives = []
    for _, row in df.iterrows():
        joined = row_join(row)
        low = joined.lower()
        if "diagnos" not in low and "disease" not in low and "disorder" not in low and "condition" not in low:
            continue
        if contains_any(low, NEGATIVE_PHRASES):
            continue
        if any(status in low for status in POSITIVE_DIAG_STATUS_TERMS):
            positives.append(joined)

    seen = set()
    out = []
    for x in positives:
        xl = x.lower()
        if xl not in seen:
            seen.add(xl)
            out.append(x)
    return out

def collect_med_rows(df):
    med_rows = []
    start_idx = None

    for i, row in df.iterrows():
        joined = row_join(row)
        low = joined.lower()
        if "drug name" in low and "dose" in low and "route" in low and "frequency" in low:
            start_idx = i + 1
            break

    if start_idx is None:
        return med_rows

    for i in range(start_idx, min(start_idx + 120, len(df))):
        joined = row_join(df.iloc[i])
        low = joined.lower()
        if not joined.strip():
            continue
        if "allergy to any drug" in low:
            break
        if "section " in low and "medication" not in low:
            break
        med_rows.append(joined)

    filtered = []
    for row in med_rows:
        low = row.lower()
        if "no known drug allergies" in low:
            continue
        if low.strip() == "drug name | dose | unit | route | frequency | prn":
            continue
        filtered.append(row)

    return filtered

def count_medications(med_rows):
    count = 0
    for row in med_rows:
        low = row.lower()
        if low.count("|") >= 2:
            first_part = row.split("|")[0].strip()
            if first_part and first_part.lower() not in ["drug name", "allergy to any drug"]:
                count += 1
                continue
        if any(x in low for x in [
            "tablet", "capsule", "solution", "cream", "ointment", "patch",
            "unit", "units", "mg", "mcg", "ml", "drop", "drops", "spray", "syrup",
            "insulin", "hfa", "diskus", "respimat"
        ]):
            count += 1
    return count

def find_rows(df, keyword):
    out = []
    for _, row in df.iterrows():
        joined = row_join(row)
        if keyword.lower() in joined.lower():
            out.append(joined)
    return out

def section_k_status(df, label):
    for _, row in df.iterrows():
        joined = row_join(row)
        low = joined.lower()
        if label.lower() in low:
            if "no selection" in low:
                return "no selection"
            if "not ordered and did not occur" in low:
                return "not ordered and did not occur"
            if "ordered, not implemented" in low:
                return "ordered, not implemented"
            if "1-2 days of last 3 days" in low:
                return "1-2 days of last 3 days"
            if "daily in last 3 days" in low:
                return "daily in last 3 days"
    return ""

def has_positive_insulin(med_rows):
    text = " ".join(med_rows).lower()
    return "insulin" in text

def has_positive_oxygen(df):
    status = section_k_status(df, "oxygen therapy")
    return status not in ["", "no selection", "not ordered and did not occur"]

def has_positive_wound(df):
    status = section_k_status(df, "wound care")
    return status not in ["", "no selection", "not ordered and did not occur"]

def build_cdpas(df, med_rows):
    triggers = []
    if has_positive_insulin(med_rows):
        triggers.append("Insulin")
    if has_positive_oxygen(df):
        triggers.append("Oxygen")
    if has_positive_wound(df):
        triggers.append("Wound Care")
    return triggers

def falls_positive(df):
    rows = find_rows(df, "falls")
    for r in rows:
        low = r.lower()
        if "no fall in last 90 days" in low:
            continue
        if "no fall in last 30 days, but fell 31-90 days ago" in low:
            return True
        if "one fall in last 30 days" in low:
            return True
        if "two or more falls in last 30 days" in low:
            return True
    return False

def pain_positive(df, positive_diagnoses):
    rows = find_rows(df, "pain")
    for r in rows:
        low = r.lower()
        if "no pain" in low:
            continue
        if "pain frequency" in low or "pain intensity" in low or "pain interferes" in low:
            return True
        if "pain" in low and not contains_any(low, NEGATIVE_PHRASES):
            return True
    for p in positive_diagnoses:
        if "pain" in p.lower():
            return True
    return False

def cognitive_positive(df, positive_diagnoses):
    for _, row in df.iterrows():
        low = row_join(row).lower()
        if "making decisions regarding tasks of daily life" in low:
            if "moderately impaired" in low or "severely impaired" in low:
                return True
    for p in positive_diagnoses:
        low = p.lower()
        if any(term in low for term in COGNITIVE_DIAG_TERMS) and any(status in low for status in POSITIVE_DIAG_STATUS_TERMS):
            return True
    return False

def build_fra(header, incontinent, positive_diagnoses, med_rows, df):
    score = 0
    reasons = []
    if header["dob"]:
        score += 1
        reasons.append("Age 65+")
    if len(positive_diagnoses) >= 3:
        score += 1
        reasons.append("3+ diagnoses")
    if falls_positive(df):
        score += 1
        reasons.append("Falls")
    if incontinent:
        score += 1
        reasons.append("Incontinence")
    score += 1
    reasons.append("Vision")
    score += 1
    reasons.append("Mobility / function impairment")
    med_count = count_medications(med_rows)
    if med_count >= 4:
        score += 1
        reasons.append("Polypharmacy")
    if pain_positive(df, positive_diagnoses):
        score += 1
        reasons.append("Pain")
    if cognitive_positive(df, positive_diagnoses):
        score += 1
        reasons.append("Cognitive impairment diagnosis")
    return score, reasons, med_count, len(positive_diagnoses)

def print_preview(header, tasks, incontinent, bladder, bowel, cdpas, fra_score, fra_reasons, med_count, diag_count, df):
    print("\\nPatient:", header["patient_name_raw"])
    print("DOB:", header["dob"])
    print("CIN:", header["cin"])
    print("Assessment Date in CSV:", header["assessment_date"])
    print("Today's Date:", header["today_date"])

    print("\\n=== TT PREVIEW ===")
    for task in TT_ORDER:
        if task in tasks:
            print(f"{task}: {tasks[task]}")

    print("\\nBladder line:", bladder if bladder else "(not found)")
    print("Bowel line:", bowel if bowel else "(not found)")
    print("Laundry row:", "Incontinent Patient" if incontinent else "Continent Patient")

    print("\\n=== CDPAS TRIGGERS ===")
    if cdpas:
        for x in cdpas:
            print("-", x)
    else:
        print("(none)")

    print("Oxygen status used:", section_k_status(df, "oxygen therapy") or "(not found)")
    print("Wound care status used:", section_k_status(df, "wound care") or "(not found)")

    print("\\n=== FRA ===")
    print("Score:", fra_score)
    print("Reasons:", ", ".join(fra_reasons) if fra_reasons else "None")
    print("Diagnosis count used:", diag_count)
    print("Medication count used:", med_count)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True)
    parser.add_argument("--membership-id", required=True)
    args = parser.parse_args()

    df = pd.read_csv(args.csv, header=None)

    header = parse_header(df)
    tasks = parse_tasks(df)
    incontinent, bladder, bowel = parse_incontinence(df)
    positive_diagnoses = collect_positive_diagnoses(df)
    med_rows = collect_med_rows(df)
    cdpas = build_cdpas(df, med_rows)
    fra_score, fra_reasons, med_count, diag_count = build_fra(
        header, incontinent, positive_diagnoses, med_rows, df
    )

    print_preview(
        header, tasks, incontinent, bladder, bowel,
        cdpas, fra_score, fra_reasons, med_count, diag_count, df
    )

    confirm = input("\\nConfirm? (y/n): ").strip().lower()
    if confirm != "y":
        print("Cancelled")
        return

    fname_base = f"ANTHEM_MLTC-{args.membership_id}-{header['patient_name_file']}-{header['cin']}"
    print("\\nWill generate:")
    print(fname_base + "-TT.xlsx")
    print(fname_base + "-CDPAS.xlsx")
    print(fname_base + "-FRA.pdf")

if __name__ == "__main__":
    main()
