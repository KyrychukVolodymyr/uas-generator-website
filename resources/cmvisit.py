from pathlib import Path
from io import BytesIO
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject
from commercial_profile import get_assessor_display_name

CMVISIT_TEMPLATE_CANDIDATES = ["CM_InPerson_Visit_Template.pdf"]

def find_cmvisit_template(app_dir=None, template_candidates=None):
    app_dir = Path(app_dir) if app_dir else Path(__file__).resolve().parent
    template_candidates = template_candidates or CMVISIT_TEMPLATE_CANDIDATES
    for name in template_candidates:
        path = app_dir / name
        if path.exists():
            return path
    for path in app_dir.glob("*6monthCMVisit.pdf"):
        if path.is_file():
            return path
    return None

def create_cmvisit_pdf(
    template_path: Path,
    output_path: Path,
    first: str,
    last: str,
    numeric_id: str,
    medicaid_id: str,
    assessment_date: str,
    equipment=None,
    dwelling_type="apartment_elevator",
    bedrooms_count="2",
    accessible_home="no",
    member_name_value=None,
):
    if equipment is None:
        equipment = {}

    reader = PdfReader(str(template_path))
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)

    try:
        acro = reader.trailer["/Root"].get("/AcroForm")
        if acro:
            writer._root_object.update({NameObject("/AcroForm"): acro})
            writer._root_object["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})
    except Exception:
        try:
            if "/AcroForm" in writer._root_object:
                writer._root_object["/AcroForm"].update({NameObject("/NeedAppearances"): BooleanObject(True)})
        except Exception:
            pass

    if not member_name_value:
        if last and first:
            member_name_value = f"{last}_{first}"
        else:
            member_name_value = f"{first}_{last}".strip("_")

    bedrooms_count = str(bedrooms_count).strip() if bedrooms_count is not None else "2"
    if not bedrooms_count:
        bedrooms_count = "2"

    accessible_home = str(accessible_home or "no").strip().lower()
    if accessible_home not in {"yes", "no"}:
        accessible_home = "no"

    field_values = {
        "Member name": member_name_value,
        "Member ID": str(numeric_id),
        "Member CIN No": str(medicaid_id),
        "Number of bedrooms please specify": bedrooms_count,
        "Completed by": get_assessor_display_name(),
        "Visit_date": str(assessment_date),
        "Grab bar  If yes specify grab bar": "shower" if equipment.get("grab_bar") else "",
        "Grab bar If yes specify grab bar": "shower" if equipment.get("grab_bar") else "",
        "Document any falls hospitalization ER visit urgent care report concerns and so onRow1": "",
    }

    for page in writer.pages:
        writer.update_page_form_field_values(page, field_values)

    def collect_widgets(pdf_reader):
        widget_map = {}
        for page_index, page in enumerate(pdf_reader.pages):
            annots = page.get("/Annots") or []
            try:
                annots = annots.get_object()
            except Exception:
                pass

            for annot_ref in annots:
                try:
                    annot = annot_ref.get_object()
                except Exception:
                    continue

                if str(annot.get("/Subtype")) != "/Widget":
                    continue

                parent = annot.get("/Parent")
                field = parent.get_object() if parent else annot

                field_name = field.get("/T") or annot.get("/T")
                if not field_name:
                    continue
                field_name = str(field_name)

                rect = annot.get("/Rect")
                if not rect or len(rect) != 4:
                    continue

                try:
                    rect_vals = [float(x) for x in rect]
                except Exception:
                    continue

                export_state = None
                ap = annot.get("/AP")
                n = ap.get("/N") if ap else None
                try:
                    keys = list(n.keys()) if n else []
                except Exception:
                    keys = []

                non_off = [str(k).lstrip("/") for k in keys if str(k) != "/Off"]
                if non_off:
                    export_state = non_off[0]

                widget_map.setdefault(field_name, []).append({
                    "page": page_index,
                    "rect": rect_vals,
                    "state": export_state,
                    "annot": annot,
                    "field": field,
                })
        return widget_map

    widget_map = collect_widgets(writer)

    desired_states = {
        "Clutter": "Yes",
        "Unobstructed": "Yes",
        "Lighting": "Yes",
        "Flooring": "Yes",
        "Handrails": "Yes",
        "Nonslip": "Yes",
        "Insects": "Yes",
        "Chemicals": "Yes",
        "Pets": "No",
        "Nonskid-mat": "Yes",
        "Reach": "Yes",
        "Nightlights": "Yes",
        "Smoke": "Yes",
        "Smoke2": "No",
        "Pers": "Yes",
        "Emerg-plan": "Yes",
        "Oxygen": "Yes" if equipment.get("oxygen") else "N/A",
        "911": "Yes",
        "Commode": "Yes" if equipment.get("bedside_commode") else "No",
        "Cane": "Yes" if equipment.get("cane") else "No",
        "Glucometer": "No",
        "Grab-bar": "Yes" if equipment.get("grab_bar") else "No",
        "Hosp-bed": "Yes" if equipment.get("hospital_bed") else "No",
        "Hoyer": "Yes" if equipment.get("hoyer_lift") else "No",
        "Motor-chair": "No",
        "Oxygen2": "Yes" if equipment.get("oxygen") else "No",
        "Toilet-seat": "Yes" if equipment.get("raised_toilet_seat") else "No",
        "Ramp": "No",
        "Shower": "Yes" if equipment.get("shower_chair") else "No",
        "Stairlift": "No",
        "Walker": "Yes" if equipment.get("walker") else "No",
        "Wheelchair": "Yes" if equipment.get("wheelchair") else "No",
        "Other": "No",
        "Single family house": "Off",
        "Multifamily house": "Off",
        "Stairs outside of dwelling": "Off",
        "Apartment building": "Off",
        "Stairs inside of dwelling": "Off",
        "Elevator": "Off",
        "Is the home accessible for people with disabilities?": "On" if accessible_home == "yes" else "Off",
        "Is the home accessible for people with disabilities": "On" if accessible_home == "yes" else "Off",
        "Accessible-home": "On" if accessible_home == "yes" else "Off",
    }

    if dwelling_type == "single_family_outside":
        desired_states["Single family house"] = "On"
        desired_states["Stairs outside of dwelling"] = "On"
    elif dwelling_type == "single_family_inside":
        desired_states["Single family house"] = "On"
        desired_states["Stairs inside of dwelling"] = "On"
    elif dwelling_type == "apartment_outside":
        desired_states["Apartment building"] = "On"
        desired_states["Stairs outside of dwelling"] = "On"
    else:
        desired_states["Apartment building"] = "On"
        desired_states["Elevator"] = "On"

    equipment_fields = {
        "Commode",
        "Cane",
        "Grab-bar",
        "Hosp-bed",
        "Hoyer",
        "Oxygen2",
        "Toilet-seat",
        "Shower",
        "Walker",
        "Wheelchair",
    }

    def set_group_state(field_name, desired_value):
        widgets = widget_map.get(field_name, [])
        if not widgets:
            return False

        normalized = str(desired_value).strip().lstrip("/")
        matched = False

        for widget in widgets:
            annot = widget.get("annot")
            field = widget.get("field")
            state = str(widget.get("state") or "").strip().lstrip("/")

            if state == normalized:
                try:
                    annot[NameObject("/AS")] = NameObject(f"/{state}")
                except Exception:
                    pass
                try:
                    field[NameObject("/V")] = NameObject(f"/{state}")
                except Exception:
                    pass
                matched = True
            else:
                try:
                    annot[NameObject("/AS")] = NameObject("/Off")
                except Exception:
                    pass

        return matched

    for field_name in equipment_fields:
        desired_value = desired_states.get(field_name)
        if desired_value in {"Yes", "No"}:
            set_group_state(field_name, desired_value)

    def draw_x(c, rect):
        x0, y0, x1, y1 = rect
        w = abs(x1 - x0)
        h = abs(y1 - y0)
        size = max(8, min(14, h * 0.9))
        c.setFont("Helvetica-Bold", size)
        c.drawCentredString(x0 + w / 2.0, y0 + (h * 0.08), "X")

    for page_index, page in enumerate(writer.pages):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)

        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))

        for field_name, desired_value in desired_states.items():
            if field_name not in widget_map:
                continue

            # v018 fix:
            # Do NOT skip DME/equipment fields during overlay drawing.
            # Adobe Acrobat may not reliably render checkbox widget appearances.
            # A real visible X on the page makes equipment selections stable.
            # if field_name in equipment_fields:
            #     continue

            for widget in widget_map[field_name]:
                if widget["page"] != page_index:
                    continue

                state = widget["state"]
                should_draw = False

                if desired_value == "On":
                    if state in (None, "", "On", "Yes"):
                        should_draw = True
                elif desired_value == "Off":
                    should_draw = False
                else:
                    if state == desired_value:
                        should_draw = True

                if should_draw:
                    draw_x(c, widget["rect"])

        c.save()
        packet.seek(0)
        overlay_reader = PdfReader(packet)
        if len(overlay_reader.pages) > 0:
            page.merge_page(overlay_reader.pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)
