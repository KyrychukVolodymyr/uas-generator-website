import json
import os
from datetime import datetime
from pathlib import Path
import re

APP_DIR = Path(__file__).resolve().parent
if os.name == "nt":
    RUNTIME_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2"
else:
    RUNTIME_DIR = Path.home() / "Library" / "Application Support" / "UASGenerator2026V2"
RUNTIME_DIR.mkdir(parents=True, exist_ok=True)

USER_PROFILE_PATH = RUNTIME_DIR / "user_profile.json"
LICENSE_PATH = RUNTIME_DIR / "license.json"
DEVICE_PATH = RUNTIME_DIR / "device.json"
TERMS_PATH = APP_DIR / "Terms_of_Use.txt"
LICENSE_TEXT_PATH = RUNTIME_DIR / "LICENSE.txt"

TERMS_VERSION = "2026-04-21-v2"

DEFAULT_DISPLAY_NAME = ""
DEFAULT_PLAIN_NAME = ""

def load_user_profile():
    if not USER_PROFILE_PATH.exists():
        return {}
    try:
        return json.loads(USER_PROFILE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_user_profile(profile: dict):
    USER_PROFILE_PATH.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")

def ensure_terms_files():
    if not TERMS_PATH.exists():
        TERMS_PATH.write_text("""Anthem BCBS Doc Generator - Terms of Use
Version: 2026-04-19-v1

1. This software is an assistive documentation tool only.
2. The software may generate incomplete, inaccurate, or unsuitable output.
3. The user must independently review, verify, and confirm every generated document before use, signature, or submission.
4. The user bears sole responsibility for any signed, submitted, relied-upon, or distributed documentation.
5. The software is provided as-is, without warranties of completeness, accuracy, merchantability, or fitness for a particular purpose.
6. This software and its templates, logic, prompts, and workflows are intellectual property of the developer.
7. Reverse engineering, redistribution, sublicensing, resale, decompilation, key sharing, and unauthorized copying are prohibited.
8. Licenses are personal, limited, revocable, and non-transferable.
9. The developer may suspend, pause, or revoke access for misuse, unauthorized sharing, reverse engineering, non-payment, or other breach of these terms.
10. No refunds are provided.
11. The software is intended to process files locally on the user device. Generated outputs are stored locally under the user's control.
12. The normal licensing workflow is not intended to remotely store patient PHI. Only license/account metadata may be stored remotely for activation, validation, support, and updates.
13. The developer may provide updates, compatibility fixes, activation checks, and support within the scope of the subscription plan.
14. Continued use of the software constitutes acceptance of the current Terms of Use version.
""", encoding="utf-8")

    if not LICENSE_TEXT_PATH.exists():
        LICENSE_TEXT_PATH.write_text("""Anthem BCBS Doc Generator - License Summary

This software is licensed, not sold.
Each license is issued to one named user.
Basic plan: 1 device maximum.
Pro plan: up to 2 devices maximum, same named user only.
All licenses are non-transferable and may be revoked for misuse, key sharing, reverse engineering, redistribution, or non-payment.
""", encoding="utf-8")

def get_assessor_display_name():
    profile = load_user_profile()

    first_name = str(profile.get("first_name", "")).strip()
    last_name = str(profile.get("last_name", "")).strip()
    display_name = str(profile.get("display_name", "")).strip()

    built_full = " ".join(x for x in [first_name, last_name] if x).strip()
    bad_values = {"", "RN", "LPN", "NP", "PA", "BSN", "MSN", "DNP"}

    if display_name and display_name.upper() not in bad_values:
        return display_name

    if built_full:
        return f"{built_full} RN"

    return DEFAULT_DISPLAY_NAME

def get_assessor_plain_name():
    profile = load_user_profile()

    first_name = str(profile.get("first_name", "")).strip()
    last_name = str(profile.get("last_name", "")).strip()
    plain_name = str(profile.get("plain_name", "")).strip()

    built_full = " ".join(x for x in [first_name, last_name] if x).strip()
    bad_values = {"", "RN", "LPN", "NP", "PA", "BSN", "MSN", "DNP"}

    if built_full:
        return built_full

    if plain_name and plain_name.upper() not in bad_values:
        return plain_name

    cleaned = re.sub(r"\b(RN|LPN|NP|PA|BSN|MSN|DNP)\b\.?$", "", get_assessor_display_name()).strip()
    return cleaned or DEFAULT_PLAIN_NAME


