from __future__ import annotations

import json
import os
import platform
import uuid
import urllib.request
import urllib.error
import ssl
import certifi
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

APP_DIR = Path(__file__).resolve().parent
if os.name == "nt":
    RUNTIME_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2"
else:
    RUNTIME_DIR = Path.home() / "Library" / "Application Support" / "UASGenerator2026V2"
RUNTIME_DIR.mkdir(parents=True, exist_ok=True)

DEVICE_PATH = RUNTIME_DIR / "device.json"
LICENSE_PATH = RUNTIME_DIR / "license.json"
USER_PROFILE_PATH = RUNTIME_DIR / "user_profile.json"
SERVER_CONFIG_PATH = APP_DIR / "license_server_config.json"
PRODUCTION_SERVER_URL = "https://uas-license-server.onrender.com"


def _read_json(path: Path, default: Any) -> Any:
    try:
        if not path.exists():
            return default
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def get_server_url() -> str:
    cfg = _read_json(SERVER_CONFIG_PATH, {})
    url = str(cfg.get("server_url") or PRODUCTION_SERVER_URL).strip().rstrip("/")
    if not url:
        url = PRODUCTION_SERVER_URL
    blocked_hosts = ["local" + "host", "127" + ".0.0.1", "0" + ".0.0.0"]
    insecure_scheme = "ht" + "tp://"
    if any(x in url for x in blocked_hosts) or url.startswith(insecure_scheme):
        raise RuntimeError("Production build blocked: invalid local or insecure license server URL.")
    return url


def get_or_create_device_id() -> str:
    existing = _read_json(DEVICE_PATH, {})
    device_id = str(existing.get("device_id", "")).strip()
    if device_id:
        return device_id
    device_id = str(uuid.uuid4())
    _write_json(DEVICE_PATH, {
        "device_id": device_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "machine": platform.machine(),
        "system": platform.system(),
        "node": platform.node(),
    })
    return device_id


def _post_json(url: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    context = ssl.create_default_context(cafile=certifi.where())
    try:
        with urllib.request.urlopen(req, timeout=30, context=context) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            data = json.loads(raw) if raw else {}
        except Exception:
            data = {"raw": raw}
        raise RuntimeError(f"License server error {exc.code}: {data.get('detail', data)}")
    except Exception as exc:
        raise RuntimeError(f"License server request failed: {exc}")


def _parse_expiry(value: str):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _license_data(data: Dict[str, Any]) -> Dict[str, Any]:
    nested = data.get("license", data)
    return nested if isinstance(nested, dict) else {}


def _local_block_reason(data: Dict[str, Any]) -> str:
    license_data = _license_data(data)
    status = str(license_data.get("status", "")).strip().lower()
    expires_at = str(license_data.get("expires_at", "")).strip()
    if bool(license_data.get("revoked", False)):
        return "License was revoked."
    if status in {"revoked", "blocked", "cancelled", "canceled", "inactive", "expired"}:
        return f"License status is {status}."
    expiry = _parse_expiry(expires_at)
    if expiry and expiry < datetime.now(timezone.utc):
        return "License subscription is expired."
    return ""


def activate_license(email: str, license_key: str, accepted_terms_version: str, app_version: str = "v1") -> Dict[str, Any]:
    email = str(email).strip()
    license_key = str(license_key).strip()
    if not email:
        raise RuntimeError("Email is required.")
    if not license_key:
        raise RuntimeError("License key is required.")
    device_id = get_or_create_device_id()
    data = _post_json(f"{get_server_url()}/activate-license", {
        "email": email,
        "license_key": license_key,
        "device_id": device_id,
        "accepted_terms_version": accepted_terms_version,
        "app_version": app_version,
    })
    record = {
        "email": email,
        "device_id": device_id,
        "license_key": license_key,
        "accepted_terms_version": accepted_terms_version,
        "app_version": app_version,
        "activated_at": datetime.now(timezone.utc).isoformat(),
        "server_url": get_server_url(),
        "license": data,
    }
    reason = _local_block_reason(record)
    _write_json(LICENSE_PATH, record)
    if reason:
        raise RuntimeError(reason)
    return record


def validate_license() -> Dict[str, Any]:
    existing = _read_json(LICENSE_PATH, {})
    email = str(existing.get("email", "")).strip()
    license_key = str(existing.get("license_key", "")).strip()
    device_id = str(existing.get("device_id", "")).strip() or get_or_create_device_id()
    if not email or not license_key or not device_id:
        return {"valid": False, "reason": "Device is not activated."}
    try:
        data = _post_json(f"{get_server_url()}/validate-license", {
            "email": email,
            "license_key": license_key,
            "device_id": device_id,
        })
    except Exception as exc:
        return {"valid": False, "reason": str(exc)}
    existing["license"] = data
    existing["validated_at"] = datetime.now(timezone.utc).isoformat()
    existing["server_url"] = get_server_url()
    _write_json(LICENSE_PATH, existing)
    reason = _local_block_reason(existing)
    if reason:
        return {"valid": False, "reason": reason, "license": data}
    if not bool(data.get("valid", True)):
        return {"valid": False, "reason": str(data.get("reason", "License validation failed.")), "license": data}
    return {"valid": True, "reason": "", "license": data}


def get_license_status() -> Dict[str, Any]:
    existing = _read_json(LICENSE_PATH, {})
    if not existing:
        return {"valid": False, "reason": "Device is not activated."}
    reason = _local_block_reason(existing)
    if reason:
        return {"valid": False, "reason": reason, "license": existing.get("license", {})}
    return validate_license()


def require_valid_license() -> Dict[str, Any]:
    status = get_license_status()
    if not status.get("valid"):
        raise RuntimeError("License check failed: " + status.get("reason", "Unknown license error."))
    return status




def require_active_license() -> dict:
    return require_valid_license()


def runtime_paths() -> Dict[str, str]:
    return {
        "runtime_dir": str(RUNTIME_DIR),
        "device_json": str(DEVICE_PATH),
        "license_json": str(LICENSE_PATH),
        "user_profile_json": str(USER_PROFILE_PATH),
        "server_config_json": str(SERVER_CONFIG_PATH),
    }



