import json
import os
import subprocess
import sys
import urllib.error
import urllib.request
import ssl
from datetime import datetime, timezone
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
RUNTIME_DIR = Path.home() / "Library" / "Application Support" / "UASGenerator2026V2"
RUNTIME_DIR.mkdir(parents=True, exist_ok=True)

VENV_DIR = RUNTIME_DIR / "runtime_venv"
VENV_PYTHON = VENV_DIR / "bin" / "python3"

REPORT_JSON = RUNTIME_DIR / "support_report.json"
REPORT_TXT = RUNTIME_DIR / "support_report.txt"

REQUIRED_FILES = [
    "system_check.py",
    "activate_license.py",
    "license_manager.py",
    "generate_outputs.py",
    "app_mac.py",
    "bootstrap_runtime.py",
    "csv_extractors.py",
    "postprocess_tt.py",
    "phq9_patch.py",
    "commercial_profile.py",
    "cmvisit.py",
    "uas_automation.py",
    "tt_service_patch.py",
    "config.json",
    "license_server_config.json",
    "Terms_of_Use.txt",
    "README.txt",
    "TT_template.xlsx",
    "CDPAS_template.xlsx",
    "FRA_template.pdf",
    "PHQ9_template.pdf",
    "CM_InPerson_Visit_Template.pdf",
]

REQUIRED_PACKAGES = ["pandas", "openpyxl", "pypdf", "reportlab", "requests", "certifi"]

IMPORT_MAP = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "pypdf": "pypdf",
    "reportlab": "reportlab",
    "requests": "requests",
    "certifi": "certifi",
}


def certifi_cafile():
    try:
        import certifi
        return certifi.where()
    except Exception:
        pass

    try:
        if VENV_PYTHON.exists():
            p = subprocess.run(
                [str(VENV_PYTHON), "-c", "import certifi; print(certifi.where())"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=30,
            )
            if p.returncode == 0:
                value = p.stdout.strip()
                if value:
                    return value
    except Exception:
        pass

    return None


def ssl_context():
    cafile = certifi_cafile()
    if cafile:
        return ssl.create_default_context(cafile=cafile)
    return ssl.create_default_context()



def now_iso():
    return datetime.now(timezone.utc).isoformat()


def load_server_url():
    cfg_path = APP_DIR / "license_server_config.json"
    default = "https://uas-license-server.onrender.com"

    try:
        data = json.loads(cfg_path.read_text(encoding="utf-8"))
        url = str(data.get("server_url") or data.get("base_url") or default).strip()
        return url.rstrip("/")
    except Exception:
        return default


def run_cmd(args, timeout=180):
    try:
        p = subprocess.run(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout,
        )
        return {
            "ok": p.returncode == 0,
            "returncode": p.returncode,
            "stdout": p.stdout[-4000:],
            "stderr": p.stderr[-4000:],
        }
    except Exception as exc:
        return {
            "ok": False,
            "returncode": None,
            "stdout": "",
            "stderr": str(exc),
        }


def ensure_runtime_venv():
    result = {
        "created": False,
        "python": str(VENV_PYTHON),
        "ok": False,
        "steps": [],
    }

    if not VENV_PYTHON.exists():
        r = run_cmd([sys.executable, "-m", "venv", str(VENV_DIR)], timeout=180)
        result["created"] = True
        result["steps"].append({"create_venv": r})
        if not r["ok"]:
            return result

    pip_cmd = [str(VENV_PYTHON), "-m", "pip", "install", "--upgrade", "pip"]
    result["steps"].append({"upgrade_pip": run_cmd(pip_cmd, timeout=240)})

    install_cmd = [str(VENV_PYTHON), "-m", "pip", "install", "--upgrade"] + REQUIRED_PACKAGES
    result["steps"].append({"install_packages": run_cmd(install_cmd, timeout=600)})

    result["ok"] = VENV_PYTHON.exists()
    return result


def import_check():
    checks = {}
    if not VENV_PYTHON.exists():
        for pkg in REQUIRED_PACKAGES:
            checks[pkg] = {"ok": False, "error": "runtime python missing"}
        return checks

    for pkg in REQUIRED_PACKAGES:
        module = IMPORT_MAP[pkg]
        code = f"import {module}; print('OK')"
        r = run_cmd([str(VENV_PYTHON), "-c", code], timeout=60)
        checks[pkg] = {
            "ok": bool(r["ok"]),
            "error": "" if r["ok"] else (r["stderr"] or r["stdout"]),
        }

    return checks


def post_empty(url):
    result = {
        "ok": False,
        "reachable": False,
        "status": None,
        "error": "",
        "body_preview": "",
    }

    try:
        req = urllib.request.Request(
            url,
            data=b"{}",
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=60, context=ssl_context()) as resp:
                status = int(resp.getcode())
                body = resp.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as exc:
            status = int(exc.code)
            body = exc.read().decode("utf-8", errors="replace")

        result["status"] = status
        result["body_preview"] = body[:500]

        # 422 is GOOD here: it means the endpoint is reachable and rejected an empty test body.
        result["reachable"] = 200 <= status < 500
        result["ok"] = result["reachable"]

    except Exception as exc:
        result["error"] = str(exc)

    return result


def get_health(server_url):
    result = {
        "ok": False,
        "status": None,
        "error": "",
        "body_preview": "",
    }

    try:
        with urllib.request.urlopen(f"{server_url}/health", timeout=60, context=ssl_context()) as resp:
            result["status"] = int(resp.getcode())
            body = resp.read().decode("utf-8", errors="replace")
            result["body_preview"] = body[:500]
            result["ok"] = 200 <= result["status"] < 300
    except Exception as exc:
        result["error"] = str(exc)

    return result


def required_file_check():
    out = {}
    for name in REQUIRED_FILES:
        p = APP_DIR / name
        out[name] = {
            "exists": p.exists(),
            "size": p.stat().st_size if p.exists() else 0,
        }
    return out


def write_reports(report):
    lines = []

    lines.append("UAS GENERATOR SYSTEM CHECK REPORT")
    lines.append(f"Generated at: {report['generated_at']}")
    lines.append(f"App resources: {report['app_resources']}")
    lines.append(f"Runtime folder: {report['runtime_folder']}")
    lines.append(f"Runtime venv: {report['runtime_venv']}")
    lines.append("")

    lines.append("SERVER")
    lines.append(f"server_url: {report['server']['server_url']}")
    lines.append(f"health_ok: {report['server']['health']['ok']}")
    lines.append(f"activate_endpoint_ok: {report['server']['activate_license_endpoint']['ok']}")
    lines.append(f"validate_endpoint_ok: {report['server']['validate_license_endpoint']['ok']}")
    lines.append("")

    lines.append("REQUIRED FILES")
    for name, info in report["required_files"].items():
        lines.append(f"{name}: exists={info['exists']} size={info['size']}")
    lines.append("")

    lines.append("IMPORTS")
    for pkg, info in report["imports"].items():
        lines.append(f"{pkg}: ok={info['ok']}")
    lines.append("")

    lines.append("PROBLEMS")
    if report["problems"]:
        for problem in report["problems"]:
            lines.append(problem)
    else:
        lines.append("none")
    lines.append("")

    lines.append("WARNINGS")
    if report["warnings"]:
        for warning in report["warnings"]:
            lines.append(warning)
    else:
        lines.append("none")
    lines.append("")

    lines.append(f"READY={report['ready']}")

    REPORT_JSON.write_text(json.dumps(report, indent=2), encoding="utf-8")
    REPORT_TXT.write_text("\n".join(lines), encoding="utf-8")


def main():
    server_url = load_server_url()

    report = {
        "generated_at": now_iso(),
        "app_resources": str(APP_DIR),
        "runtime_folder": str(RUNTIME_DIR),
        "runtime_venv": str(VENV_DIR),
        "server": {
            "server_url": server_url,
            "health": {},
            "activate_license_endpoint": {},
            "validate_license_endpoint": {},
        },
        "required_files": {},
        "venv": {},
        "imports": {},
        "warnings": [],
        "problems": [],
        "ready": False,
    }

    report["required_files"] = required_file_check()
    report["venv"] = ensure_runtime_venv()
    report["imports"] = import_check()

    report["server"]["health"] = get_health(server_url)
    report["server"]["activate_license_endpoint"] = post_empty(f"{server_url}/activate-license")
    report["server"]["validate_license_endpoint"] = post_empty(f"{server_url}/validate-license")

    # If health is good, do not fail System Check because empty POST validation behaves differently.
    # Actual activation/validation is tested later with real payload.
    if report["server"]["health"]["ok"]:
        for endpoint_name in ["activate_license_endpoint", "validate_license_endpoint"]:
            endpoint = report["server"][endpoint_name]
            status = endpoint.get("status")
            endpoint["ok"] = bool(endpoint.get("ok") or endpoint.get("reachable") or (isinstance(status, int) and 200 <= status < 500))
            endpoint["reachable"] = endpoint["ok"]

    for name, info in report["required_files"].items():
        if not info["exists"]:
            report["problems"].append(f"Required file missing: {name}")

    if not report["venv"].get("ok"):
        report["problems"].append("Runtime virtual environment is not ready.")

    for pkg, info in report["imports"].items():
        if not info["ok"]:
            report["problems"].append(f"Python package import failed: {pkg}")

    if not report["server"]["health"]["ok"]:
        report["problems"].append("License server health endpoint is not reachable.")

    if not report["server"]["activate_license_endpoint"]["ok"]:
        report["problems"].append("License server activate endpoint is not reachable.")

    if not report["server"]["validate_license_endpoint"]["ok"]:
        report["problems"].append("License server validate endpoint is not reachable.")

    report["warnings"].append("Running from app resources. Mutable files must remain in runtime_dir.")
    report["warnings"].append("No configured output folder yet. User will choose output folder during normal use.")

    report["ready"] = len(report["problems"]) == 0

    write_reports(report)

    print(f"READY={report['ready']}")
    print(f"REPORT_JSON={REPORT_JSON}")
    print(f"REPORT_TXT={REPORT_TXT}")

    raise SystemExit(0 if report["ready"] else 1)


if __name__ == "__main__":
    main()
