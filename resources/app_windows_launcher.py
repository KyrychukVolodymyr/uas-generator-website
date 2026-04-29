import tempfile
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

APP_DIR = Path(__file__).resolve().parent
APP_VERSION = "v019-windows-ui"

if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))


def get_runtime_resource_dir():
    candidates = []

    try:
        if getattr(sys, "frozen", False):
            if hasattr(sys, "_MEIPASS"):
                candidates.append(Path(sys._MEIPASS))
            candidates.append(Path(sys.executable).resolve().parent / "_internal")
            candidates.append(Path(sys.executable).resolve().parent)
    except Exception:
        pass

    candidates.append(APP_DIR)

    required = [
        "app_windows_engine.py",
        "generate_outputs.py",
        "uas_automation.py",
        "TT_template.xlsx",
        "CDPAS_template.xlsx",
        "FRA_template.pdf",
    ]

    for folder in candidates:
        try:
            if folder.exists() and all((folder / item).exists() for item in required):
                return folder
        except Exception:
            pass

    return APP_DIR


def prepare_runtime_paths():
    resource_dir = get_runtime_resource_dir()

    if str(resource_dir) not in sys.path:
        sys.path.insert(0, str(resource_dir))

    os.chdir(str(resource_dir))
    os.environ["UAS_RESOURCES_DIR"] = str(resource_dir)

    return resource_dir


def maybe_run_subprocess_mode():
    args = list(sys.argv[1:])
    lower_args = [str(a).lower() for a in args]

    has_engine_script = any(a.endswith("app_windows_engine.py") for a in lower_args)
    has_generate_script = any(a.endswith("generate_outputs.py") for a in lower_args)

    has_engine_cli = "--csv" in lower_args and "--output" in lower_args

    has_generate_cli = (
        "--csv" in lower_args
        and "--case-id" in lower_args
        and (
            "--tt-template" in lower_args
            or "--cdpas-template" in lower_args
            or "--fra-template" in lower_args
        )
    )

    if not has_engine_script and not has_engine_cli and not has_generate_script and not has_generate_cli:
        return False

    resource_dir = prepare_runtime_paths()

    if has_generate_script or has_generate_cli:
        filtered_args = []
        for item in args:
            if str(item).lower().endswith("generate_outputs.py"):
                continue
            filtered_args.append(item)

        generator_script = str(resource_dir / "generate_outputs.py")
        sys.argv = [generator_script] + filtered_args

        print("GENERATE_OUTPUTS MODE ACTIVE")
        print("Resource dir:", str(resource_dir))
        print("Generate argv:", sys.argv)

        import generate_outputs
        generate_outputs.main()
        return True

    filtered_args = []
    for item in args:
        if str(item).lower().endswith("app_windows_engine.py"):
            continue
        filtered_args.append(item)

    engine_script = str(resource_dir / "app_windows_engine.py")
    sys.argv = [engine_script] + filtered_args

    print("ENGINE MODE ACTIVE")
    print("Resource dir:", str(resource_dir))
    print("Engine argv:", sys.argv)

    import app_windows_engine
    app_windows_engine.main()
    return True


def load_modules():
    prepare_runtime_paths()
    import commercial_profile
    import license_manager
    return commercial_profile, license_manager


def is_profile_complete(commercial_profile):
    try:
        profile = commercial_profile.load_user_profile()
    except Exception:
        return False

    required = [
        "first_name",
        "last_name",
        "display_name",
        "email",
        "accepted_terms_version",
        "accepted_at",
    ]

    for key in required:
        if not str(profile.get(key, "")).strip():
            return False

    return profile.get("accepted_terms_version") == commercial_profile.TERMS_VERSION


def is_license_active(license_manager):
    try:
        if hasattr(license_manager, "require_active_license"):
            license_manager.require_active_license()
            return True

        if hasattr(license_manager, "validate_license"):
            result = license_manager.validate_license()
            if isinstance(result, dict):
                return bool(result.get("ok") or result.get("valid") or result.get("active"))
            return True

        return False
    except Exception:
        return False


def open_main_app():
    prepare_runtime_paths()
    import app_windows_ui_tk
    app_windows_ui_tk.main()


class OnboardingApp:
    def __init__(self, root, commercial_profile, license_manager):
        self.root = root
        self.commercial_profile = commercial_profile
        self.license_manager = license_manager
        self.activated = False

        self.first_name = tk.StringVar()
        self.last_name = tk.StringVar()
        self.display_name = tk.StringVar()
        self.email = tk.StringVar()
        self.license_key = tk.StringVar()
        self.accept_terms = tk.BooleanVar(value=False)

        self.build()

    def build(self):
        self.root.title("UAS Generator - Activation")
        self.root.geometry("640x560")
        self.root.resizable(False, False)

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Activate UAS Generator",
            font=("Segoe UI", 18, "bold")
        ).pack(pady=(0, 12))

        form = ttk.Frame(frame)
        form.pack(fill="x")

        labels = [
            ("First name:", self.first_name),
            ("Last name:", self.last_name),
            ("Professional display name:", self.display_name),
            ("Email:", self.email),
            ("License key:", self.license_key),
        ]

        for row, item in enumerate(labels):
            label, var = item
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=6)
            show_value = "*" if label == "License key:" else ""
            ttk.Entry(form, textvariable=var, width=46, show=show_value).grid(row=row, column=1, sticky="we", padx=4, pady=6)

        form.columnconfigure(1, weight=1)

        terms_frame = ttk.LabelFrame(frame, text="Terms of Use", padding=10)
        terms_frame.pack(fill="x", pady=14)

        terms_text = (
            "I understand this software is assistive only. "
            "I will verify all outputs before use or signature. "
            "I understand no refunds are provided and key sharing is prohibited."
        )

        terms_row = ttk.Frame(terms_frame)
        terms_row.pack(fill="x", anchor="w")

        ttk.Checkbutton(
            terms_row,
            variable=self.accept_terms
        ).pack(side="left", anchor="n")

        ttk.Label(
            terms_row,
            text=terms_text,
            wraplength=520,
            justify="left"
        ).pack(side="left", anchor="w", padx=(6, 0))

        ttk.Button(
            terms_frame,
            text="Open Terms of Use",
            command=self.open_terms
        ).pack(anchor="w", pady=(8, 0))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=10)

        self.activate_button = ttk.Button(
            buttons,
            text="Activate and Open App",
            command=self.activate
        )
        self.activate_button.pack(side="left", padx=4)

        ttk.Button(buttons, text="Cancel", command=self.cancel).pack(side="left", padx=4)

        self.status = ttk.Label(frame, text="", foreground="blue")
        self.status.pack(anchor="w", pady=(8, 0))

    def open_terms(self):
        try:
            self.commercial_profile.ensure_terms_files()
            terms_path = Path(self.commercial_profile.TERMS_PATH)
            if terms_path.exists():
                subprocess.Popen(["notepad.exe", str(terms_path)])
            else:
                messagebox.showwarning("Terms not found", "Terms file was not found.")
        except Exception as e:
            messagebox.showwarning("Could not open terms", str(e))

    def validate(self):
        first = self.first_name.get().strip()
        last = self.last_name.get().strip()
        display = self.display_name.get().strip()
        email = self.email.get().strip()
        key = self.license_key.get().strip()

        if not first:
            raise ValueError("First name is required.")
        if not last:
            raise ValueError("Last name is required.")
        if not display:
            raise ValueError("Professional display name is required.")
        if not email or "@" not in email:
            raise ValueError("Valid email is required.")
        if not key:
            raise ValueError("License key is required.")
        if not self.accept_terms.get():
            raise ValueError("You must accept the Terms of Use.")

        return first, last, display, email, key

    def activate(self):
        try:
            first, last, display, email, key = self.validate()
        except Exception as e:
            messagebox.showwarning("Missing information", str(e))
            return

        self.activate_button.config(state="disabled")
        self.status.config(text="Activating license...")
        self.root.update_idletasks()

        try:
            now = datetime.now().isoformat(timespec="seconds")

            profile = self.commercial_profile.load_user_profile()
            profile.update({
                "first_name": first,
                "last_name": last,
                "display_name": display,
                "email": email,
                "accepted_terms_version": self.commercial_profile.TERMS_VERSION,
                "accepted_at": now,
                "updated_at": now,
            })
            self.commercial_profile.save_user_profile(profile)

            if not hasattr(self.license_manager, "activate_license"):
                raise RuntimeError("activate_license function was not found in license_manager.")

            self.license_manager.activate_license(
                email=email,
                license_key=key,
                accepted_terms_version=self.commercial_profile.TERMS_VERSION,
                app_version=APP_VERSION,
            )

            if not is_license_active(self.license_manager):
                raise RuntimeError("License activation was submitted, but active license check failed.")

            self.activated = True
            messagebox.showinfo("Activated", "License activated successfully.")
            self.root.destroy()

        except Exception as e:
            self.activate_button.config(state="normal")
            self.status.config(text="")
            messagebox.showerror("Activation failed", str(e))

    def cancel(self):
        self.activated = False
        self.root.destroy()



def run_pre_onboarding_system_check():
    import os
    import sys
    import platform
    import subprocess
    import tempfile
    import zipfile
    from pathlib import Path
    from datetime import datetime

    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        tk = None
        messagebox = None

    app_dir = Path(__file__).resolve().parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    desktop = Path(os.path.expanduser("~")) / "Desktop"
    report_path = desktop / f"UAS_WINDOWS_PRE_ONBOARDING_SYSTEM_CHECK_{timestamp}.txt"

    lines = []
    failed_required = []

    def section(title):
        lines.append("")
        lines.append("=" * 72)
        lines.append(title)
        lines.append("=" * 72)

    def check(name, ok, detail="", required=True):
        status = "PASS" if ok else "FAIL"
        lines.append(f"{status}: {name}")
        if detail:
            lines.append(f"      {detail}")
        if required and not ok:
            failed_required.append(name)

    lines.append("UAS GENERATOR WINDOWS PRE-ONBOARDING SYSTEM CHECK")
    lines.append(f"Created: {timestamp}")
    lines.append("This check runs before onboarding.")
    lines.append("")

    section("Windows Environment")
    check("Running on Windows", os.name == "nt", f"os.name={os.name}, platform={platform.platform()}")
    check("Application folder exists", app_dir.exists(), str(app_dir))

    section("Required App Files")
    required_files = [
        "app_windows_launcher.py",
        "app_windows_ui.py",
        "app_windows_engine.py",
        "TT_template.xlsx",
        "CDPAS_template.xlsx",
        "FRA_template.pdf",
        "anthem_logo.jpeg",
    ]

    for name in required_files:
        p = app_dir / name
        check(f"Required file present: {name}", p.exists(), str(p))

    section("Required Python Modules")
    modules = ["openpyxl", "pandas", "pypdf", "reportlab"]
    for mod in modules:
        try:
            __import__(mod)
            check(f"Module available: {mod}", True)
        except Exception as e:
            check(f"Module available: {mod}", False, str(e))

    section("Runtime Write Access")
    try:
        appdata_root = Path(os.environ.get("APPDATA", "")) / "UASGenerator2026V2"
        appdata_root.mkdir(parents=True, exist_ok=True)
        test_file = appdata_root / f"pre_onboarding_write_test_{timestamp}.txt"
        test_file.write_text("ok", encoding="utf-8")
        test_file.unlink(missing_ok=True)
        check("APPDATA runtime folder writable", True, str(appdata_root))
    except Exception as e:
        check("APPDATA runtime folder writable", False, str(e))

    try:
        with tempfile.TemporaryDirectory(prefix="uas_precheck_") as tmp:
            temp_test = Path(tmp) / "test.txt"
            temp_test.write_text("ok", encoding="utf-8")
            check("Temporary folder writable", temp_test.exists(), str(temp_test))
    except Exception as e:
        check("Temporary folder writable", False, str(e))

    section("PDF Export Support")
    soffice_candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]

    soffice_found = None
    for candidate in soffice_candidates:
        if candidate.exists():
            soffice_found = candidate
            break

    check(
        "LibreOffice available for Excel-to-PDF export",
        soffice_found is not None,
        str(soffice_found) if soffice_found else "LibreOffice not found in standard locations."
    )

    section("Optional Microsoft Excel Check")
    excel_detail = "Not available or not tested."
    try:
        ps = "try { $e = New-Object -ComObject Excel.Application; $v=$e.Version; $e.Quit(); Write-Output $v; exit 0 } catch { Write-Error $_; exit 1 }"
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode == 0:
            excel_detail = "Excel COM available. Version: " + result.stdout.strip()
        else:
            excel_detail = "Excel COM not available. " + (result.stderr.strip() or result.stdout.strip())
    except Exception as e:
        excel_detail = str(e)

    check("Microsoft Excel desktop COM optional", True, excel_detail, required=False)

    section("Result")
    if failed_required:
        lines.append("OVERALL RESULT: FAIL")
        lines.append("Required checks failed:")
        for item in failed_required:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("You can continue, but the app may not generate documents correctly.")
    else:
        lines.append("OVERALL RESULT: PASS")
        lines.append("Required environment checks passed. Continue onboarding.")

    report_path.write_text("\n".join(lines), encoding="utf-8")

    try:
        subprocess.Popen(["notepad.exe", str(report_path)])
    except Exception:
        pass

    proceed = True

    try:
        if tk is not None and messagebox is not None:
            root = tk.Tk()
            root.withdraw()

            if failed_required:
                proceed = messagebox.askyesno(
                    "System Check Found Problems",
                    "System Check found required problems.\n\nA report was opened on your Desktop.\n\nDo you still want to continue onboarding?"
                )
            else:
                messagebox.showinfo(
                    "System Check Passed",
                    "System Check passed.\n\nA report was opened on your Desktop.\n\nOnboarding will continue now."
                )
                proceed = True

            root.destroy()
    except Exception:
        proceed = True

    if not proceed:
        sys.exit(0)

    return True




def should_run_pre_onboarding_system_check_once():
    import os
    from pathlib import Path

    runtime_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    marker = runtime_dir / "pre_onboarding_system_check_completed.flag"

    return not marker.exists()


def mark_pre_onboarding_system_check_completed():
    import os
    from pathlib import Path
    from datetime import datetime

    runtime_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    marker = runtime_dir / "pre_onboarding_system_check_completed.flag"
    marker.write_text(datetime.now().isoformat(), encoding="utf-8")


def run_pre_onboarding_system_check_forced():
    print("V020D15_PRE_ONBOARDING_SYSTEM_CHECK_RUNNING", flush=True)

    import os
    import platform
    import subprocess
    import tempfile
    from pathlib import Path
    from datetime import datetime

    app_dir = Path(__file__).resolve().parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    desktop = Path(os.path.expanduser("~")) / "Desktop"
    report_path = desktop / f"UAS_WINDOWS_PRE_ONBOARDING_SYSTEM_CHECK_{timestamp}.txt"

    lines = []
    failed = []

    def section(title):
        lines.append("")
        lines.append("=" * 72)
        lines.append(title)
        lines.append("=" * 72)

    def check(name, ok, detail="", required=True):
        status = "PASS" if ok else "FAIL"
        lines.append(f"{status}: {name}")
        if detail:
            lines.append(f"      {detail}")
        if required and not ok:
            failed.append(name)

    lines.append("UAS GENERATOR WINDOWS PRE-ONBOARDING SYSTEM CHECK")
    lines.append(f"Created: {timestamp}")
    lines.append("This check runs before activation/onboarding.")
    lines.append("")

    section("Windows Environment")
    check("Running on Windows", os.name == "nt", f"os.name={os.name}, platform={platform.platform()}")
    check("Application folder exists", app_dir.exists(), str(app_dir))

    section("Required App Files")
    required_files = [
        "app_windows_launcher.py",
        "app_windows_ui.py",
        "app_windows_engine.py",
        "TT_template.xlsx",
        "CDPAS_template.xlsx",
        "FRA_template.pdf",
        "anthem_logo.jpeg",
    ]

    for name in required_files:
        p = app_dir / name
        check(f"Required file present: {name}", p.exists(), str(p))

    section("Required Modules")
    for mod in ["openpyxl", "pandas", "pypdf", "reportlab"]:
        try:
            __import__(mod)
            check(f"Module available: {mod}", True)
        except Exception as e:
            check(f"Module available: {mod}", False, str(e))

    section("Write Access")
    try:
        runtime_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2"
        runtime_dir.mkdir(parents=True, exist_ok=True)
        test = runtime_dir / f"write_test_{timestamp}.txt"
        test.write_text("ok", encoding="utf-8")
        test.unlink(missing_ok=True)
        check("APPDATA runtime folder writable", True, str(runtime_dir))
    except Exception as e:
        check("APPDATA runtime folder writable", False, str(e))

    try:
        with tempfile.TemporaryDirectory(prefix="uas_precheck_") as tmp:
            test = Path(tmp) / "test.txt"
            test.write_text("ok", encoding="utf-8")
            check("Temporary folder writable", test.exists(), str(test))
    except Exception as e:
        check("Temporary folder writable", False, str(e))

    section("PDF Export")
    soffice_paths = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    soffice = next((p for p in soffice_paths if p.exists()), None)
    check(
        "LibreOffice available for Excel-to-PDF export",
        soffice is not None,
        str(soffice) if soffice else "LibreOffice not found in standard locations."
    )

    section("Result")
    if failed:
        lines.append("OVERALL RESULT: FAIL")
        lines.append("Required checks failed:")
        for item in failed:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("The app will continue, but document generation may fail.")
    else:
        lines.append("OVERALL RESULT: PASS")
        lines.append("Required checks passed. Activation/onboarding will continue.")

    report_path.write_text("\n".join(lines), encoding="utf-8")

    try:
        subprocess.Popen(["notepad.exe", str(report_path)])
    except Exception:
        pass

    try:
        import ctypes
        if failed:
            message = "System Check found problems.\n\nA report was opened on your Desktop.\n\nClick OK to continue onboarding anyway."
            title = "System Check Found Problems"
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x30)
        else:
            message = "System Check passed.\n\nA report was opened on your Desktop.\n\nClick OK to continue onboarding."
            title = "System Check Passed"
            ctypes.windll.user32.MessageBoxW(0, message, title, 0x40)
    except Exception:
        pass

    mark_pre_onboarding_system_check_completed()

    print("V020D15_PRE_ONBOARDING_SYSTEM_CHECK_FINISHED_CONTINUING", flush=True)
    return True


def main():
    print('V020D15_PRE_ONBOARDING_CALL_SITE', flush=True)
    if should_run_pre_onboarding_system_check_once():
        run_pre_onboarding_system_check_forced()
    if maybe_run_subprocess_mode():
        return

    commercial_profile, license_manager = load_modules()

    if is_profile_complete(commercial_profile) and is_license_active(license_manager):
        open_main_app()
        return

    root = tk.Tk()
    app = OnboardingApp(root, commercial_profile, license_manager)
    root.mainloop()

    if is_profile_complete(commercial_profile) and is_license_active(license_manager):
        open_main_app()
        return

    return
if __name__ == "__main__":
    main()



