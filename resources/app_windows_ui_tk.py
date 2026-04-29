import os
import sys
import re
import subprocess
import threading
import json
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


APP_DIR = Path(__file__).resolve().parent

def subprocess_creationflags():
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return subprocess.CREATE_NO_WINDOW
    return 0

WORK_ROOT = APP_DIR.parent
ENGINE = APP_DIR / "app_windows_engine.py"
PYTHON_EXE = WORK_ROOT / ".venv" / "Scripts" / "python.exe"
UI_CONFIG_FILE = WORK_ROOT / "windows_ui_config.json"
DEFAULT_OUTPUT_FOLDER = str(Path.home() / "Documents" / "Anthem Blue Cross Blue Shield Assessment Documentation")



def get_uas_support_log_path():
    import os
    from pathlib import Path

    support_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2" / "support_logs"
    support_dir.mkdir(parents=True, exist_ok=True)
    return support_dir / "uas_windows_ui_last_engine_error.txt"


class UASGeneratorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("UAS Generator for Windows")
        self.root.geometry("1040x760")
        self.root.minsize(980, 720)

        self.csv_path = tk.StringVar()
        self.output_path = tk.StringVar(value=self.load_saved_output_folder())
        self.patient_text = tk.StringVar()

        self.days_per_week = tk.IntVar(value=5)
        self.hours_per_day = tk.IntVar(value=8)
        self.dwelling_type = tk.StringVar(value="Elevator")
        self.bedrooms_count = tk.IntVar(value=2)
        self.accessible_home = tk.StringVar(value="no")

        self.equipment_vars = {}
        self.cdpas_aspiration = tk.BooleanVar(value=False)
        self.cdpas_tube_feeding = tk.BooleanVar(value=False)
        self.cdpas_oxygen = tk.BooleanVar(value=False)
        self.cdpas_catheter_type = tk.StringVar()
        self.service_hours_discussion = tk.BooleanVar(value=False)

        self.build()

    def load_saved_output_folder(self):
        try:
            if UI_CONFIG_FILE.exists():
                data = json.loads(UI_CONFIG_FILE.read_text(encoding="utf-8"))
                saved = str(data.get("output_folder", "")).strip()
                if saved:
                    return saved
        except Exception:
            pass
        return DEFAULT_OUTPUT_FOLDER

    def save_output_folder(self):
        try:
            folder = str(self.output_path.get()).strip()
            if folder:
                UI_CONFIG_FILE.write_text(json.dumps({"output_folder": folder}, indent=2), encoding="utf-8")
        except Exception:
            pass

    def build(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        title = ttk.Label(main, text="UAS Generator for Windows", font=("Segoe UI", 20, "bold"))
        title.pack(pady=(0, 12))

        files = ttk.LabelFrame(main, text="Files", padding=10)
        files.pack(fill="x", pady=6)

        ttk.Label(files, text="CSV file:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(files, textvariable=self.csv_path, width=95).grid(row=0, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(files, text="Select CSV", command=self.select_csv).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(files, text="Output folder:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(files, textvariable=self.output_path, width=95).grid(row=1, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(files, text="Select Output", command=self.select_output).grid(row=1, column=2, padx=4, pady=4)

        ttk.Label(files, text="Member ID / Case ID:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(files, textvariable=self.patient_text, width=95).grid(row=2, column=1, columnspan=2, sticky="we", padx=4, pady=4)

        files.columnconfigure(1, weight=1)

        schedule = ttk.LabelFrame(main, text="Service Schedule / Home", padding=10)
        schedule.pack(fill="x", pady=6)

        ttk.Label(schedule, text="Days/week:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Spinbox(schedule, from_=1, to=7, textvariable=self.days_per_week, width=10).grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(schedule, text="Hours/day:").grid(row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Spinbox(schedule, from_=1, to=24, textvariable=self.hours_per_day, width=10).grid(row=0, column=3, sticky="w", padx=4, pady=4)

        ttk.Label(schedule, text="Dwelling:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        dwelling = ttk.Combobox(schedule, textvariable=self.dwelling_type, width=28, state="readonly")
        dwelling["values"] = ("Elevator", "No elevator", "Private house", "Private house with stairs")
        dwelling.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(schedule, text="Bedrooms:").grid(row=1, column=2, sticky="w", padx=4, pady=4)
        ttk.Spinbox(schedule, from_=0, to=10, textvariable=self.bedrooms_count, width=10).grid(row=1, column=3, sticky="w", padx=4, pady=4)

        ttk.Label(schedule, text="Accessible home:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        accessible = ttk.Combobox(schedule, textvariable=self.accessible_home, width=10, state="readonly")
        accessible["values"] = ("no", "yes")
        accessible.grid(row=2, column=1, sticky="w", padx=4, pady=4)

        equipment = ttk.LabelFrame(main, text="Equipment / DME", padding=10)
        equipment.pack(fill="x", pady=6)

        equipment_items = [
            ("cane", "Cane", "--cane"),
            ("walker", "Walker", "--walker"),
            ("wheelchair", "Wheelchair", "--wheelchair"),
            ("hospital_bed", "Hospital bed", "--hospital-bed"),
            ("catheter", "Catheter", "--catheter"),
            ("grab_bar", "Grab bar", "--grab-bar"),
            ("bedside_commode", "Bedside commode", "--bedside-commode"),
            ("oxygen", "Oxygen in use", "--oxygen-in-use"),
            ("hoyer_lift", "Hoyer lift", "--hoyer-lift"),
            ("raised_toilet_seat", "Raised toilet seat", "--raised-toilet-seat"),
            ("shower_chair", "Shower chair", "--shower-chair"),
        ]

        for i, (key, label, arg) in enumerate(equipment_items):
            var = tk.BooleanVar(value=False)
            self.equipment_vars[key] = (var, arg)
            ttk.Checkbutton(equipment, text=label, variable=var).grid(row=i // 4, column=i % 4, sticky="w", padx=8, pady=4)

        tt_service = ttk.LabelFrame(main, text="TT Service Hours Discussion", padding=10)
        tt_service.pack(fill="x", pady=6)

        ttk.Checkbutton(
            tt_service,
            text="Member requests to discuss additional service hours",
            variable=self.service_hours_discussion
        ).grid(row=0, column=0, sticky="w", padx=8, pady=4)

        cdpas = ttk.LabelFrame(main, text="CDPAS Special Options", padding=10)
        cdpas.pack(fill="x", pady=6)

        ttk.Checkbutton(cdpas, text="Aspiration precautions", variable=self.cdpas_aspiration).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(cdpas, text="Tube feeding / G-tube care", variable=self.cdpas_tube_feeding).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(cdpas, text="Oxygen care", variable=self.cdpas_oxygen).grid(row=0, column=2, sticky="w", padx=8, pady=4)

        ttk.Label(cdpas, text="Catheter type:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        catheter_combo = ttk.Combobox(cdpas, textvariable=self.cdpas_catheter_type, width=48, state="readonly")
        catheter_combo["values"] = ("", "Foley catheter", "Suprapubic catheter", "Condom catheter", "Straight catheter")
        catheter_combo.grid(row=1, column=1, columnspan=2, sticky="w", padx=8, pady=4)

        buttons = ttk.Frame(main)
        buttons.pack(fill="x", pady=10)

        self.generate_button = ttk.Button(buttons, text="Generate Documents", command=self.generate)
        self.generate_button.pack(side="left", padx=4)

        ttk.Button(buttons, text="Open Work Folder", command=lambda: self.open_folder(WORK_ROOT)).pack(side="left", padx=4)

        self.log = tk.Text(main, height=16, wrap="word")
        self.log.pack(fill="both", expand=True, pady=6)

    def select_csv(self):
        path = filedialog.askopenfilename(title="Select UAS CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.csv_path.set(path)

    def select_output(self):
        path = filedialog.askdirectory(title="Select output folder", initialdir=self.output_path.get() or str(Path.home()))
        if path:
            self.output_path.set(path)
            self.save_output_folder()
            self.save_output_folder()

    def log_line(self, text):
        self.log.insert("end", text)
        self.log.see("end")
        self.root.update_idletasks()

    def dwelling_engine_value(self):
        mapping = {
            "Elevator": "apartment_elevator",
            "No elevator": "apartment_no_elevator",
            "Private house": "private_house",
            "Private house with stairs": "private_house_stairs",
        }
        return mapping.get(self.dwelling_type.get(), "apartment_elevator")

    def build_command(self):
        csv = self.csv_path.get().strip()
        output = self.output_path.get().strip()
        member_id = self.patient_text.get().strip()

        if not csv or not Path(csv).exists():
            raise ValueError("Please select a valid CSV file.")

        if not member_id:
            raise ValueError("Member ID / Case ID is required. Please type the member ID before generating documents.")

        if not output:
            raise ValueError("Please select an output folder.")

        self.save_output_folder()

        Path(output).mkdir(parents=True, exist_ok=True)

        python_exe = PYTHON_EXE if PYTHON_EXE.exists() else Path(sys.executable)

        cmd = [
            str(python_exe),
            str(ENGINE),
            "--csv", csv,
            "--patient", member_id,
            "--output", output,
            "--days-per-week", str(self.days_per_week.get()),
            "--hours-per-day", str(self.hours_per_day.get()),
            "--dwelling-type", self.dwelling_engine_value(),
            "--bedrooms-count", str(self.bedrooms_count.get()),
            "--accessible-home", self.accessible_home.get(),
            "--no-dialogs",
        ]

        for var, arg in self.equipment_vars.values():
            if var.get():
                cmd.append(arg)

        if self.service_hours_discussion.get():
            cmd.append("--service")

        if self.cdpas_aspiration.get():
            cmd.append("--cdpas-aspiration-precautions")
        if self.cdpas_tube_feeding.get():
            cmd.append("--cdpas-tube-feeding")
        if self.cdpas_oxygen.get():
            cmd.append("--cdpas-oxygen-care")

        catheter_type = self.cdpas_catheter_type.get().strip()
        if catheter_type:
            cmd.extend(["--cdpas-catheter-type", catheter_type])

        return cmd, output

    def generated_docs_present(self, output, full_log, run_started_at):
        """
        Validate only files created/modified during this current run.
        Uses robust normalized matching so 6monthCMVisit is recognized correctly.
        """
        output_path = Path(output)

        recent_files = []
        cutoff = float(run_started_at) - 3.0

        try:
            for p in output_path.rglob("*"):
                if p.is_file():
                    try:
                        if p.stat().st_mtime >= cutoff:
                            recent_files.append(p)
                    except Exception:
                        pass
        except Exception:
            recent_files = []

        names = [p.name.lower() for p in recent_files]
        compact_names = [re.sub(r"[^a-z0-9]", "", n.lower()) for n in names]

        required = ["TT", "CDPAS", "FRA"]

        match = re.search(r"Required=\[(.*?)\]", full_log)
        if match:
            parsed_required = []
            for token in re.findall(r"'([^']+)'|\"([^\"]+)\"", match.group(1)):
                value = (token[0] or token[1] or "").strip()
                if value:
                    parsed_required.append(value)
            if parsed_required:
                required = parsed_required

        def norm_doc(value):
            return re.sub(r"[^a-z0-9]", "", str(value).lower())

        def has_doc(doc):
            d = norm_doc(doc)

            if d == "tt":
                return any(
                    n.endswith("-tt.xlsx") or
                    n.endswith("-tt.pdf") or
                    n == "tt.xlsx" or
                    n == "tt.pdf" or
                    cn.endswith("ttxlsx") or
                    cn.endswith("ttpdf")
                    for n, cn in zip(names, compact_names)
                )

            if d == "cdpas":
                return any(
                    ("cdpas" in n and (n.endswith(".xlsx") or n.endswith(".pdf"))) or
                    ("cdpas" in cn and (cn.endswith("xlsx") or cn.endswith("pdf")))
                    for n, cn in zip(names, compact_names)
                )

            if d == "fra":
                return any(
                    ("fra" in n and n.endswith(".pdf")) or
                    ("fra" in cn and cn.endswith("pdf"))
                    for n, cn in zip(names, compact_names)
                )

            if d == "6monthcmvisit":
                return any(
                    ("6monthcmvisit" in cn and cn.endswith("pdf")) or
                    ("6month" in cn and "cmvisit" in cn and cn.endswith("pdf"))
                    for cn in compact_names
                )

            if d == "phq9":
                return any(
                    ("phq9" in n and n.endswith(".pdf")) or
                    ("phq9" in cn and cn.endswith("pdf"))
                    for n, cn in zip(names, compact_names)
                )

            return False

        missing = [doc for doc in required if not has_doc(doc)]

        return len(missing) == 0, required, missing, recent_files

    def generate(self):
        try:
            cmd, output = self.build_command()
        except Exception as e:
            messagebox.showwarning("Missing information", str(e))
            return

        self.generate_button.config(state="disabled")
        self.log.delete("1.0", "end")
        self.log_line("Running engine...\n\n")

        thread = threading.Thread(target=self.run_engine, args=(cmd, output), daemon=True)
        thread.start()

    def run_engine(self, cmd, output):
        try:
            self.log_line("Command:\n" + " ".join(cmd) + "\n\n")

            run_started_at = time.time()

            proc = subprocess.Popen(
                cmd,
                cwd=str(APP_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            lines = []
            for line in proc.stdout:
                lines.append(line)
                self.log_line(line)

            code = proc.wait()
            full_log = "".join(lines)

            try:
                error_log = get_uas_support_log_path()
                error_log.write_text(full_log, encoding="utf-8")
            except Exception:
                error_log = None

            if code == 0:
                self.root.after(0, lambda: self.success(output, "Documents generated successfully."))
                return

            ok_by_files, required, missing, recent_files = self.generated_docs_present(output, full_log, run_started_at)
            if ok_by_files:
                self.log_line("\nEngine returned exit code 1, but all required documents from this current run were found. Treating as success.\n")
                self.root.after(0, lambda: self.success(output, "Documents generated successfully. Fields reset for next member."))
            else:
                recent_names = [str(p) for p in recent_files]
                msg = "Generation failed with exit code " + str(code)
                if error_log:
                    msg += "\n\nFull log saved to:\n" + str(error_log)
                msg += "\n\nRequired: " + str(required)
                msg += "\nMissing from current run: " + str(missing)
                msg += "\n\nRecent files from this run:\n" + "\n".join(recent_names[-20:])
                msg += "\n\n" + full_log[-2200:]
                self.root.after(0, lambda: messagebox.showerror("Generation failed", msg[:3000]))

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Generation failed", str(e)))
        finally:
            self.root.after(0, lambda: self.generate_button.config(state="normal"))

    def success(self, output, message):
        self.log_line("\nDone. Opening output folder.\n")
        self.open_folder(output)
        self.reset_patient_fields()
        messagebox.showinfo("Done", message)

    def reset_patient_fields(self):
        # Output folder is intentionally not reset. It stays saved until user changes it.
        self.csv_path.set("")
        self.patient_text.set("")
        self.days_per_week.set(5)
        self.hours_per_day.set(8)
        self.dwelling_type.set("Elevator")
        self.bedrooms_count.set(2)
        self.accessible_home.set("no")

        for var, arg in self.equipment_vars.values():
            var.set(False)

        self.cdpas_aspiration.set(False)
        self.cdpas_tube_feeding.set(False)
        self.cdpas_oxygen.set(False)
        self.cdpas_catheter_type.set("")
        self.service_hours_discussion.set(False)

        self.log_line("\nReady for next member. Patient-specific fields were reset.\n")

    def open_folder(self, folder):
        try:
            Path(folder).mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))
        except Exception as e:
            messagebox.showwarning("Could not open folder", str(e))


def main():
    root = tk.Tk()
    UASGeneratorUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

