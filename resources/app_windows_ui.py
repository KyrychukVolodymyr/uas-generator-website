import os
import sys
import subprocess
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QPushButton,
    QFileDialog,
    QLineEdit,
    QSpinBox,
    QComboBox,
    QCheckBox,
    QTextEdit,
    QMessageBox,
    QGridLayout,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout)


APP_DIR = Path(__file__).resolve().parent
WORK_ROOT = APP_DIR.parent
ENGINE = APP_DIR / "app_windows_engine.py"
PYTHON_EXE = WORK_ROOT / ".venv" / "Scripts" / "python.exe"



def get_uas_support_log_path():
    import os
    from pathlib import Path
    runtime_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "UASGenerator2026V2" / "support_logs"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir / "uas_windows_ui_last_engine_error.txt"


class EngineRunner(QThread):
    finished_ok = Signal(str)
    failed = Signal(str)
    output = Signal(str)

    def __init__(self, command, output_dir):
        super().__init__()
        self.command = command
        self.output_dir = output_dir

    def run(self):
        try:
            self.output.emit("Running engine...\n")
            self.output.emit(" ".join([str(x) for x in self.command]) + "\n\n")

            proc = subprocess.Popen(
                self.command,
                cwd=str(APP_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            lines = []
            if proc.stdout:
                for line in proc.stdout:
                    lines.append(line)
                    self.output.emit(line)

            code = proc.wait()

            if code == 0:
                self.finished_ok.emit(str(self.output_dir))
            else:
                self.failed.emit("Engine failed with exit code " + str(code) + "\n\n" + "".join(lines))

        except Exception as e:
            self.failed.emit(str(e))


class UASWindowsUI(QWidget):
    def __init__(self):
        super().__init__()
        self.runner = None
        self.setWindowTitle("UAS Generator for Windows")
        self.setMinimumWidth(980)
        self.setMinimumHeight(760)
        self.build_ui()

    def build_ui(self):
        root = QVBoxLayout(self)

        title = QLabel("UAS Generator for Windows")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; padding: 10px;")
        root.addWidget(title)

        files_box = QGroupBox("Files")
        files_layout = QGridLayout(files_box)

        self.csv_path = QLineEdit()
        self.csv_path.setPlaceholderText("Select UAS CSV file")

        csv_btn = QPushButton("Select CSV")
        csv_btn.clicked.connect(self.select_csv)

        self.output_path = QLineEdit()
        default_output = Path.home() / "Documents" / "Anthem Blue Cross Blue Shield Assessment Documentation"
        self.output_path.setText(str(default_output))

        output_btn = QPushButton("Select Output Folder")
        output_btn.clicked.connect(self.select_output)

        self.patient_text = QLineEdit()
        self.patient_text.setPlaceholderText("Optional. Example: HABER_TEST_99999")

        files_layout.addWidget(QLabel("CSV file:"), 0, 0)
        files_layout.addWidget(self.csv_path, 0, 1)
        files_layout.addWidget(csv_btn, 0, 2)

        files_layout.addWidget(QLabel("Output folder:"), 1, 0)
        files_layout.addWidget(self.output_path, 1, 1)
        files_layout.addWidget(output_btn, 1, 2)

        files_layout.addWidget(QLabel("Patient field:"), 2, 0)
        files_layout.addWidget(self.patient_text, 2, 1, 1, 2)

        root.addWidget(files_box)

        schedule_box = QGroupBox("Service Schedule / Home")
        schedule_layout = QGridLayout(schedule_box)

        self.days_per_week = QSpinBox()
        self.days_per_week.setRange(1, 7)
        self.days_per_week.setValue(5)

        self.hours_per_day = QSpinBox()
        self.hours_per_day.setRange(1, 24)
        self.hours_per_day.setValue(8)

        self.dwelling_type = QComboBox()
        self.dwelling_type.addItem("Apartment with elevator", "apartment_elevator")
        self.dwelling_type.addItem("Apartment without elevator", "apartment_no_elevator")
        self.dwelling_type.addItem("Private house", "private_house")
        self.dwelling_type.addItem("Other", "other")

        self.bedrooms_count = QSpinBox()
        self.bedrooms_count.setRange(0, 10)
        self.bedrooms_count.setValue(2)

        self.accessible_home = QComboBox()
        self.accessible_home.addItem("No", "no")
        self.accessible_home.addItem("Yes", "yes")

        schedule_layout.addWidget(QLabel("Days per week:"), 0, 0)
        schedule_layout.addWidget(self.days_per_week, 0, 1)

        schedule_layout.addWidget(QLabel("Hours per day:"), 0, 2)
        schedule_layout.addWidget(self.hours_per_day, 0, 3)

        schedule_layout.addWidget(QLabel("Dwelling type:"), 1, 0)
        schedule_layout.addWidget(self.dwelling_type, 1, 1)

        schedule_layout.addWidget(QLabel("Bedrooms:"), 1, 2)
        schedule_layout.addWidget(self.bedrooms_count, 1, 3)

        schedule_layout.addWidget(QLabel("Accessible home:"), 2, 0)
        schedule_layout.addWidget(self.accessible_home, 2, 1)

        root.addWidget(schedule_box)

        equipment_box = QGroupBox("Equipment / DME")
        equipment_layout = QGridLayout(equipment_box)

        self.cane = QCheckBox("Cane")
        self.walker = QCheckBox("Walker")
        self.wheelchair = QCheckBox("Wheelchair")
        self.hospital_bed = QCheckBox("Hospital bed")
        self.catheter = QCheckBox("Catheter")
        self.grab_bar = QCheckBox("Grab bar")
        self.bedside_commode = QCheckBox("Bedside commode")
        self.oxygen = QCheckBox("Oxygen")
        self.service = QCheckBox("Service animal")
        self.hoyer_lift = QCheckBox("Hoyer lift")
        self.raised_toilet_seat = QCheckBox("Raised toilet seat")
        self.shower_chair = QCheckBox("Shower chair")

        equipment_items = [
            self.cane,
            self.walker,
            self.wheelchair,
            self.hospital_bed,
            self.catheter,
            self.grab_bar,
            self.bedside_commode,
            self.oxygen,
            self.service,
            self.hoyer_lift,
            self.raised_toilet_seat,
            self.shower_chair,
        ]

        for i, item in enumerate(equipment_items):
            equipment_layout.addWidget(item, i // 4, i % 4)

        root.addWidget(equipment_box)

        cdpas_box = QGroupBox("CDPAS Special Options")
        cdpas_layout = QGridLayout(cdpas_box)

        self.cdpas_aspiration = QCheckBox("Aspiration precautions")
        self.cdpas_tube_feeding = QCheckBox("Tube feeding / G-tube care")
        self.cdpas_oxygen = QCheckBox("Oxygen care")
        self.cdpas_catheter_type = QLineEdit()
        self.cdpas_catheter_type.setPlaceholderText("Catheter type, optional. Example: Foley catheter")

        cdpas_layout.addWidget(self.cdpas_aspiration, 0, 0)
        cdpas_layout.addWidget(self.cdpas_tube_feeding, 0, 1)
        cdpas_layout.addWidget(self.cdpas_oxygen, 0, 2)

        cdpas_layout.addWidget(QLabel("Catheter type:"), 1, 0)
        cdpas_layout.addWidget(self.cdpas_catheter_type, 1, 1, 1, 2)

        root.addWidget(cdpas_box)

        buttons = QHBoxLayout()

        self.generate_btn = QPushButton("Generate Documents")
        self.generate_btn.setMinimumHeight(44)
        self.generate_btn.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.generate_btn.clicked.connect(self.generate)

        open_work_btn = QPushButton("Open Work Folder")
        open_work_btn.clicked.connect(lambda: self.open_folder(WORK_ROOT))

        buttons.addWidget(self.generate_btn)
        buttons.addWidget(open_work_btn)
        root.addLayout(buttons)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Run log will appear here.")
        root.addWidget(self.log)

    def select_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select UAS CSV", str(Path.home()), "CSV files (*.csv);;All files (*.*)")
        if path:
            self.csv_path.setText(path)

    def select_output(self):
        path = QFileDialog.getExistingDirectory(self, "Select output folder", self.output_path.text() or str(Path.home()))
        if path:
            self.output_path.setText(path)

    def add_checkbox_arg(self, command, checkbox, arg):
        if checkbox.isChecked():
            command.append(arg)

    def generate(self):
        csv = self.csv_path.text().strip()
        output = self.output_path.text().strip()
        patient = self.patient_text.text().strip()

        if not csv or not Path(csv).exists():
            QMessageBox.warning(self, "Missing CSV", "Please select a valid CSV file.")
            return

        if not output:
            QMessageBox.warning(self, "Missing output folder", "Please select an output folder.")
            return

        Path(output).mkdir(parents=True, exist_ok=True)

        if not patient:
            patient = "WINDOWS_UI_PATIENT"

        if not ENGINE.exists():
            QMessageBox.critical(self, "Missing engine", "app_windows_engine.py was not found.")
            return

        python_exe = PYTHON_EXE if PYTHON_EXE.exists() else Path(sys.executable)

        command = [
            str(python_exe),
            str(ENGINE),
            "--csv",
            csv,
            "--patient",
            patient,
            "--output",
            output,
            "--days-per-week",
            str(self.days_per_week.value()),
            "--hours-per-day",
            str(self.hours_per_day.value()),
            "--dwelling-type",
            self.dwelling_type.currentData(),
            "--bedrooms-count",
            str(self.bedrooms_count.value()),
            "--accessible-home",
            self.accessible_home.currentData(),
            "--no-dialogs",
        ]

        self.add_checkbox_arg(command, self.cane, "--cane")
        self.add_checkbox_arg(command, self.walker, "--walker")
        self.add_checkbox_arg(command, self.wheelchair, "--wheelchair")
        self.add_checkbox_arg(command, self.hospital_bed, "--hospital-bed")
        self.add_checkbox_arg(command, self.catheter, "--catheter")
        self.add_checkbox_arg(command, self.grab_bar, "--grab-bar")
        self.add_checkbox_arg(command, self.bedside_commode, "--bedside-commode")
        self.add_checkbox_arg(command, self.oxygen, "--oxygen")
        self.add_checkbox_arg(command, self.service, "--service")
        self.add_checkbox_arg(command, self.hoyer_lift, "--hoyer-lift")
        self.add_checkbox_arg(command, self.raised_toilet_seat, "--raised-toilet-seat")
        self.add_checkbox_arg(command, self.shower_chair, "--shower-chair")

        self.add_checkbox_arg(command, self.cdpas_aspiration, "--cdpas-aspiration-precautions")
        self.add_checkbox_arg(command, self.cdpas_tube_feeding, "--cdpas-tube-feeding")
        self.add_checkbox_arg(command, self.cdpas_oxygen, "--cdpas-oxygen-care")

        catheter_type = self.cdpas_catheter_type.text().strip()
        if catheter_type:
            command.extend(["--cdpas-catheter-type", catheter_type])

        self.log.clear()
        self.generate_btn.setEnabled(False)

        self.runner = EngineRunner(command, Path(output))
        self.runner.output.connect(self.log.append)
        self.runner.finished_ok.connect(self.on_success)
        self.runner.failed.connect(self.on_failed)
        self.runner.start()

    def on_success(self, output_dir):
        self.generate_btn.setEnabled(True)
        self.log.append("\nDone. Opening output folder.")
        self.open_folder(Path(output_dir))
        QMessageBox.information(self, "Done", "Documents generated successfully.")

    def on_failed(self, message):
        self.generate_btn.setEnabled(True)
        self.log.append("\nERROR:\n" + message)
        QMessageBox.critical(self, "Generation failed", message[:3000])

    def open_folder(self, folder):
        try:
            folder = Path(folder)
            folder.mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))
        except Exception as e:
            QMessageBox.warning(self, "Could not open folder", str(e))


def main():
    app = QApplication(sys.argv)
    win = UASWindowsUI()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

