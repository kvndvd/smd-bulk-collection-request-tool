from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
import ctypes

from PyQt5.QtCore import Qt, QEvent, QObject, QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QApplication
)

from appLogger import setup_logger
from emailBody import build_email_payload, send_email_via_outlook
from networkDrop import copy_outputs_to_network
from smdRequest import COUNSEL_FOLDER, COURT_FOLDER, generate_outputs


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_STYLESHEET = """
QWidget#root {
    background: #232323;
    color: white;
    font-family: "Calibri";
    font-size: 10px;
    border-radius: 20px;
}

QFrame#card {
    background: transparent;
}

QLabel#title {
    font-size: 15px;
    font-weight: 700;
    color: #f3f3f3;
}

QLabel#subtitle {
    font-size: 11px;
    font-weight: 500;
    color: #999999;
}

QLabel#sectionTitle {
    font-size: 12px;
    font-weight: 700;
    color: rgba(255,255,255,0.84);
}

QLabel#statusText {
    font-size: 12px;
    font-weight: 600;
    color: rgba(255,255,255,0.92);
}

QLabel#dot {
    background: #32d98a;
    border-radius: 10px;
    min-width: 20px;
    min-height: 20px;
    max-width: 20px;
    max-height: 20px;
}

QPushButton {
    border: none;
    border-radius: 20px;
    padding: 10px 16px;
    font-weight: 700;
    background: #006986;
    color: white;
}

QPushButton#autoBtn:hover,
QPushButton#checkBtn:hover,
QMessageBox QPushButton:hover {
    background: qlineargradient(
        x1: 0, y1: 0, x2: 1, y2: 0,
        stop: 0 #ad59ff, stop: 1 #ff4960
    );
}

QPushButton:disabled {
    background: rgba(255,255,255,0.08);
    color: rgba(255,255,255,0.35);
}

QPushButton#macMinBtn,
QPushButton#macCloseBtn,
QPushButton#devCloseBtn {
    border-radius: 10px;
    font-size: 11px;
    font-weight: 900;
    padding: 0px;
    color: rgba(0, 0, 0, 0.72);
}

QPushButton#macMinBtn {
    background: #f5c542;
}

QPushButton#macMinBtn:hover {
    background: #ffd45c;
}

QPushButton#macCloseBtn,
QPushButton#devCloseBtn {
    background: #ff5f57;
}

QPushButton#macCloseBtn:hover,
QPushButton#devCloseBtn:hover {
    background: #ff7b74;
}

QFrame#statusFrame,
QFrame#logShell {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.08);
}

QFrame#statusFrame {
    border-radius: 15px;
    min-height: 40px;
}

QFrame#logShell {
    border-radius: 18px;
}

QLineEdit,
QPlainTextEdit#logViewer {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    color: rgba(255,255,255,0.92);
    selection-background-color: rgba(155, 61, 255, 0.35);
}

QLineEdit {
    padding: 10px 12px;
}

QPlainTextEdit#logViewer {
    padding: 8px;
}

QLineEdit:focus,
QPlainTextEdit#logViewer:focus {
    border: 1px solid rgba(0, 132, 168, 0.75);
}

QCheckBox#extraCourtRequestCheck,
QCheckBox#networkDropCheck,
QCheckBox#sendEmailCheck {
    color: #f3f3f3;
    font-size: 13px;
    font-weight: 600;
    spacing: 10px;
}

QCheckBox#extraCourtRequestCheck::indicator,
QCheckBox#networkDropCheck::indicator,
QCheckBox#sendEmailCheck::indicator {
    width: 18px;
    height: 18px;
    border-radius: 9px;
    border: 1px solid rgba(255,255,255,0.35);
    background: rgba(255,255,255,0.06);
}

QCheckBox#extraCourtRequestCheck::indicator:checked,
QCheckBox#networkDropCheck::indicator:checked,
QCheckBox#sendEmailCheck::indicator:checked {
    background: #32d98a;
    border: 1px solid #32d98a;
}

QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 4px 2px 4px 2px;
}

QScrollBar::handle:vertical {
    background: rgba(255,255,255,0.18);
    border-radius: 4px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background: rgba(255,255,255,0.28);
}

QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical {
    background: none;
    border: none;
}

QMessageBox {
    background-color: #232323;
}

QMessageBox QLabel {
    color: #f3f3f3;
    font-size: 13px;
}

QMessageBox QPushButton {
    border-radius: 12px;
    padding: 8px 14px;
    min-width: 80px;
}
"""


class GenerateWorker(QObject):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    log = pyqtSignal(str)
    status = pyqtSignal(str)

    def __init__(
        self,
        template_path: Path,
        output_dir: Path,
        create_example_ct: bool,
        network_drop: bool,
        send_email: bool,
    ) -> None:
        super().__init__()
        self.template_path = template_path
        self.output_dir = output_dir
        self.create_example_ct = create_example_ct
        self.network_drop = network_drop
        self.send_email = send_email

    def run(self) -> None:
        try:
            self.status.emit("Generating files...")
            self.log.emit(f"Template: {self.template_path}")
            self.log.emit(f"Base output folder: {self.output_dir}")
            self.log.emit(f"CNL CSV folder: {self.output_dir / COUNSEL_FOLDER}")
            self.log.emit(f"CT CSV folder: {self.output_dir / COURT_FOLDER}")
            self.log.emit(f"Create extra CT example file: {'Yes' if self.create_example_ct else 'No'}")
            self.log.emit(f"Auto-drop to network path: {'Yes' if self.network_drop else 'No'}")
            self.log.emit(f"Send completion email: {'Yes' if self.send_email else 'No'}")

            paths = generate_outputs(
                self.template_path,
                self.output_dir,
                create_example_ct=self.create_example_ct,
            )
            self.log.emit("Local files created successfully.")

            if self.network_drop:
                self.status.emit("Copying files to network path...")
                self.log.emit("Copying files to network path...")
                copy_outputs_to_network(paths)
                self.log.emit("Network copy completed successfully.")

            if self.send_email:
                self.status.emit("Preparing completion email...")
                self.log.emit("Preparing completion email...")
                payload = build_email_payload(paths)
                send_email_via_outlook(payload, send_now=True)
                self.log.emit(
                    f"Completion email sent. Counsel count: {payload.counsel_count}, "
                    f"Court count: {payload.court_count}"
                )

            self.finished.emit(paths)

        except Exception as exc:
            self.error.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        self.logger = setup_logger()
        self.logger.info("MainWindow initialized")

        super().__init__()
        self._drag_pos = None
        self._dragging = False
        self.thread: QThread | None = None
        self.worker: GenerateWorker | None = None

        self.setWindowTitle("SMD Bulk Collection Request Tool")
        self.setFixedSize(450, 640)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self._build_ui()
        self.setStyleSheet(APP_STYLESHEET)
        self._set_status("Ready")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(16)

        card = QFrame()
        card.setObjectName("card")
        root_layout.addWidget(card)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        self.header_bar = QFrame()
        self.header_bar.setObjectName("card")
        header_layout = QHBoxLayout(self.header_bar)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(0)

        self.window_btn_wrap = QWidget()
        self.window_btn_wrap.setFixedWidth(60)
        self.window_btn_row = QHBoxLayout(self.window_btn_wrap)
        self.window_btn_row.setContentsMargins(0, 0, 0, 0)
        self.window_btn_row.setSpacing(8)

        self.close_btn = QPushButton("✕")
        self.close_btn.setObjectName("macCloseBtn")
        self.close_btn.setFixedSize(20, 20)
        self.close_btn.clicked.connect(self.close)

        self.min_btn = QPushButton("—")
        self.min_btn.setObjectName("macMinBtn")
        self.min_btn.setFixedSize(20, 20)
        self.min_btn.clicked.connect(self.showMinimized)

        self.window_btn_row.addWidget(self.close_btn)
        self.window_btn_row.addWidget(self.min_btn)
        self.window_btn_row.addStretch()

        self.title_wrap = QWidget()
        title_layout = QVBoxLayout(self.title_wrap)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(0)

        self.title_label = QLabel("SMD Bulk Collection Request Tool")
        self.title_label.setObjectName("title")
        self.title_label.setAlignment(Qt.AlignCenter)

        self.subtitle_label = QLabel("Court & Counsel Bulk Collection Request Tool")
        self.subtitle_label.setObjectName("subtitle")
        self.subtitle_label.setAlignment(Qt.AlignCenter)

        title_layout.addWidget(self.title_label)
        title_layout.addWidget(self.subtitle_label)

        self.header_right_spacer = QWidget()
        self.header_right_spacer.setFixedWidth(60)

        header_layout.addWidget(self.window_btn_wrap, 0, Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addStretch(1)
        header_layout.addWidget(self.title_wrap, 0, Qt.AlignCenter)
        header_layout.addStretch(1)
        header_layout.addWidget(self.header_right_spacer, 0, Qt.AlignRight | Qt.AlignVCenter)

        status_frame = QFrame()
        status_frame.setObjectName("statusFrame")
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(12, 10, 12, 10)
        status_layout.setSpacing(10)

        self.dot = QLabel()
        self.dot.setObjectName("dot")

        self.status_text = QLabel("Ready")
        self.status_text.setObjectName("statusText")

        status_layout.addWidget(self.dot)
        status_layout.addWidget(self.status_text, 1)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)
        grid.setColumnStretch(0, 1)

        template_label = QLabel("Template File")
        template_label.setObjectName("sectionTitle")

        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("Select the SMD Appeals Template (.xlsx or .xlsm)")

        browse_button = QPushButton("Browse...")
        browse_button.setObjectName("autoBtn")
        browse_button.clicked.connect(self.select_template)
        browse_button.setMinimumHeight(42)

        output_note_title = QLabel("Output Folders")
        output_note_title.setObjectName("sectionTitle")

        output_note = QLabel(
            "Created automatically beside the selected template file:\n"
            f"CNL-{COUNSEL_FOLDER}\n"
            f"CT-{COURT_FOLDER}"
        )
        output_note.setObjectName("subtitle")
        output_note.setWordWrap(True)

        extra_note = QLabel("Creates a new CT request template")
        extra_note.setObjectName("subtitle")
        extra_note.setWordWrap(True)

        self.extra_ct_checkbox = QCheckBox("Create new request template")
        self.extra_ct_checkbox.setObjectName("extraCourtRequestCheck")

        network_note = QLabel("Automatically copy generated files to the designated network path")
        network_note.setObjectName("subtitle")
        network_note.setWordWrap(True)

        self.network_drop_checkbox = QCheckBox("Auto-drop to network path")
        self.network_drop_checkbox.setObjectName("networkDropCheck")

        email_note = QLabel("Send the completion email after all processing is done")
        email_note.setObjectName("subtitle")
        email_note.setWordWrap(True)

        self.send_email_checkbox = QCheckBox("Send completion email (works only in Classic Outlook)")
        self.send_email_checkbox.setObjectName("sendEmailCheck")

        self.view_folder_button = QPushButton("View Folder")
        self.view_folder_button.setObjectName("autoBtn")
        self.view_folder_button.clicked.connect(self.view_output_folder)
        self.view_folder_button.setMinimumHeight(42)

        self.generate_button = QPushButton("Generate Files")
        self.generate_button.setObjectName("autoBtn")
        self.generate_button.clicked.connect(self.generate_files)
        self.generate_button.setMinimumHeight(42)

        grid.addWidget(template_label, 0, 0, 1, 2)
        grid.addWidget(self.template_edit, 1, 0)
        grid.addWidget(browse_button, 1, 1)
        grid.addWidget(output_note_title, 2, 0, 1, 2)
        grid.addWidget(output_note, 3, 0, 1, 2)
        grid.addWidget(extra_note, 4, 0, 1, 2)
        grid.addWidget(self.extra_ct_checkbox, 5, 0, 1, 2)
        grid.addWidget(network_note, 6, 0, 1, 2)
        grid.addWidget(self.network_drop_checkbox, 7, 0, 1, 2)
        grid.addWidget(email_note, 8, 0, 1, 2)
        grid.addWidget(self.send_email_checkbox, 9, 0, 1, 2)

        button_row = QHBoxLayout()
        button_row.addWidget(self.generate_button)
        button_row.addWidget(self.view_folder_button)

        log_shell = QFrame()
        log_shell.setObjectName("logShell")
        log_layout = QVBoxLayout(log_shell)
        log_layout.setContentsMargins(12, 12, 12, 12)
        log_layout.setSpacing(8)

        log_title = QLabel("Activity Log")
        log_title.setObjectName("sectionTitle")

        self.log_box = QPlainTextEdit()
        self.log_box.setObjectName("logViewer")
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Status messages will appear here.")

        log_layout.addWidget(log_title)
        log_layout.addWidget(self.log_box)

        layout.addWidget(self.header_bar)
        layout.addWidget(status_frame)
        layout.addLayout(grid)
        layout.addLayout(button_row)
        layout.addWidget(log_shell, 1)

        for widget in (
            self.header_bar,
            self.window_btn_wrap,
            self.title_wrap,
            self.title_label,
            self.subtitle_label,
            self.header_right_spacer,
        ):
            widget.installEventFilter(self)

    def set_controls_enabled(self, enabled: bool) -> None:
        self.generate_button.setEnabled(enabled)
        self.view_folder_button.setEnabled(enabled)
        self.extra_ct_checkbox.setEnabled(enabled)
        self.network_drop_checkbox.setEnabled(enabled)
        self.send_email_checkbox.setEnabled(enabled)
        self.template_edit.setEnabled(enabled)

    def view_output_folder(self) -> None:
        output_dir = get_app_dir()

        if not output_dir.exists():
            self.show_frameless_message(
                "Folder Not Found",
                f"Output folder was not found:\n\n{output_dir}",
                QMessageBox.Warning,
            )
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(output_dir)], check=True)
            else:
                subprocess.run(["xdg-open", str(output_dir)], check=True)
        except Exception as exc:
            self.show_frameless_message(
                "Open Folder Failed",
                f"Failed to open folder:\n\n{exc}",
                QMessageBox.Critical,
            )

    def show_frameless_message(self, title: str, message: str, icon=QMessageBox.Information) -> None:
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(icon)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        msg.exec_()

    def _set_status(self, message: str) -> None:
        self.status_text.setText(message)

    def log(self, message: str) -> None:
        self.log_box.appendPlainText(message)
        self.log_box.moveCursor(QTextCursor.End)
        self.logger.info(message)

    def select_template(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select the SMD Appeals Template",
            "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*.*)",
        )
        if file_path:
            self.template_edit.setText(file_path)
            self._set_status("Template selected")
            self.log(f"Selected template: {file_path}")

    def generate_files(self) -> None:
        if self.thread is not None and self.thread.isRunning():
            return

        template_text = self.template_edit.text().strip()

        if not template_text:
            self.logger.warning("Generate clicked without selecting a template")
            self.show_frameless_message("Missing Template", "Please select a template file.", QMessageBox.Warning)
            self._set_status("Waiting for template selection")
            return

        template_path = Path(template_text).expanduser().resolve()
        output_dir = get_app_dir()

        if not template_path.exists():
            self.logger.error("Template file not found: %s", template_path)
            self.show_frameless_message(
                "Template Not Found",
                f"Template file was not found:\n\n{template_path}",
                QMessageBox.Critical,
            )
            self._set_status("Template file not found")
            return

        self.log_box.clear()
        self.set_controls_enabled(False)
        self._set_status("Starting...")

        self.thread = QThread()
        self.worker = GenerateWorker(
            template_path=template_path,
            output_dir=output_dir,
            create_example_ct=self.extra_ct_checkbox.isChecked(),
            network_drop=self.network_drop_checkbox.isChecked(),
            send_email=self.send_email_checkbox.isChecked(),
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log.connect(self.log)
        self.worker.status.connect(self._set_status)
        self.worker.finished.connect(self.on_generate_finished)
        self.worker.error.connect(self.on_generate_error)

        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)

        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self._clear_worker_refs)

        self.thread.start()

    def on_generate_finished(self, paths) -> None:
        self.log("")
        self.log("Files created successfully:")
        self.log(f"- {paths.cnl_path}")
        self.log(f"- {paths.ct_path}")
        if paths.ct_example_path is not None:
            self.log(f"- {paths.ct_example_path}")
        self.log(f"- {paths.xlsm_path}")

        if self.network_drop_checkbox.isChecked():
            self.log("- Files were also copied to the network path.")

        if self.send_email_checkbox.isChecked():
            self.log("- Completion email was sent.")

        self._set_status("Completed successfully")
        self.set_controls_enabled(True)
        self.show_frameless_message("Success", "Done!", QMessageBox.Information)

    def on_generate_error(self, error_message: str) -> None:
        self.logger.exception("Generation failed")
        self.log(f"Error: {error_message}")
        self._set_status("Generation failed")
        self.set_controls_enabled(True)
        self.show_frameless_message(
            "Generation Failed",
            f"Failed to generate files:\n\n{error_message}",
            QMessageBox.Critical,
        )

    def _clear_worker_refs(self) -> None:
        self.worker = None
        self.thread = None

    def eventFilter(self, obj, event):
        draggable_widgets = {
            self.header_bar,
            self.window_btn_wrap,
            self.title_wrap,
            self.title_label,
            self.subtitle_label,
            self.header_right_spacer,
        }

        if obj in draggable_widgets:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                self._dragging = True
                self._drag_pos = event.globalPos() - self.frameGeometry().topLeft()
                event.accept()
                return True

            if event.type() == QEvent.MouseMove and self._dragging and event.buttons() & Qt.LeftButton:
                self.move(event.globalPos() - self._drag_pos)
                event.accept()
                return True

            if event.type() == QEvent.MouseButtonRelease:
                self._dragging = False
                self._drag_pos = None
                event.accept()
                return True

        return super().eventFilter(obj, event)


def run_cli(args: list[str]) -> int:
    if len(args) not in {2, 3, 4, 5}:
        print(
            "Usage:\n"
            "  python main.py <template.xlsx/xlsm> [--with-example] [--network-drop] [--send-email]\n"
            "  python main.py    # opens the PyQt GUI"
        )
        return 1

    template_path = Path(args[1]).expanduser().resolve()
    output_dir = get_app_dir()

    create_example = "--with-example" in args[2:]
    network_drop = "--network-drop" in args[2:]
    send_email = "--send-email" in args[2:]

    if not template_path.exists():
        print(f"Template not found: {template_path}")
        return 1

    try:
        paths = generate_outputs(
            template_path,
            output_dir,
            create_example_ct=create_example,
        )
        if network_drop:
            copy_outputs_to_network(paths)
        if send_email:
            payload = build_email_payload(paths)
            send_email_via_outlook(payload, send_now=True)
    except Exception as exc:
        print(f"Failed to generate, copy, or email files: {exc}")
        return 1

    print(f"Created: {paths.cnl_path}")
    print(f"Created: {paths.ct_path}")
    if paths.ct_example_path is not None:
        print(f"Created: {paths.ct_example_path}")
    print(f"Created: {paths.xlsm_path}")
    if network_drop:
        print("Copied files to network path.")
    if send_email:
        print("Completion email sent.")
    return 0

def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def run_gui() -> int:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "smd.bulk.collection.request.tool"
    )

    app = QApplication(sys.argv)

    icon_path = get_app_dir() / "icon.ico"
    app_icon = QIcon(str(icon_path))

    app.setWindowIcon(app_icon)

    window = MainWindow()
    window.setWindowIcon(app_icon)
    window.show()

    return app.exec_()

if __name__ == "__main__":
    if len(sys.argv) == 1:
        raise SystemExit(run_gui())
    raise SystemExit(run_cli(sys.argv))
