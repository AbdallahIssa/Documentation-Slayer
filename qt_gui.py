"""
PyQt6 GUI for Documentation Slayer
Modern, professional interface with tabbed layout and sortable tables
"""

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QTabWidget, QPushButton, QLabel, QLineEdit, QCheckBox,
                              QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
                              QProgressDialog, QDialog, QDialogButtonBox, QGroupBox,
                              QGridLayout, QHeaderView, QStyle, QComboBox, QAbstractItemView)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont, QColor
from pathlib import Path
import sys
import subprocess


class DocumentationSlayerGUI(QMainWindow):
    """Main window for Documentation Slayer"""

    def __init__(self, password_manager, parse_file_func, write_excel_func, write_docx_func, write_markdown_func, open_file_func):
        super().__init__()

        # Store references to functions
        self.password_manager = password_manager
        self.parse_file = parse_file_func
        self.write_excel = write_excel_func
        self.write_docx = write_docx_func
        self.write_markdown = write_markdown_func
        self.open_file_func = open_file_func

        # Field definitions
        self.function_fields = [
            "Line Number", "Name", "Description", "Syntax", "Triggers", "In-Parameters", "Out-Parameters",
            "Return Value", "Function Type", "Inputs", "Outputs",
            "Invoked Operations", "Used Data Types", "Sync/Async", "Reentrancy"
        ]
        self.macro_fields = ["Line Number", "Name", "Value"]
        self.variable_fields = ["Line Number", "Name", "Data Type", "Initial Value", "Scope"]
        self.formats = ["Excel", "Word", "MD"]

        # Worker thread (keep reference to prevent premature destruction)
        self.worker_thread = None

        # Initialize UI
        self.init_ui()

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Documentation Slayer")
        self.setGeometry(100, 100, 1000, 700)

        # Set application icon
        try:
            if getattr(sys, "frozen", False):
                base_path = Path(sys._MEIPASS)
            else:
                base_path = Path(__file__).parent
            icon_path = base_path / "DocSlayerLogo.ico"
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except:
            pass

        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Create tab widget
        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self.on_tab_changed)
        main_layout.addWidget(self.tabs)

        # Create tabs
        self.create_functions_tab()
        self.create_macros_tab()
        self.create_variables_tab()
        self.create_activity_diagram_tab()

        # Bottom settings panel
        self.create_settings_panel(main_layout)

        # Center window on screen
        self.center_window()

    def closeEvent(self, event):
        """Handle window close event - clean up threads"""
        if self.worker_thread is not None and self.worker_thread.isRunning():
            reply = QMessageBox.question(self, 'Confirm Exit',
                                        'A parsing operation is in progress. Are you sure you want to exit?',
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                        QMessageBox.StandardButton.No)

            if reply == QMessageBox.StandardButton.Yes:
                # Force thread termination
                self.worker_thread.terminate()
                self.worker_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def center_window(self):
        """Center the window on the screen"""
        frame_geometry = self.frameGeometry()
        screen_center = self.screen().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())

    def create_functions_tab(self):
        """Create the Functions tab with checkboxes"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Header with select all/deselect all buttons
        header_layout = QHBoxLayout()
        header_label = QLabel("Select function fields:")
        header_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        select_all_btn.clicked.connect(lambda: self.select_all_checkboxes(self.function_checkboxes))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_checkboxes(self.function_checkboxes))
        header_layout.addWidget(select_all_btn)
        header_layout.addWidget(deselect_all_btn)
        layout.addLayout(header_layout)

        # Checkboxes in grid
        grid = QGridLayout()
        self.function_checkboxes = {}
        for i, field in enumerate(self.function_fields):
            checkbox = QCheckBox(field)
            checkbox.setChecked(field != "Line Number")
            self.function_checkboxes[field] = checkbox
            grid.addWidget(checkbox, i // 3, i % 3)

        layout.addLayout(grid)
        layout.addStretch()

        self.tabs.addTab(tab, "Functions")

    def create_macros_tab(self):
        """Create the Macros tab with checkboxes"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel("Select macro fields:")
        header_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        select_all_btn.clicked.connect(lambda: self.select_all_checkboxes(self.macro_checkboxes))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_checkboxes(self.macro_checkboxes))
        header_layout.addWidget(select_all_btn)
        header_layout.addWidget(deselect_all_btn)
        layout.addLayout(header_layout)

        # Checkboxes
        grid = QGridLayout()
        self.macro_checkboxes = {}
        for i, field in enumerate(self.macro_fields):
            checkbox = QCheckBox(field)
            checkbox.setChecked(field != "Line Number")
            self.macro_checkboxes[field] = checkbox
            grid.addWidget(checkbox, i // 3, i % 3)

        layout.addLayout(grid)
        layout.addStretch()

        self.tabs.addTab(tab, "Macros")

    def create_variables_tab(self):
        """Create the Variables tab with checkboxes"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Header
        header_layout = QHBoxLayout()
        header_label = QLabel("Select variable fields:")
        header_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        select_all_btn.clicked.connect(lambda: self.select_all_checkboxes(self.variable_checkboxes))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_checkboxes(self.variable_checkboxes))
        header_layout.addWidget(select_all_btn)
        header_layout.addWidget(deselect_all_btn)
        layout.addLayout(header_layout)

        # Checkboxes
        grid = QGridLayout()
        self.variable_checkboxes = {}
        for i, field in enumerate(self.variable_fields):
            checkbox = QCheckBox(field)
            checkbox.setChecked(field != "Line Number")
            self.variable_checkboxes[field] = checkbox
            grid.addWidget(checkbox, i // 3, i % 3)

        layout.addLayout(grid)
        layout.addStretch()

        self.tabs.addTab(tab, "Variables")

    def create_activity_diagram_tab(self):
        """Create the password-protected Activity Diagram tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Title
        title = QLabel("Activity Diagram Generator")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Description
        desc = QLabel("Click the button below to generate activity diagrams from your C source file.")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(desc)

        # Generate button
        generate_btn = QPushButton("Generate Activity Diagram")
        generate_btn.setMinimumHeight(40)
        generate_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        generate_btn.clicked.connect(self.generate_activity_diagram)
        layout.addWidget(generate_btn)

        self.tabs.addTab(tab, "Activity Diagram")

    def create_settings_panel(self, parent_layout):
        """Create the bottom settings panel"""
        settings_group = QGroupBox("Export Settings")
        settings_layout = QVBoxLayout()

        # Format selection
        format_layout = QHBoxLayout()
        format_label = QLabel("Select formats:")
        format_layout.addWidget(format_label)

        self.format_checkboxes = {}
        for fmt in self.formats:
            checkbox = QCheckBox(fmt)
            checkbox.setChecked(True)
            self.format_checkboxes[fmt] = checkbox
            format_layout.addWidget(checkbox)

        format_layout.addStretch()
        settings_layout.addLayout(format_layout)

        # Output directory selection
        dir_layout = QHBoxLayout()
        dir_label = QLabel("Output directory:")
        dir_layout.addWidget(dir_label)

        self.save_dir_input = QLineEdit(str(Path.cwd()))
        self.save_dir_input.setReadOnly(True)
        dir_layout.addWidget(self.save_dir_input)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.choose_directory)
        dir_layout.addWidget(browse_btn)

        settings_layout.addLayout(dir_layout)

        # Run button
        run_btn = QPushButton("Run Documentation Slayer")
        run_btn.setMinimumHeight(40)
        run_btn.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        run_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; }")
        run_btn.clicked.connect(self.on_run)
        settings_layout.addWidget(run_btn)

        settings_group.setLayout(settings_layout)
        parent_layout.addWidget(settings_group)

    def select_all_checkboxes(self, checkbox_dict):
        """Select all checkboxes in a dictionary"""
        for checkbox in checkbox_dict.values():
            checkbox.setChecked(True)

    def deselect_all_checkboxes(self, checkbox_dict):
        """Deselect all checkboxes in a dictionary"""
        for checkbox in checkbox_dict.values():
            checkbox.setChecked(False)

    def choose_directory(self):
        """Open directory chooser dialog"""
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory",
                                                     str(Path.cwd()))
        if directory:
            self.save_dir_input.setText(directory)

    def on_tab_changed(self, index):
        """Handle tab change event for password protection"""
        if self.tabs.tabText(index) == "Activity Diagram" and not self.password_manager.is_authenticated:
            from parser import ask_password
            if not ask_password(self):
                # Go back to previous tab (Variables)
                self.tabs.setCurrentIndex(2)

    def generate_activity_diagram(self):
        """Generate activity diagram using CodeSmasher.exe"""
        if not self.password_manager.is_authenticated:
            from parser import ask_password
            if not ask_password(self):
                return

        try:
            if getattr(sys, "frozen", False):
                base_path = Path(sys._MEIPASS)
            else:
                base_path = Path(__file__).parent

            exe_path = base_path / "CodeSmasher.exe"
            if exe_path.exists():
                result = subprocess.run([str(exe_path)], capture_output=True, text=True)
                if result.returncode == 0:
                    QMessageBox.information(self, "Success",
                                          "Activity diagrams got Slayed (generated) successfully!")
                else:
                    QMessageBox.critical(self, "Error",
                                       f"Failed to generate activity diagrams:\n{result.stderr}")
            else:
                QMessageBox.critical(self, "Error",
                                   "CodeSmasher.exe not found! Please ensure it's in the same directory as this script.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error running CodeSmasher.exe: {str(e)}")

    def on_run(self):
        """Handle the Run button click"""
        # Get selected fields
        sel_function_fields = [f for f, cb in self.function_checkboxes.items() if cb.isChecked()]
        sel_macro_fields = [f for f, cb in self.macro_checkboxes.items() if cb.isChecked()]
        sel_variable_fields = [f for f, cb in self.variable_checkboxes.items() if cb.isChecked()]
        sel_formats = [f for f, cb in self.format_checkboxes.items() if cb.isChecked()]
        outdir = Path(self.save_dir_input.text())

        # Ask for C source file
        cfile, _ = QFileDialog.getOpenFileName(self, "Select C Source File",
                                               str(Path.cwd()),
                                               "C Files (*.c);;All Files (*)")
        if not cfile:
            return

        # Create progress dialog
        from parser import CancellationToken, ParserThread

        cancel_token = CancellationToken()
        progress = QProgressDialog("Processing file...", "Cancel", 0, 0, self)
        progress.setWindowTitle("Documentation Slayer")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)

        # Clean up previous worker thread if exists
        if self.worker_thread is not None:
            if self.worker_thread.isRunning():
                self.worker_thread.wait()
            self.worker_thread.deleteLater()
            self.worker_thread = None

        # Create worker thread and store reference
        self.worker_thread = ParserThread(cfile, cancel_token)

        def update_progress(text):
            progress.setLabelText(text)

        def on_finished(success, error, functions, macros, variables):
            # Wait for thread to fully complete
            if self.worker_thread:
                self.worker_thread.wait()
            progress.close()

            if not success:
                if error and error != "Operation cancelled":
                    QMessageBox.critical(self, "Error", f"An error occurred:\n{error}")
                elif error == "Operation cancelled":
                    QMessageBox.warning(self, "Cancelled", "Operation was cancelled by user.")
                return

            # Export files
            stem = Path(cfile).stem
            xlsx_path = outdir / f"{stem}.xlsx"

            try:
                if "Excel" in sel_formats:
                    self.write_excel(str(xlsx_path), functions, macros, variables,
                                   sel_function_fields, sel_macro_fields, sel_variable_fields)

                if "Word" in sel_formats:
                    self.write_docx(str(xlsx_path), stem, functions, macros, variables,
                                  sel_function_fields, sel_macro_fields, sel_variable_fields)

                if "MD" in sel_formats:
                    md_path = outdir / f"{stem}.md"
                    self.write_markdown(str(md_path), functions, macros, variables,
                                      sel_function_fields, sel_macro_fields, sel_variable_fields)

                # Open generated files
                if "Excel" in sel_formats:
                    self.open_file_func(str(xlsx_path))
                if "Word" in sel_formats:
                    self.open_file_func(str(xlsx_path.with_suffix('.docx')))
                if "MD" in sel_formats:
                    self.open_file_func(str(outdir / f"{stem}.md"))

                # Show success message
                QMessageBox.information(self, "Success",
                                      f"Documentation got Slayed (generated) successfully!\n\n"
                                      f"File: {stem}\nFormats: {', '.join(sel_formats)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

        def on_cancel():
            cancel_token.cancel()
            # Wait for thread to finish after cancellation
            if self.worker_thread:
                self.worker_thread.wait(2000)  # Wait up to 2 seconds

        self.worker_thread.progress.connect(update_progress)
        self.worker_thread.finished.connect(on_finished)
        progress.canceled.connect(on_cancel)

        self.worker_thread.start()
