"""
Modern Fancy PyQt6 GUI for Documentation Slayer
Beautiful dark theme with smooth animations and professional design
"""

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QTabWidget, QPushButton, QLabel, QLineEdit, QCheckBox,
                              QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem,
                              QProgressDialog, QDialog, QDialogButtonBox, QGroupBox,
                              QGridLayout, QHeaderView, QFrame, QScrollArea)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QRect, QSize, pyqtProperty
from PyQt6.QtGui import QIcon, QFont, QColor, QPalette, QLinearGradient, QPainter, QPixmap, QPen
from pathlib import Path
import sys
import subprocess


class ModernToggleSwitch(QCheckBox):
    """Custom toggle switch widget with smooth animation"""
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setFixedSize(50, 25)
        self._circle_position = 3
        self._animation_progress = 0.0

        # Initialize animations after properties are set
        self.animation = QPropertyAnimation(self, b"circle_position", self)
        self.animation.setEasingCurve(QEasingCurve.Type.OutBounce)
        self.animation.setDuration(300)
        self.animation.setStartValue(3)
        self.animation.setEndValue(3)

        # Add color animation
        self.color_animation = QPropertyAnimation(self, b"animation_progress", self)
        self.color_animation.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self.color_animation.setDuration(300)
        self.color_animation.setStartValue(0.0)
        self.color_animation.setEndValue(0.0)
    
        self.stateChanged.connect(self.animate_toggle)

    def animate_toggle(self, state):
        self.animation.stop()

        # Set start values based on current state
        self.animation.setStartValue(self._circle_position)
        self.color_animation.setStartValue(self._animation_progress)
        
        if state:
            self.animation.setEndValue(25)
            self.color_animation.setEndValue(1.0)
        else:
            self.animation.setEndValue(3)
            self.color_animation.setEndValue(0.0)
        
        self.animation.start()
        self.color_animation.start()


    # Use PyQt's pyqtProperty for proper property binding
    @pyqtProperty(float)
    def animation_progress(self):
        return self._animation_progress
    
    @animation_progress.setter
    def animation_progress(self, value):
        self._animation_progress = value
        self.update()
    
    @pyqtProperty(int)
    def circle_position(self):
        return self._circle_position
    
    @circle_position.setter
    def circle_position(self, pos):
        self._circle_position = pos
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Animated background color
        gray = QColor(80, 80, 80)
        blue = QColor(33, 150, 243)
        # Interpolate between gray and blue based on animation progress
        r = int(gray.red() + (blue.red() - gray.red()) * self._animation_progress)
        g = int(gray.green() + (blue.green() - gray.green()) * self._animation_progress)
        b = int(gray.blue() + (blue.blue() - gray.blue()) * self._animation_progress)
        painter.setBrush(QColor(r, g, b))


        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, 50, 25, 12, 12)

        # Animated circle with glow effect when active
        if self._animation_progress > 0.5:
            # Add glow effect
            glow_size = int(2 * self._animation_progress)
            painter.setBrush(QColor(33, 150, 243, 50))  # Semi-transparent blue
            painter.drawEllipse(int(self._circle_position) - glow_size, 
                              2 - glow_size, 
                              21 + glow_size * 2, 
                              21 + glow_size * 2)
        
        # Main circle with color transition
        gray_circle = QColor(220, 220, 220)
        cyan_circle = QColor(200, 240, 255)
        cr = int(gray_circle.red() + (cyan_circle.red() - gray_circle.red()) * self._animation_progress)
        cg = int(gray_circle.green() + (cyan_circle.green() - gray_circle.green()) * self._animation_progress)
        cb = int(gray_circle.blue() + (cyan_circle.blue() - gray_circle.blue()) * self._animation_progress)
        painter.setBrush(QColor(cr, cg, cb))
        painter.drawEllipse(int(self._circle_position), 2, 21, 21)
    
        # Inner circle with animation
        inner_gray = QColor(180, 180, 180)
        inner_cyan = QColor(150, 220, 255)
        ir = int(inner_gray.red() + (inner_cyan.red() - inner_gray.red()) * self._animation_progress)
        ig = int(inner_gray.green() + (inner_cyan.green() - inner_gray.green()) * self._animation_progress)
        ib = int(inner_gray.blue() + (inner_cyan.blue() - inner_gray.blue()) * self._animation_progress)
        painter.setBrush(QColor(ir, ig, ib))
        painter.drawEllipse(int(self._circle_position) + 5, 7, 11, 11)


    def hitButton(self, pos):
        """Make entire widget clickable"""
        return self.rect().contains(pos)



class ModernCard(QFrame):
    """Modern card widget with shadow effect"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            ModernCard {
                background-color: #2d2d2d;
                border-radius: 12px;
                border: 1px solid #3a3a3a;
            }
            ModernCard:hover {
                border: 1px solid #2196F3;
            }
        """)


class DocumentationSlayerModernGUI(QMainWindow):
    """Ultra modern main window for Documentation Slayer"""

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
            "Name", "Syntax", "Description", "Used Data Types", "Triggers", "Sync/Async",
            "Function Type", "Invoked Operations", "Out-Parameters", "Return Value",
            "Outputs", "In-Parameters", "Reentrancy", "Line Number"
        ]
        self.macro_fields = ["Name", "Value", "Line Number"]
        self.variable_fields = ["Name", "Data Type", "Initial Value", "Scope", "Line Number"]
        self.formats = ["Excel", "Word", "MD"]

        # Worker thread
        self.worker_thread = None

        # Initialize UI
        self.init_ui()
        self.apply_modern_stylesheet()

    def init_ui(self):
        """Initialize the ultra modern user interface"""
        self.setWindowTitle("Documentation Slayer")
        self.setGeometry(100, 100, 1100, 750)
        self.setMinimumSize(1100, 650)

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

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Sidebar
        sidebar = self.create_sidebar()
        main_layout.addWidget(sidebar)

        # Main content area
        content_area = QWidget()
        content_layout = QVBoxLayout(content_area)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(20)

        # Header
        header = self.create_header()
        content_layout.addWidget(header)

        # Tab widget
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self.on_tab_changed)
        content_layout.addWidget(self.tabs)

        # Create tabs
        self.create_functions_tab()
        self.create_macros_tab()
        self.create_variables_tab()
        self.create_activity_diagram_tab()

        # Bottom panel
        self.create_bottom_panel(content_layout)

        main_layout.addWidget(content_area, 1)

        # Set initial button state
        if self.nav_buttons:
            self.switch_tab(0)

        # Center window
        # self.center_window()

    def create_sidebar(self):
        """Create modern sidebar navigation"""
        sidebar = QWidget()
        sidebar.setFixedWidth(280)
        sidebar.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                           stop:0 #1a1a1a, stop:1 #2d2d2d);
                border-right: 1px solid #3a3a3a;
            }
        """)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(20, 30, 20, 20)
        layout.setSpacing(20)

        # Logo and title
        logo_container = QWidget()
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Try to load actual logo
        try:
            if getattr(sys, "frozen", False):
                base_path = Path(sys._MEIPASS)
            else:
                base_path = Path(__file__).parent

            icon_path = base_path / "DocSlayerLogo.ico"
            if icon_path.exists():
                logo_pixmap = QPixmap(str(icon_path))
                if not logo_pixmap.isNull():
                    # Scale logo to fit
                    scaled_logo = logo_pixmap.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio,
                                                    Qt.TransformationMode.SmoothTransformation)
                    logo_label = QLabel()
                    logo_label.setPixmap(scaled_logo)
                    logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    logo_layout.addWidget(logo_label)
                else:
                    raise Exception("Invalid pixmap")
            else:
                raise Exception("Logo not found")
        except:
            # Fallback to emoji
            logo_label = QLabel("⚔️")
            logo_label.setFont(QFont("Segoe UI", 32))
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            logo_layout.addWidget(logo_label)

        # Title
        title = QLabel("Documentation")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #2196F3;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_layout.addWidget(title)

        subtitle = QLabel("Slayer")
        subtitle.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        subtitle.setStyleSheet("color: #2196F3;")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_layout.addWidget(subtitle)

        layout.addWidget(logo_container)

        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #3a3a3a; max-height: 1px;")
        layout.addWidget(separator)

        # Navigation buttons
        nav_buttons = [
            (" ", "Functions", 0),
            (" ", "Macros", 1),
            (" ", "Variables", 2),
            (" ", "Activity Diagram", 3)
        ]

        self.nav_buttons = []
        for idx, (icon, text, tab_index) in enumerate(nav_buttons):
            btn = QPushButton(f"{icon}  {text}")
            btn.setFixedHeight(45)
            btn.setMinimumWidth(240)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setCheckable(False)
            btn.setAutoDefault(False)
            btn.setDefault(False)
            btn.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, False)
            btn.setMouseTracking(True)

            # Create a proper closure for the lambda
            def make_handler(index):
                return lambda: self.switch_tab(index)

            btn.clicked.connect(make_handler(tab_index))
            # Also connect to pressed for immediate feedback
            btn.pressed.connect(make_handler(tab_index))

            btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: #b0b0b0;
                    border: none;
                    border-radius: 8px;
                    padding: 12px 20px;
                    text-align: left;
                    font-size: 14px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #2a2a2a;
                    color: #2196F3;
                }
                QPushButton:pressed {
                    background-color: #3a3a3a;
                }
            """)
            layout.addWidget(btn)
            self.nav_buttons.append(btn)

        layout.addStretch()

        # Version info
        version_label = QLabel("v3.3.0")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_label.setStyleSheet("color: #666666; font-size: 11px;")
        layout.addWidget(version_label)

        return sidebar

    def create_header(self):
        """Create modern header with title"""
        header = QWidget()
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 10)

        title = QLabel("Slay your C code 🗡️")
        title.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title.setStyleSheet("color: #ffffff;")
        header_layout.addWidget(title)

        subtitle = QLabel("Select the fields you want to export")
        subtitle.setFont(QFont("Segoe UI", 12))
        subtitle.setStyleSheet("color: #888888;")
        header_layout.addWidget(subtitle)

        return header

    def create_functions_tab(self):
        """Create modern Functions tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(15)

        # Action buttons
        actions = QWidget()
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)

        select_all_btn = QPushButton("✓ Select All")
        deselect_all_btn = QPushButton("✗ Deselect All")

        for btn in [select_all_btn, deselect_all_btn]:
            btn.setFixedHeight(32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 8px 20px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #1976D2;
                }
                QPushButton:pressed {
                    background-color: #0D47A1;
                }
            """)

        select_all_btn.clicked.connect(lambda: self.select_all_toggles(self.function_toggles))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_toggles(self.function_toggles))

        actions_layout.addWidget(select_all_btn)
        actions_layout.addWidget(deselect_all_btn)
        actions_layout.addStretch()

        layout.addWidget(actions)

        # Scroll area for toggles
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        scroll_content = QWidget()
        grid = QGridLayout(scroll_content)
        grid.setSpacing(10)
        grid.setContentsMargins(0, 0, 10, 0)  # Right margin for scrollbar

        self.function_toggles = {}
        for i, field in enumerate(self.function_fields):
            card = ModernCard()
            card.setFixedHeight(65)
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(15, 10, 15, 10)

            label = QLabel(field)
            label.setFont(QFont("Segoe UI", 11, QFont.Weight.Normal))
            label.setStyleSheet("color: #ffffff; background-color: transparent;")
            label.setWordWrap(False)
            label.setMinimumWidth(180)
            card_layout.addWidget(label, 1)

            toggle = ModernToggleSwitch()
            toggle.setFixedSize(50, 25)
            toggle.setChecked(field != "Line Number")
            self.function_toggles[field] = toggle
            card_layout.addWidget(toggle)

            grid.addWidget(card, i // 2, i % 2)

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        self.tabs.addTab(tab, "Functions")

    def create_macros_tab(self):
        """Create modern Macros tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(15)

        # Action buttons
        actions = QWidget()
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)

        select_all_btn = QPushButton("✓ Select All")
        deselect_all_btn = QPushButton("✗ Deselect All")

        for btn in [select_all_btn, deselect_all_btn]:
            btn.setFixedHeight(32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 8px 20px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #1976D2;
                }
                QPushButton:pressed {
                    background-color: #0D47A1;
                }
            """)

        select_all_btn.clicked.connect(lambda: self.select_all_toggles(self.macro_toggles))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_toggles(self.macro_toggles))

        actions_layout.addWidget(select_all_btn)
        actions_layout.addWidget(deselect_all_btn)
        actions_layout.addStretch()

        layout.addWidget(actions)

        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        scroll_content = QWidget()
        grid = QGridLayout(scroll_content)
        grid.setSpacing(10)
        grid.setContentsMargins(0, 0, 10, 0)

        self.macro_toggles = {}
        for i, field in enumerate(self.macro_fields):
            card = ModernCard()
            card.setFixedHeight(65)
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(15, 10, 15, 10)

            label = QLabel(field)
            label.setFont(QFont("Segoe UI", 11, QFont.Weight.Normal))
            label.setStyleSheet("color: #ffffff; background-color: transparent;")
            label.setWordWrap(False)
            label.setMinimumWidth(180)
            card_layout.addWidget(label, 1)

            toggle = ModernToggleSwitch()
            toggle.setFixedSize(50, 25)
            toggle.setChecked(field != "Line Number")
            self.macro_toggles[field] = toggle
            card_layout.addWidget(toggle)

            grid.addWidget(card, i // 2, i % 2)

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        self.tabs.addTab(tab, "Macros")

    def create_variables_tab(self):
        """Create modern Variables tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 20, 0, 0)
        layout.setSpacing(15)

        # Action buttons
        actions = QWidget()
        actions_layout = QHBoxLayout(actions)
        actions_layout.setContentsMargins(0, 0, 0, 0)

        select_all_btn = QPushButton("✓ Select All")
        deselect_all_btn = QPushButton("✗ Deselect All")

        for btn in [select_all_btn, deselect_all_btn]:
            btn.setFixedHeight(32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 8px 20px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #1976D2;
                }
                QPushButton:pressed {
                    background-color: #0D47A1;
                }
            """)

        select_all_btn.clicked.connect(lambda: self.select_all_toggles(self.variable_toggles))
        deselect_all_btn.clicked.connect(lambda: self.deselect_all_toggles(self.variable_toggles))

        actions_layout.addWidget(select_all_btn)
        actions_layout.addWidget(deselect_all_btn)
        actions_layout.addStretch()

        layout.addWidget(actions)

        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        scroll_content = QWidget()
        grid = QGridLayout(scroll_content)
        grid.setSpacing(10)
        grid.setContentsMargins(0, 0, 10, 0)

        self.variable_toggles = {}
        for i, field in enumerate(self.variable_fields):
            card = ModernCard()
            card.setFixedHeight(65)
            card_layout = QHBoxLayout(card)
            card_layout.setContentsMargins(15, 10, 15, 10)

            label = QLabel(field)
            label.setFont(QFont("Segoe UI", 11, QFont.Weight.Normal))
            label.setStyleSheet("color: #ffffff; background-color: transparent;")
            label.setWordWrap(False)
            label.setMinimumWidth(180)
            card_layout.addWidget(label, 1)

            toggle = ModernToggleSwitch()
            toggle.setFixedSize(50, 25)
            toggle.setChecked(field != "Line Number")
            self.variable_toggles[field] = toggle
            card_layout.addWidget(toggle)

            grid.addWidget(card, i // 2, i % 2)

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        self.tabs.addTab(tab, "Variables")

    def create_activity_diagram_tab(self):
        """Create modern Activity Diagram tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(30)

        # Icon
        icon_label = QLabel("📈")
        icon_label.setFont(QFont("Segoe UI", 72))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        # Title
        title = QLabel("Activity Diagram Generator")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #ffffff;")
        layout.addWidget(title)

        # Description
        desc = QLabel("Generate beautiful activity diagrams from your C source code")
        desc.setFont(QFont("Segoe UI", 12))
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setStyleSheet("color: #888888;")
        layout.addWidget(desc)

        # Generate button
        generate_btn = QPushButton("Generate Activity Diagram")
        generate_btn.setFixedSize(280, 50)
        generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        generate_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        generate_btn.clicked.connect(self.generate_activity_diagram)
        generate_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #2196F3, stop:1 #21CBF3);
                color: white;
                border: none;
                border-radius: 25px;
                font-size: 13px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #1976D2, stop:1 #00BCD4);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #0D47A1, stop:1 #0097A7);
            }
        """)
        layout.addWidget(generate_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addStretch()

        self.tabs.addTab(tab, "Activity Diagram")

    def create_bottom_panel(self, parent_layout):
        """Create modern bottom panel"""
        # Create scroll area for bottom panel
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setMinimumHeight(200)

        panel = ModernCard()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setSpacing(20)

        # Formats section
        formats_label = QLabel("Export Formats")
        formats_label.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        formats_label.setStyleSheet("color: #ffffff;")
        panel_layout.addWidget(formats_label)

        formats_container = QWidget()
        formats_container.setMinimumHeight(40)
        formats_layout = QHBoxLayout(formats_container)
        formats_layout.setContentsMargins(0, 10, 0, 0)
        formats_layout.setSpacing(50) # More space between format options

        self.format_toggles = {}
        format_labels = {"Excel": "Excel", "Word": "Word", "MD": "Markdown"}

        for fmt in ["Excel", "Word", "MD"]:
            # Simple horizontal layout for each format
            fmt_widget = QWidget()
            fmt_layout = QHBoxLayout(fmt_widget)
            fmt_layout.setContentsMargins(0, 0, 0, 0)
            fmt_layout.setSpacing(15)
            
            # Label
            label = QLabel(format_labels[fmt])
            label.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
            label.setStyleSheet("color: #2196F3;")
            label.setMinimumWidth(80)  # Ensure labels have minimum width
            fmt_layout.addWidget(label)
            
            # Toggle
            toggle = ModernToggleSwitch()
            toggle.setChecked(True)
            self.format_toggles[fmt] = toggle
            fmt_layout.addWidget(toggle)
            
            formats_layout.addWidget(fmt_widget)

        formats_layout.addStretch()
        panel_layout.addWidget(formats_container)

        # Add separator
        panel_layout.addSpacing(10)

        # Output directory
        dir_container = QWidget()
        dir_layout = QHBoxLayout(dir_container)
        dir_layout.setContentsMargins(0, 10, 0, 10)
        dir_layout.setSpacing(10)

        dir_icon = QLabel("📁")
        dir_icon.setFont(QFont("Segoe UI", 16))
        dir_layout.addWidget(dir_icon)

        self.save_dir_input = QLineEdit(str(Path.cwd()))
        self.save_dir_input.setReadOnly(True)
        self.save_dir_input.setFixedHeight(40)
        self.save_dir_input.setStyleSheet("""
            QLineEdit {
                background-color: #1f1f1f;
                color: #e0e0e0;
                border: 1px solid #3a3a3a;
                border-radius: 8px;
                padding: 10px 15px;
                font-size: 11px;
            }
        """)
        dir_layout.addWidget(self.save_dir_input, 1)

        browse_btn = QPushButton("Browse...")
        browse_btn.setFixedHeight(40)
        browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        browse_btn.clicked.connect(self.choose_directory)
        browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #3a3a3a;
                color: #e0e0e0;
                border: none;
                border-radius: 8px;
                padding: 10px 25px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #4a4a4a;
            }
        """)
        dir_layout.addWidget(browse_btn)

        panel_layout.addWidget(dir_container)

        # Run button
        panel_layout.addSpacing(15)
        run_btn = QPushButton("Run Documentation Slayer")
        run_btn.setFixedHeight(60)
        run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        run_btn.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        run_btn.clicked.connect(self.on_run)
        run_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #4CAF50, stop:1 #45a049);
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 14px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #45a049, stop:1 #3d8b40);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                           stop:0 #3d8b40, stop:1 #2E7D32);
            }
        """)
        panel_layout.addWidget(run_btn)
        parent_layout.addWidget(panel)
        scroll_area.setWidget(panel)
        parent_layout.addWidget(scroll_area)

    def apply_modern_stylesheet(self):
        """Apply the ultra modern dark theme stylesheet"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1a1a1a;
            }
            QWidget {
                background-color: #1a1a1a;
                color: #e0e0e0;
            }
            QTabWidget::pane {
                border: none;
                background-color: #1a1a1a;
            }
            QTabBar::tab {
                background-color: transparent;
                color: #888888;
                padding: 12px 24px;
                margin-right: 5px;
                border: none;
                border-bottom: 2px solid transparent;
                font-size: 12px;
                font-weight: 500;
            }
            QTabBar::tab:selected {
                color: #2196F3;
                border-bottom: 2px solid #2196F3;
            }
            QTabBar::tab:hover {
                color: #64B5F6;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #2b2b2b;
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #4a4a4a;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #5a5a5a;
            }
        """)

    def select_all_toggles(self, toggle_dict):
        """Select all toggles"""
        for toggle in toggle_dict.values():
            toggle.setChecked(True)

    def deselect_all_toggles(self, toggle_dict):
        """Deselect all toggles"""
        for toggle in toggle_dict.values():
            toggle.setChecked(False)

    def choose_directory(self):
        """Choose output directory"""
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory", str(Path.cwd()))
        if directory:
            self.save_dir_input.setText(directory)

    # No longer needed - starting maximized instead
    # def center_window(self):
    #     """Center window on screen"""
    #     frame_geometry = self.frameGeometry()
    #     screen_center = self.screen().availableGeometry().center()
    #     frame_geometry.moveCenter(screen_center)
    #     self.move(frame_geometry.topLeft())

    def switch_tab(self, tab_index):
        """Switch to the specified tab"""
        self.tabs.setCurrentIndex(tab_index)
        # Update button states
        for i, btn in enumerate(self.nav_buttons):
            if i == tab_index:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #2196F3;
                        color: #ffffff;
                        border: none;
                        border-radius: 8px;
                        padding: 12px 20px;
                        text-align: left;
                        font-size: 14px;
                        font-weight: 500;
                    }
                """)
            else:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        color: #b0b0b0;
                        border: none;
                        border-radius: 8px;
                        padding: 12px 20px;
                        text-align: left;
                        font-size: 14px;
                        font-weight: 500;
                    }
                    QPushButton:hover {
                        background-color: #2a2a2a;
                        color: #2196F3;
                    }
                    QPushButton:pressed {
                        background-color: #3a3a3a;
                    }
                """)

    def on_tab_changed(self, index):
        """Handle tab change"""
        if self.tabs.tabText(index) == "Activity Diagram" and not self.password_manager.is_authenticated:
            from parser import ask_password
            if not ask_password(self):
                self.tabs.setCurrentIndex(2)

    def generate_activity_diagram(self):
        """Generate activity diagram"""
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
                    QMessageBox.information(self, "Success", "Activity diagrams got Slayed (generated) successfully!")
                else:
                    QMessageBox.critical(self, "Error", f"Failed to generate activity diagrams:\n{result.stderr}")
            else:
                QMessageBox.critical(self, "Error", "CodeSmasher.exe not found!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error running CodeSmasher.exe: {str(e)}")

    def closeEvent(self, event):
        """Handle window close"""
        if self.worker_thread is not None and self.worker_thread.isRunning():
            reply = QMessageBox.question(self, 'Confirm Exit',
                                        'A parsing operation is in progress. Exit anyway?',
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.worker_thread.terminate()
                self.worker_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def on_run(self):
        """Handle Run button - same logic as before"""
        from parser import CancellationToken, ParserThread

        # Get selections
        sel_function_fields = [f for f, t in self.function_toggles.items() if t.isChecked()]
        sel_macro_fields = [f for f, t in self.macro_toggles.items() if t.isChecked()]
        sel_variable_fields = [f for f, t in self.variable_toggles.items() if t.isChecked()]
        sel_formats = [f for f, t in self.format_toggles.items() if t.isChecked()]
        outdir = Path(self.save_dir_input.text())

        # Select file
        cfile, _ = QFileDialog.getOpenFileName(self, "Select C Source File", str(Path.cwd()), "C Files (*.c);;All Files (*)")
        if not cfile:
            return

        # Clean up previous thread
        if self.worker_thread is not None:
            if self.worker_thread.isRunning():
                self.worker_thread.wait()
            self.worker_thread.deleteLater()
            self.worker_thread = None

        # Create worker and progress
        cancel_token = CancellationToken()
        progress = QProgressDialog("Processing file...", "Cancel", 0, 0, self)
        progress.setWindowTitle("Documentation Slayer")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)

        self.worker_thread = ParserThread(cfile, cancel_token)

        def on_finished(success, error, functions, macros, variables):
            if self.worker_thread:
                self.worker_thread.wait()
            progress.close()

            if not success:
                if error and error != "Operation cancelled":
                    QMessageBox.critical(self, "Error", f"An error occurred:\n{error}")
                elif error == "Operation cancelled":
                    QMessageBox.warning(self, "Cancelled", "Operation was cancelled.")
                return

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

                # Open files
                if "Excel" in sel_formats:
                    self.open_file_func(str(xlsx_path))
                if "Word" in sel_formats:
                    self.open_file_func(str(xlsx_path.with_suffix('.docx')))
                if "MD" in sel_formats:
                    self.open_file_func(str(outdir / f"{stem}.md"))

                QMessageBox.information(self, "Success",
                                      f"Documentation got Slayed successfully!\n\nFile: {stem}\nFormats: {', '.join(sel_formats)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed:\n{str(e)}")

        self.worker_thread.progress.connect(lambda text: progress.setLabelText(text))
        self.worker_thread.finished.connect(on_finished)
        progress.canceled.connect(lambda: (cancel_token.cancel(), self.worker_thread.wait(2000) if self.worker_thread else None))

        self.worker_thread.start()
