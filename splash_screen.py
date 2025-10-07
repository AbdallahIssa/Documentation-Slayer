"""
Splash Screen for Documentation Slayer
Beautiful loading screen with animated logo
"""

from PyQt6.QtWidgets import QSplashScreen, QApplication, QLabel, QVBoxLayout, QWidget
from PyQt6.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, QRect, QSize
from PyQt6.QtGui import QPixmap, QPainter, QColor, QFont, QLinearGradient, QPen
from pathlib import Path
import sys


class ModernSplashScreen(QSplashScreen):
    """Modern animated splash screen with logo"""

    def __init__(self, logo_path=None):
        # Create a pixmap for the splash screen
        pixmap = QPixmap(600, 400)
        pixmap.fill(Qt.GlobalColor.transparent)

        super().__init__(pixmap, Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)

        self.logo_path = logo_path
        self.opacity_value = 0.0
        self.progress = 0

        # Set window attributes
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # Animation for fade in
        self.fade_timer = QTimer()
        self.fade_timer.timeout.connect(self.fade_in)
        self.fade_timer.start(20)

        # Auto close timer
        self.close_timer = QTimer()
        self.close_timer.timeout.connect(self.start_fade_out)
        self.close_timer.start(5000)  # Show for 5 seconds

    def fade_in(self):
        """Fade in animation"""
        if self.opacity_value < 1.0:
            self.opacity_value += 0.05
            self.setWindowOpacity(self.opacity_value)
            self.repaint()
        else:
            self.fade_timer.stop()

    def start_fade_out(self):
        """Start fade out animation"""
        self.close_timer.stop()
        self.fade_out_timer = QTimer()
        self.fade_out_timer.timeout.connect(self.fade_out)
        self.fade_out_timer.start(20)

    def fade_out(self):
        """Fade out animation"""
        if self.opacity_value > 0.0:
            self.opacity_value -= 0.05
            self.setWindowOpacity(self.opacity_value)
        else:
            self.fade_out_timer.stop()
            self.close()

    def drawContents(self, painter):
        """Draw the splash screen contents"""
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Background with gradient
        gradient = QLinearGradient(0, 0, 600, 400)
        gradient.setColorAt(0, QColor(10, 10, 15))
        gradient.setColorAt(0.5, QColor(15, 20, 30))
        gradient.setColorAt(1, QColor(10, 10, 15))
        painter.fillRect(0, 0, 600, 400, gradient)

        # Draw border
        painter.setPen(QPen(QColor(33, 150, 243), 2))
        painter.drawRoundedRect(2, 2, 596, 396, 15, 15)

        # Draw logo if available
        if self.logo_path and Path(self.logo_path).exists():
            logo_pixmap = QPixmap(str(self.logo_path))
            if not logo_pixmap.isNull():
                # Scale logo to fit
                scaled_logo = logo_pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio,
                                                Qt.TransformationMode.SmoothTransformation)
                # Center logo
                logo_x = (600 - scaled_logo.width()) // 2
                logo_y = 80
                painter.drawPixmap(logo_x, logo_y, scaled_logo)
        else:
            # Draw triangle logo with cyan lines
            self.draw_triangle_logo(painter)

        # Draw title
        title_font = QFont("Segoe UI", 32, QFont.Weight.Bold)
        painter.setFont(title_font)
        painter.setPen(QColor(33, 203, 243))
        painter.drawText(0, 300, 600, 50, Qt.AlignmentFlag.AlignCenter, "Documentation Slayer")

        # Draw subtitle
        subtitle_font = QFont("Segoe UI", 12)
        painter.setFont(subtitle_font)
        painter.setPen(QColor(180, 180, 180))
        painter.drawText(0, 340, 600, 30, Qt.AlignmentFlag.AlignCenter, "Automate Your Documentation – and Much More!")

        # Draw version
        version_font = QFont("Segoe UI", 9)
        painter.setFont(version_font)
        painter.setPen(QColor(120, 120, 120))
        painter.drawText(0, 370, 600, 20, Qt.AlignmentFlag.AlignCenter, "v3.3.0")

    def draw_triangle_logo(self, painter):
        """Draw the triangle logo with cyan lines"""
        center_x = 300
        center_y = 160

        # Draw outer triangle
        painter.setPen(QPen(QColor(100, 100, 110), 3))
        painter.setBrush(QColor(20, 25, 35))

        points = [
            (center_x, center_y - 80),      # Top
            (center_x - 90, center_y + 60),  # Bottom left
            (center_x + 90, center_y + 60)   # Bottom right
        ]

        from PyQt6.QtGui import QPolygon
        from PyQt6.QtCore import QPoint

        triangle = QPolygon([QPoint(int(x), int(y)) for x, y in points])
        painter.drawPolygon(triangle)

        # Draw cyan lines inside triangle
        painter.setPen(QPen(QColor(33, 203, 243), 2))

        # Horizontal lines with circles
        line_spacing = 15
        start_y = center_y - 40

        for i in range(6):
            y = start_y + (i * line_spacing)
            # Left line
            painter.drawLine(center_x - 50, y, center_x - 10, y)
            # Right line
            painter.drawLine(center_x + 10, y, center_x + 50, y)

            # Circles
            painter.setBrush(QColor(33, 203, 243))
            painter.drawEllipse(center_x - 55, y - 4, 8, 8)
            painter.drawEllipse(center_x + 47, y - 4, 8, 8)

        # Draw diagonal slash
        gradient_pen = QPen(QColor(100, 240, 255), 4)
        painter.setPen(gradient_pen)
        painter.drawLine(center_x + 30, center_y - 60, center_x - 30, center_y + 40)


def show_splash_screen(logo_path=None):
    """Show the splash screen"""
    # Create application if needed
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    # Create and show splash
    splash = ModernSplashScreen(logo_path)
    splash.show()

    # Process events to show splash
    app.processEvents()

    return splash


if __name__ == "__main__":
    """Test the splash screen"""
    app = QApplication(sys.argv)

    # Try to find logo
    logo_path = Path(__file__).parent / "DocSlayerLogo.ico"

    splash = ModernSplashScreen(str(logo_path) if logo_path.exists() else None)
    splash.show()

    sys.exit(app.exec())
