from PyQt5.QtCore import QVariantAnimation
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout

class AnimatedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        # Define the default and hover colors
        self.default_bg_color = QColor(44, 44, 44)  # #2C2C2C background color
        self.hover_bg_color = QColor(52, 152, 219)  # #3498db background color
        self.border_color = QColor(52, 152, 219)    # Border color on hover
        self.setStyleSheet(f"""
            QPushButton {{
                border: none;
                color: white;
                background-color: {self.default_bg_color.name()};
                border-radius: 20px;
                padding: 10px;
            }}
        """)

        # Setting up the hover animation for background color
        self.bg_animation = QVariantAnimation(self)
        self.bg_animation.setDuration(150)  # Duration in milliseconds
        self.bg_animation.valueChanged.connect(self.on_bg_value_changed)

    def enterEvent(self, event):
        # Animate background color change on hover
        self.animate_bg_color(self.default_bg_color, self.hover_bg_color)
        super().enterEvent(event)

    def leaveEvent(self, event):
        # Animate back to default background color on mouse leave
        self.animate_bg_color(self.hover_bg_color, self.default_bg_color)
        super().leaveEvent(event)

    def animate_bg_color(self, start_color, end_color):
        self.bg_animation.setStartValue(start_color)
        self.bg_animation.setEndValue(end_color)
        self.bg_animation.start()

    def on_bg_value_changed(self, value):
        color = value.name()  # Get the color as a hex string
        self.setStyleSheet(f"""
            QPushButton {{
                border: none;
                color: white;
                background-color: {color};
                border-radius: 20px;
                padding: 10px;
            }}
        """)

# Main application
app = QApplication([])
window = QWidget()
layout = QVBoxLayout()

# Create the animated button with the desired text
button = AnimatedButton('Create Presentation')
layout.addWidget(button)
window.setLayout(layout)

window.show()
app.exec_()
