import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QFileDialog, QComboBox, QColorDialog, QTextEdit
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

class ScriptureSlidesApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Scripture Slides")

        layout = QVBoxLayout()

        # Title input
        self.title_input = QLineEdit(self)
        self.title_input.setPlaceholderText("Enter Slide Title (Optional)")
        layout.addWidget(self.title_input)

        # Verse input (now a multi-line text edit for multiple verses)
        self.verse_input = QTextEdit(self)
        self.verse_input.setPlaceholderText("Enter Bible Verses (each verse on a new line)")
        layout.addWidget(self.verse_input)

        # Font size selection for title
        self.font_size_title = QComboBox(self)
        self.font_size_title.addItems([str(i) for i in range(24, 73, 2)])  # Sizes 24 to 72
        layout.addWidget(self.font_size_title)

        # Font size selection for verse
        self.font_size_verse = QComboBox(self)
        self.font_size_verse.addItems([str(i) for i in range(16, 49, 2)])  # Sizes 16 to 48
        layout.addWidget(self.font_size_verse)

        # Button to change title color
        self.title_color_button = QPushButton('Change Title Color', self)
        self.title_color_button.clicked.connect(self.pick_title_color)
        layout.addWidget(self.title_color_button)

        # Button to change verse color
        self.verse_color_button = QPushButton('Change Verse Color', self)
        self.verse_color_button.clicked.connect(self.pick_verse_color)
        layout.addWidget(self.verse_color_button)

        # Label to display selected background
        self.label = QLabel("No background selected", self)
        layout.addWidget(self.label)

        # Button to upload background image
        self.bg_button = QPushButton('Upload Background Image', self)
        self.bg_button.clicked.connect(self.upload_background)
        layout.addWidget(self.bg_button)

        # Button to create the PowerPoint
        self.ppt_button = QPushButton('Create PowerPoint', self)
        self.ppt_button.clicked.connect(self.create_presentation)
        layout.addWidget(self.ppt_button)

        self.setLayout(layout)

        # Color placeholders
        self.title_color = None
        self.verse_color = None

    def upload_background(self):
        options = QFileDialog.Options()
        file, _ = QFileDialog.getOpenFileName(self, "Select Background Image", "", "Image Files (*.png *.jpg *.jpeg)", options=options)
        if file:
            self.label.setText(f"Selected Background: {file}")
            self.background_image = file

    def pick_title_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.title_color = color

    def pick_verse_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.verse_color = color

    def create_presentation(self):
        prs = Presentation()
        prs.slide_width = 18288000  # Set width to 1920 pixels
        prs.slide_height = 10281600  # Set height to 1080 pixels

        slide_layout = prs.slide_layouts[5]  # blank slide

        # Get user input for verses (split by new lines)
        verses = self.verse_input.toPlainText().splitlines()

        for verse in verses:
            slide = prs.slides.add_slide(slide_layout)

            if hasattr(self, 'background_image'):
                # Apply background image
                self.set_slide_background(slide, self.background_image, prs)

            # Add title and verse text boxes for each slide
            title = self.title_input.text()

            if title:
                self.add_text(slide, title, 1000, 500, int(self.font_size_title.currentText()), self.title_color, prs.slide_width, prs.slide_height)

            if verse:
                self.add_text(slide, verse, 1000, 2000, int(self.font_size_verse.currentText()), self.verse_color, prs.slide_width, prs.slide_height)

        prs.save('scripture_slides_from_gui.pptx')
        self.label.setText("Presentation created successfully!")

    def set_slide_background(self, slide, img_path, prs):
        # Get slide dimensions
        prs_width = prs.slide_width
        prs_height = prs.slide_height
        slide.shapes.add_picture(img_path, 0, 0, prs_width, prs_height)

    def add_text(self, slide, text, left, top, font_size, color, width, height):
        textbox = slide.shapes.add_textbox(left, top, width, height)
        text_frame = textbox.text_frame
        p = text_frame.add_paragraph()
        p.text = text
        p.font.size = Pt(font_size)
        if color:
            p.font.color.rgb = RGBColor(color.red(), color.green(), color.blue())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ScriptureSlidesApp()
    ex.show()
    sys.exit(app.exec_())
