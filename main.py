import os
import sys
import requests
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QFileDialog, QComboBox, QColorDialog, QTextEdit, QFontComboBox, QMainWindow, QAction, QMessageBox, QHBoxLayout
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

class ScriptureSlidesApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.api_token = os.getenv("ESV_API_TOKEN")
        self.initUI()
        

    def initUI(self):
        self.setWindowTitle("Scripture Slides")

        # Set the main widget
        main_widget = QWidget(self)
        main_layout = QVBoxLayout(main_widget)

        # Menu bar
        menubar = self.menuBar()

        # File menu (Presets, Help, Version)
        fileMenu = menubar.addMenu('File')

        # Presets action
        presets_action = QAction('Presets', self)
        presets_action.triggered.connect(self.show_presets)
        fileMenu.addAction(presets_action)

        # Help action
        help_action = QAction('Help', self)
        help_action.triggered.connect(self.show_help)
        fileMenu.addAction(help_action)

        # Version action
        version_action = QAction('Version', self)
        version_action.triggered.connect(self.show_version)
        fileMenu.addAction(version_action)

        # Bible search section
        search_layout = QHBoxLayout()

        self.book_input = QLineEdit(self)
        self.book_input.setPlaceholderText("Enter Book (e.g., Genesis)")
        search_layout.addWidget(self.book_input)

        self.chapter_input = QLineEdit(self)
        self.chapter_input.setPlaceholderText("Enter Chapter (e.g., 1)")
        search_layout.addWidget(self.chapter_input)

        self.verse_input = QLineEdit(self)
        self.verse_input.setPlaceholderText("Enter Verse (e.g., 1)")
        search_layout.addWidget(self.verse_input)

        self.search_button = QPushButton("Search Bible Verse", self)
        self.search_button.clicked.connect(self.search_bible_verse)
        search_layout.addWidget(self.search_button)

        main_layout.addLayout(search_layout)

        # Result of the Bible verse search
        self.search_result = QTextEdit(self)
        self.search_result.setPlaceholderText("Search result will appear here")
        main_layout.addWidget(self.search_result)

        # Title input
        self.title_input = QLineEdit(self)
        self.title_input.setPlaceholderText("Enter Slide Title (Optional)")
        main_layout.addWidget(self.title_input)

        # Font family and size selection for title and verse
        self.font_family_title = QFontComboBox(self)
        main_layout.addWidget(self.font_family_title)

        self.font_size_title = QComboBox(self)
        self.font_size_title.addItems([str(i) for i in range(24, 73, 2)])  # Sizes 24 to 72
        main_layout.addWidget(self.font_size_title)

        self.font_family_verse = QFontComboBox(self)
        main_layout.addWidget(self.font_family_verse)

        self.font_size_verse = QComboBox(self)
        self.font_size_verse.addItems([str(i) for i in range(16, 49, 2)])  # Sizes 16 to 48
        main_layout.addWidget(self.font_size_verse)

        # Button to change title color
        self.title_color_button = QPushButton('Change Title Color', self)
        self.title_color_button.clicked.connect(self.pick_title_color)
        main_layout.addWidget(self.title_color_button)

        # Button to change verse color
        self.verse_color_button = QPushButton('Change Verse Color', self)
        self.verse_color_button.clicked.connect(self.pick_verse_color)
        main_layout.addWidget(self.verse_color_button)

        # Label to display selected background
        self.label = QLabel("No backgrounds selected", self)
        main_layout.addWidget(self.label)

        # Button to upload background images
        self.bg_button = QPushButton('Upload Background Images', self)
        self.bg_button.clicked.connect(self.upload_backgrounds)
        main_layout.addWidget(self.bg_button)

        # LineEdit for specifying the file name
        self.file_name_input = QLineEdit(self)
        self.file_name_input.setPlaceholderText("Enter file name for the PowerPoint (without extension)")
        main_layout.addWidget(self.file_name_input)

        # Button to create the PowerPoint
        self.ppt_button = QPushButton('Create PowerPoint', self)
        self.ppt_button.clicked.connect(self.create_presentation)
        main_layout.addWidget(self.ppt_button)

        self.setCentralWidget(main_widget)

        # Color placeholders
        self.title_color = None
        self.verse_color = None

        # List to hold multiple background images
        self.background_images = []

    def search_bible_verse(self):
        book = self.book_input.text().strip()
        chapter = self.chapter_input.text().strip()
        verse = self.verse_input.text().strip()

        if book and chapter and verse:
            verse_text = self.fetch_bible_verse(book, chapter, verse)
            if verse_text:
                self.search_result.setText(f"{book} {chapter}:{verse} - {verse_text}")
            else:
                self.search_result.setText(f"Verse not found: {book} {chapter}:{verse}")
        else:
            self.search_result.setText("Please enter valid book, chapter, and verse.")

    def fetch_bible_verse(self, book, chapter, verse):
        """
        Fetches a Bible verse from the ESV API.
        """
        url = f"https://api.esv.org/v3/passage/text/"
        params = {
            "q": f"{book} {chapter}:{verse}",
            "include-passage-references": "true",
            "include-verse-numbers": "false",
            "include-footnotes": "false",
            "include-headings": "false"
        }
        headers = {
            "Authorization": f"Token {self.api_token}"
        }

        response = requests.get(url, headers=headers, params=params)

        if response.status_code == 200:
            data = response.json()
            return data["passages"][0].strip() if data["passages"] else None
        else:
            QMessageBox.warning(self, "Error", f"Failed to fetch verse: {response.status_code}")
            return None

    def upload_backgrounds(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "Select Background Images", "", "Image Files (*.png *.jpg *.jpeg)", options=options)
        if files:
            self.background_images = files
            self.label.setText(f"Selected {len(files)} backgrounds")

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

        # Get the verse from search result
        verse = self.search_result.toPlainText()

        slide = prs.slides.add_slide(slide_layout)

        if self.background_images:
            self.set_slide_background(slide, self.background_images[0], prs)

        title = self.title_input.text()

        if title:
            self.add_text(slide, title, Inches(1), Inches(0.5), int(self.font_size_title.currentText()), self.font_family_title.currentText(), self.title_color, prs.slide_width, prs.slide_height, is_title=True)

        if verse:
            self.add_text(slide, verse, Inches(1), Inches(2), int(self.font_size_verse.currentText()), self.font_family_verse.currentText(), self.verse_color, prs.slide_width, prs.slide_height, is_title=False)

        # Get the file name from the input, or use a default name if none is provided
        file_name = self.file_name_input.text() or 'scripture_slides'
        prs.save(f'{file_name}.pptx')
        self.label.setText(f"Presentation '{file_name}.pptx' created successfully!")

    def set_slide_background(self, slide, img_path, prs):
        # Get slide dimensions
        prs_width = prs.slide_width
        prs_height = prs.slide_height
        slide.shapes.add_picture(img_path, 0, 0, prs_width, prs_height)

    def add_text(self, slide, text, left, top, font_size, font_family, color, width, height, is_title):
        # Create a textbox with appropriate width/height
        textbox = slide.shapes.add_textbox(left, top, width - left * 2, height - top * 2)
        text_frame = textbox.text_frame
        text_frame.word_wrap = True  # Enable word wrap

        # Add the paragraph
        p = text_frame.add_paragraph()
        p.text = text
        p.font.size = Pt(font_size)
        p.font.name = font_family  # Set the font family

        if color:
            p.font.color.rgb = RGBColor(color.red(), color.green(), color.blue())

        # Ensure the text doesn't overlap with the screen bounds
        if not is_title:
            max_height = height - Inches(2.5)  # Limit the text box height for verses
            if textbox.height > max_height:
                textbox.height = max_height  # Adjust the height to prevent overflow

    # Presets placeholder (add actual functionality as needed)
    def show_presets(self):
        QMessageBox.information(self, 'Presets', 'This is where you can add or select presets for slides.')

    # Help placeholder
    def show_help(self):
        QMessageBox.information(self, 'Help', 'Instructions:\n1. Search for a Bible verse\n2. Customize title, fonts, and background.\n3. Click Create PowerPoint.')

    # Version information
    def show_version(self):
        QMessageBox.information(self, 'Version', 'Scripture Slides v1.0')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ScriptureSlidesApp()
    ex.show()
    sys.exit(app.exec_())
