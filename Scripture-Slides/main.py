import sys
import os
import warnings
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QFontComboBox, QPushButton,
                             QListWidget, QColorDialog, QGraphicsScene, QGraphicsTextItem, QGraphicsView, QTextEdit)
from PyQt5.QtGui import QFont, QColor, QPixmap, QBrush
from PIL import ImageGrab
from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

warnings.simplefilter("ignore", DeprecationWarning)

from PIL import Image, ImageDraw, ImageFont
import os

class ScriptureSlides(QMainWindow):
    def __init__(self):
        super(ScriptureSlides, self).__init__()
        loadUi("./Scripture-Slides/ScriptureSlides.ui", self)
        
        # Initialize presentation object
        self.prs = Presentation()

        self.prs.slide_width = Inches(20)
        self.prs.slide_height = Inches(11.25) 

        self.slide_count = 0
        
        # Dictionary to keep track of slides and their preview images
        self.slide_previews = {}

        # Initialize QGraphicsScene for the preview
        self.scene = QGraphicsScene(self.graphicsView)
        self.graphicsView.setScene(self.scene)

        self.setStyleSheet("""
        QComboBox#fontComboBox {
            font-family: 'Arial', sans-serif;
            font-size: 14px;
            color: black;
            background-color: white;
            border: 1px solid #ccc;
            padding: 5px;
            border-radius: 8px;
        }
        QComboBox#fontComboBox::drop-down {
            background-color: transparent;
            width: 30px;
        }
        QComboBox#fontComboBox::drop-down:hover {
            background-color: #EFEFEF;
            border-radius: 8px;
        }
        QComboBox#fontComboBox::down-arrow {
            image: url(C:/Users/markt/Scripture-Slides/Scripture-Slides/assets/arrowDown.png);
            width: 18px;
            height: 18px;
        }
        QComboBox#fontComboBox QAbstractItemView {
            font-family: 'Inter 18pt', sans-serif;
            font-size: 5px;
            color: black;
            background-color: white;
            border: none;
        }
        """)

        # Connect buttons to their respective functions
        self.addSlideBtn.clicked.connect(self.add_slide)
        self.addBackgroundImageBtn.clicked.connect(self.add_background_image)
        self.createPresentationBtn.clicked.connect(self.create_presentation)
        self.Slide_List_Widget.itemSelectionChanged.connect(self.display_slide_in_graphics_view)

        self.current_slide = None

    def add_slide(self):
        """Add a new blank slide and save its preview."""
        slide_layout = self.prs.slide_layouts[6]  # Layout for a blank slide
        self.current_slide = self.prs.slides.add_slide(slide_layout)
        self.slide_count += 1

        slide_item = f"Slide {self.slide_count}"
        self.Slide_List_Widget.addItem(slide_item)
        
        # Capture preview and save it (using placeholder data for now)
        self.save_slide_preview()

        print(f"{slide_item} added!")

    def add_background_image(self):
        """Add background image to the selected slide and save the preview."""
        selected_items = self.Slide_List_Widget.selectedItems()
        if not selected_items:
            print("No slide selected. Add a slide first.")
            return
        
        # Get the selected slide index
        selected_index = self.Slide_List_Widget.currentRow()
        self.current_slide = self.prs.slides[selected_index]  # Get the actual slide from the selected index
        
        image_path, _ = QFileDialog.getOpenFileName(self, 'Open Background Image', '', 'Image files (*.jpg *.png)')
        if image_path:
            slide_width = self.prs.slide_width
            slide_height = self.prs.slide_height
            self.current_slide.shapes.add_picture(image_path, 0, 0, width=slide_width, height=slide_height)
            print(f"Background image {image_path} added to the selected slide.")
            
            # Update preview
            self.save_slide_preview(image_path)
            self.display_image_in_graphics_view(image_path)

    def save_slide_preview(self, image_path=None):
        """Simulate the current slide preview and save it as an image."""
        # Specify the path to save the slide previews
        preview_folder = "./slide_previews/"
        os.makedirs(preview_folder, exist_ok=True)
        
        # Generate a file name based on the slide number
        slide_number = self.Slide_List_Widget.currentRow() + 1  # Use the selected slide index +1 as the slide number
        preview_image_path = f"{preview_folder}slide_{slide_number}.png"
        
        # Create an image with the slide's content using Pillow (simulating the slide)
        slide_width = 1280  # Example width (adjust as needed)
        slide_height = 720  # Example height (adjust as needed)

        # Create a blank canvas
        image = Image.new("RGB", (slide_width, slide_height), "white")
        draw = ImageDraw.Draw(image)

        # If a background image is present, add it to the canvas
        if image_path:
            background = Image.open(image_path)
            background = background.resize((slide_width, slide_height))
            image.paste(background, (0, 0))

        # Simulate adding text to the slide (replace this with actual text logic)
        text = f"Slide {slide_number}"
        font = ImageFont.load_default()
        text_color = (0, 0, 0)  # Black text
        draw.text((50, 50), text, fill=text_color, font=font)

        # Save the image as a preview
        image.save(preview_image_path)
        print(f"Preview for Slide {slide_number} saved at {preview_image_path}")
        
        # Store the preview path
        self.slide_previews[f"Slide {slide_number}"] = preview_image_path

    def display_image_in_graphics_view(self, image_path):
        """Display an image in the QGraphicsView."""
        pixmap = QPixmap(image_path)
        self.scene.clear()  # Clear any existing items in the scene
        self.scene.addPixmap(pixmap)  # Add the new image
        self.graphicsView.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)

    def display_slide_in_graphics_view(self):
        """Display the selected slide in QGraphicsView."""
        selected_items = self.Slide_List_Widget.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text()  # Get the selected slide (e.g., 'Slide 1')
            
            # Get the preview image for the selected slide
            image_path = self.slide_previews.get(selected_item)
            if image_path:
                self.display_image_in_graphics_view(image_path)
            else:
                print(f"No preview available for {selected_item}")

    def create_presentation(self):
        """Save the PowerPoint presentation."""
        save_path, _ = QFileDialog.getSaveFileName(self, 'Save Presentation', '', 'PowerPoint files (*.pptx)')
        if save_path:
            self.prs.save(save_path)
            print(f"Presentation saved at {save_path}")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = ScriptureSlides()
    ui.show()
    sys.exit(app.exec_())
