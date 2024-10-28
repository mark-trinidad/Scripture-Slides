import sys
import os
import time
import warnings
from PyQt5.QtWidgets import (QMainWindow, QApplication, QFileDialog, QFontComboBox, QPushButton,
                             QListWidget, QColorDialog, QGraphicsScene, QGraphicsTextItem, QGraphicsView, QTextEdit, QMenu)
from PyQt5.QtGui import QFont, QColor, QPixmap, QBrush, QIcon
from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PyQt5.QtWidgets import QGraphicsPixmapItem, QDialog, QLabel, QVBoxLayout, QPushButton, QSpinBox # Added for DraggableImageItem
import win32com.client as win32
import msvcrt

warnings.simplefilter("ignore", DeprecationWarning)

from PIL import Image, ImageDraw, ImageFont

class DraggableTextItem(QGraphicsTextItem):
    def __init__(self, text):
        super().__init__(text)
        self.setFlags(QGraphicsTextItem.ItemIsMovable | QGraphicsTextItem.ItemIsSelectable | QGraphicsTextItem.ItemIsFocusable)
        self.setTextInteractionFlags(Qt.TextEditorInteraction)  # Allow editing on double-click


class DraggableImageItem(QGraphicsPixmapItem):  # QGraphicsPixmapItem imported here
    def __init__(self, pixmap):
        super().__init__(pixmap)
        self.setFlags(QGraphicsPixmapItem.ItemIsMovable | QGraphicsPixmapItem.ItemIsSelectable)

class VerseRepeatWindow(QDialog):
    def __init__(self, parent=None):
        super(VerseRepeatWindow, self).__init__(parent)
        self.setWindowTitle("Add Verse")

        # Load the UI file for the VerseRepeatWindow (optional if you don't have one)
        loadUi("Scripture-Slides/AddVerse.ui", self)


class ScriptureSlides(QMainWindow):
    def __init__(self):
        super(ScriptureSlides, self).__init__()
        self.setWindowFlags(Qt.Window | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        loadUi("Scripture-Slides/ScriptureSlides.ui", self)

        self.setStyleSheet(""" 
                QComboBox#fontComboBox {
            font-family: 'Inter 18pt', sans-serif;
            font-size: 10pt; /* Change font size to points */
            color: black;
            background-color: white;
            border: 1px solid #ccc;
            padding: 5px;
            border-radius: 8px;
        }

        QComboBox#fontComboBox::drop-down {
            background-color: transparent; /* Change drop-down background */
            width: 30px; /* Adjust drop-down width */
        }

        QComboBox#fontComboBox::drop-down:hover {
        background-color: #EFEFEF;
            border-radius: 8px;
        }

        QComboBox#fontComboBox::down-arrow {
            image: url(C:/Users/markt/Scripture-Slides/Scripture-Slides/assets/arrowDown.png); /* Path to your custom arrow image */
            width: 18px; /* Adjust width if needed */
            height: 18px; /* Adjust height if needed */
        }

        QComboBox#fontComboBox QAbstractItemView {
            font-family: 'Inter 18pt', sans-serif;
            font-size: 10pt; /* Set point size for list items */
            color: black;
            background-color: white;
            border: none;
        }

        """)

        # Initialize presentation object and QGraphicsScene only once
        self.prs = Presentation()
        self.prs.slide_width = Inches(20)
        self.prs.slide_height = Inches(11.25)
        self.slide_count = 0
        self.slide_previews = {}

        # Initialize QGraphicsScene for preview
        self.scene = QGraphicsScene(self.graphicsView)
        self.graphicsView.setScene(self.scene)

        # Connect buttons to functions
        self.addSlideBtn.clicked.connect(self.add_slide)
        self.addBackgroundImageBtn.clicked.connect(self.add_background_image)
        self.createPresentationBtn.clicked.connect(self.create_presentation)
        self.Slide_List_Widget.itemSelectionChanged.connect(self.display_slide_in_graphics_view)
        self.addTextBtn.clicked.connect(self.add_text_item) 
        self.Slide_List_Widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.Slide_List_Widget.customContextMenuRequested.connect(self.open_context_menu)
        self.current_slide = None
        self.VerseRepeatBtn.clicked.connect(self.open_verse_repeat_window)
        self.verse_repeat_window = None
        self.SlideShowBtn.clicked.connect(self.start_slideshow_without_pptx)
        

    def start_slideshow_without_pptx(self):
        """Generate a new PowerPoint presentation from the current slides in the app and start the slideshow with keyboard control."""
        try:
            slide_width = Inches(20).pt
            slide_height = Inches(11.25).pt

            powerpoint = win32.Dispatch("PowerPoint.Application")
            powerpoint.Visible = True
            presentation = powerpoint.Presentations.Add()

            # Add slides from the graphics view items
            for slide_num in range(self.Slide_List_Widget.count()):
                slide_layout = 1  # Use title + content layout or change to blank if needed
                slide = presentation.Slides.Add(slide_num + 1, slide_layout)
                self.Slide_List_Widget.setCurrentRow(slide_num)
                rect = self.graphicsView.viewport().rect()
                screen = QApplication.primaryScreen()
                pixmap = screen.grabWindow(self.graphicsView.winId(), rect.x(), rect.y(), rect.width(), rect.height())
                temp_image_path = os.path.abspath(f"./temp_slide_{slide_num + 1}.png")
                pixmap.save(temp_image_path)

                if os.path.exists(temp_image_path):
                    # Open image to get dimensions for scaling, and ensure it is closed after use
                    with Image.open(temp_image_path) as img:
                        img_width, img_height = img.size
                        img_ratio = img_width / img_height
                        slide_ratio = slide_width / slide_height

                        # Scale and center the image to fit the slide
                        if img_ratio > slide_ratio:
                            new_width = slide_width
                            new_height = slide_width / img_ratio
                        else:
                            new_height = slide_height
                            new_width = slide_height * img_ratio
                        
                        left = (slide_width - new_width) / 2
                        top = (slide_height - new_height) / 2
                    
                    # Insert scaled image into the slide
                    slide.Shapes.AddPicture(temp_image_path, 0, 1, left, top, new_width, new_height)
                    
                    # Remove the image after it has been added to the slide
                    os.remove(temp_image_path)

            # Start the slideshow
            slideshow = presentation.SlideShowSettings
            slideshow.StartingSlide = 1
            slideshow.EndingSlide = presentation.Slides.Count
            slideshow.AdvanceMode = 1  # Manual advance
            slideshow.Run()

            # Control the slideshow with keyboard input
            slide_show_window = powerpoint.SlideShowWindows(1)
            running = True
            print("Slideshow started. Press 'n' for next, 'p' for previous, or 'q' to quit.")

            while running:
                if msvcrt.kbhit():
                    key = msvcrt.getch().decode('utf-8').lower()
                    if key == 'n':  # Next slide
                        slide_show_window.View.Next()
                    elif key == 'p':  # Previous slide
                        slide_show_window.View.Previous()
                    elif key == 'q':  # Quit slideshow
                        print("Exiting slideshow.")
                        running = False

            presentation.Close()
            powerpoint.Quit()

        except Exception as e:
            print(f"Error during slideshow: {e}")
        finally:
            # Ensure PowerPoint closes after completion
            presentation.Close()
            powerpoint.Quit()

    def open_verse_repeat_window(self):
        """Open the Verse Repeat window."""
        # Check if window already exists to prevent multiple instances
        if self.verse_repeat_window is None or not self.verse_repeat_window.isVisible():
            self.verse_repeat_window = VerseRepeatWindow(self)
            self.verse_repeat_window.finished.connect(self.on_verse_repeat_window_closed)
            self.verse_repeat_window.show()
        else:
            self.verse_repeat_window.raise_()  # Bring the existing window to the front

    def on_verse_repeat_window_closed(self):
        """Handle the closing of the Verse Repeat window."""
        print("Verse Repeat window closed.")
        self.verse_repeat_window = None  # Reset the window instance when closed

    def add_text_item(self):
        """Add a new draggable text item to the graphics view."""
        text_item = DraggableTextItem("Editable Text")
        text_item.setFont(QFont("Arial", 20))
        self.scene.addItem(text_item)

    def add_slide(self):
        """Add a new blank slide and save its preview."""
        slide_layout = self.prs.slide_layouts[6]  # Layout for a blank slide
        self.current_slide = self.prs.slides.add_slide(slide_layout)
        self.slide_count += 1

        slide_item = f"Slide {self.slide_count}"
        self.Slide_List_Widget.addItem(slide_item)
        
        self.save_slide_preview()  # Capture preview (using placeholder data)
        print(f"{slide_item} added!")

    def add_background_image(self):
        """Add a draggable background image to the QGraphicsView and PowerPoint slide."""
        selected_items = self.Slide_List_Widget.selectedItems()
        if not selected_items:
            print("No slide selected. Add a slide first.")
            return

        selected_index = self.Slide_List_Widget.currentRow()
        self.current_slide = self.prs.slides[selected_index]  # Get the actual slide from the selected index

        # Open file dialog to select an image
        image_path, _ = QFileDialog.getOpenFileName(self, 'Open Background Image', '', 'Image files (*.jpg *.png)')
        if image_path:
            # Add image to QGraphicsView for a preview
            pixmap = QPixmap(image_path)
            image_item = DraggableImageItem(pixmap)
            
            # Clear any previous background image to avoid stacking images
            self.scene.clear()
            
            # Add the new image as a draggable item
            self.scene.addItem(image_item)
            
            # Adjust the view to fit the new image
            self.graphicsView.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)
            print(f"Draggable background image {image_path} added to the graphics view.")

            # Update preview with new image path
            self.save_slide_preview(image_path)

            # Set the background image in PowerPoint slide
            left, top, width, height = 0, 0, self.prs.slide_width, self.prs.slide_height
            self.current_slide.shapes.add_picture(image_path, left, top, width, height)
            print(f"Background image {image_path} added to slide in PowerPoint presentation.")


    def save_slide_preview(self, image_path=None):
        """Simulate the current slide preview and save it as an image."""
        preview_folder = "./slide_previews/"
        os.makedirs(preview_folder, exist_ok=True)
        
        slide_number = self.Slide_List_Widget.currentRow() + 1
        preview_image_path = f"{preview_folder}slide_{slide_number}.png"
        
        slide_width, slide_height = 1280, 720
        image = Image.new("RGB", (slide_width, slide_height), "white")
        draw = ImageDraw.Draw(image)

        if image_path:
            background = Image.open(image_path)
            background = background.resize((slide_width, slide_height))
            image.paste(background, (0, 0))

        text = f"Slide {slide_number}"
        font = ImageFont.load_default()
        draw.text((50, 50), text, fill=(0, 0, 0), font=font)

        image.save(preview_image_path)
        print(f"Preview for Slide {slide_number} saved at {preview_image_path}")
        
        self.slide_previews[f"Slide {slide_number}"] = preview_image_path

    def display_image_in_graphics_view(self, image_path):
        """Display an image in the QGraphicsView."""
        pixmap = QPixmap(image_path)
        self.scene.clear()
        self.scene.addPixmap(pixmap)
        self.graphicsView.fitInView(self.scene.itemsBoundingRect(), Qt.KeepAspectRatio)

    def display_slide_in_graphics_view(self):
        """Display the selected slide in QGraphicsView."""
        selected_items = self.Slide_List_Widget.selectedItems()
        if selected_items:
            selected_item = selected_items[0].text()
            image_path = self.slide_previews.get(selected_item)
            if image_path:
                self.display_image_in_graphics_view(image_path)
            else:
                print(f"No preview available for {selected_item}")

    def open_context_menu(self, position):
        """Open a custom context menu for deleting slides."""
        context_menu = QMenu(self)
        delete_action = context_menu.addAction("Delete Slide")
        action = context_menu.exec_(self.Slide_List_Widget.mapToGlobal(position))
        
        if action == delete_action:
            self.delete_slide()

    def delete_slide(self):
        """Delete the selected slide from the list and presentation."""
        selected_row = self.Slide_List_Widget.currentRow()
        if selected_row >= 0:
            selected_item = self.Slide_List_Widget.takeItem(selected_row)
            slide_name = selected_item.text()
            
            preview_path = self.slide_previews.pop(slide_name, None)
            if preview_path and os.path.exists(preview_path):
                os.remove(preview_path)
                print(f"Deleted preview image: {preview_path}")

            new_prs = Presentation()
            new_prs.slide_width = self.prs.slide_width
            new_prs.slide_height = self.prs.slide_height
            
            for i, slide in enumerate(self.prs.slides):
                if i != selected_row:
                    new_slide = new_prs.slides.add_slide(slide.slide_layout)
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            new_shape = new_slide.shapes.add_shape(
                                shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height
                            )
                            new_shape.text = shape.text_frame.text
                        elif hasattr(shape, "image"):
                            image_path = shape.image.filename
                            if os.path.exists(image_path):
                                new_slide.shapes.add_picture(image_path, shape.left, shape.top)
                            else:
                                print(f"Warning: Image file {image_path} not found. Skipping this image.")
            
            self.prs = new_prs
            print(f"{slide_name} deleted.")

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
