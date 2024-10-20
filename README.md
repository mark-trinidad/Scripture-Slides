# Scripture Slides

## Overview
This project is an **Automatic PowerPoint Maker** that allows users to:
- Create PowerPoint presentations.
- Change slide backgrounds using custom images.
- Add Bible verses to slides.
- Edit and customize slide titles and text.
- Automatically format and export the final presentation as `.pptx`.

## Features
- **Create PowerPoint Presentations**: Automatically generate slides with customizable content.
- **Change Backgrounds**: Upload or select custom images for the slide backgrounds.
- **Bible Verses Integration**: Add Bible verses, with support for different translations and searchable by book, chapter, and verse.
- **Title and Text Customization**: Edit titles and text with customizable fonts, sizes, and styles.
- **Preview and Export**: Real-time slide preview and export as `.pptx` or `.pdf`.
- **Optional Cloud Support**: Save to or open from cloud storage (Google Drive).

## Tech Stack
### Language
- **Python**

## Key Libraries
- **[python-pptx](https://python-pptx.readthedocs.io/en/latest/)**: For creating and manipulating PowerPoint presentations.
- **[Pillow (PIL)](https://pillow.readthedocs.io/)**: For image manipulation (cropping, resizing, etc.).
- **[PyQt5](https://riverbankcomputing.com/software/pyqt/intro)**: To create a desktop GUI.
- **[SQLite](https://www.sqlite.org/index.html)**: (Optional) For storing Bible verses locally.
- **[PyInstaller](https://pyinstaller.org/en/stable/)**: To package the app for macOS and Windows as standalone executables.

## Installation

### 1. Clone the repository:
  ```bash
  git clone https://github.com/mark-trinidad/Scripture-Slides.git
  cd Scripture-Slides
  ```
### 2. Clone the repository:
  ```bash
  pip install -r requirements.txt
  ```
### 3. Clone the repository:
  ```bash
  python main.py
  ```
## Packaging the Application
### To create a standalone executable (e.g., .app for macOS):
  ```bash
  pyinstaller --onefile --windowed main.py
  ```
