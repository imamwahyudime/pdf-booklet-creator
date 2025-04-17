# PDF Booklet Creator

[![Version](https://img.shields.io/badge/Version-1.1-blue.svg)](https://github.com/imamwahyudime/pdf-booklet-creator/releases/tag/v1.1)
[![Release Date](https://img.shields.io/badge/Release-April%2017,%202025-brightgreen.svg)](https://github.com/imamwahyudime/pdf-booklet-creator/releases/tag/v1.1)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)

**PDF Booklet Creator** is a simple and user-friendly Python application built with Tkinter (ttk) and PyMuPDF (fitz) to convert a regular PDF file into a booklet format, optimized for printing on A4 landscape paper. It arranges the pages so that when printed double-sided and folded, they form a correctly ordered booklet.

## Features:

* **Converts to A4 Landscape Booklet:** Arranges pages for printing as an A4 landscape booklet.
* **Adjustable Margins:** Allows setting custom center and outer margins in millimeters for optimal layout.
* **Add Blank Pages:** Automatically adds blank pages if the total number of pages in the input PDF is not a multiple of 4, ensuring correct booklet folding.
* **User-Friendly GUI:** Intuitive graphical interface for easy file selection and option configuration.
* **Progress and Status Updates:** Provides real-time feedback on the conversion process.
* **About Window:** Displays program information and links to the author's profiles.

## How to Use:

1.  **Select Input PDF:** Click the "Browse..." button under "Input PDF" to choose the PDF file you want to convert.
2.  **Select Output Booklet PDF:** Click the "Browse..." button under "Output Booklet PDF (A4 Landscape)" to specify where to save the generated booklet PDF. A default name based on the input file is suggested.
3.  **Configure Options:**
    * **Center Margin (mm):** Enter the desired width of the central margin between the two halves of each A4 landscape page (in millimeters). Default is 10.0 mm.
    * **Outer Margin (mm):** Enter the desired width of the outer margins on the edges of the A4 landscape page (in millimeters). Default is 5.0 mm.
    * **Add blank pages if needed (for multiple of 4):** Check this box to automatically add blank pages to the end of the document if the total page count is not divisible by 4. This is recommended for proper booklet folding.
4.  **Create Booklet:** Click the "Create Booklet" button to start the conversion process. A progress bar and status updates will keep you informed.
5.  **Printing:** Once the booklet PDF is created, open it with your PDF viewer and print it **double-sided**, making sure to **flip on the LONG edge** of the paper.

## Requirements:

* Python 3.x
* PyMuPDF (`fitz`) library: Install using `pip install pymupdf`
* Tkinter (usually included with standard Python installations)

## Installation:

1.  **Install PyMuPDF:**
    ```bash
    pip install pymupdf
    ```
2.  **Download the script:** Download the `pdf_booklet_creator.py` file from this repository.
3.  **Run the application:** Execute the script from your terminal:
    ```bash
    python pdf_booklet_creator.py
    ```
