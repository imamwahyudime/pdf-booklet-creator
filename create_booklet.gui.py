import fitz  # PyMuPDF
# from PIL import Image # Keep if planning image support later
import tkinter as tk
from tkinter import ttk  # Themed widgets
from tkinter import filedialog
from tkinter import messagebox
import os
import math
from pathlib import Path
import io
import sys
import threading
import queue
import webbrowser # For About window links

# --- Constants ---
A4_WIDTH_PT, A4_HEIGHT_PT = 595.276, 841.890
A4_LANDSCAPE_WIDTH_PT, A4_LANDSCAPE_HEIGHT_PT = A4_HEIGHT_PT, A4_WIDTH_PT # Swapped for landscape
A5_WIDTH_PT, A5_HEIGHT_PT = 420.945, 595.276
MM_TO_PT = 2.83465 # Conversion factor from mm to points

# Program Info
PROGRAM_VERSION = "1.1"
RELEASE_DATE = "April 17, 2025"
AUTHOR_INFO = {
    "GitHub": "https://github.com/imamwahyudime",
    "LinkedIn": "https://linkedin.com/in/imam-wahyudi/"
}

# --- Helper Functions ---

def mm_to_points(mm):
    """Converts millimeters to points."""
    try:
        return float(mm) * MM_TO_PT
    except ValueError:
        return 0.0

# (get_pdf_page_count remains the same)
def get_pdf_page_count(input_path):
    """Gets the page count of a PDF file."""
    path = Path(input_path)
    if not path.is_file() or path.suffix.lower() != '.pdf':
        return 0, "Invalid PDF file path."

    doc = None
    try:
        doc = fitz.open(input_path)
        if doc.is_encrypted:
            return 0, "Error: PDF file is encrypted."
        count = doc.page_count
        doc.close()
        return count, None # Return count and no error
    except Exception as e:
        if doc:
            doc.close()
        return 0, f"Error opening PDF: {e}"

# (calculate_inset_rect remains the same)
def calculate_inset_rect(rect, margin):
    """Calculates an inset rectangle manually."""
    if not rect or rect.is_empty or rect.is_infinite:
        return fitz.Rect(margin, margin, margin+1, margin+1) # Default small rect if original is invalid
    return fitz.Rect(rect.x0 + margin, rect.y0 + margin, rect.x1 - margin, rect.y1 - margin)

def open_link(url):
    """Opens a URL in the default web browser."""
    try:
        webbrowser.open_new(url)
    except Exception as e:
        print(f"Could not open link {url}: {e}") # Log to console if fails

# --- Core Booklet Logic ---

def create_booklet(input_pdf_path, output_pdf_path, central_margin_mm, outer_margin_mm, add_blanks, status_callback=None, progress_callback=None):
    """
    Creates the booklet PDF from a single input PDF (Outputting A4 Landscape).
    Args: [Same as before]
    Returns: bool: True on success, False on failure.
    """
    def update_status(message):
        if status_callback:
            status_callback(message)

    def update_progress(value):
        if progress_callback:
            progress_callback(value)

    update_status("Starting booklet creation (Landscape Output)...")
    update_progress(0)

    central_margin_pt = mm_to_points(central_margin_mm)
    outer_margin_pt = mm_to_points(outer_margin_mm)

    # 1. Open Input PDF
    source_doc = None
    try:
        source_doc = fitz.open(input_pdf_path)
        if source_doc.is_encrypted:
            update_status("Error: Input PDF is encrypted.")
            return False
        total_pages = source_doc.page_count
        if total_pages == 0:
             update_status("Error: Input PDF has no pages.")
             source_doc.close()
             return False
        update_status(f"Input PDF loaded: {total_pages} pages found.")
    except Exception as e:
        update_status(f"Error opening input PDF: {e}")
        if source_doc: source_doc.close()
        return False

    # 2. Handle page count
    target_pages = total_pages
    blanks_added = 0
    if total_pages % 4 != 0:
        blanks_to_add = 4 - (total_pages % 4)
        if add_blanks:
            target_pages = total_pages + blanks_to_add
            blanks_added = blanks_to_add
            update_status(f"Adding {blanks_to_add} blank page(s) for a total of {target_pages}.")
        else:
            update_status(f"Warning: Total pages ({total_pages}) not a multiple of 4. Booklet may not fold correctly.")

    if target_pages == 0:
        update_status("Error: No pages to process.")
        source_doc.close()
        return False

    # 3. Create output document
    output_doc = fitz.open() # New empty PDF

    # 4. Calculate A5 layout areas on A4 *Landscape* sheet
    # Use Landscape dimensions
    a4_page_width = A4_LANDSCAPE_WIDTH_PT
    a4_page_height = A4_LANDSCAPE_HEIGHT_PT

    usable_width = a4_page_width - 2 * outer_margin_pt
    usable_height = a4_page_height - 2 * outer_margin_pt

    if usable_width < central_margin_pt + 20 or usable_height < 20:
        update_status(f"Error: Margins ({outer_margin_mm}mm outer, {central_margin_mm}mm center) are too large for A4 Landscape page.")
        source_doc.close()
        output_doc.close()
        return False

    # Width for each A5 content block (side-by-side on landscape)
    a5_area_width = (usable_width - central_margin_pt) / 2
    # Height for each A5 content block
    a5_area_height = usable_height

    if a5_area_width <= 0 or a5_area_height <= 0:
        update_status(f"Error: Calculated A5 area has non-positive dimensions. Check margins.")
        source_doc.close()
        output_doc.close()
        return False

    # Define rectangles relative to A4 Landscape page origin (top-left)
    left_rect = fitz.Rect(
        outer_margin_pt,                      # x0
        outer_margin_pt,                      # y0
        outer_margin_pt + a5_area_width,      # x1
        outer_margin_pt + a5_area_height      # y1
    )
    right_rect = fitz.Rect(
        left_rect.x1 + central_margin_pt,    # x0
        outer_margin_pt,                     # y0
        left_rect.x1 + central_margin_pt + a5_area_width, # x1
        outer_margin_pt + a5_area_height     # y1
    )

    # 5. Arrange pages in booklet order
    update_status(f"Arranging {target_pages} pages (including {blanks_added} blanks)...")
    num_output_sheets = target_pages // 2
    processed_sheets = 0

    for i in range(num_output_sheets // 2):
        output_page_idx_front = i * 2
        left_page_num_front = target_pages - output_page_idx_front    # 1-based
        right_page_num_front = output_page_idx_front + 1              # 1-based

        output_page_idx_back = i * 2 + 1
        left_page_num_back = output_page_idx_back + 1                 # 1-based
        right_page_num_back = target_pages - output_page_idx_back     # 1-based

        # --- Create Front Side (A4 Landscape) ---
        sheet_front = output_doc.new_page(width=a4_page_width, height=a4_page_height)
        update_status(f"Processing Output Page {output_page_idx_front + 1}/{num_output_sheets} (Input {left_page_num_front} | {right_page_num_front})")

        # Place Left Page (Front)
        page_idx_0based_left = left_page_num_front - 1
        if 0 <= page_idx_0based_left < total_pages:
            try:
                sheet_front.show_pdf_page(left_rect, source_doc, page_idx_0based_left)
            except Exception as e:
                update_status(f"Warning placing page {page_idx_0based_left + 1}: {e}")
                error_rect = calculate_inset_rect(left_rect, 5)
                sheet_front.draw_rect(left_rect, color=(1, 0, 0), width=1)
                sheet_front.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_left + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)

        # Place Right Page (Front)
        page_idx_0based_right = right_page_num_front - 1
        if 0 <= page_idx_0based_right < total_pages:
            try:
                sheet_front.show_pdf_page(right_rect, source_doc, page_idx_0based_right)
            except Exception as e:
                update_status(f"Warning placing page {page_idx_0based_right + 1}: {e}")
                error_rect = calculate_inset_rect(right_rect, 5)
                sheet_front.draw_rect(right_rect, color=(1, 0, 0), width=1)
                sheet_front.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_right + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)
        processed_sheets += 1
        update_progress(int(100 * processed_sheets / num_output_sheets))

        # --- Create Back Side (A4 Landscape) ---
        sheet_back = output_doc.new_page(width=a4_page_width, height=a4_page_height)
        update_status(f"Processing Output Page {output_page_idx_back + 1}/{num_output_sheets} (Input {left_page_num_back} | {right_page_num_back})")

        # Place Left Page (Back)
        page_idx_0based_left = left_page_num_back - 1
        if 0 <= page_idx_0based_left < total_pages:
             try:
                 sheet_back.show_pdf_page(left_rect, source_doc, page_idx_0based_left)
             except Exception as e:
                update_status(f"Warning placing page {page_idx_0based_left + 1}: {e}")
                error_rect = calculate_inset_rect(left_rect, 5)
                sheet_back.draw_rect(left_rect, color=(1, 0, 0), width=1)
                sheet_back.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_left + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)

        # Place Right Page (Back)
        page_idx_0based_right = right_page_num_back - 1
        if 0 <= page_idx_0based_right < total_pages:
             try:
                 sheet_back.show_pdf_page(right_rect, source_doc, page_idx_0based_right)
             except Exception as e:
                update_status(f"Warning placing page {page_idx_0based_right + 1}: {e}")
                error_rect = calculate_inset_rect(right_rect, 5)
                sheet_back.draw_rect(right_rect, color=(1, 0, 0), width=1)
                sheet_back.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_right + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)
        processed_sheets += 1
        update_progress(int(100 * processed_sheets / num_output_sheets))


    # Handle odd number of output sheets (last front page) - logic remains the same
    if num_output_sheets % 2 != 0:
        i = num_output_sheets // 2
        output_page_idx_front = i * 2
        left_page_num_front = target_pages - output_page_idx_front    # 1-based
        right_page_num_front = output_page_idx_front + 1              # 1-based

        sheet_front = output_doc.new_page(width=a4_page_width, height=a4_page_height)
        update_status(f"Processing Output Page {output_page_idx_front + 1}/{num_output_sheets} (Input {left_page_num_front} | {right_page_num_front})")

        page_idx_0based_left = left_page_num_front - 1
        if 0 <= page_idx_0based_left < total_pages:
            try:
                sheet_front.show_pdf_page(left_rect, source_doc, page_idx_0based_left)
            except Exception as e:
                 update_status(f"Warning placing page {page_idx_0based_left + 1}: {e}")
                 error_rect = calculate_inset_rect(left_rect, 5)
                 sheet_front.draw_rect(left_rect, color=(1, 0, 0), width=1)
                 sheet_front.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_left + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)

        page_idx_0based_right = right_page_num_front - 1
        if 0 <= page_idx_0based_right < total_pages:
            try:
                sheet_front.show_pdf_page(right_rect, source_doc, page_idx_0based_right)
            except Exception as e:
                update_status(f"Warning placing page {page_idx_0based_right + 1}: {e}")
                error_rect = calculate_inset_rect(right_rect, 5)
                sheet_front.draw_rect(right_rect, color=(1, 0, 0), width=1)
                sheet_front.insert_textbox(error_rect, f"Error\nPage {page_idx_0based_right + 1}", fontsize=8, color=(1,0,0), align=fitz.TEXT_ALIGN_CENTER)
        processed_sheets += 1
        update_progress(int(100 * processed_sheets / num_output_sheets))


    # 6. Save the output PDF
    update_status(f"Saving booklet to {output_pdf_path}...")
    success = False
    try:
        output_doc.save(output_pdf_path, garbage=4, deflate=True, clean=True)
        update_status(f"Booklet created: {output_pdf_path}. Print double-sided, flip on LONG edge.") # Added printing hint
        success = True
        update_progress(100)
    except Exception as e:
        update_status(f"Error saving output PDF: {e}")
        success = False

    # 7. Clean up
    output_doc.close()
    source_doc.close()
    update_status("Process finished.") # Keep generic finish message here
    return success


# --- GUI Application ---

class BookletApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Booklet Creator v1.1") # Updated title
        master.geometry("550x480") # Slightly taller for About button

        self.style = ttk.Style()
        self.style.theme_use('clam')

        # --- Input File ---
        self.input_frame = ttk.LabelFrame(master, text="Input PDF", padding=(10, 5))
        self.input_frame.pack(padx=10, pady=5, fill=tk.X)
        self.input_path_var = tk.StringVar()
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path_var, width=50)
        self.input_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.input_button = ttk.Button(self.input_frame, text="Browse...", command=self.select_input_file)
        self.input_button.pack(side=tk.LEFT, padx=5)

        # --- Output File ---
        self.output_frame = ttk.LabelFrame(master, text="Output Booklet PDF (A4 Landscape)", padding=(10, 5)) # Updated Label
        self.output_frame.pack(padx=10, pady=5, fill=tk.X)
        self.output_path_var = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path_var, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.output_button = ttk.Button(self.output_frame, text="Browse...", command=self.select_output_file)
        self.output_button.pack(side=tk.LEFT, padx=5)

        # --- Options ---
        self.options_frame = ttk.LabelFrame(master, text="Options", padding=(10, 5))
        self.options_frame.pack(padx=10, pady=5, fill=tk.X)
        self.margins_frame = ttk.Frame(self.options_frame)
        self.margins_frame.pack(fill=tk.X, pady=2)
        ttk.Label(self.margins_frame, text="Center Margin (mm):").pack(side=tk.LEFT, padx=5)
        self.center_margin_var = tk.StringVar(value="10.0")
        self.center_margin_entry = ttk.Entry(self.margins_frame, textvariable=self.center_margin_var, width=6)
        self.center_margin_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(self.margins_frame, text="Outer Margin (mm):").pack(side=tk.LEFT, padx=15)
        self.outer_margin_var = tk.StringVar(value="5.0")
        self.outer_margin_entry = ttk.Entry(self.margins_frame, textvariable=self.outer_margin_var, width=6)
        self.outer_margin_entry.pack(side=tk.LEFT, padx=5)
        self.add_blanks_var = tk.BooleanVar(value=True)
        self.add_blanks_check = ttk.Checkbutton(self.options_frame, text="Add blank pages if needed (for multiple of 4)", variable=self.add_blanks_var)
        self.add_blanks_check.pack(anchor=tk.W, pady=2)

        # --- Progress & Status ---
        self.progress_frame = ttk.Frame(master, padding=(10, 5))
        self.progress_frame.pack(padx=10, pady=5, fill=tk.X)

        # Run and About Buttons side-by-side
        self.buttons_frame = ttk.Frame(self.progress_frame)
        self.buttons_frame.pack(pady=5)
        self.run_button = ttk.Button(self.buttons_frame, text="Create Booklet", command=self.run_process_threaded)
        self.run_button.pack(side=tk.LEFT, padx=10)
        self.about_button = ttk.Button(self.buttons_frame, text="About", command=self.show_about_window) # New About button
        self.about_button.pack(side=tk.LEFT, padx=10)

        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_bar.pack(pady=5, fill=tk.X, expand=True)
        self.status_label_var = tk.StringVar(value="Status: Ready")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_label_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(pady=5, fill=tk.X, expand=True)

        # --- Threading & Queue ---
        self.status_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.master.after(100, self.check_queues) # Start queue checker

    # --- GUI Methods ---
    # (select_input_file, select_output_file remain largely the same)
    def select_input_file(self):
        filepath = filedialog.askopenfilename(
            title="Select Input PDF",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))
        )
        if filepath:
            self.input_path_var.set(filepath)
            if not self.output_path_var.get():
                 base, ext = os.path.splitext(filepath)
                 self.output_path_var.set(base + "_booklet_landscape" + ext) # Suggest landscape in name

    def select_output_file(self):
        filepath = filedialog.asksaveasfilename(
            title="Save Landscape Booklet As",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")),
            defaultextension=".pdf"
        )
        if filepath:
            self.output_path_var.set(filepath)

    def update_status(self, message):
        self.status_queue.put(message)

    def update_progress(self, value):
        self.progress_queue.put(value)

    def check_queues(self):
        try:
            while True:
                message = self.status_queue.get_nowait()
                self.status_label_var.set(f"Status: {message}")
        except queue.Empty:
            pass
        try:
            while True:
                progress = self.progress_queue.get_nowait()
                self.progress_bar['value'] = progress
        except queue.Empty:
            pass
        self.master.after(100, self.check_queues)

    def run_process_threaded(self):
        input_pdf = self.input_path_var.get()
        output_pdf = self.output_path_var.get()
        center_margin_str = self.center_margin_var.get()
        outer_margin_str = self.outer_margin_var.get()
        add_blanks = self.add_blanks_var.get()

        if not input_pdf or not Path(input_pdf).is_file():
            messagebox.showerror("Input Error", "Please select a valid input PDF file.")
            return
        if not output_pdf:
            messagebox.showerror("Output Error", "Please specify an output PDF file path.")
            return
        try:
            center_margin_mm = float(center_margin_str)
            outer_margin_mm = float(outer_margin_str)
            if center_margin_mm < 0 or outer_margin_mm < 0:
                raise ValueError("Margins cannot be negative.")
        except ValueError as e:
            messagebox.showerror("Margin Error", f"Invalid margin value: {e}. Please enter numbers.")
            return

        self.run_button.config(state=tk.DISABLED)
        self.about_button.config(state=tk.DISABLED) # Disable about button too
        self.progress_bar['value'] = 0

        thread = threading.Thread(
            target=self.worker_thread_task,
            args=(input_pdf, output_pdf, center_margin_mm, outer_margin_mm, add_blanks),
            daemon=True
        )
        thread.start()

    def worker_thread_task(self, input_pdf, output_pdf, center_mm, outer_mm, add_blanks_flag):
        """ The actual task run by the thread """
        success = False
        try:
            success = create_booklet(
                input_pdf_path=input_pdf,
                output_pdf_path=output_pdf,
                central_margin_mm=center_mm,
                outer_margin_mm=outer_mm,
                add_blanks=add_blanks_flag,
                status_callback=self.update_status,
                progress_callback=self.update_progress
            )
            # Status updated within create_booklet
        except Exception as e:
             self.update_status(f"An unexpected error occurred: {e}")
        finally:
            # Re-enable buttons via main thread
            self.master.after(0, lambda: self.run_button.config(state=tk.NORMAL))
            self.master.after(0, lambda: self.about_button.config(state=tk.NORMAL))
            # Optionally show final message box after completion check
            if success:
                 # self.master.after(10, lambda: messagebox.showinfo("Success", f"Booklet created!\n{output_pdf}\n\nRemember to print double-sided, flipping on the LONG edge."))
                 pass # Status bar already shows success message with hint
            else:
                 self.master.after(10, lambda: messagebox.showerror("Failed", "Booklet creation failed. Check status messages."))


    def show_about_window(self):
        """ Displays the About window """
        about_win = tk.Toplevel(self.master)
        about_win.title("About PDF Booklet Creator")
        about_win.geometry("380x220")
        about_win.resizable(False, False)
        # Center window (optional)
        # self.master.eval(f'tk::PlaceWindow {str(about_win)} center')

        about_frame = ttk.Frame(about_win, padding=15)
        about_frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(about_frame, text="PDF Booklet Creator", font=('Helvetica', 14, 'bold')).pack(pady=(0, 5))
        ttk.Label(about_frame, text=f"Version: {PROGRAM_VERSION}").pack()
        ttk.Label(about_frame, text=f"Release Date: {RELEASE_DATE}").pack(pady=(0, 15))

        ttk.Label(about_frame, text="Author: Imam Wahyudi", font=('Helvetica', 10, 'bold')).pack()

        # Clickable Links
        gh_link = ttk.Label(about_frame, text="GitHub", foreground="blue", cursor="hand2")
        gh_link.pack()
        gh_link.bind("<Button-1>", lambda e: open_link(AUTHOR_INFO["GitHub"]))

        li_link = ttk.Label(about_frame, text="LinkedIn", foreground="blue", cursor="hand2")
        li_link.pack()
        li_link.bind("<Button-1>", lambda e: open_link(AUTHOR_INFO["LinkedIn"]))


        ttk.Button(about_frame, text="OK", command=about_win.destroy).pack(pady=(20, 0))

        about_win.transient(self.master) # Keep on top of main window
        about_win.grab_set() # Make modal (wait for this window)
        self.master.wait_window(about_win) # Wait until closed


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = BookletApp(root)
    root.mainloop()
