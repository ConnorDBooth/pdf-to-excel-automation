import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from mold_processing import (
    find_mold_values, insert_into_excel, total_count, mean_count, stdv_count,
    display_mold_type_frequency, find_min, fifth_percentile, find_median,
    find_ninety_fifth_percentile, find_max, find_count, clear_old_stats
)

def show_progress(root, message="Processing..."):
    progress_win = tk.Toplevel(root)
    progress_win.title("Please wait")
    progress_win.geometry("300x80")
    progress_win.resizable(False, False)
    label = tk.Label(progress_win, text=message, padx=20, pady=20)
    label.pack(expand=True)
    progress_win.protocol("WM_DELETE_WINDOW", lambda: None)
    progress_win.update()
    return progress_win
def main_gui():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    pdf_path = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    excel_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not excel_path:
        return

    progress_win = show_progress(root, "Processing files, please wait...")
    root.after(100, lambda: process_files(pdf_path, excel_path, progress_win, root))
    root.mainloop()

def process_files(pdf_path, excel_path, progress_win, root):
    try:
        info = find_mold_values(pdf_path)
        if info is not None:
            mold_dict, lab_reference_number = info
            try:
                workbook = load_workbook(excel_path)
                sheet = workbook.active
                clear_old_stats(sheet)  # Clear previous stats before recalculating
                insert_into_excel(mold_dict, sheet, lab_reference_number)
                total_count(sheet)
                mean_count(sheet)
                stdv_count(sheet)
                display_mold_type_frequency(sheet)
                find_min(sheet)
                fifth_percentile(sheet)
                find_median(sheet)
                find_ninety_fifth_percentile(sheet)
                find_max(sheet)
                find_count(sheet)
                workbook.save(excel_path)
                progress_win.destroy()
                messagebox.showinfo("Success", f"Saved updated Excel file as '{excel_path}'")
            except PermissionError:
                progress_win.destroy()
                messagebox.showerror("Error", f"Permission denied: Unable to save to '{excel_path}'. Please close the file if it is open.")
            except Exception as e:
                progress_win.destroy()
                messagebox.showerror("Error", f"An error occurred while processing the Excel file: {e}")
        else:
            progress_win.destroy()
            messagebox.showerror("Error", "No 'Outdoor' section found in the PDF. No data to insert into Excel.")
    finally:
        root.quit()

if __name__ == "__main__":
    main_gui()