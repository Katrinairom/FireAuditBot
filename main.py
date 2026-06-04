import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Import our custom modules
from src.excel_reader import read_store_data
from src.validators import run_preflight_checks
from src.formula_engine import calculate_loads
from src.nbc_checker import check_compliance
from src.report_builder import generate_report


class FireAuditBotApp:
    def __init__(self, root):
        self.root = root  # <-- Always assign the root window first thing!

        # This dynamically finds the folder main.py is sitting in
        base_dir = os.path.dirname(os.path.abspath(__file__))

        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar(value=os.path.join(
            base_dir, "data", "FA0XX_template.docx"))
        self.output_dir = tk.StringVar(value=os.path.join(base_dir, "output"))

        self.root.title("FireAuditBot - Auto Report Generator")
        self.root.geometry("500x350")

        self.setup_ui()

    def setup_ui(self):
        # Excel Input
        tk.Label(self.root, text="1. Select Excel Data File:").pack(
            pady=(15, 0))
        frame1 = tk.Frame(self.root)
        frame1.pack(pady=5)
        tk.Entry(frame1, textvariable=self.excel_path, width=40,
                 state='readonly').pack(side=tk.LEFT, padx=5)
        tk.Button(frame1, text="Browse",
                  command=self.browse_excel).pack(side=tk.LEFT)

        # Word Template Input
        tk.Label(self.root, text="2. Word Template File:").pack(pady=(10, 0))
        frame2 = tk.Frame(self.root)
        frame2.pack(pady=5)
        tk.Entry(frame2, textvariable=self.template_path, width=40,
                 state='readonly').pack(side=tk.LEFT, padx=5)
        tk.Button(frame2, text="Browse",
                  command=self.browse_template).pack(side=tk.LEFT)

        # Generate Button
        tk.Button(self.root, text="Generate Report", bg="green", fg="white", font=("Arial", 12, "bold"),
                  command=self.process_audit).pack(pady=30, ipadx=10, ipady=5)

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.excel_path.set(filename)

    def browse_template(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")])
        if filename:
            self.template_path.set(filename)

    def process_audit(self):
        excel_file = self.excel_path.get()
        template_file = self.template_path.get()

        # Explicit path pointing to your background Inspection calculator sheet
        inspection_template = os.path.join(os.path.dirname(os.path.abspath(
            __file__)), "data", "Inspection_Sheet_Drafters_Copy.xlsx")

        if not excel_file:
            messagebox.showwarning(
                "Missing File", "Please select an Excel file first.")
            return

        try:
            # 1. Run the customized workflow engine (Reads raw, maps to inspector, evaluates)
            extracted_data = read_store_data(excel_file, inspection_template)

            # 2. Setup paths and execute report assembly
            out_filename = f"Report_{extracted_data.get('store_code', 'UNKNOWN')}.docx"
            out_path = os.path.join(self.output_dir.get(), out_filename)
            os.makedirs(self.output_dir.get(), exist_ok=True)

            generate_report(template_file, out_path, extracted_data)

            messagebox.showinfo(
                "Success", f"Report generated successfully!\nSaved to: {out_path}")

        except Exception as e:
            messagebox.showerror(
                "Error", f"An error occurred during processing:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FireAuditBotApp(root)
    root.mainloop()
