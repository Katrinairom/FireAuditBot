# main.py
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import date

from src.excel_reader import read_store_data
from src.validators import run_preflight_checks
from src.formula_engine import calculate_loads
from src.nbc_checker import check_compliance
from src.report_builder import generate_report


class FireAuditBotApp:
    def __init__(self, root):
        self.root = root
        base_dir = os.path.dirname(os.path.abspath(__file__))

        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar(
            value=os.path.join(base_dir, "data", "FA0XX_template.docx"))
        self.output_dir = tk.StringVar(
            value=os.path.join(base_dir, "output"))

        # Date range for the audit period
        self.date_start = tk.StringVar(value="28th February 2026")
        self.date_end = tk.StringVar(value="08th March 2026")

        self.root.title("FireAuditBot - Auto Report Generator")
        self.root.geometry("540x420")
        self.root.resizable(False, False)

        self._setup_ui()

    def _setup_ui(self):
        pad = {"pady": (10, 0)}

        # --- Excel file picker ---
        tk.Label(self.root, text="1. Select Store Excel File:",
                 font=("Calibri", 10, "bold")).pack(**pad)
        f1 = tk.Frame(self.root)
        f1.pack(pady=4)
        tk.Entry(f1, textvariable=self.excel_path,    width=42,
                 state="readonly").pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="Browse", command=self._browse_excel).pack(
            side=tk.LEFT)

        # --- Word template picker ---
        tk.Label(self.root, text="2. Word Template File:",
                 font=("Calibri", 10, "bold")).pack(**pad)
        f2 = tk.Frame(self.root)
        f2.pack(pady=4)
        tk.Entry(f2, textvariable=self.template_path, width=42,
                 state="readonly").pack(side=tk.LEFT, padx=5)
        tk.Button(f2, text="Browse", command=self._browse_template).pack(
            side=tk.LEFT)

        # --- Audit date range ---
        tk.Label(self.root, text="3. Audit Period:",
                 font=("Calibri", 10, "bold")).pack(**pad)
        f3 = tk.Frame(self.root)
        f3.pack(pady=4)
        tk.Label(f3, text="From:").pack(side=tk.LEFT, padx=(10, 2))
        tk.Entry(f3, textvariable=self.date_start, width=18).pack(side=tk.LEFT)
        tk.Label(f3, text="  To:").pack(side=tk.LEFT, padx=(10, 2))
        tk.Entry(f3, textvariable=self.date_end,   width=18).pack(side=tk.LEFT)

        # --- Output directory ---
        tk.Label(self.root, text="4. Output Folder:",
                 font=("Calibri", 10, "bold")).pack(**pad)
        f4 = tk.Frame(self.root)
        f4.pack(pady=4)
        tk.Entry(f4, textvariable=self.output_dir,  width=42,
                 state="readonly").pack(side=tk.LEFT, padx=5)
        tk.Button(f4, text="Browse", command=self._browse_output).pack(
            side=tk.LEFT)

        # --- Generate button ---
        tk.Button(
            self.root, text="Generate Report",
            bg="#1a7f3c", fg="white", font=("Calibri", 12, "bold"),
            command=self._process_audit
        ).pack(pady=28, ipadx=14, ipady=6)

    # ------------------------------------------------------------------ #
    #  BROWSE CALLBACKS
    # ------------------------------------------------------------------ #

    def _browse_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.excel_path.set(f)

    def _browse_template(self):
        f = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")])
        if f:
            self.template_path.set(f)

    def _browse_output(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir.set(d)

    # ------------------------------------------------------------------ #
    #  MAIN WORKFLOW
    # ------------------------------------------------------------------ #

    def _process_audit(self):
        excel_file = self.excel_path.get()
        template_file = self.template_path.get()

        if not excel_file:
            messagebox.showwarning(
                "Missing File", "Please select a Store Excel file first.")
            return
        if not os.path.exists(template_file):
            messagebox.showwarning("Missing Template",
                                   f"Word template not found:\n{template_file}\n\nPlease browse to your template file.")
            return

        try:
            # Step 1 — Read Excel (no Drafter's Copy needed)
            store_data = read_store_data(excel_file)

            # Step 2 — Pre-flight validation
            warnings = run_preflight_checks(store_data, excel_file)
            if warnings:
                warn_text = "\n".join(f"• {w}" for w in warnings)
                proceed = messagebox.askyesno(
                    "Warnings Found",
                    f"The following issues were found:\n\n{warn_text}"
                    "\n\nDo you want to continue generating the report anyway?"
                )
                if not proceed:
                    return

            # Step 3 — Calculate occupancy load + fire load density
            formula_results = calculate_loads(store_data)
            store_data.update(formula_results)

            # Step 4 — NBC compliance check (uses property_area, not store_area)
            nbc_results = check_compliance(store_data)
            store_data["nbc"] = nbc_results

            # Step 5 — Add date range to store_data for report
            store_data["audit_date_start"] = self.date_start.get()
            store_data["audit_date_end"] = self.date_end.get()
            store_data["report_date"] = date.today().strftime("%d/%m/%Y")

            # Step 6 — Generate report
            os.makedirs(self.output_dir.get(), exist_ok=True)
            out_filename = f"{store_data.get('store_code', 'UNKNOWN')}-FIRE_AUDIT_REPORT.docx"
            out_path = os.path.join(self.output_dir.get(), out_filename)

            generate_report(template_file, out_path, store_data)

            messagebox.showinfo(
                "Success",
                f"Report generated successfully!\n\nSaved to:\n{out_path}"
            )

        except FileNotFoundError as e:
            messagebox.showerror("File Not Found", str(e))
        except KeyError as e:
            messagebox.showerror("Sheet Not Found",
                                 f"Could not find sheet or column: {e}\n\n"
                                 "Check that the Excel sheet names match exactly:\n"
                                 "- 'Property and Store Details'\n"
                                 "- 'Firefighting Equipment Details'\n"
                                 "- 'Checklist'\n"
                                 "- 'Firefighting Equipment Store Wi'")
        except Exception as e:
            messagebox.showerror("Error",
                                 f"An error occurred during processing:\n\n{str(e)}\n\n"
                                 "If this keeps happening, please note the error message and raise it.")


if __name__ == "__main__":
    root = tk.Tk()
    app = FireAuditBotApp(root)
    root.mainloop()
