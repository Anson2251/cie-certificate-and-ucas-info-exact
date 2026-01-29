import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from parse_cie_certificate import CambridgeOCRExtractor
from parse_ucas_statement import UCASExtractor
from datetime import datetime


class ExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CIE Certificate & UCAS Extractor")
        self.root.geometry("600x400")

        self.statement_dir = tk.StringVar()
        self.ucas_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        self._create_widgets()

    def _create_widgets(self):
        padding = {"padx": 10, "pady": 5}

        ttk.Label(self.root, text="CIE Certificate/Statement Directory:").pack(
            fill=tk.X, **padding
        )
        statement_frame = ttk.Frame(self.root)
        statement_frame.pack(fill=tk.X, **padding)
        ttk.Entry(statement_frame, textvariable=self.statement_dir, width=50).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(
            statement_frame, text="Browse", command=self._browse_statement_dir
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Label(self.root, text="UCAS PDF Directory:").pack(fill=tk.X, **padding)
        ucas_frame = ttk.Frame(self.root)
        ucas_frame.pack(fill=tk.X, **padding)
        ttk.Entry(ucas_frame, textvariable=self.ucas_dir, width=50).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(ucas_frame, text="Browse", command=self._browse_ucas_dir).pack(
            side=tk.RIGHT, padx=5
        )

        ttk.Label(self.root, text="Output Directory:").pack(fill=tk.X, **padding)
        output_frame = ttk.Frame(self.root)
        output_frame.pack(fill=tk.X, **padding)
        ttk.Entry(output_frame, textvariable=self.output_dir, width=50).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(output_frame, text="Browse", command=self._browse_output_dir).pack(
            side=tk.RIGHT, padx=5
        )

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)

        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(
            button_frame,
            text="Generate CIE Certificate XLSX",
            command=self._generate_cie_xlsx,
        ).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        ttk.Button(
            button_frame,
            text="Generate UCAS XLSX",
            command=self._generate_ucas_xlsx,
        ).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            self.root, textvariable=self.status_var, relief=tk.SUNKEN
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _browse_statement_dir(self):
        path = filedialog.askdirectory(
            title="Select CIE Certificate/Statement Directory"
        )
        if path:
            self.statement_dir.set(path)

    def _browse_ucas_dir(self):
        path = filedialog.askdirectory(title="Select UCAS PDF Directory")
        if path:
            self.ucas_dir.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.output_dir.set(path)

    def _get_output_path(self, filename):
        output_dir = self.output_dir.get()
        if output_dir:
            return os.path.join(output_dir, filename)
        return filename

    def _generate_cie_xlsx(self):
        directory = self.statement_dir.get()
        if not directory:
            messagebox.showwarning(
                "Warning", "Please select a directory for CIE certificates/statements."
            )
            return

        self.status_var.set("Processing CIE certificates/statements...")
        self.root.update()

        try:
            pdf_files = [f for f in os.listdir(directory) if f.lower().endswith(".pdf")]
            if not pdf_files:
                messagebox.showinfo(
                    "Info", "No PDF files found in the selected directory."
                )
                return

            pdf_paths = [os.path.join(directory, f) for f in pdf_files]
            extractor = CambridgeOCRExtractor(dpi=300)
            all_records = extractor.extract_all(pdf_paths)

            if all_records:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = self._get_output_path(f"cie_results_{timestamp}.xlsx")
                extractor.write_to_xlsx(all_records, output_path)
                self.status_var.set(f"Done! Written to: {output_path}")
                messagebox.showinfo("Success", f"Generated {output_path}")
            else:
                messagebox.showerror("Error", "No records extracted from PDFs.")
                self.status_var.set("No records extracted.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process: {e}")
            self.status_var.set("Error occurred.")

    def _generate_ucas_xlsx(self):
        directory = self.ucas_dir.get()
        if not directory:
            messagebox.showwarning(
                "Warning", "Please select a directory for UCAS PDFs."
            )
            return

        self.status_var.set("Processing UCAS PDFs...")
        self.root.update()

        try:
            pdf_files = [f for f in os.listdir(directory) if f.lower().endswith(".pdf")]
            if not pdf_files:
                messagebox.showinfo(
                    "Info", "No PDF files found in the selected directory."
                )
                return

            all_data = []

            for pdf_file in pdf_files:
                pdf_path = os.path.join(directory, pdf_file)
                self.status_var.set(f"Processing: {pdf_file}")
                self.root.update()

                try:
                    extractor = UCASExtractor(pdf_path)
                    data = extractor.extract()
                    print(
                        f"Extracted {len(data.education)} education entries from {pdf_file}"
                    )
                    all_data.append(data)
                except Exception as e:
                    print(f"Error processing {pdf_file}: {e}")

            print(f"Total PDFs processed: {len(all_data)}")
            total_entries = sum(len(data.education) for data in all_data)
            print(f"Total education entries collected: {total_entries}")

            if all_data:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = self._get_output_path(f"ucas_results_{timestamp}.xlsx")
                UCASExtractor.write_combined_to_xlsx(
                    UCASExtractor(), all_data, output_path
                )
                print(f"Written to: {output_path}")
                self.status_var.set(f"Done! Written to: {output_path}")
                messagebox.showinfo("Success", f"Generated {output_path}")
            else:
                messagebox.showerror("Error", "No records extracted from PDFs.")
                self.status_var.set("No records extracted.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process: {e}")
            self.status_var.set("Error occurred.")


def main():
    root = tk.Tk()
    app = ExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
