import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from parse_cie_statement import CambridgeOCRExtractor
from parse_predicted_grade_statement import PredictedGradeExtractor
from parse_ucas_statement import UCASExtractor


class ExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CIE Statement & UCAS Extractor")
        self.root.geometry("800x400")

        self.statement_dir = tk.StringVar()
        self.ucas_dir = tk.StringVar()
        self.predicted_grade_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        self._create_widgets()

    def _create_widgets(self):
        padding = {"padx": 10, "pady": 5}

        ttk.Label(self.root, text="CIE Statement Directory:").pack(fill=tk.X, **padding)
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

        ttk.Label(self.root, text="Predicted Grade PDF Directory:").pack(
            fill=tk.X, **padding
        )
        predicted_frame = ttk.Frame(self.root)
        predicted_frame.pack(fill=tk.X, **padding)
        ttk.Entry(
            predicted_frame, textvariable=self.predicted_grade_dir, width=50
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(
            predicted_frame, text="Browse", command=self._browse_predicted_grade_dir
        ).pack(side=tk.RIGHT, padx=5)

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
            text="Generate CIE Statement XLSX",
            command=self._generate_cie_xlsx,
        ).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        ttk.Button(
            button_frame,
            text="Generate UCAS XLSX",
            command=self._generate_ucas_xlsx,
        ).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        ttk.Button(
            button_frame,
            text="Generate Predicted Grade XLSX",
            command=self._generate_predicted_xlsx,
        ).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            self.root, textvariable=self.status_var, relief=tk.SUNKEN
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _browse_statement_dir(self):
        path = filedialog.askdirectory(title="Select CIE Statement/Statement Directory")
        if path:
            self.statement_dir.set(path)

    def _browse_ucas_dir(self):
        path = filedialog.askdirectory(title="Select UCAS PDF Directory")
        if path:
            self.ucas_dir.set(path)

    def _browse_predicted_grade_dir(self):
        path = filedialog.askdirectory(title="Select Predicted Grade PDF Directory")
        if path:
            self.predicted_grade_dir.set(path)

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.output_dir.set(path)

    def _get_output_path(self, filename):
        output_dir = self.output_dir.get()
        if output_dir:
            return os.path.join(output_dir, filename)
        return filename

    def _validate_directory(self, directory, label):
        if not directory:
            messagebox.showwarning("Warning", f"Please select a directory for {label}.")
            return False
        if not os.path.isdir(directory):
            messagebox.showerror(
                "Error", f"Selected {label} does not exist:\n{directory}"
            )
            return False
        return True

    def _validate_output_directory(self):
        output_dir = self.output_dir.get().strip()
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror(
                "Error", f"Selected output directory does not exist:\n{output_dir}"
            )
            return False
        return True

    def _list_pdf_files(self, directory):
        return sorted(f for f in os.listdir(directory) if f.lower().endswith(".pdf"))

    def _start_generation(self, status_message, worker, directory):
        self.status_var.set(status_message)
        self._set_buttons_enabled(False)
        thread = threading.Thread(target=worker, args=(directory,), daemon=True)
        thread.start()

    def _set_status(self, message):
        self.root.after(0, lambda: self.status_var.set(message))

    def _show_info(self, title, message):
        self.root.after(0, lambda: messagebox.showinfo(title, message))

    def _show_error_dialog(self, title, message, errors):
        self.root.after(0, lambda: self._show_error_summary(title, message, errors))

    def _build_progress_callback(self, current_file_idx, total_files, filename):
        def progress_callback(page_num, total_pages):
            self._set_status(
                f"[{current_file_idx}/{total_files}] Processing: {filename} [page {page_num}/{total_pages}]"
            )

        return progress_callback

    def _run_batch_job(
        self,
        directory,
        *,
        empty_status,
        output_filename_prefix,
        process_file,
        write_output,
        count_records,
        build_error_context=None,
    ):
        errors = []
        try:
            pdf_files = self._list_pdf_files(directory)
            if not pdf_files:
                self._show_info("Info", "No PDF files found in the selected directory.")
                self._set_status("No PDF files found.")
                return

            total_files = len(pdf_files)
            all_data = []

            for idx, pdf_file in enumerate(pdf_files, 1):
                pdf_path = os.path.join(directory, pdf_file)
                self._set_status(f"[{idx}/{total_files}] Processing: {pdf_file}")

                try:
                    result = process_file(
                        pdf_path,
                        self._build_progress_callback(idx, total_files, pdf_file),
                    )
                    all_data.extend(result)
                except Exception as e:
                    import traceback

                    debug_context = ""
                    if build_error_context is not None:
                        try:
                            debug_context = build_error_context(pdf_path)
                        except Exception as debug_error:
                            debug_context = (
                                f"<failed to extract debug text>\n{debug_error}"
                            )

                    errors.append(
                        (pdf_file, str(e), traceback.format_exc(), debug_context)
                    )

            if all_data:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = self._get_output_path(
                    f"{output_filename_prefix}_{timestamp}.xlsx"
                )
                write_output(all_data, output_path)

                extracted_count = count_records(all_data)
                self._set_status(f"Done! Written to: {output_path}")
                if errors:
                    self._show_error_dialog(
                        "Completed with Errors",
                        f"Generated {output_path}\n\nExtracted {extracted_count} record(s).\n{len(errors)} file(s) failed to process.",
                        errors,
                    )
                else:
                    self._show_info(
                        "Success",
                        f"Generated {output_path}\n\nExtracted {extracted_count} record(s).",
                    )
            else:
                error_msg = empty_status
                if errors:
                    error_msg += f"\n\n{len(errors)} file(s) had errors."
                self._show_error_dialog("Error", error_msg, errors)
                self._set_status(empty_status)

        except Exception as e:
            import traceback

            self._show_error_dialog(
                "Error",
                f"Failed to process: {e}",
                [(directory, str(e), traceback.format_exc(), "")],
            )
            self._set_status("Error occurred.")
        finally:
            self.root.after(0, lambda: self._set_buttons_enabled(True))

    def _generate_cie_xlsx(self):
        directory = self.statement_dir.get()
        if not self._validate_directory(directory, "CIE statements/statements"):
            return
        if not self._validate_output_directory():
            return

        self._start_generation(
            "Processing CIE statements/statements...",
            self._generate_cie_xlsx_thread,
            directory,
        )

    def _generate_cie_xlsx_thread(self, directory):
        extractor = CambridgeOCRExtractor(dpi=300)
        self._run_batch_job(
            directory,
            empty_status="No records extracted.",
            output_filename_prefix="cie_results",
            process_file=lambda pdf_path, progress_callback: extractor.extract(
                pdf_path, progress_callback
            ),
            write_output=extractor.write_to_xlsx,
            count_records=lambda records: sum(
                len(record.subjects) for record in records
            ),
            build_error_context=None,
        )

    def _generate_ucas_xlsx(self):
        directory = self.ucas_dir.get()
        if not self._validate_directory(directory, "UCAS PDFs"):
            return
        if not self._validate_output_directory():
            return

        self._start_generation(
            "Processing UCAS PDFs...", self._generate_ucas_xlsx_thread, directory
        )

    def _generate_ucas_xlsx_thread(self, directory):
        writer = UCASExtractor()
        self._run_batch_job(
            directory,
            empty_status="No records extracted.",
            output_filename_prefix="ucas_results",
            process_file=lambda pdf_path, progress_callback: [
                UCASExtractor(pdf_path).extract(progress_callback)
            ],
            write_output=writer.write_to_xlsx,
            count_records=lambda all_data: sum(
                len(data.education) for data in all_data
            ),
            build_error_context=lambda pdf_path: UCASExtractor(
                pdf_path
            ).debug_dump_text(),
        )

    def _generate_predicted_xlsx(self):
        directory = self.predicted_grade_dir.get()
        if not self._validate_directory(directory, "Predicted Grade PDFs"):
            return
        if not self._validate_output_directory():
            return

        self._start_generation(
            "Processing Predicted Grade PDFs...",
            self._generate_predicted_xlsx_thread,
            directory,
        )

    def _generate_predicted_xlsx_thread(self, directory):
        extractor = PredictedGradeExtractor(dpi=300)
        self._run_batch_job(
            directory,
            empty_status="No records extracted.",
            output_filename_prefix="predicted_grades",
            process_file=lambda pdf_path, progress_callback: extractor.extract(
                pdf_path, progress_callback
            ),
            write_output=extractor.write_to_xlsx,
            count_records=lambda records: sum(
                len(record.subjects) for record in records
            ),
            build_error_context=None,
        )

    def _show_error_summary(self, title, message, errors):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text=message, wraplength=560).pack(anchor=tk.W)

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        ttk.Label(frame, text="Error Details:").pack(anchor=tk.W)

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)

        for filename, error_msg, traceback_detail, debug_context in errors:
            text.insert(tk.END, f"File: {filename}\n")
            text.insert(tk.END, f"Error: {error_msg}\n")
            text.insert(tk.END, f"Traceback:\n{traceback_detail}\n")
            if debug_context:
                text.insert(tk.END, f"Extracted Text:\n{debug_context}\n")
            text.insert(tk.END, "-" * 60 + "\n\n")

        text.config(state=tk.DISABLED)

        ttk.Button(frame, text="Close", command=dialog.destroy).pack(pady=10)

    def _set_buttons_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Button):
                        child.configure(state=state)


def main():
    root = tk.Tk()
    app = ExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
