import os
import pdfkit
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
import re

class EnhancedPDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced HTML to PDF Converter")
        self.root.geometry("800x600")

        # Enter path to wkhtmltopdf executable here
        self.wkhtmltopdf_path = [INSERT PATH TO EXECUTABLE HERE]
        # for example:
        # self.wkhtmltopdf_path = r"C:\Scripts\wkhtmltopdf\bin\wkhtmltopdf.exe"

        if not os.path.exists(self.wkhtmltopdf_path):
            messagebox.showwarning("Setup Required", 
                                   f"wkhtmltopdf not found at {self.wkhtmltopdf_path}.\n"
                                   f"Please install it and update the path in the script.")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.main_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)
        self.log_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.main_tab, text="Conversion")
        self.notebook.add(self.settings_tab, text="Settings")
        self.notebook.add(self.log_tab, text="Log")

        self.setup_main_tab()
        self.setup_settings_tab()
        self.setup_log_tab()

        self.conversion_running = False
        self.pause_conversion = False
        self.processed_files = []
        self.failed_files = []

        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "converted_html_files")
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        self.settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf_converter_settings.json")
        self.load_settings()

    def setup_main_tab(self):
        folder_frame = ttk.LabelFrame(self.main_tab, text="HTML Files Source", padding="10")
        folder_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(folder_frame, text="Folder Path:").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.folder_path = tk.StringVar()
        ttk.Entry(folder_frame, width=50, textvariable=self.folder_path).grid(column=1, row=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(folder_frame, text="Browse...", command=self.browse_folder).grid(column=2, row=0, sticky=tk.W, pady=5, padx=5)

        self.include_subdirs = tk.BooleanVar(value=True)
        ttk.Checkbutton(folder_frame, text="Include subdirectories", variable=self.include_subdirs).grid(column=1, row=1, sticky=tk.W)

        ttk.Label(folder_frame, text="File Filter:").grid(column=0, row=2, sticky=tk.W, pady=5)
        self.file_filter = tk.StringVar(value="*.html;*.htm")
        ttk.Entry(folder_frame, width=50, textvariable=self.file_filter).grid(column=1, row=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(folder_frame, text="(Separate patterns with semicolons)").grid(column=2, row=2, sticky=tk.W, padx=5)

        progress_frame = ttk.LabelFrame(self.main_tab, text="Conversion Progress", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

        status_frame = ttk.Frame(progress_frame)
        status_frame.pack(fill=tk.X)

        ttk.Label(status_frame, text="Current File:").grid(column=0, row=0, sticky=tk.W, pady=2)
        self.current_file_var = tk.StringVar(value="")
        ttk.Label(status_frame, textvariable=self.current_file_var, width=60).grid(column=1, row=0, sticky=tk.W, pady=2)

        ttk.Label(status_frame, text="Progress:").grid(column=0, row=1, sticky=tk.W, pady=2)
        self.progress_stats_var = tk.StringVar(value="0 of 0 files (0%)")
        ttk.Label(status_frame, textvariable=self.progress_stats_var).grid(column=1, row=1, sticky=tk.W, pady=2)

        ttk.Label(status_frame, text="Results:").grid(column=0, row=2, sticky=tk.W, pady=2)
        self.results_stats_var = tk.StringVar(value="0 succeeded, 0 failed")
        ttk.Label(status_frame, textvariable=self.results_stats_var).grid(column=1, row=2, sticky=tk.W, pady=2)

        ttk.Label(status_frame, text="Est. Time Remaining:").grid(column=0, row=3, sticky=tk.W, pady=2)
        self.time_remaining_var = tk.StringVar(value="--")
        ttk.Label(status_frame, textvariable=self.time_remaining_var).grid(column=1, row=3, sticky=tk.W, pady=2)

        buttons_frame = ttk.Frame(self.main_tab)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_btn = ttk.Button(buttons_frame, text="Start Conversion", command=self.start_conversion)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(buttons_frame, text="Pause", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.cancel_btn = ttk.Button(buttons_frame, text="Cancel", command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

        self.open_output_btn = ttk.Button(buttons_frame, text="Open Output Folder", command=self.open_output_folder)
        self.open_output_btn.pack(side=tk.RIGHT, padx=5)

        folder_frame.columnconfigure(1, weight=1)

    def setup_settings_tab(self):
        settings_frame = ttk.Frame(self.settings_tab, padding="10")
        settings_frame.pack(fill=tk.BOTH, expand=True)

        pdf_frame = ttk.LabelFrame(settings_frame, text="PDF Output Settings", padding="10")
        pdf_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(pdf_frame, text="Page Size:").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.page_size = tk.StringVar(value="Letter")
        page_sizes = ["A4", "Letter", "Legal", "Tabloid"]
        ttk.Combobox(pdf_frame, textvariable=self.page_size, values=page_sizes, width=15).grid(column=1, row=0, sticky=tk.W, pady=5)

        ttk.Label(pdf_frame, text="Orientation:").grid(column=0, row=1, sticky=tk.W, pady=5)
        self.orientation = tk.StringVar(value="Portrait")
        orientations = ["Portrait", "Landscape"]
        ttk.Combobox(pdf_frame, textvariable=self.orientation, values=orientations, width=15).grid(column=1, row=1, sticky=tk.W, pady=5)

        margin_frame = ttk.Frame(pdf_frame)
        margin_frame.grid(column=0, row=2, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(margin_frame, text="Margins (mm):").grid(column=0, row=0, sticky=tk.W)

        ttk.Label(margin_frame, text="Top:").grid(column=0, row=1, sticky=tk.W, pady=2)
        self.margin_top = tk.StringVar(value="10")
        ttk.Spinbox(margin_frame, from_=0, to=50, textvariable=self.margin_top, width=5).grid(column=1, row=1, sticky=tk.W, pady=2, padx=5)

        ttk.Label(margin_frame, text="Right:").grid(column=2, row=1, sticky=tk.W, pady=2)
        self.margin_right = tk.StringVar(value="10")
        ttk.Spinbox(margin_frame, from_=0, to=50, textvariable=self.margin_right, width=5).grid(column=3, row=1, sticky=tk.W, pady=2, padx=5)

        ttk.Label(margin_frame, text="Bottom:").grid(column=0, row=2, sticky=tk.W, pady=2)
        self.margin_bottom = tk.StringVar(value="10")
        ttk.Spinbox(margin_frame, from_=0, to=50, textvariable=self.margin_bottom, width=5).grid(column=1, row=2, sticky=tk.W, pady=2, padx=5)

        ttk.Label(margin_frame, text="Left:").grid(column=2, row=2, sticky=tk.W, pady=2)
        self.margin_left = tk.StringVar(value="10")
        ttk.Spinbox(margin_frame, from_=0, to=50, textvariable=self.margin_left, width=5).grid(column=3, row=2, sticky=tk.W, pady=2, padx=5)

        perf_frame = ttk.LabelFrame(settings_frame, text="Performance Settings", padding="10")
        perf_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(perf_frame, text="Worker Threads:").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.worker_threads = tk.IntVar(value=4)
        ttk.Spinbox(perf_frame, from_=1, to=16, textvariable=self.worker_threads, width=5).grid(column=1, row=0, sticky=tk.W, pady=5)

        filename_frame = ttk.LabelFrame(settings_frame, text="Filename Settings", padding="10")
        filename_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(filename_frame, text="Output Filename:").grid(column=0, row=0, sticky=tk.W, pady=5)
        self.filename_pattern = tk.StringVar(value="{filename}")
        ttk.Entry(filename_frame, textvariable=self.filename_pattern).grid(column=1, row=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(filename_frame, text="Use {filename}, {date}, {time}").grid(column=2, row=0, sticky=tk.W, pady=5, padx=5)

        self.overwrite_files = tk.BooleanVar(value=True)
        ttk.Checkbutton(filename_frame, text="Overwrite existing files", variable=self.overwrite_files).grid(column=1, row=1, sticky=tk.W)

        ttk.Button(settings_frame, text="Save Settings", command=self.save_settings).pack(pady=10)

    def setup_log_tab(self):
        log_frame = ttk.Frame(self.log_tab, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)

        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.pack(fill=tk.X, pady=5)

        ttk.Button(log_buttons_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_buttons_frame, text="Save Log", command=self.save_log).pack(side=tk.LEFT, padx=5)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def open_output_folder(self):
        if os.path.exists(self.output_dir):
            if os.name == 'nt':  # Windows
                os.startfile(self.output_dir)
            elif os.name == 'posix':  # macOS or Linux
                import subprocess
                if os.uname().sysname == 'Darwin':  # macOS
                    subprocess.call(['open', self.output_dir])
                else:  # Linux
                    subprocess.call(['xdg-open', self.output_dir])

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        self.root.after(0, lambda: self._update_log(log_entry))

    def _update_log(self, entry):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, entry)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def save_log(self):
        log_content = self.log_text.get(1.0, tk.END)
        if not log_content.strip():
            messagebox.showinfo("Save Log", "Log is empty. Nothing to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialdir=os.path.dirname(os.path.abspath(__file__)),
            initialfile=f"pdf_conversion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )

        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(log_content)
            messagebox.showinfo("Save Log", f"Log saved to {file_path}")

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)

                if 'folder_path' in settings and os.path.exists(settings['folder_path']):
                    self.folder_path.set(settings['folder_path'])
                if 'include_subdirs' in settings:
                    self.include_subdirs.set(settings['include_subdirs'])
                if 'file_filter' in settings:
                    self.file_filter.set(settings['file_filter'])
                if 'page_size' in settings:
                    self.page_size.set(settings['page_size'])
                if 'orientation' in settings:
                    self.orientation.set(settings['orientation'])
                if 'margin_top' in settings:
                    self.margin_top.set(settings['margin_top'])
                if 'margin_right' in settings:
                    self.margin_right.set(settings['margin_right'])
                if 'margin_bottom' in settings:
                    self.margin_bottom.set(settings['margin_bottom'])
                if 'margin_left' in settings:
                    self.margin_left.set(settings['margin_left'])
                if 'worker_threads' in settings:
                    self.worker_threads.set(settings['worker_threads'])
                if 'filename_pattern' in settings:
                    self.filename_pattern.set(settings['filename_pattern'])
                if 'overwrite_files' in settings:
                    self.overwrite_files.set(settings['overwrite_files'])

                self.log_message("Settings loaded from file.")
            except Exception as e:
                self.log_message(f"Error loading settings: {str(e)}")

    def save_settings(self):
        settings = {
            'folder_path': self.folder_path.get(),
            'include_subdirs': self.include_subdirs.get(),
            'file_filter': self.file_filter.get(),
            'page_size': self.page_size.get(),
            'orientation': self.orientation.get(),
            'margin_top': self.margin_top.get(),
            'margin_right': self.margin_right.get(),
            'margin_bottom': self.margin_bottom.get(),
            'margin_left': self.margin_left.get(),
            'worker_threads': self.worker_threads.get(),
            'filename_pattern': self.filename_pattern.get(),
            'overwrite_files': self.overwrite_files.get()
        }

        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
            messagebox.showinfo("Settings", "Settings saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def start_conversion(self):
        input_folder = self.folder_path.get()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid folder")
            return

        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.NORMAL)
        self.conversion_running = True
        self.pause_conversion = False

        self.progress.config(value=0)
        self.current_file_var.set("")
        self.progress_stats_var.set("0 of 0 files (0%)")
        self.results_stats_var.set("0 succeeded, 0 failed")
        self.time_remaining_var.set("Calculating...")

        self.processed_files = []
        self.failed_files = []

        self.log_message(f"Starting conversion from {input_folder}")
        self.log_message(f"Output directory: {self.output_dir}")

        conversion_thread = threading.Thread(target=self.convert_files, args=(input_folder,))
        conversion_thread.daemon = True
        conversion_thread.start()

    def toggle_pause(self):
        if self.conversion_running:
            self.pause_conversion = not self.pause_conversion
            if self.pause_conversion:
                self.pause_btn.config(text="Resume")
                self.log_message("Conversion paused. Click Resume to continue.")
            else:
                self.pause_btn.config(text="Pause")
                self.log_message("Conversion resumed.")

    def cancel_conversion(self):
        if self.conversion_running:
            if messagebox.askyesno("Cancel Conversion", "Are you sure you want to cancel the conversion?"):
                self.conversion_running = False
                self.log_message("Conversion cancelled by user.")
                self.pause_btn.config(state=tk.DISABLED)
                self.cancel_btn.config(state=tk.DISABLED)
                self.start_btn.config(state=tk.NORMAL)

    def get_all_html_files(self, html_dir):
        html_files = []
        patterns = self.file_filter.get().split(';')

        regex_patterns = []
        for pattern in patterns:
            pattern = pattern.strip()
            if pattern:
                regex = pattern.replace('.', '\\.').replace('*', '.*')
                regex_patterns.append(f"^{regex}$")

        if regex_patterns:
            combined_regex = re.compile('|'.join(regex_patterns), re.IGNORECASE)
        else:
            combined_regex = re.compile(r"^.*\.html?$", re.IGNORECASE)

        if self.include_subdirs.get():
            for root, _, files in os.walk(html_dir):
                for file in files:
                    if combined_regex.match(file):
                        html_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(html_dir):
                file_path = os.path.join(html_dir, file)
                if os.path.isfile(file_path) and combined_regex.match(file):
                    html_files.append(file_path)

        return html_files

    def convert_files(self, html_dir):
        try:
            config = pdfkit.configuration(wkhtmltopdf=self.wkhtmltopdf_path)

            html_files = self.get_all_html_files(html_dir)

            if not html_files:
                self.root.after(0, lambda: messagebox.showinfo("Information", "No matching HTML files found in the selected folder"))
                self.root.after(0, lambda: self.log_message("No matching HTML files found"))
                self.root.after(0, lambda: self.reset_ui())
                return
            
            total_files = len(html_files)
            self.root.after(0, lambda: self.log_message(f"Found {total_files} HTML files to convert"))

            self.root.after(0, lambda: self.progress.config(maximum=total_files))

            options = {
                'page-size': self.page_size.get(),
                'orientation': self.orientation.get().lower(),
                'margin-top': f"{self.margin_top.get()}mm",
                'margin-right': f"{self.margin_right.get()}mm",
                'margin-bottom': f"{self.margin_bottom.get()}mm",
                'margin-left': f"{self.margin_left.get()}mm",
                'encoding': 'UTF-8',
                'quiet': True
            }

            start_time = time.time()
            processed_count = 0
            success_count = 0
            failure_count = 0
            
            with ThreadPoolExecutor(max_workers=self.worker_threads.get()) as executor:
                future_to_file = {executor.submit(self.convert_file, html_path, config, options): html_path for html_path in html_files}

                for future in future_to_file:
                    if not self.conversion_running:
                        executor.shutdown(wait=False, cancel_futures=True)
                        break

                    html_path = future_to_file[future]

                    self.root.after(0, lambda file=html_path: self.current_file_var.set(os.path.basename(file)))

                    while self.pause_conversion and self.conversion_running:
                        time.sleep(0.1)

                    try:
                        result = future.result()
                        if result:
                            success_count += 1
                        else:
                            failure_count += 1
                    except Exception as e:
                        self.root.after(0, lambda file=html_path, error=e: self.log_message(f"Error processing {os.path.basename(file)}: {str(error)}"))
                        failure_count += 1

                    processed_count += 1
                    self.update_progress(processed_count, total_files, success_count, failure_count)

                    if processed_count > 0:
                        elapsed_time = time.time() - start_time
                        files_per_second = processed_count / elapsed_time if elapsed_time > 0 else 0
                        remaining_files = total_files - processed_count
                        if files_per_second > 0:
                            remaining_seconds = remaining_files / files_per_second
                            mins = int(remaining_seconds // 60)
                            secs = int(remaining_seconds % 60)
                            eta = f"{mins}m {secs}s"
                            self.root.after(0, lambda t=eta: self.time_remaining_var.set(t))

            if self.conversion_running:
                final_message = f"Conversion complete! Successfully converted {success_count} of {total_files} files."
                self.root.after(0, lambda: self.log_message(final_message))
                self.root.after(0, lambda: messagebox.showinfo("Conversion Complete", final_message))
            else:
                self.root.after(0, lambda: self.log_message(f"Conversion cancelled after processing {processed_count} of {total_files} files."))

            self.root.after(0, lambda: self.reset_ui())

        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Error during conversion: {str(e)}"))
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred during conversion: {str(e)}"))
            self.root.after(0, lambda: self.reset_ui())

    def convert_file(self, html_path, config, options):
        try:
            if not self.conversion_running:
                return False

            base_filename = os.path.basename(html_path)
            name_without_ext = os.path.splitext(base_filename)[0]

            output_filename = self.filename_pattern.get()
            now = datetime.now()
            output_filename = output_filename.replace("{filename}", name_without_ext)
            output_filename = output_filename.replace("{date}", now.strftime("%Y%m%d"))
            output_filename = output_filename.replace("{time}", now.strftime("%H%M%S"))

            if not output_filename.lower().endswith(".pdf"):
                output_filename += ".pdf"

            pdf_path = os.path.join(self.output_dir, output_filename)

            if os.path.exists(pdf_path) and not self.overwrite_files.get():
                self.root.after(0, lambda: self.log_message(f"Skipped {base_filename} (output file already exists)"))
                return False

            pdfkit.from_file(html_path, pdf_path, options=options, configuration=config)

            self.root.after(0, lambda: self.log_message(f"Converted {base_filename} -> {output_filename}"))
            self.processed_files.append(pdf_path)
            return True

        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Failed to convert {os.path.basename(html_path)}: {str(e)}"))
            self.failed_files.append(html_path)
            return False

    def update_progress(self, processed_count, total_files, success_count, failure_count):

        self.root.after(0, lambda: self.progress.config(value=processed_count))

        percent = round((processed_count / total_files) * 100) if total_files > 0 else 0
        self.root.after(0, lambda: self.progress_stats_var.set(f"{processed_count} of {total_files} files ({percent}%)"))
        self.root.after(0, lambda: self.results_stats_var.set(f"{success_count} succeeded, {failure_count} failed"))

    def reset_ui(self):
        self.conversion_running = False
        self.pause_conversion = False
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.pause_btn.config(text="Pause")
        self.cancel_btn.config(state=tk.DISABLED)
        self.time_remaining_var.set("--")
        self.current_file_var.set("")

if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedPDFConverterApp(root)
    root.mainloop()
