"""
PV Elite Extractor - Desktop GUI

A thin Tkinter front end over the existing extraction pipeline
(pvelite_dynamic_extractor.py + bulk_extract.py). No extraction logic
lives here - this only adds folder/file pickers, a progress log, and
error handling so the tool can be used without a command line.

Packaging (run on a WINDOWS machine, not this build environment):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name PVElite_Extractor ^
        --hidden-import=pymupdf --hidden-import=openpyxl --hidden-import=pandas ^
        pvelite_gui.py

This produces dist\\PVElite_Extractor.exe - a single file that can be
copied to any Windows machine and double-clicked, no Python required.
"""
import os
import sys
import glob
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# bulk_extract.py contains all the actual extraction logic, unchanged -
# this GUI only calls it and displays progress.
import bulk_extract
import db_writer
from db_manager_gui import DatabaseManagerWindow


def scan_projects(input_root):
    """
    Scans an Input/ folder shaped like:
        Input/{Project}/{Category}/*.pdf
    Returns {project_name: [category_name, ...]} for every project
    folder that has at least one category subfolder.
    """
    projects = {}
    if not os.path.isdir(input_root):
        return projects
    for project_name in sorted(os.listdir(input_root)):
        project_path = os.path.join(input_root, project_name)
        if not os.path.isdir(project_path):
            continue
        categories = sorted(
            d for d in os.listdir(project_path)
            if os.path.isdir(os.path.join(project_path, d))
        )
        if categories:
            projects[project_name] = categories
    return projects


# Categories known to work reliably (vertical, skirt-supported vessels -
# what this pipeline has actually been built and validated against).
# Anything else is still selectable, just labeled as unvalidated so
# nobody mistakes incomplete results for a finished extraction.
_VALIDATED_CATEGORIES = {"vessels_vert", "columns"}


class PVEliteExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("PV Elite Extractor")
        root.geometry("720x520")
        root.resizable(True, True)

        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()
        self.save_to_db = tk.BooleanVar(value=False)
        self.db_path = tk.StringVar(value="pvelite_data.db")
        self.check_vars = {}  # (project, category) -> tk.BooleanVar

        pad = {"padx": 10, "pady": 6}

        # --- Input root folder row ---
        frame1 = tk.Frame(root)
        frame1.pack(fill="x", **pad)
        tk.Label(frame1, text="Input folder:", width=14, anchor="w").pack(side="left")
        tk.Entry(frame1, textvariable=self.input_dir).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame1, text="Browse...", command=self.pick_input_dir).pack(side="left")

        # --- Project/Category tree (checkboxes) ---
        tk.Label(root, text="Select project(s) and categories to import:", anchor="w").pack(fill="x", padx=10)
        tree_frame = tk.Frame(root, relief="sunken", borderwidth=1)
        tree_frame.pack(fill="both", expand=False, padx=10, pady=(0, 6))
        tree_canvas = tk.Canvas(tree_frame, height=180)
        tree_vscroll = tk.Scrollbar(tree_frame, orient="vertical", command=tree_canvas.yview)
        tree_hscroll = tk.Scrollbar(tree_frame, orient="horizontal", command=tree_canvas.xview)
        self.tree_inner = tk.Frame(tree_canvas)
        self.tree_inner.bind("<Configure>", lambda e: tree_canvas.configure(scrollregion=tree_canvas.bbox("all")))
        tree_canvas.create_window((0, 0), window=self.tree_inner, anchor="nw")
        tree_canvas.configure(yscrollcommand=tree_vscroll.set, xscrollcommand=tree_hscroll.set)
        tree_canvas.grid(row=0, column=0, sticky="nsew")
        tree_vscroll.grid(row=0, column=1, sticky="ns")
        tree_hscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # --- Output file row ---
        frame2 = tk.Frame(root)
        frame2.pack(fill="x", **pad)
        tk.Label(frame2, text="Output file:", width=14, anchor="w").pack(side="left")
        tk.Entry(frame2, textvariable=self.output_file).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame2, text="Browse...", command=self.pick_output_file).pack(side="left")

        # --- Database row ---
        frame_db = tk.Frame(root)
        frame_db.pack(fill="x", **pad)
        tk.Checkbutton(frame_db, text="Also save to database:", variable=self.save_to_db,
                        width=20, anchor="w").pack(side="left")
        tk.Entry(frame_db, textvariable=self.db_path).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(frame_db, text="Browse...", command=self.pick_db_path).pack(side="left")
        tk.Button(frame_db, text="Manage Database...", command=self.open_db_manager).pack(side="left", padx=(6, 0))
        tk.Button(frame_db, text="Export DB to Excel...", command=self.export_db_to_excel).pack(side="left", padx=(6, 0))

        # --- Run button + progress bar ---
        frame3 = tk.Frame(root)
        frame3.pack(fill="x", **pad)
        self.run_button = tk.Button(frame3, text="Run Extraction", command=self.start_run,
                                     bg="#1F4E78", fg="white", font=("Segoe UI", 10, "bold"))
        self.run_button.pack(side="left")
        self.progress = ttk.Progressbar(frame3, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # --- Log output ---
        tk.Label(root, text="Log:", anchor="w").pack(fill="x", padx=10)
        log_frame = tk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", bg="#F5F5F5")
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- Bottom buttons ---
        frame4 = tk.Frame(root)
        frame4.pack(fill="x", padx=10, pady=(0, 10))
        self.open_output_button = tk.Button(frame4, text="Open Output File", command=self.open_output,
                                             state="disabled")
        self.open_output_button.pack(side="right")

    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    def pick_input_dir(self):
        path = filedialog.askdirectory(title="Select the Input folder (contains Project subfolders)")
        if path:
            self.input_dir.set(path)
            if not self.output_file.get():
                self.output_file.set(os.path.join(path, "Master_Element_Data.xlsx"))
            self.populate_tree(path)

    def pick_db_path(self):
        path = filedialog.asksaveasfilename(
            title="Database file (existing or new)", defaultextension=".db",
            filetypes=[("SQLite database", "*.db")], initialfile="pvelite_data.db",
        )
        if path:
            self.db_path.set(path)

    def populate_tree(self, input_root):
        for widget in self.tree_inner.winfo_children():
            widget.destroy()
        self.check_vars.clear()

        projects = scan_projects(input_root)
        if not projects:
            tk.Label(self.tree_inner, text="No Project/Category subfolders found here.",
                      fg="gray").pack(anchor="w", padx=5)
            return

        # Each project is its own vertical column: a bold project-name
        # header, then its categories listed as checkboxes stacked
        # underneath (not laid out horizontally next to the header) -
        # columns sit side by side left to right.
        for project_name, categories in projects.items():
            col = tk.Frame(self.tree_inner, relief="groove", borderwidth=1)
            col.pack(side="left", fill="y", anchor="n", padx=4, pady=4)

            tk.Label(col, text=project_name, font=("Segoe UI", 9, "bold"),
                     bg="#E8EEF4", anchor="w").pack(fill="x", padx=4, pady=(2, 4))

            for cat in categories:
                var = tk.BooleanVar(value=True)
                self.check_vars[(project_name, cat)] = var
                label = cat
                if cat.lower().replace(" ", "_") not in _VALIDATED_CATEGORIES:
                    label += "  (unvalidated)"
                cb = tk.Checkbutton(col, text=label, variable=var, anchor="w", justify="left")
                cb.pack(fill="x", anchor="w", padx=4)

    def selected_pdf_entries(self):
        """Builds the (path, project, category) list from checked boxes."""
        input_root = self.input_dir.get().strip()
        entries = []
        for (project, category), var in self.check_vars.items():
            if not var.get():
                continue
            folder = os.path.join(input_root, project, category)
            for pdf_path in sorted(glob.glob(os.path.join(folder, "*.pdf"))):
                entries.append({"path": pdf_path, "project": project, "category": category})
        return entries

    def pick_output_file(self):
        path = filedialog.asksaveasfilename(
            title="Save output as", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile="Master_Element_Data.xlsx",
        )
        if path:
            self.output_file.set(path)

    def start_run(self):
        output_file = self.output_file.get().strip()
        entries = self.selected_pdf_entries()

        if not output_file:
            messagebox.showerror("PV Elite Extractor", "Please choose where to save the output file.")
            return
        if not entries:
            messagebox.showwarning("PV Elite Extractor",
                                    "No PDFs found under the checked project/category folders.")
            return

        db_config = None
        if self.save_to_db.get():
            db_path = self.db_path.get().strip()
            if not db_path:
                messagebox.showerror("PV Elite Extractor", "Please choose a database file path.")
                return
            db_config = db_writer.DBConfig(db_path)
            try:
                db_writer.ensure_schema(db_config)
            except Exception as exc:
                messagebox.showerror("PV Elite Extractor", f"Could not set up the database:\n{exc}")
                return

        self.run_button.configure(state="disabled")
        self.open_output_button.configure(state="disabled")
        self.progress.start(12)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.log(f"Found {len(entries)} PDF file(s) across selected projects/categories")
        self.log("Starting extraction...\n")

        thread = threading.Thread(
            target=self._run_extraction, args=(entries, output_file, db_config), daemon=True)
        thread.start()

    def _run_extraction(self, entries, output_file, db_config):
        try:
            imported_by = os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"
            bulk_extract.build_workbook(entries, output_file, log=self._threadsafe_log,
                                         db_config=db_config, imported_by=imported_by)
            self.root.after(0, self._on_success, output_file)
        except Exception:
            err = traceback.format_exc()
            self.root.after(0, self._on_failure, err)

    def _threadsafe_log(self, message):
        # bulk_extract.main() runs on a background thread; tkinter widgets
        # must only be touched from the main thread, so hand the log line
        # back via .after() rather than calling self.log() directly here.
        self.root.after(0, self.log, str(message))

    def _on_success(self, output_file):
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.open_output_button.configure(state="normal")
        self.log(f"\nDone. Saved to: {output_file}")
        messagebox.showinfo("PV Elite Extractor", "Extraction complete.")

    def _on_failure(self, error_text):
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.log("\nERROR:\n" + error_text)
        messagebox.showerror("PV Elite Extractor",
                              "Extraction failed. See the log panel for details.")

    def open_output(self):
        path = self.output_file.get()
        if path and os.path.isfile(path):
            os.startfile(path)  # Windows-only, matches deployment target

    def open_db_manager(self):
        db_path = self.db_path.get().strip()
        if not db_path:
            messagebox.showerror("PV Elite Extractor", "Please choose a database file path first.")
            return
        db_config = db_writer.DBConfig(db_path)
        try:
            db_writer.ensure_schema(db_config)
        except Exception as exc:
            messagebox.showerror("PV Elite Extractor", f"Could not open the database:\n{exc}")
            return
        DatabaseManagerWindow(self.root, db_path)

    def export_db_to_excel(self):
        db_path = self.db_path.get().strip()
        if not db_path or not os.path.isfile(db_path):
            messagebox.showerror("PV Elite Extractor", "Please select an existing database file first.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Export database as", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")], initialfile="Master_Element_Data_export.xlsx",
        )
        if not out_path:
            return
        try:
            import db_export
            ok = db_export.export_to_excel(db_path, out_path, log=self.log)
            if ok:
                messagebox.showinfo("PV Elite Extractor", f"Exported to:\n{out_path}")
            else:
                messagebox.showwarning("PV Elite Extractor", "Database has no records to export.")
        except Exception as exc:
            messagebox.showerror("PV Elite Extractor", f"Export failed:\n{exc}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PVEliteExtractorApp(root)
    root.mainloop()
