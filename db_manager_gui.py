"""
db_manager_gui.py

A record browser/editor for equipment_master - view, add, edit, and
delete records directly, independent of the PDF-import workflow.
Launched as its own window from pvelite_gui.py ("Manage Database"
button) so the import flow and manual record management stay separate.

Deleting a record cascades to its element_data/platform_data/
weight_summation/mechanical_details/documents rows automatically (the
schema's ON DELETE CASCADE foreign keys - already tested in
db_writer.py's re-import path, reused here for manual delete).
"""
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox

# Columns shown in the list view - a readable subset, not all 50+
# equipment_master columns (the edit form below shows everything).
_LIST_COLUMNS = ["id", "project", "category", "tag_no", "equipment_name",
                  "orientation", "fabricated_mt", "imported_by"]

# Columns nobody should hand-edit - system-managed or structural.
_READONLY_COLUMNS = {"id", "imported_at"}


def _table_columns(db_path, table="equipment_master"):
    conn = sqlite3.connect(db_path)
    cols = [row[1] for row in conn.execute(f"PRAGMA table_info({table})")]
    conn.close()
    return cols


class DatabaseManagerWindow(tk.Toplevel):
    def __init__(self, parent, db_path):
        super().__init__(parent)
        self.db_path = db_path
        self.title("PV Elite Database Manager")
        self.geometry("900x500")

        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        tk.Button(top, text="Refresh", command=self.refresh).pack(side="left")
        tk.Button(top, text="Add New Record", command=self.add_record).pack(side="left", padx=6)
        tk.Button(top, text="Edit Selected", command=self.edit_selected).pack(side="left")
        tk.Button(top, text="Delete Selected", command=self.delete_selected,
                  fg="white", bg="#8B0000").pack(side="left", padx=6)

        self.tree = ttk.Treeview(self, columns=_LIST_COLUMNS, show="headings")
        for col in _LIST_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        self.refresh()

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cols_sql = ", ".join(_LIST_COLUMNS)
        for row in conn.execute(f"SELECT {cols_sql} FROM equipment_master ORDER BY project, tag_no"):
            self.tree.insert("", "end", values=[row[c] for c in _LIST_COLUMNS])
        conn.close()

    def _selected_id(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.item(sel[0])["values"][0]  # "id" is always the first list column

    def add_record(self):
        RecordForm(self, self.db_path, record_id=None, on_saved=self.refresh)

    def edit_selected(self):
        record_id = self._selected_id()
        if record_id is None:
            messagebox.showinfo("PV Elite Database Manager", "Select a record first.")
            return
        RecordForm(self, self.db_path, record_id=record_id, on_saved=self.refresh)

    def delete_selected(self):
        record_id = self._selected_id()
        if record_id is None:
            messagebox.showinfo("PV Elite Database Manager", "Select a record first.")
            return
        sel = self.tree.selection()[0]
        tag_no = self.tree.item(sel)["values"][3]
        if not messagebox.askyesno(
            "Confirm delete",
            f"Delete record for tag '{tag_no}'?\n\n"
            "This also deletes all its elements, platforms, weight "
            "breakdown, mechanical details, and document sections. "
            "This cannot be undone.",
        ):
            return
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("DELETE FROM equipment_master WHERE id = ?", (record_id,))
        conn.commit()
        conn.close()
        self.refresh()


class RecordForm(tk.Toplevel):
    """Add (record_id=None) or edit (record_id=<id>) one equipment_master row."""

    def __init__(self, parent, db_path, record_id, on_saved):
        super().__init__(parent)
        self.db_path = db_path
        self.record_id = record_id
        self.on_saved = on_saved
        self.title("Edit Record" if record_id else "Add New Record")
        self.geometry("480x600")

        self.columns = _table_columns(db_path)
        existing = self._load_existing() if record_id else {}

        canvas = tk.Canvas(self)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        form_frame = tk.Frame(canvas)
        form_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.entries = {}
        for col in self.columns:
            row = tk.Frame(form_frame)
            row.pack(fill="x", padx=8, pady=2)
            tk.Label(row, text=col, width=22, anchor="w").pack(side="left")
            var = tk.StringVar(value="" if existing.get(col) is None else str(existing.get(col)))
            state = "disabled" if col in _READONLY_COLUMNS else "normal"
            entry = tk.Entry(row, textvariable=var, state=state)
            entry.pack(side="left", fill="x", expand=True)
            self.entries[col] = var

        btn_row = tk.Frame(self)
        btn_row.pack(fill="x", padx=8, pady=8)
        tk.Button(btn_row, text="Save", command=self.save,
                  bg="#1F4E78", fg="white").pack(side="right", padx=4)
        tk.Button(btn_row, text="Cancel", command=self.destroy).pack(side="right")

    def _load_existing(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM equipment_master WHERE id = ?", (self.record_id,)).fetchone()
        conn.close()
        return dict(row) if row else {}

    def save(self):
        values = {}
        for col, var in self.entries.items():
            if col in _READONLY_COLUMNS:
                continue
            raw = var.get().strip()
            values[col] = raw if raw != "" else None

        if not values.get("project") or not values.get("category") or not values.get("tag_no"):
            messagebox.showerror("PV Elite Database Manager",
                                  "Project, Category, and Tag No are required.")
            return

        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            if self.record_id:
                set_clause = ", ".join(f"{c}=?" for c in values)
                conn.execute(
                    f"UPDATE equipment_master SET {set_clause} WHERE id=?",
                    list(values.values()) + [self.record_id],
                )
            else:
                cols = list(values.keys())
                placeholders = ", ".join(["?"] * len(cols))
                conn.execute(
                    f"INSERT INTO equipment_master ({', '.join(cols)}) VALUES ({placeholders})",
                    list(values.values()),
                )
            conn.commit()
        except sqlite3.IntegrityError as exc:
            messagebox.showerror("PV Elite Database Manager",
                                  f"Could not save - {exc}\n\n"
                                  "(A record for this Project + Tag No combination may already exist.)")
            conn.close()
            return
        conn.close()
        self.on_saved()
        self.destroy()
