import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import os
from datetime import datetime


class ExcelColumnExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Column Extractor")
        self.root.geometry("960x820")

        # {file_path: DataFrame}
        self.loaded_files = {}
        # {file_path: {column_name: tk.BooleanVar}}
        self.column_vars = {}

        self.setup_ui()

    def setup_ui(self):
        # ── File Management ──────────────────────────────────────────────────
        file_frame = tk.LabelFrame(self.root, text="File Management",
                                   font=("Arial", 12, "bold"))
        file_frame.pack(fill="x", padx=10, pady=(10, 5))

        btn_row = tk.Frame(file_frame)
        btn_row.pack(pady=8)

        tk.Button(btn_row, text="+ Add Excel File(s)",
                  command=self.add_files,
                  bg="darkgreen", fg="white", width=18).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_row, text="✕ Remove Selected",
                  command=self.remove_selected_file,
                  bg="darkred", fg="white", width=18).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_row, text="Clear All",
                  command=self.clear_all_files,
                  bg="gray", fg="white", width=12).pack(side=tk.LEFT, padx=5)

        # File list
        list_row = tk.Frame(file_frame)
        list_row.pack(fill="x", padx=10, pady=(0, 8))

        file_sb = tk.Scrollbar(list_row, orient="vertical")
        file_sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(list_row, height=4,
                                       selectmode=tk.SINGLE,
                                       yscrollcommand=file_sb.set)
        self.file_listbox.pack(side=tk.LEFT, fill="x", expand=True)
        file_sb.config(command=self.file_listbox.yview)

        # ── Column Selection Panel ───────────────────────────────────────────
        col_outer = tk.LabelFrame(self.root,
                                  text="Column Selection — All Files",
                                  font=("Arial", 12, "bold"))
        col_outer.pack(fill="both", expand=True, padx=10, pady=5)

        # Global select/deselect row
        global_row = tk.Frame(col_outer)
        global_row.pack(pady=(8, 4))

        tk.Button(global_row, text="✓ Select All (All Files)",
                  command=self.select_all_global,
                  width=22).pack(side=tk.LEFT, padx=5)

        tk.Button(global_row, text="✗ Deselect All (All Files)",
                  command=self.deselect_all_global,
                  width=22).pack(side=tk.LEFT, padx=5)

        # Scrollable canvas that holds all file sections + checkboxes
        canvas_container = tk.Frame(col_outer)
        canvas_container.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        self.col_canvas = tk.Canvas(canvas_container, borderwidth=0)
        v_sb = tk.Scrollbar(canvas_container, orient="vertical",
                            command=self.col_canvas.yview)
        h_sb = tk.Scrollbar(canvas_container, orient="horizontal",
                            command=self.col_canvas.xview)

        self.columns_frame = tk.Frame(self.col_canvas)
        self.columns_frame.bind(
            "<Configure>",
            lambda e: self.col_canvas.configure(
                scrollregion=self.col_canvas.bbox("all"))
        )
        self.col_canvas.create_window((0, 0), window=self.columns_frame,
                                      anchor="nw")
        self.col_canvas.configure(yscrollcommand=v_sb.set,
                                  xscrollcommand=h_sb.set)

        h_sb.pack(side=tk.BOTTOM, fill=tk.X)
        v_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.col_canvas.pack(side=tk.LEFT, fill="both", expand=True)

        # Mousewheel scrolling
        self.col_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.col_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units")
        )

        # Initial placeholder
        self.placeholder = tk.Label(
            self.columns_frame,
            text="Add Excel files above to see all their columns here.",
            font=("Arial", 10, "italic"), fg="gray"
        )
        self.placeholder.pack(pady=30)

        # ── Extract & Export ─────────────────────────────────────────────────
        extract_frame = tk.LabelFrame(self.root, text="Extract & Export",
                                      font=("Arial", 12, "bold"))
        extract_frame.pack(fill="x", padx=10, pady=5)

        tk.Button(extract_frame,
                  text="🗂  Extract Selected Columns to New Excel File",
                  command=self.extract_columns,
                  bg="darkblue", fg="white",
                  width=40, height=2,
                  font=("Arial", 11, "bold")).pack(pady=10)

        # ── Log ──────────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self.root, text="Log",
                                  font=("Arial", 12, "bold"))
        log_frame.pack(fill="x", padx=10, pady=(0, 10))

        log_sb = tk.Scrollbar(log_frame)
        log_sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(log_frame, height=6,
                                yscrollcommand=log_sb.set,
                                state="disabled", wrap="word")
        self.log_text.pack(side=tk.LEFT, fill="both", expand=True,
                           padx=5, pady=5)
        log_sb.config(command=self.log_text.yview)

        self.log("Ready. Add Excel files to begin.")

    # ── File Management ──────────────────────────────────────────────────────

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Excel File(s)",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm"),
                ("All files", "*.*")
            ]
        )
        if not paths:
            return

        added = 0
        for path in paths:
            if path in self.loaded_files:
                self.log(f"⚠️ Already loaded: {os.path.basename(path)}")
                continue
            try:
                df = pd.read_excel(path, engine="openpyxl")
                self.loaded_files[path] = df
                self.file_listbox.insert(tk.END, os.path.basename(path))
                self.log(f"✓ Loaded: {os.path.basename(path)} "
                         f"({len(df)} rows, {len(df.columns)} columns)")
                added += 1
            except Exception as e:
                self.log(f"✗ Failed to load {os.path.basename(path)}: {e}")

        if added:
            self.rebuild_column_panel()

    def remove_selected_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a file to remove.")
            return

        display = self.file_listbox.get(sel[0])
        path = self._path_from_display(display)

        if path:
            del self.loaded_files[path]
            self.column_vars.pop(path, None)
            self.file_listbox.delete(sel[0])
            self.log(f"✗ Removed: {display}")
            self.rebuild_column_panel()

    def clear_all_files(self):
        if not self.loaded_files:
            return
        if messagebox.askyesno("Clear All", "Remove all files and selections?"):
            self.loaded_files.clear()
            self.column_vars.clear()
            self.file_listbox.delete(0, tk.END)
            self.rebuild_column_panel()
            self.log("✓ All files cleared.")

    # ── Column Panel Builder ─────────────────────────────────────────────────

    def rebuild_column_panel(self):
        """
        Destroy and fully recreate the checkbox panel so every loaded file
        and all of its columns are visible at the same time.
        Each file gets its own labelled section with a 4-column checkbox grid.
        """
        for widget in self.columns_frame.winfo_children():
            widget.destroy()

        if not self.loaded_files:
            tk.Label(self.columns_frame,
                     text="Add Excel files above to see all their columns here.",
                     font=("Arial", 10, "italic"), fg="gray").pack(pady=30)
            return

        for file_path, df in self.loaded_files.items():
            file_name = os.path.basename(file_path)
            columns = df.columns.tolist()

            # Preserve any existing BooleanVar state; create new ones for
            # columns we haven't seen before
            if file_path not in self.column_vars:
                self.column_vars[file_path] = {}
            for col in columns:
                if col not in self.column_vars[file_path]:
                    self.column_vars[file_path][col] = tk.BooleanVar(value=False)

            # ── File header bar ──────────────────────────────────────────────
            header = tk.Frame(self.columns_frame,
                              relief="raised", bd=1, bg="#cce0ff")
            header.pack(fill="x", padx=5, pady=(12, 0))

            tk.Label(header,
                     text=f"📄  {file_name}",
                     font=("Arial", 10, "bold"),
                     bg="#cce0ff", anchor="w").pack(side=tk.LEFT, padx=8, pady=5)

            tk.Label(header,
                     text=f"{len(columns)} columns",
                     font=("Arial", 9), fg="#444", bg="#cce0ff").pack(
                         side=tk.LEFT, padx=4)

            # Per-file select/deselect buttons (right side of header)
            tk.Button(header, text="Deselect All",
                      command=lambda p=file_path: self.deselect_all_for_file(p),
                      width=11, pady=1).pack(side=tk.RIGHT, padx=5, pady=4)

            tk.Button(header, text="Select All",
                      command=lambda p=file_path: self.select_all_for_file(p),
                      width=11, pady=1).pack(side=tk.RIGHT, padx=2, pady=4)

            # ── Checkbox grid (4 columns wide) ───────────────────────────────
            grid = tk.Frame(self.columns_frame, relief="sunken", bd=1, bg="white")
            grid.pack(fill="x", padx=5, pady=(0, 2))

            GRID_COLS = 4
            for i, col in enumerate(columns):
                var = self.column_vars[file_path][col]
                cb = tk.Checkbutton(
                    grid,
                    text=str(col),
                    variable=var,
                    anchor="w",
                    wraplength=195,
                    bg="white"
                )
                cb.grid(row=i // GRID_COLS, column=i % GRID_COLS,
                        sticky="w", padx=8, pady=3)

        # Reset scroll to top after rebuild
        self.col_canvas.yview_moveto(0)

    # ── Select / Deselect Helpers ────────────────────────────────────────────

    def select_all_for_file(self, file_path):
        for var in self.column_vars.get(file_path, {}).values():
            var.set(True)

    def deselect_all_for_file(self, file_path):
        for var in self.column_vars.get(file_path, {}).values():
            var.set(False)

    def select_all_global(self):
        for file_vars in self.column_vars.values():
            for var in file_vars.values():
                var.set(True)

    def deselect_all_global(self):
        for file_vars in self.column_vars.values():
            for var in file_vars.values():
                var.set(False)

    # ── Extraction ───────────────────────────────────────────────────────────

    def extract_columns(self):
        if not self.loaded_files:
            messagebox.showwarning("No Files", "Load at least one Excel file first.")
            return

        any_selected = any(
            var.get()
            for file_vars in self.column_vars.values()
            for var in file_vars.values()
        )
        if not any_selected:
            messagebox.showwarning("No Columns Selected",
                                   "Tick at least one column checkbox first.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save Extracted Data As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"extracted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                written = 0
                for i, (file_path, df) in enumerate(self.loaded_files.items(), 1):
                    if file_path not in self.column_vars:
                        continue

                    selected = [
                        col for col, var in self.column_vars[file_path].items()
                        if var.get() and col in df.columns
                    ]

                    if not selected:
                        self.log(f"⚠️  {os.path.basename(file_path)} — "
                                 f"no columns selected, skipping.")
                        continue

                    sheet = self._make_sheet_name(
                        os.path.splitext(os.path.basename(file_path))[0], i)

                    df[selected].to_excel(writer, sheet_name=sheet, index=False)
                    self.log(f"✓  {os.path.basename(file_path)}  →  "
                             f"sheet '{sheet}'  "
                             f"({len(selected)} columns, {len(df)} rows)")
                    written += 1

            if written == 0:
                messagebox.showwarning("Nothing Exported",
                                       "No columns were selected for any file.")
            else:
                self.log(f"\n✓✓  DONE — {written} sheet(s) saved to:\n{save_path}\n")
                messagebox.showinfo("Success",
                                    f"Extraction complete!\n"
                                    f"{written} sheet(s) saved to:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Export Error", f"Export failed:\n{e}")
            self.log(f"✗  Export error: {e}")

    # ── Utilities ────────────────────────────────────────────────────────────

    def _path_from_display(self, display_name):
        for path in self.loaded_files:
            if os.path.basename(path) == display_name:
                return path
        return None

    def _make_sheet_name(self, name, number):
        for ch in [':', '\\', '/', '?', '*', '[', ']']:
            name = name.replace(ch, '_')
        return f"{name[:24]}_{number}"[:31]

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(
            tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        import pandas
        import openpyxl
    except ImportError:
        print("Install requirements: pip install pandas openpyxl")
        exit(1)

    app = ExcelColumnExtractor()
    app.run()
