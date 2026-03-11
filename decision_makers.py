"""
Decision maker interfaces for DSS (Decision Support System).
Each decision maker has its own window: matrix view, weight, Introduce Preferences.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import List

from config import DEFAULT_ALTERNATIVES, DEFAULT_CRITERIA
from table_style import apply_excel_style


class DecisionMakerWindow:
    """Window for one decision maker: matrix display, weight, preferences table."""

    def __init__(self, root, name: str, weight: float = 0.0):
        self.criteria = [
            "Nuisances", "Bruit", "Impacts", "Géotechnique",
            "Equipements", "Accessibilité", "Climat"
        ]
        self.name = name
        self.matrix = None
        self.weight = weight

        self.window = tk.Toplevel(root)
        self.window.title(f"Decision maker : {name}")
        self.window.minsize(400, 300)
        apply_excel_style()

        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        self.dm_file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=self.dm_file_menu)
        self.dm_file_menu.add_command(label="New", command=self._dm_file_new, state=tk.DISABLED)
        self.dm_file_menu.add_command(label="Open", command=self._dm_file_open, state=tk.DISABLED)
        self.dm_file_menu.add_command(label="Save", command=self._dm_file_save, state=tk.DISABLED)
        self.dm_file_menu.add_separator()
        self.dm_file_menu.add_command(label="Exit", command=self.window.destroy)

        main = ttk.Frame(self.window, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        self.weight_label = ttk.Label(main, text=f"Weight : {weight:.1f} %")
        self.weight_label.pack(anchor="w", pady=(0, 5))
        # Excel-like bordered grid
        table_border = tk.Frame(main, bg="gray65", padx=2, pady=2)
        table_border.pack(fill=tk.BOTH, expand=True)
        tree_container = ttk.Frame(table_border)
        tree_container.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="browse")
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self._dm_on_cell_double_click)

        self.pref_button = ttk.Button(
            main, text="Introduce Preferences",
            state=tk.DISABLED, command=self._add_preferences
        )
        self.pref_button.pack(pady=5)

    def _dm_file_new(self):
        config = self._dm_configure_new_matrix()
        if not config:
            return
        n_alt, n_crit, col_names = config
        self.matrix = pd.DataFrame(
            [[1.0] * n_crit for _ in range(n_alt)],
            index=[f"Alternative {i + 1}" for i in range(n_alt)],
            columns=col_names,
        )
        self._refresh_matrix_tree()

    def _dm_file_open(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files (*.xlsx)", "*.xlsx")],
            defaultextension=".xlsx", parent=self.window,
        )
        if not path:
            return
        try:
            self.matrix = pd.read_excel(path, index_col=0)
            self._refresh_matrix_tree()
            messagebox.showinfo("Open", "Matrix loaded.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot open file.\n{exc}", parent=self.window)

    def _dm_file_save(self):
        if self.matrix is None or self.matrix.empty:
            messagebox.showwarning("Save", "No matrix to save.", parent=self.window)
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")], parent=self.window,
        )
        if not path:
            return
        try:
            self._matrix_from_tree_dm()
            self.matrix.to_excel(path)
            messagebox.showinfo("Save", "File saved.", parent=self.window)
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot save.\n{exc}", parent=self.window)

    def _dm_configure_new_matrix(self):
        dialog = tk.Toplevel(self.window)
        dialog.title("New matrix")
        dialog.transient(self.window)
        dialog.grab_set()
        frm = ttk.Frame(dialog, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Number of alternatives :").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        alt_var = tk.IntVar(value=DEFAULT_ALTERNATIVES)
        tk.Spinbox(frm, from_=1, to=100, textvariable=alt_var, width=5).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(frm, text="Number of criteria :").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        crit_var = tk.IntVar(value=DEFAULT_CRITERIA)
        tk.Spinbox(frm, from_=1, to=50, textvariable=crit_var, width=5).grid(row=1, column=1, sticky="w", pady=4)
        names_frame = ttk.LabelFrame(frm, text="Criteria names", padding=8)
        names_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        frm.rowconfigure(2, weight=1)
        frm.columnconfigure(1, weight=1)
        crit_vars: List[tk.StringVar] = []

        def rebuild_entries(*_):
            for child in names_frame.winfo_children():
                child.destroy()
            crit_vars.clear()
            for i in range(max(1, crit_var.get())):
                ttk.Label(names_frame, text=f"Criterion {i + 1} :").grid(row=i, column=0, sticky="w", padx=(0, 8), pady=2)
                v = tk.StringVar(value=f"Criterion {i + 1}")
                crit_vars.append(v)
                ttk.Entry(names_frame, textvariable=v, width=25).grid(row=i, column=1, sticky="ew", pady=2)
            names_frame.columnconfigure(1, weight=1)
        crit_var.trace_add("write", rebuild_entries)
        rebuild_entries()
        result = {"value": None}

        def on_ok():
            try:
                n_alt = int(alt_var.get())
                n_crit = int(crit_var.get())
            except ValueError:
                messagebox.showerror("Error", "Enter valid numbers.", parent=dialog)
                return
            if n_alt <= 0 or n_crit <= 0:
                messagebox.showerror("Error", "Values must be positive.", parent=dialog)
                return
            names = [crit_vars[i].get().strip() or f"Criterion {i + 1}" for i in range(n_crit)]
            result["value"] = (n_alt, n_crit, names)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side=tk.RIGHT)
        dialog.bind("<Return>", lambda e: on_ok())
        dialog.bind("<Escape>", lambda e: on_cancel())
        self.window.wait_window(dialog)
        return result["value"]

    def _refresh_matrix_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.matrix is None or self.matrix.empty:
            self.pref_button.configure(state=tk.DISABLED)
            return
        cols = list(self.matrix.columns)
        self.tree["columns"] = ["_index"] + cols
        self.tree.heading("_index", text="Alternative")
        self.tree.column("_index", width=120)
        for col in cols:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=80)
        for idx in self.matrix.index:
            row = [str(idx)] + [self.matrix.loc[idx, c] for c in cols]
            self.tree.insert("", tk.END, values=row, iid=str(idx))
        self.pref_button.configure(state=tk.NORMAL)

    def _matrix_from_tree_dm(self):
        if self.matrix is None:
            return
        cols = list(self.matrix.columns)
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if len(vals) != len(cols) + 1:
                continue
            idx = vals[0]
            if idx not in self.matrix.index:
                continue
            for j, c in enumerate(cols):
                try:
                    self.matrix.loc[idx, c] = float(str(vals[j + 1]).strip().replace(",", "."))
                except (ValueError, TypeError):
                    pass

    def _dm_on_cell_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        cols = self.tree["columns"]
        if col_idx >= len(cols) or cols[col_idx] == "_index":
            return
        vals = list(self.tree.item(item, "values"))
        if col_idx >= len(vals):
            return
        x_pos, y_pos, width, height = self.tree.bbox(item, col)
        entry = ttk.Entry(self.tree, width=max(1, width // 8))
        entry.place(x=x_pos, y=y_pos, width=width, height=height)
        entry.insert(0, str(vals[col_idx]))
        entry.select_range(0, tk.END)
        entry.focus_set()

        def commit(_event=None):
            try:
                new_val = float(str(entry.get()).strip().replace(",", "."))
            except ValueError:
                new_val = vals[col_idx]
            vals[col_idx] = new_val
            self.tree.item(item, values=vals)
            entry.destroy()
        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    def receive_matrix(self, matrix):
        self.matrix = matrix.copy()
        self._refresh_matrix_tree()
        for i in (0, 1, 2):
            self.dm_file_menu.entryconfig(i, state=tk.NORMAL)

    def update_weight(self, weight: float) -> None:
        self.weight = weight
        if hasattr(self, "weight_label"):
            self.weight_label.configure(text=f"Weight : {weight:.1f} %")

    def _add_preferences(self):
        root = self.window.master
        self.pref_window = tk.Toplevel(root)
        self.pref_window.title(f"Preferences - {self.name}")
        self.pref_window.geometry("700x450")
        self.pref_window.minsize(500, 350)

        menubar = tk.Menu(self.pref_window)
        self.pref_window.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self._pref_new)
        file_menu.add_command(label="Open", command=self._pref_open)
        file_menu.add_command(label="Save", command=self._pref_save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.pref_window.destroy)

        frame = ttk.LabelFrame(self.pref_window, text="Preferences matrix (Critère, Poids, Q, P, V)", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        # Excel-like bordered grid
        table_border = tk.Frame(frame, bg="gray65", padx=2, pady=2)
        table_border.pack(fill=tk.BOTH, expand=True)
        pref_container = ttk.Frame(table_border)
        pref_container.pack(fill=tk.BOTH, expand=True)
        self.pref_tree = ttk.Treeview(
            pref_container,
            columns=("Critère", "Poids", "Q", "P", "V"),
            show="headings", height=14
        )
        pref_vsb = ttk.Scrollbar(pref_container, orient=tk.VERTICAL, command=self.pref_tree.yview)
        pref_hsb = ttk.Scrollbar(pref_container, orient=tk.HORIZONTAL, command=self.pref_tree.xview)
        self.pref_tree.configure(yscrollcommand=pref_vsb.set, xscrollcommand=pref_hsb.set)
        self.pref_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pref_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        pref_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.pref_tree.heading("Critère", text="Critère")
        self.pref_tree.heading("Poids", text="Poids")
        self.pref_tree.heading("Q", text="Q (Indifference threshold)")
        self.pref_tree.heading("P", text="P (Preference threshold)")
        self.pref_tree.heading("V", text="V")
        for col in ("Critère", "Poids", "Q", "P", "V"):
            self.pref_tree.column(col, width=120, anchor="center")
        criteria_list = list(self.matrix.columns) if self.matrix is not None and not self.matrix.empty else self.criteria
        for crit in criteria_list:
            self.pref_tree.insert("", "end", values=(crit, "", "", "", ""))
        self.pref_tree.bind("<Double-1>", self.edit_cell)

    def _pref_new(self):
        for item in self.pref_tree.get_children():
            values = self.pref_tree.item(item)["values"]
            self.pref_tree.item(item, values=(values[0], "", "", "", ""))

    def _pref_open(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files (*.xlsx)", "*.xlsx")],
            defaultextension=".xlsx",
            parent=getattr(self, "pref_window", self.window),
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            for item in self.pref_tree.get_children():
                self.pref_tree.delete(item)
            for _, row in df.iterrows():
                self.pref_tree.insert("", "end", values=list(row))
        except Exception as exc:
            messagebox.showerror("Error", f"Cannot open file.\n{exc}", parent=getattr(self, "pref_window", self.window))

    def _pref_save_file(self):
        data = []
        for item in self.pref_tree.get_children():
            data.append(self.pref_tree.item(item)["values"])
        df = pd.DataFrame(data, columns=["Critère", "Poids", "Q", "P", "V"])
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files (*.xlsx)", "*.xlsx")],
            parent=getattr(self, "pref_window", self.window),
        )
        if file_path:
            df.to_excel(file_path, index=False)

    def edit_cell(self, event):
        item = self.pref_tree.identify_row(event.y)
        column = self.pref_tree.identify_column(event.x)
        if not item:
            return
        col_index = int(column.replace("#", "")) - 1
        if col_index == 0:
            return
        x, y, width, height = self.pref_tree.bbox(item, column)
        entry = tk.Entry(self.pref_tree)
        entry.place(x=x, y=y, width=width, height=height)
        current_value = self.pref_tree.item(item)["values"][col_index]
        entry.insert(0, current_value)
        entry.focus()

        def save_value(event):
            values = list(self.pref_tree.item(item, "values"))
            values[col_index] = entry.get()
            self.pref_tree.item(item, values=values)
            entry.destroy()

        entry.bind("<Return>", save_value)
        entry.bind("<FocusOut>", lambda e: entry.destroy())
