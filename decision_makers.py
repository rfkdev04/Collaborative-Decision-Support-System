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

        ttk.Label(frm, text="Number of alternatives :").grid(row=0, column=0, sticky="w")
        alt_var = tk.IntVar(value=DEFAULT_ALTERNATIVES)
        tk.Spinbox(frm, from_=1, to=100, textvariable=alt_var, width=5).grid(row=0, column=1)

        ttk.Label(frm, text="Number of criteria :").grid(row=1, column=0, sticky="w")
        crit_var = tk.IntVar(value=DEFAULT_CRITERIA)
        tk.Spinbox(frm, from_=1, to=50, textvariable=crit_var, width=5).grid(row=1, column=1)

        result = {"value": None}

        def on_ok():
            result["value"] = (
                int(alt_var.get()),
                int(crit_var.get()),
                [f"Criterion {i+1}" for i in range(int(crit_var.get()))]
            )
            dialog.destroy()

        ttk.Button(frm, text="OK", command=on_ok).grid(row=3, column=1)
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

        for col in cols:
            self.tree.heading(col, text=str(col))

        for idx in self.matrix.index:
            row = [str(idx)] + [self.matrix.loc[idx, c] for c in cols]
            self.tree.insert("", tk.END, values=row)

        self.pref_button.configure(state=tk.NORMAL)

    def _matrix_from_tree_dm(self):
        pass

    def _dm_on_cell_double_click(self, event):
        pass

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

        menubar = tk.Menu(self.pref_window)
        self.pref_window.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)

        file_menu.add_command(label="New", command=self._pref_new)
        file_menu.add_command(label="Open", command=self._pref_open)
        file_menu.add_command(label="Save", command=self._pref_save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.pref_window.destroy)

        frame = ttk.LabelFrame(self.pref_window, text="Preferences matrix", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        table_border = tk.Frame(frame, bg="gray65", padx=2, pady=2)
        table_border.pack(fill=tk.BOTH, expand=True)

        pref_container = ttk.Frame(table_border)
        pref_container.pack(fill=tk.BOTH, expand=True)

        self.pref_tree = ttk.Treeview(
            pref_container,
            columns=("Critère", "Poids", "Q", "P", "V"),
            show="headings"
        )

        self.pref_tree.pack(fill=tk.BOTH, expand=True)

        self.pref_tree.heading("Critère", text="Critère")
        self.pref_tree.heading("Poids", text="Poids")
        self.pref_tree.heading("Q", text="Q (Indifference threshold)")
        self.pref_tree.heading("P", text="P (Preference threshold)")
        self.pref_tree.heading("V", text="V")

        criteria_list = list(self.matrix.columns) if self.matrix is not None else self.criteria

        for crit in criteria_list:
            self.pref_tree.insert("", "end", values=(crit, "", "", "", ""))

        self.pref_tree.bind("<Double-1>", self.edit_cell)

        buttons_frame = ttk.Frame(self.pref_window)
        buttons_frame.pack(fill=tk.X, padx=8, pady=5)

        self.ranger_button = ttk.Button(
            buttons_frame,
            text="Ranger",
            state=tk.DISABLED
        )
        self.ranger_button.pack(side=tk.LEFT)

    def _pref_new(self):
        for item in self.pref_tree.get_children():
            values = self.pref_tree.item(item)["values"]
            self.pref_tree.item(item, values=(values[0], "", "", "", ""))

    def _pref_open(self):
        pass

    def _pref_save_file(self):
        pass

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