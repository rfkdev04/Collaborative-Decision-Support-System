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

    def __init__(self, root, name: str, weight: float = 0.0, mode: str = "dark"):
        self.criteria = [
            "Nuisances", "Bruit", "Impacts", "Géotechnique",
            "Equipements", "Accessibilité", "Climat"
        ]
        self.name = name
        self.matrix = None
        self.weight = weight
        self.current_mode = mode
        self.palette = apply_excel_style(self.current_mode)

        self.window = tk.Toplevel(root)
        self.window.configure(bg=self.palette["bg"])
        self.window.title(f"Decision maker : {name}")
        self.window.geometry("980x620")
        self.window.minsize(760, 500)

        menubar = tk.Menu(self.window)
        self.window.config(menu=menubar)
        self.dm_file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=self.dm_file_menu)
        self.dm_file_menu.add_command(label="New", command=self._dm_file_new, state=tk.DISABLED)
        self.dm_file_menu.add_command(label="Open", command=self._dm_file_open, state=tk.DISABLED)
        self.dm_file_menu.add_command(label="Save", command=self._dm_file_save, state=tk.DISABLED)
        self.dm_file_menu.add_separator()
        self.dm_file_menu.add_command(label="Exit", command=self.window.destroy)

        shell = ttk.Frame(self.window, padding=10, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(14, 10))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=f"Decision maker — {name}", style="HeroCompactTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Review the matrix, check your assigned weight, and manage your preferences.",
            style="HeroCompactSubtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.mode_button = ttk.Button(header, text="Mode clair", style="Secondary.TButton", command=self._toggle_mode)
        self.mode_button.grid(row=0, column=1, rowspan=2, sticky="e")

        main = ttk.Frame(shell, style="Card.TFrame", padding=10)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        topbar = ttk.Frame(main, style="CardInner.TFrame")
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        topbar.columnconfigure(0, weight=1)

        left_info = ttk.Frame(topbar, style="CardInner.TFrame")
        left_info.grid(row=0, column=0, sticky="w")
        ttk.Label(left_info, text="Assigned weight", style="MutedCapsCompact.TLabel").pack(anchor="w")
        self.weight_label = ttk.Label(left_info, text=f"Weight : {weight:.1f} %", style="MetricCompact.TLabel")
        self.weight_label.pack(anchor="w", pady=(1, 0))

        self.pref_button = ttk.Button(
            topbar,
            text="Introduce Preferences",
            state=tk.DISABLED,
            command=self._add_preferences,
            style="Accent.TButton",
        )
        self.pref_button.grid(row=0, column=1, sticky="e")

        self.table_border = tk.Frame(main, bg=self.palette["border"], padx=1, pady=1)
        self.table_border.grid(row=1, column=0, sticky="nsew")

        tree_container = ttk.Frame(self.table_border, style="TableWrap.TFrame")
        tree_container.pack(fill=tk.BOTH, expand=True)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._dm_on_cell_double_click)

    def _toggle_mode(self):
        self.apply_mode("light" if self.current_mode == "dark" else "dark")

    def apply_mode(self, mode: str):
        self.current_mode = mode
        self.palette = apply_excel_style(self.current_mode)
        self.window.configure(bg=self.palette["bg"])
        self.table_border.configure(bg=self.palette["border"])
        if hasattr(self, "pref_window") and self.pref_window.winfo_exists():
            self.pref_window.configure(bg=self.palette["bg"])
            if hasattr(self, "pref_table_border"):
                self.pref_table_border.configure(bg=self.palette["border"])
        self.mode_button.configure(text="Mode clair" if self.current_mode == "dark" else "Mode sombre")

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
        dialog.configure(bg=self.palette["bg"])
        apply_excel_style(self.current_mode)

        frm = ttk.Frame(dialog, padding=12, style="Dialog.TFrame")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Number of alternatives :", style="DialogLabel.TLabel").grid(row=0, column=0, sticky="w")
        alt_var = tk.IntVar(value=DEFAULT_ALTERNATIVES)
        tk.Spinbox(frm, from_=1, to=100, textvariable=alt_var, width=5).grid(row=0, column=1)

        ttk.Label(frm, text="Number of criteria :", style="DialogLabel.TLabel").grid(row=1, column=0, sticky="w")
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

        ttk.Button(frm, text="OK", command=on_ok, style="Accent.TButton").grid(row=3, column=1, pady=(10, 0), sticky="e")
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
        self.tree.column("_index", width=150)

        for col in cols:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=110)

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
        self.pref_window.configure(bg=self.palette["bg"])
        self.pref_window.title(f"Preferences - {self.name}")
        self.pref_window.geometry("980x620")
        self.pref_window.minsize(760, 500)
        apply_excel_style(self.current_mode)

        menubar = tk.Menu(self.pref_window)
        self.pref_window.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)

        file_menu.add_command(label="New", command=self._pref_new)
        file_menu.add_command(label="Open", command=self._pref_open)
        file_menu.add_command(label="Save", command=self._pref_save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.pref_window.destroy)

        shell = ttk.Frame(self.pref_window, padding=10, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(14, 10))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text=f"Preferences — {self.name}", style="HeroCompactTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Set weights and thresholds in a compact editable grid.", style="HeroCompactSubtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 0))

        frame = ttk.LabelFrame(shell, text="Preferences matrix", padding=10, style="Card.TLabelframe")
        frame.grid(row=1, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.pref_table_border = tk.Frame(frame, bg=self.palette["border"], padx=1, pady=1)
        self.pref_table_border.grid(row=0, column=0, sticky="nsew")

        pref_container = ttk.Frame(self.pref_table_border, style="TableWrap.TFrame")
        pref_container.pack(fill=tk.BOTH, expand=True)
        pref_container.columnconfigure(0, weight=1)
        pref_container.rowconfigure(0, weight=1)

        self.pref_tree = ttk.Treeview(
            pref_container,
            columns=("Critère", "Poids", "Q", "P", "V"),
            show="headings"
        )
        vsb = ttk.Scrollbar(pref_container, orient=tk.VERTICAL, command=self.pref_tree.yview)
        hsb = ttk.Scrollbar(pref_container, orient=tk.HORIZONTAL, command=self.pref_tree.xview)
        self.pref_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in ("Critère", "Poids", "Q", "P", "V"):
            self.pref_tree.column(col, width=120)

        self.pref_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.pref_tree.heading("Critère", text="Critère")
        self.pref_tree.heading("Poids", text="Poids")
        self.pref_tree.heading("Q", text="Q (Indifference threshold)")
        self.pref_tree.heading("P", text="P (Preference threshold)")
        self.pref_tree.heading("V", text="V")

        criteria_list = list(self.matrix.columns) if self.matrix is not None else self.criteria

        for crit in criteria_list:
            self.pref_tree.insert("", "end", values=(crit, "", "", "", ""))

        self.pref_tree.bind("<Double-1>", self.edit_cell)

        buttons_frame = ttk.Frame(shell, style="CardInner.TFrame")
        buttons_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        self.ranger_button = ttk.Button(
            buttons_frame,
            text="Ranger",
            state=tk.DISABLED,
            style="Secondary.TButton",
        )
        self.ranger_button.pack(side=tk.LEFT)

    def _pref_new(self):
        for item in self.pref_tree.get_children():
            values = self.pref_tree.item(item)["values"]
            self.pref_tree.item(item, values=(values[0], "", "", "", ""))

    def _pref_open(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files (*.xlsx)", "*.xlsx")],
            defaultextension=".xlsx",
            parent=self.pref_window
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)

            for i, item in enumerate(self.pref_tree.get_children()):
                current_values = list(self.pref_tree.item(item)["values"])
                if not current_values:
                    current_values = ["", "", "", "", ""]

                if i < len(df):
                    row = df.iloc[i]

                    criterion_value = row.get("Critère", current_values[0])
                    if pd.isna(criterion_value) or str(criterion_value).strip() == "":
                        criterion_value = current_values[0]

                    poids_value = row.get("Poids", "")
                    q_value = row.get("Q", "")
                    p_value = row.get("P", "")
                    v_value = row.get("V", "")

                    values = (
                        criterion_value,
                        "" if pd.isna(poids_value) else poids_value,
                        "" if pd.isna(q_value) else q_value,
                        "" if pd.isna(p_value) else p_value,
                        "" if pd.isna(v_value) else v_value,
                    )
                    self.pref_tree.item(item, values=values)

            messagebox.showinfo("Open", "Preferences loaded.", parent=self.pref_window)

        except Exception as exc:
            messagebox.showerror("Error", f"Cannot open file.\n{exc}", parent=self.pref_window)

    def _pref_save_file(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            parent=self.pref_window
        )
        if not path:
            return

        try:
            data = []
            for item in self.pref_tree.get_children():
                values = self.pref_tree.item(item)["values"]
                data.append(values)

            df = pd.DataFrame(data, columns=["Critère", "Poids", "Q", "P", "V"])
            df.to_excel(path, index=False)

            messagebox.showinfo("Save", "Preferences saved.", parent=self.pref_window)

        except Exception as exc:
            messagebox.showerror("Error", f"Cannot save.\n{exc}", parent=self.pref_window)

    def edit_cell(self, event):
        item = self.pref_tree.identify_row(event.y)
        column = self.pref_tree.identify_column(event.x)
        if not item:
            return

        col_index = int(column.replace("#", "")) - 1
        if col_index == 0:
            return

        x, y, width, height = self.pref_tree.bbox(item, column)

        entry = tk.Entry(
            self.pref_tree,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#CFD8EA",
            highlightcolor=self.palette["primary"],
            bg="#FFFFFF",
            fg="#0F172A",
            font=("Segoe UI", 9),
        )
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
