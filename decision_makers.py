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
        if self.matrix is not None and not self.matrix.empty:
            confirm = messagebox.askyesno(
                "Nouveau",
                "Une matrice a déjà été reçue du coordinateur.\nVoulez-vous vraiment la remplacer ?",
                parent=self.window
            )
            if not confirm:
                return
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
            self.tree.insert("", tk.END, values=row, iid=str(idx))

        self.pref_button.configure(state=tk.NORMAL)

    # ── CORRECTION 1 : synchronise self.matrix depuis le Treeview ──────────
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

    # ── CORRECTION 2 : édition inline identique au coordinateur ────────────
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
        x_pos, y_pos, width, height = self.tree.bbox(item, col)
        entry = ttk.Entry(self.tree, width=width // 8, style="Modern.TEntry")
        entry.place(x=x_pos, y=y_pos, width=width, height=height)
        entry.insert(0, str(vals[col_idx]))
        entry.select_range(0, tk.END)
        entry.focus_set()

        def commit(_event=None):
            try:
                new_val = entry.get().strip().replace(",", ".")
                new_val = float(new_val) if new_val else 0.0
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
        self.pref_window = tk.Toplevel(self.window)
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
            columns=("Critère", "Poids", "P", "Q", "V"),
            show="headings"
        )
        vsb = ttk.Scrollbar(pref_container, orient=tk.VERTICAL, command=self.pref_tree.yview)
        hsb = ttk.Scrollbar(pref_container, orient=tk.HORIZONTAL, command=self.pref_tree.xview)
        self.pref_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        for col in ("Critère", "Poids", "P", "Q", "V"):
            self.pref_tree.column(col, width=120)

        self.pref_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.pref_tree.heading("Critère", text="Critère")
        self.pref_tree.heading("Poids", text="Poids")
        self.pref_tree.heading("P", text="P (Preference threshold)")
        self.pref_tree.heading("Q", text="Q (Indifference threshold)")
        self.pref_tree.heading("V", text="V")

        criteria_list = list(self.matrix.columns) if self.matrix is not None else self.criteria

        for crit in criteria_list:
            self.pref_tree.insert("", "end", values=(crit, "", "", "", ""))

        self.pref_tree.bind("<Double-1>", self.edit_cell)

        buttons_frame = ttk.Frame(shell, style="CardInner.TFrame")
        buttons_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        # ── CORRECTION 3 : bouton Ranger activé + handler ──────────────────
        self.ranger_button = ttk.Button(
            buttons_frame,
            text="Ranger",
            style="Accent.TButton",
            command=self._ranger,
        )
        self.ranger_button.pack(side=tk.LEFT)

    # ── CORRECTION 4 : logique du bouton Ranger ────────────────────────────
    def _ranger(self):
        """Valide les préférences saisies et affiche un classement simplifié."""

        # Synchronise self.matrix depuis le Treeview avant tout calcul
        self._matrix_from_tree_dm()

        if self.matrix is None or self.matrix.empty:
            messagebox.showwarning("Ranger", "Aucune matrice disponible.", parent=self.pref_window)
            return

        def _parse(v):
            try:
                return float(str(v).replace(",", ".")) if str(v).strip() != "" else 0.0
            except (ValueError, TypeError):
                return 0.0

        rows = []
        for item in self.pref_tree.get_children():
            vals = self.pref_tree.item(item)["values"]
            crit = str(vals[0]).strip()
            rows.append({
                "crit": crit,
                "poids": _parse(vals[1]),
                "p": _parse(vals[2]),
                "q": _parse(vals[3]),
                "v": _parse(vals[4]),
            })

        # Vérifie que les critères correspondent bien aux colonnes de la matrice
        # Matching insensible à la casse + tolérance aux noms tronqués
        matrix_cols_raw = list(self.matrix.columns)
        matrix_cols_lower = [str(c).strip().lower() for c in matrix_cols_raw]

        def find_matrix_col(crit_name: str):
            """Retourne le nom de colonne réel dans self.matrix ou None."""
            needle = crit_name.strip().lower()
            # 1) Égalité exacte
            if needle in matrix_cols_lower:
                return matrix_cols_raw[matrix_cols_lower.index(needle)]
            # 2) La colonne matrice commence par le nom du critère (tronqué)
            for raw, low in zip(matrix_cols_raw, matrix_cols_lower):
                if needle.startswith(low) or low.startswith(needle):
                    return raw
            return None

        matched = []
        for r in rows:
            col = find_matrix_col(r["crit"])
            if col is not None:
                matched.append({**r, "col": col})  # on stocke le vrai nom de colonne

        if not matched:
            messagebox.showerror(
                "Ranger",
                f"Aucun critère de préférences ne correspond aux colonnes de la matrice.\n\n"
                f"Critères préférences : {[r['crit'] for r in rows]}\n"
                f"Colonnes matrice     : {matrix_cols}",
                parent=self.pref_window,
            )
            return

        total_poids = sum(r["poids"] for r in matched)
        if total_poids <= 0:
            # Si aucun poids saisi, on les égalise automatiquement
            for r in matched:
                r["poids"] = 1.0
            total_poids = float(len(matched))

        # Calcul PROMETHEE avec fonction de préférence
        alternatives = list(self.matrix.index)
        n = len(alternatives)
        scores = {a: 0.0 for a in alternatives}

        for a in alternatives:
            for b in alternatives:
                if a == b:
                    continue
                pi_ab = 0.0
                for r in matched:
                    col = r["col"]
                    w = r["poids"] / total_poids
                    q, p = r["q"], r["p"]
                    d = float(self.matrix.loc[a, col]) - float(self.matrix.loc[b, col])
                    if d <= 0:
                        h = 0.0
                    elif (p - q) > 1e-9:
                        if d <= q:
                            h = 0.0
                        elif d >= p:
                            h = 1.0
                        else:
                            h = (d - q) / (p - q)
                    elif q > 1e-9:
                        h = 0.0 if d <= q else 1.0
                    else:
                        h = 1.0
                    pi_ab += w * h
                scores[a] += pi_ab
                scores[b] -= pi_ab

        ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)

        # Calcul de la matrice de préférence π(a,b) pour affichage
        n = len(alternatives)
        pi_matrix = {a: {b: 0.0 for b in alternatives} for a in alternatives}
        for a in alternatives:
            for b in alternatives:
                if a == b:
                    continue
                pi_ab = 0.0
                for r in matched:
                    col = r["col"]
                    w = r["poids"] / total_poids
                    q, p = r["q"], r["p"]
                    d = float(self.matrix.loc[a, col]) - float(self.matrix.loc[b, col])
                    if d <= 0:
                        h = 0.0
                    elif (p - q) > 1e-9:
                        if d <= q:
                            h = 0.0
                        elif d >= p:
                            h = 1.0
                        else:
                            h = (d - q) / (p - q)
                    elif q > 1e-9:
                        h = 0.0 if d <= q else 1.0
                    else:
                        h = 1.0
                    pi_ab += w * h
                pi_matrix[a][b] = pi_ab

        result_win = tk.Toplevel(self.pref_window)
        result_win.title(f"Classement — {self.name}")
        result_win.configure(bg=self.palette["bg"])
        result_win.geometry("420x320")
        apply_excel_style(self.current_mode)

        shell = ttk.Frame(result_win, padding=14, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        ttk.Label(shell, text="Classement des alternatives", style="HeroCompactTitle.TLabel").pack(anchor="w", pady=(0, 10))

        tree = ttk.Treeview(shell, columns=("Rang", "Alternative", "Score net"), show="headings", height=n)
        tree.heading("Rang", text="Rang")
        tree.heading("Alternative", text="Alternative")
        tree.heading("Score net", text="Score net")
        tree.column("Rang", width=60, anchor="center")
        tree.column("Alternative", width=180)
        tree.column("Score net", width=110, anchor="center")

        for rank, (alt, score) in enumerate(ranked, start=1):
            tree.insert("", tk.END, values=(rank, alt, f"{score:.4f}"))

        tree.pack(fill=tk.BOTH, expand=True)

        btn_row = ttk.Frame(shell, style="App.TFrame")
        btn_row.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(
            btn_row, text="Voir matrice π(a,b)",
            style="Secondary.TButton",
            command=lambda: self._show_pi_matrix(pi_matrix, alternatives)
        ).pack(side=tk.LEFT)

        ttk.Button(shell, text="Fermer", command=result_win.destroy, style="Secondary.TButton").pack(anchor="e", pady=(10, 0))

    def _show_pi_matrix(self, pi_matrix: dict, alternatives: list):
        """Affiche la matrice de préférence agrégée π(a,b) — entrée Phase 2 PROMETHEE."""
        win = tk.Toplevel(self.window)
        win.title(f"Matrice π(a,b) — {self.name}")
        win.configure(bg=self.palette["bg"])
        win.geometry("900x480")
        apply_excel_style(self.current_mode)

        shell = ttk.Frame(win, padding=14, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        # En-tête
        header = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(12, 8))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(header, text="Matrice de préférence agrégée π(a,b)",
                  style="HeroCompactTitle.TLabel").pack(anchor="w")
        ttk.Label(header,
                  text="π(a,b) = Σ wj · Pj(a,b)  —  entrée de la Phase 2 PROMETHEE (calcul des flux Φ+, Φ-, Φnet)",
                  style="HeroCompactSubtitle.TLabel").pack(anchor="w", pady=(2, 0))

        # Tableau
        border = tk.Frame(shell, bg=self.palette["border"], padx=1, pady=1)
        border.grid(row=1, column=0, sticky="nsew")
        container = ttk.Frame(border, style="TableWrap.TFrame")
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        cols = ["π(a,b)"] + [str(b) for b in alternatives]
        tree = ttk.Treeview(container, columns=cols, show="headings")
        vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # En-têtes colonnes
        tree.heading("π(a,b)", text="a \\ b")
        tree.column("π(a,b)", width=100, anchor="w")
        for b in alternatives:
            tree.heading(str(b), text=str(b))
            tree.column(str(b), width=90, anchor="center")

        # Remplissage lignes
        for a in alternatives:
            row_vals = [str(a)]
            for b in alternatives:
                if a == b:
                    row_vals.append("—")
                else:
                    row_vals.append(f"{pi_matrix[a][b]:.4f}")
            tree.insert("", tk.END, values=row_vals)

        # Ligne Φ+ (somme des lignes / n-1)
        n = len(alternatives)
        phi_plus_row = ["Φ+ (sortant)"]
        for b in alternatives:
            phi_p = sum(pi_matrix[a][b] for a in alternatives if a != b) / (n - 1)
            phi_plus_row.append(f"{phi_p:.4f}")
        tree.insert("", tk.END, values=phi_plus_row, tags=("phi",))

        # Ligne Φ- (somme des colonnes / n-1)
        phi_minus_row = ["Φ- (entrant)"]
        for b in alternatives:
            phi_m = sum(pi_matrix[b][a] for a in alternatives if a != b) / (n - 1)
            phi_minus_row.append(f"{phi_m:.4f}")
        tree.insert("", tk.END, values=phi_minus_row, tags=("phi",))

        tree.tag_configure("phi", background=self.palette["heading_bg"], font=("Segoe UI Semibold", 9))

        ttk.Button(shell, text="Fermer", command=win.destroy,
                   style="Secondary.TButton").grid(row=2, column=0, sticky="e", pady=(8, 0))
        for item in self.pref_tree.get_children():
            values = self.pref_tree.item(item)["values"]
            self.pref_tree.item(item, values=(values[0], "", "", "", ""))

    def _pref_new(self):
        if not hasattr(self, "pref_tree"):
            return
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
                    p_value = row.get("P", "")
                    q_value = row.get("Q", "")
                    v_value = row.get("V", "")

                    values = (
                        criterion_value,
                        "" if pd.isna(poids_value) else poids_value,
                        "" if pd.isna(p_value) else p_value,
                        "" if pd.isna(q_value) else q_value,
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

            df = pd.DataFrame(data, columns=["Critère", "Poids", "P", "Q", "V"])
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

    def get_preferences(self) -> dict:
        """Retourne les préférences saisies sous forme de dict utilisable par Promethee."""
        if not hasattr(self, "pref_tree"):
            return {}
        prefs = {}
        for item in self.pref_tree.get_children():
            vals = self.pref_tree.item(item)["values"]
            crit = vals[0]
            def _f(v):
                try:
                    return float(str(v).replace(",", ".")) if v != "" else 0.0
                except (ValueError, TypeError):
                    return 0.0
            prefs[crit] = {"poids": _f(vals[1]), "p": _f(vals[2]), "q": _f(vals[3]), "v": _f(vals[4])}
        return prefs