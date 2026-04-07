"""
Coordinator interface for DSS (Decision Support System).
Python 3 + Tkinter + pandas + openpyxl.
"""

import tkinter as tk
import math
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Dict, List, Optional

from config import DEFAULT_ALTERNATIVES, DEFAULT_CRITERIA, MIN_DECISION_MAKERS, MAX_DECISION_MAKERS
from decision_makers import DecisionMakerWindow
from table_style import apply_excel_style


class CoordinatorApp:
    def __init__(self):
        self.sent_matrix = None
        self.root = tk.Tk()
        self.root.title("DSS - Interface Coordinateur")
        self.root.minsize(500, 400)
        self.root.configure(bg="#F5F7FA")

        # Data model (prepared for future decision maker integration)
        self.matrix: Optional[pd.DataFrame] = None
        self.matrix_structure: Optional[pd.DataFrame] = None  # base structure (alternatives × critères)
        self.weights: List[float] = []
        self.num_decision_makers = 4
        self.decisions: Dict[str, pd.DataFrame] = {}  # ex: {"D1": df, "D2": df, ...}
        self.expected_decisions = 4
        self.decision_maker_names = [
            "Politician",
            "Economist",
            "Environment representative",
            "Public representative"
        ]

        # stockage des fenêtres décideurs
        self.decision_windows = {}

        # Pie chart for the decision-makers weights
        self._weights_pie_colors = ["#4F81BD", "#C0504D", "#9BBB59", "#8064A2"]
        self._weights_pie_diameter = 120
        self._weights_pie_padding = 10
        self.weights_pie_canvas = None

        self._build_ui()


    def _build_ui(self):
        apply_excel_style()
        self._build_menubar()
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # Split layout (Parameters left, matrix right) like the reference UI
        paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        left_panel = ttk.Frame(paned)
        right_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=1)
        paned.add(right_panel, weight=3)

        # --- Frame 1: Parameters ---
        params_frame = ttk.LabelFrame(left_panel, text="Parameters", padding=8)
        params_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(params_frame, text="Introduce the weight of decision-makers.").pack(anchor="w", pady=(0, 4))
        row1 = ttk.Frame(params_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="Decision makers : 4 ").pack(side=tk.LEFT, padx=(0, 8))

        self.weights_frame = ttk.Frame(params_frame)
        self.weights_frame.pack(fill=tk.X, pady=(8, 0))
        self.weights_status_label = ttk.Label(params_frame, text="Total weight : 0%")
        self.weights_status_label.pack(anchor="w", pady=(4, 0))

        # Circle chart below the weights inputs
        self.weights_pie_canvas = tk.Canvas(
            params_frame,
            width=self._weights_pie_diameter,
            height=self._weights_pie_diameter,
            highlightthickness=0,
            bg="white",
        )
        self.weights_pie_canvas.pack(pady=(8, 4))

        self._rebuild_weight_fields()

        # --- Frame 2: Matrix (Excel-like grid) ---
        matrix_frame = ttk.LabelFrame(right_panel, text="Performance matrix", padding=8)
        matrix_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        # Bordered container like Excel
        table_border = tk.Frame(matrix_frame, bg="#D6DCE5", padx=1, pady=1)
        table_border.pack(fill=tk.BOTH, expand=True)
        tree_container = ttk.Frame(table_border)
        tree_container.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="browse", height=10)
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.bind("<Double-1>", self._on_cell_double_click)

        # --- Frame 3: Action ---
        action_frame = ttk.Frame(main)
        action_frame.pack(fill=tk.X)
        self.open_decision_makers_button = ttk.Button(
            action_frame,
            text="Decision makers",
            command=self._open_decision_makers,
            state=tk.NORMAL,
        )
        self.open_decision_makers_button.pack(side=tk.LEFT, padx=(0, 4), pady=4)

        self.send_button = ttk.Button(
            action_frame,
            text="Send",
            command=self._send_matrix,
            state=tk.DISABLED,
        )
        self.send_button.pack(side=tk.LEFT, padx=(0, 4), pady=4)

    def _build_menubar(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self._file_new)
        file_menu.add_command(label="Open", command=self._file_open)
        file_menu.add_command(label="Save", command=self._file_save)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._file_exit)

    def _on_decision_maker_count_changed(self) -> None:
        """Clear stored decisions when the number of decision makers changes."""
        self.decisions.clear()

    def _rebuild_weight_fields(self):
        for w in self.weights_frame.winfo_children():
            w.destroy()
        n = self.num_decision_makers
        self.weights = [0.0] * n
        for i in range(n):
            row = ttk.Frame(self.weights_frame)
            row.pack(fill=tk.X, pady=2)
            label_text = self.decision_maker_names[i] if i < len(self.decision_maker_names) else f"Decision maker {i + 1}"
            ttk.Label(row, text=f"{label_text} (%) :").pack(side=tk.LEFT, padx=(0, 8))
            var = tk.StringVar(value="0")
            var.trace_add("write", lambda *_, idx=i, v=var: self._on_weight_edited(idx, v))
            sb = tk.Spinbox(row, from_=0, to=100, textvariable=var, width=8)
            sb.pack(side=tk.LEFT)
            self.weights[i] = self._parse_weight(var.get())
        self._check_weights_sum()
        self._update_weights_pie_chart()

    def _on_weight_edited(self, idx: int, var: tk.StringVar):
        try:
            self.weights[idx] = self._parse_weight(var.get())
        except (ValueError, TypeError):
            pass
        self._check_weights_sum()
        self._update_weights_pie_chart()

    def _update_weights_pie_chart(self) -> None:
        """Draw the weights as a pie chart."""
        if self.weights_pie_canvas is None:
            return

        canvas = self.weights_pie_canvas
        canvas.delete("all")

        total = sum(self.weights) if self.weights else 0.0

        x0 = self._weights_pie_padding
        y0 = self._weights_pie_padding
        x1 = self._weights_pie_diameter - self._weights_pie_padding
        y1 = self._weights_pie_diameter - self._weights_pie_padding

        # background circle
        canvas.create_oval(x0, y0, x1, y1, fill="#E6E6E6", outline="#BFBFBF", width=1)
        if total <= 0.0:
            return

        cx = (x0 + x1) / 2.0
        cy = (y0 + y1) / 2.0
        radius = (x1 - x0) / 2.0

        start_angle = 90.0  # start at top
        for i, w in enumerate(self.weights):
            if w <= 0.0:
                continue

            extent = - (w / total) * 360.0
            color = self._weights_pie_colors[i % len(self._weights_pie_colors)]

            # Slice arc
            canvas.create_arc(
                x0, y0, x1, y1,
                start=start_angle,
                extent=extent,
                fill=color,
                outline="white",
                width=2,
            )

            # Place the decision-maker name inside the slice
            mid_angle = start_angle + extent / 2.0
            theta = math.radians(mid_angle)
            label_radius = radius * 0.65
            tx = cx + label_radius * math.cos(theta)
            ty = cy - label_radius * math.sin(theta)  # canvas y grows down

            label_text = self.decision_maker_names[i]
            if len(label_text) > 14 and " " in label_text:
                first, rest = label_text.split(" ", 1)
                label_text = f"{first}\n{rest}"

            canvas.create_text(
                tx, ty,
                text=label_text,
                fill="white",
                font=("Segoe UI", 8, "bold"),
                anchor="center",
            )

            start_angle += extent

    def _parse_weight(self, s: str) -> float:
        s = (s or "0").strip().replace(",", ".")
        return float(s) if s else 0.0

    def _get_weights_from_ui(self) -> List[float]:
        out = []
        for child in self.weights_frame.winfo_children():
            for w in child.winfo_children():
                if isinstance(w, tk.Spinbox):
                    try:
                        out.append(self._parse_weight(w.get()))
                    except (ValueError, TypeError):
                        out.append(0.0)
                    break
        return out

    def _validate_weights_sum_100(self) -> bool:
        """Return True if the sum of weights equals 100%."""
        weights = self._get_weights_from_ui()
        total = sum(weights)
        return abs(total - 100.0) <= 1e-6

    def _check_weights_sum(self):
        weights = self._get_weights_from_ui()
        total = sum(weights)
        if hasattr(self, "weights_status_label"):
            if abs(total - 100.0) <= 1e-6:
                self.weights_status_label.configure(
                    text=f"Total weight : {total:.2f} % (OK)",
                    foreground="green",
                )
            else:
                self.weights_status_label.configure(
                    text=f"Total weight : {total:.2f} % (must be 100%)",
                    foreground="red",
                )

    def _configure_new_matrix(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("New matrix")
        dialog.transient(self.root)
        dialog.grab_set()

        frm = ttk.Frame(dialog, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Nombre d'alternatives :").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        alt_var = tk.IntVar(value=DEFAULT_ALTERNATIVES)
        alt_spin = tk.Spinbox(frm, from_=1, to=100, textvariable=alt_var, width=5)
        alt_spin.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Nombre de critères :").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        crit_var = tk.IntVar(value=DEFAULT_CRITERIA)
        crit_spin = tk.Spinbox(frm, from_=1, to=50, textvariable=crit_var, width=5)
        crit_spin.grid(row=1, column=1, sticky="w", pady=4)

        names_frame = ttk.LabelFrame(frm, text="Noms des critères", padding=8)
        names_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

        frm.rowconfigure(2, weight=1)
        frm.columnconfigure(1, weight=1)

        crit_vars: List[tk.StringVar] = []

        def rebuild_criteria_entries(*_):
            for child in names_frame.winfo_children():
                child.destroy()
            crit_vars.clear()
            n_local = max(1, crit_var.get())
            for i_local in range(n_local):
                ttk.Label(names_frame, text=f"Critère {i_local + 1} :").grid(row=i_local, column=0, sticky="w", padx=(0, 8), pady=2)
                v_local = tk.StringVar(value=f"Critère {i_local + 1}")
                crit_vars.append(v_local)
                ttk.Entry(names_frame, textvariable=v_local, width=25).grid(row=i_local, column=1, sticky="ew", pady=2)
            names_frame.columnconfigure(1, weight=1)

        crit_var.trace_add("write", rebuild_criteria_entries)
        rebuild_criteria_entries()

        result = {"value": None}

        def on_ok():
            try:
                n_alt = int(alt_var.get())
                n_crit = int(crit_var.get())
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez saisir des nombres valides.")
                return
            if n_alt <= 0 or n_crit <= 0:
                messagebox.showerror("Erreur", "Les valeurs doivent être strictement positives.")
                return
            names = []
            for i_local in range(n_crit):
                if i_local < len(crit_vars):
                    name = crit_vars[i_local].get().strip() or f"Critère {i_local + 1}"
                else:
                    name = f"Critère {i_local + 1}"
                names.append(name)
            result["value"] = (n_alt, n_crit, names)
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(btn_frame, text="Annuler", command=on_cancel).pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side=tk.RIGHT)

        dialog.bind("<Return>", lambda e: on_ok())
        dialog.bind("<Escape>", lambda e: on_cancel())

        self.root.wait_window(dialog)
        return result["value"]

    def _file_new(self):
        config = self._configure_new_matrix()
        if not config:
            return
        n_alt, n_crit, col_names = config
        self.matrix = pd.DataFrame(
            [[1.0] * n_crit for _ in range(n_alt)],
            index=[f"Alternative {i + 1}" for i in range(n_alt)],
            columns=col_names,
        )
        # Base structure is the newly created matrix
        self.matrix_structure = self.matrix.copy(deep=True)
        self._on_matrix_modified()
        self._refresh_tree()

    def _file_open(self):
        path = filedialog.askopenfilename(
            filetypes=[("Fichiers Excel (*.xlsx)", "*.xlsx")],
            defaultextension=".xlsx",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            messagebox.showerror("Erreur", "Seuls les fichiers .xlsx sont autorisés.")
            return
        try:
            self.matrix = pd.read_excel(path, index_col=0)
            # Base structure comes from the loaded file
            self.matrix_structure = self.matrix.copy(deep=True)
            self._on_matrix_modified()
            self._refresh_tree()
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier.\n{exc}")

    def _file_save(self):
        if self.matrix is None:
            messagebox.showwarning("Enregistrer", "Aucune matrice chargée.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        try:
            self._matrix_from_tree()
            self.matrix.to_excel(path)
            messagebox.showinfo("Enregistrer", "Fichier enregistré.")
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer.\n{exc}")

    def _file_exit(self):
        self.root.quit()
        self.root.destroy()

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = []
        if self.matrix is None or self.matrix.empty:
            if hasattr(self, "send_button"):
                self.send_button.configure(state=tk.DISABLED)
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
        if hasattr(self, "send_button"):
            self.send_button.configure(state=tk.NORMAL)

    def _matrix_from_tree(self):
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
                    self.matrix.loc[idx, c] = float(
                        str(vals[j + 1]).strip().replace(",", ".")
                    )
                except (ValueError, TypeError):
                    pass
        # Any update to the base matrix invalidates stored decisions
        self._on_matrix_modified()

    def _on_cell_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        cols = self.tree["columns"]
        if col_idx >= len(cols):
            return
        col_id = cols[col_idx]
        if col_id == "_index":
            return
        vals = list(self.tree.item(item, "values"))
        if col_idx >= len(vals):
            return
        x_pos, y_pos, width, height = self.tree.bbox(item, col)
        entry = ttk.Entry(self.tree, width=width // 8)
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
            # Editing the visible matrix should clear any previously collected decisions
            self._on_matrix_modified()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)

    def _send_matrix(self):

        if self.matrix is None:
            messagebox.showwarning("Error", "No matrix.")
            return

        self._matrix_from_tree()

        # stocker la matrice
        self.sent_matrix = self.matrix.copy()

        for name, window in self.decision_windows.items():
            window.receive_matrix(self.sent_matrix)
            window.update_weight(self._get_weight_for_decision_maker(name))

        messagebox.showinfo("Success", "Matrix sent to decision makers.")

    def _on_matrix_modified(self) -> None:
        """Clear stored decisions when the base matrix content/structure changes."""
        self.decisions.clear()

    def _can_open_decision_makers(self) -> bool:
        """Check if prerequisites are met to open future decision maker windows."""
        return (
            self.matrix is not None
            and not self.matrix.empty
            and self._validate_weights_sum_100()
            and self.num_decision_makers == self.expected_decisions
        )

    def _all_decisions_received(self) -> bool:
        """Return True when all expected decision maker matrices have been collected."""
        return len(self.decisions) >= self.expected_decisions

    def _open_decision_makers(self):

        win = tk.Toplevel(self.root)
        win.title("Decision makers list")
        win.transient(self.root)

        frame = ttk.Frame(win, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        listbox = tk.Listbox(frame)
        listbox.pack(fill=tk.BOTH, expand=True)

        for name in self.decision_maker_names:
            listbox.insert(tk.END, name)

        def open_selected(event=None):
            idx = listbox.curselection()
            if not idx:
                return
            name = listbox.get(idx[0])
            if name not in self.decision_windows:
                weight = self._get_weight_for_decision_maker(name)
                window = DecisionMakerWindow(self.root, name, weight)
                self.decision_windows[name] = window
                if self.sent_matrix is not None:
                    window.receive_matrix(self.sent_matrix)
        listbox.bind("<Double-1>", open_selected)
        ttk.Button(frame, text="Open selected", command=open_selected).pack(pady=5)

    def _get_weight_for_decision_maker(self, name: str) -> float:
        weights = self._get_weights_from_ui()
        if name in self.decision_maker_names:
            i = self.decision_maker_names.index(name)
            return weights[i] if i < len(weights) else 0.0
        return 0.0

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = CoordinatorApp()
    app.run()

