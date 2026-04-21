"""
Coordinator interface for DSS (Decision Support System).
Python 3 + Tkinter + pandas + openpyxl.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Dict, List, Optional

from config import DEFAULT_ALTERNATIVES, DEFAULT_CRITERIA
from decision_makers import DecisionMakerWindow
from table_style import apply_excel_style
from promethee import Promethee

class CoordinatorApp:
    def __init__(self):
        self.sent_matrix = None
        self.current_mode = "dark"
        self.palette = None
        self.root = tk.Tk()
        self.root.title("DSS - Interface Coordinateur")
        self.root.geometry("1180x700")
        self.root.minsize(900, 560)

        self.matrix: Optional[pd.DataFrame] = None
        self.matrix_structure: Optional[pd.DataFrame] = None
        self.weights: List[float] = []
        self.num_decision_makers = 4
        self.expected_decisions = 4
        self.decision_maker_names = [
            "Politician",
            "Economist",
            "Environment representative",
            "Public representative"
        ]

        self.decision_windows: Dict[str, DecisionMakerWindow] = {}
        self.weight_vars: List[tk.StringVar] = []
        self.weight_spinboxes: List[tk.Spinbox] = []
        self.legend_rows: List[tuple] = []
        self._building_weights = False

        self._weights_pie_colors = ["#054A91", "#3E7CB1", "#81A4CD", "#DBE4EE"]
        self._weights_pie_diameter = 112
        self._weights_pie_padding = 10
        self.weights_pie_canvas = None
        self.mode_button = None
        self.weight_canvas = None
        self.weight_scrollbar = None
        self.weight_scroll_inner = None
        self.weight_canvas_window = None
        self.table_border = None
        self.send_button = None
        self.promethee_result = None

        self._build_ui()
        self._apply_mode(self.current_mode)

    def _build_ui(self):
        self.palette = apply_excel_style(self.current_mode)
        self.root.configure(bg=self.palette["bg"])
        self._build_menubar()

        shell = ttk.Frame(self.root, padding=10, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(14, 10))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)

        title_box = ttk.Frame(header, style="HeroCompact.TFrame")
        title_box.grid(row=0, column=0, sticky="w")
        ttk.Label(title_box, text="Decision Support System", style="HeroCompactTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            title_box,
            text="Coordinator workspace for matrix editing, decision-maker weighting, and distribution.",
            style="HeroCompactSubtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.mode_button = ttk.Button(header, text="Mode clair", style="Secondary.TButton", command=self._toggle_mode)
        self.mode_button.grid(row=0, column=1, rowspan=2, sticky="e")

        content = ttk.Frame(shell, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        left_panel = ttk.Frame(content, style="Card.TFrame", padding=10)
        left_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 10))

        right_panel = ttk.Frame(content, style="Card.TFrame", padding=10)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(0, weight=1)

        params_frame = ttk.LabelFrame(left_panel, text="Parameters", padding=10, style="Card.TLabelframe")
        params_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(params_frame, text="Introduce the weight of decision-makers.", style="SectionHintCompact.TLabel").pack(anchor="w", pady=(0, 8))

        stat_row = ttk.Frame(params_frame, style="CardInner.TFrame")
        stat_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(stat_row, text="Decision makers", style="MutedCapsCompact.TLabel").pack(anchor="w")
        ttk.Label(stat_row, text=str(self.num_decision_makers), style="MetricCompact.TLabel").pack(anchor="w", pady=(1, 0))

        weights_scroll_wrap = ttk.Frame(params_frame, style="CardInner.TFrame")
        weights_scroll_wrap.pack(fill=tk.BOTH, expand=False, pady=(0, 8))
        weights_scroll_wrap.columnconfigure(0, weight=1)
        weights_scroll_wrap.rowconfigure(0, weight=1)

        self.weight_canvas = tk.Canvas(weights_scroll_wrap, width=300, height=190, highlightthickness=0, bd=0)
        self.weight_canvas.grid(row=0, column=0, sticky="nsew")
        self.weight_scrollbar = ttk.Scrollbar(weights_scroll_wrap, orient=tk.VERTICAL, command=self.weight_canvas.yview)
        self.weight_scrollbar.grid(row=0, column=1, sticky="ns")
        self.weight_canvas.configure(yscrollcommand=self.weight_scrollbar.set)

        self.weight_scroll_inner = ttk.Frame(self.weight_canvas, style="CardInner.TFrame")
        self.weight_canvas_window = self.weight_canvas.create_window((0, 0), window=self.weight_scroll_inner, anchor="nw")
        self.weight_scroll_inner.bind("<Configure>", self._sync_weight_scrollregion)
        self.weight_canvas.bind("<Configure>", self._resize_weight_canvas_window)

        self.weights_frame = ttk.Frame(self.weight_scroll_inner, style="CardInner.TFrame")
        self.weights_frame.pack(fill=tk.X, expand=True)

        self.weights_status_label = ttk.Label(params_frame, text="Total weight : 0%", style="StatusBadCompact.TLabel")
        self.weights_status_label.pack(anchor="w", pady=(2, 4))

        chart_card = ttk.Frame(params_frame, style="ChartCard.TFrame", padding=8)
        chart_card.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(chart_card, text="Weight distribution", style="ChartTitleCompact.TLabel").pack(anchor="w", pady=(0, 4))

        chart_row = ttk.Frame(chart_card, style="ChartCard.TFrame")
        chart_row.pack(fill=tk.X)

        self.weights_pie_canvas = tk.Canvas(chart_row, width=self._weights_pie_diameter, height=self._weights_pie_diameter, highlightthickness=0, bd=0)
        self.weights_pie_canvas.pack(side=tk.LEFT, pady=(2, 0))

        self.legend_frame = ttk.Frame(chart_row, style="ChartCard.TFrame")
        self.legend_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        matrix_frame = ttk.LabelFrame(right_panel, text="Performance matrix", padding=10, style="Card.TLabelframe")
        matrix_frame.grid(row=0, column=0, sticky="nsew")
        matrix_frame.columnconfigure(0, weight=1)
        matrix_frame.rowconfigure(1, weight=1)

        matrix_top = ttk.Frame(matrix_frame, style="CardInner.TFrame")
        matrix_top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        matrix_top.columnconfigure(0, weight=1)

        text_block = ttk.Frame(matrix_top, style="CardInner.TFrame")
        text_block.grid(row=0, column=0, sticky="w")
        ttk.Label(text_block, text="Structured evaluation grid", style="SectionTitleCompact.TLabel").pack(anchor="w")
        ttk.Label(text_block, text="Double-click a numeric cell to edit values.", style="SectionHintCompact.TLabel").pack(anchor="w", pady=(2, 0))

        button_bar = ttk.Frame(matrix_top, style="CardInner.TFrame")
        button_bar.grid(row=0, column=1, sticky="e")

        ttk.Button(button_bar, text="Decision makers", command=self._open_decision_makers, style="Secondary.TButton").pack(side=tk.LEFT, padx=(0, 6))

        self.send_button = ttk.Button(button_bar, text="Send", command=self._send_matrix, state=tk.DISABLED, style="Accent.TButton")
        self.send_button.pack(side=tk.LEFT)

        self.table_border = tk.Frame(matrix_frame, padx=1, pady=1)
        self.table_border.grid(row=1, column=0, sticky="nsew")
        tree_container = ttk.Frame(self.table_border, style="TableWrap.TFrame")
        tree_container.pack(fill=tk.BOTH, expand=True)
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="browse", height=12)
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_cell_double_click)

        self._rebuild_weight_fields()

    def _sync_weight_scrollregion(self, event=None):
        self.weight_canvas.update_idletasks()
        self.weight_canvas.configure(scrollregion=self.weight_canvas.bbox("all"))

    def _resize_weight_canvas_window(self, event):
        if self.weight_canvas_window is not None:
            self.weight_canvas.itemconfigure(self.weight_canvas_window, width=event.width)
        self._sync_weight_scrollregion()

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

    def _toggle_mode(self):
        self.current_mode = "light" if self.current_mode == "dark" else "dark"
        self._apply_mode(self.current_mode)

    def _apply_mode(self, mode: str):
        self.palette = apply_excel_style(mode)
        self.root.configure(bg=self.palette["bg"])
        if self.weights_pie_canvas is not None:
            self.weights_pie_canvas.configure(bg=self.palette["panel_alt"])
        if self.weight_canvas is not None:
            self.weight_canvas.configure(bg=self.palette["card"])
        if self.table_border is not None:
            self.table_border.configure(bg=self.palette["border"])
        if self.mode_button is not None:
            self.mode_button.configure(text="Mode clair" if mode == "dark" else "Mode sombre")
        for window in self.decision_windows.values():
            try:
                window.apply_mode(mode)
            except Exception:
                pass
        self._update_weights_pie_chart()
        self._refresh_legend_colors()

    def _rebuild_weight_fields(self):
        self._building_weights = True
        for w in self.weights_frame.winfo_children():
            w.destroy()
        self.weight_vars.clear()
        self.weight_spinboxes.clear()
        self.weights = [0.0] * self.num_decision_makers

        for i in range(self.num_decision_makers):
            row = ttk.Frame(self.weights_frame, style="WeightRowCompact.TFrame", padding=(8, 6))
            row.pack(fill=tk.X, pady=3)
            row.columnconfigure(0, weight=1)
            label_text = self.decision_maker_names[i]
            ttk.Label(row, text=label_text, style="WeightNameCompact.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(row, text="Weight (%)", style="WeightHintCompact.TLabel").grid(row=1, column=0, sticky="w", pady=(1, 0))

            dot = tk.Canvas(row, width=14, height=14, highlightthickness=0, bd=0, bg=self.palette["panel_alt"])
            dot.grid(row=0, column=1, rowspan=2, padx=(0, 6))
            dot.create_oval(2, 2, 12, 12, fill=self._weights_pie_colors[i % len(self._weights_pie_colors)], outline="")

            var = tk.StringVar(value="0")
            var.trace_add("write", lambda *_, idx=i, v=var: self._on_weight_edited(idx, v))
            sb = tk.Spinbox(
                row,
                from_=0,
                to=100,
                textvariable=var,
                width=6,
                relief="flat",
                highlightthickness=1,
                bg="#FFFFFF",
                fg="#0F172A",
                buttonbackground="#E8EEFF",
                highlightbackground="#CFD8EA",
                highlightcolor=self.palette["primary"],
                font=("Segoe UI", 10, "bold"),
                command=lambda idx=i, v=var: self._on_weight_edited(idx, v),
            )
            sb.grid(row=0, column=2, rowspan=2, sticky="e")

            self.weight_vars.append(var)
            self.weight_spinboxes.append(sb)

        self._refresh_legend()
        self._building_weights = False
        self._check_weights_sum()
        self._update_weights_pie_chart()
        self.root.after(50, self._sync_weight_scrollregion)

    def _refresh_legend(self):
        for child in self.legend_frame.winfo_children():
            child.destroy()
        self.legend_rows.clear()
        for i, name in enumerate(self.decision_maker_names[:self.num_decision_makers]):
            row = ttk.Frame(self.legend_frame, style="ChartCard.TFrame")
            row.pack(fill=tk.X, pady=2)
            dot = tk.Canvas(row, width=14, height=14, highlightthickness=0, bd=0, bg=self.palette["panel_alt"])
            dot.pack(side=tk.LEFT, padx=(0, 6))
            dot.create_oval(2, 2, 12, 12, fill=self._weights_pie_colors[i % len(self._weights_pie_colors)], outline="")
            label = ttk.Label(row, text=name, style="WeightHintCompact.TLabel")
            label.pack(side=tk.LEFT)
            self.legend_rows.append((dot, label))

    def _refresh_legend_colors(self):
        for dot, _label in self.legend_rows:
            dot.configure(bg=self.palette["panel_alt"])

    def _on_weight_edited(self, idx: int, var: tk.StringVar):
        if self._building_weights or idx >= len(self.weights):
            return
        try:
            self.weights[idx] = self._parse_weight(var.get())
        except (ValueError, TypeError):
            self.weights[idx] = 0.0
        self._check_weights_sum()
        self._update_weights_pie_chart()

    def _update_weights_pie_chart(self):
        if self.weights_pie_canvas is None:
            return
        canvas = self.weights_pie_canvas
        canvas.delete("all")
        canvas.configure(bg=self.palette["panel_alt"])

        total = sum(self.weights) if self.weights else 0.0
        x0 = self._weights_pie_padding
        y0 = self._weights_pie_padding
        x1 = self._weights_pie_diameter - self._weights_pie_padding
        y1 = self._weights_pie_diameter - self._weights_pie_padding

        canvas.create_oval(x0, y0, x1, y1, fill=self.palette["heading_bg"], outline=self.palette["border"], width=2)
        canvas.create_oval(x0 + 22, y0 + 22, x1 - 22, y1 - 22, fill=self.palette["panel_alt"], outline="")

        if total <= 0.0:
            canvas.create_text(self._weights_pie_diameter / 2, self._weights_pie_diameter / 2, text="0%", fill=self.palette["text"], font=("Segoe UI", 11, "bold"))
            return

        start_angle = 90.0
        for i, w in enumerate(self.weights):
            if w <= 0.0:
                continue
            extent = -(w / total) * 360.0
            color = self._weights_pie_colors[i % len(self._weights_pie_colors)]
            canvas.create_arc(x0, y0, x1, y1, start=start_angle, extent=extent, fill=color, outline=self.palette["panel_alt"], width=2)
            start_angle += extent

        canvas.create_text(self._weights_pie_diameter / 2, self._weights_pie_diameter / 2, text=f"{total:.0f}%", fill=self.palette["text"], font=("Segoe UI", 11, "bold"))

    def _parse_weight(self, s: str) -> float:
        s = (s or "0").strip().replace(",", ".")
        return float(s) if s else 0.0

    def _get_weights_from_ui(self) -> List[float]:
        values = []
        for var in self.weight_vars:
            try:
                values.append(self._parse_weight(var.get()))
            except (ValueError, TypeError):
                values.append(0.0)
        return values

    def _validate_weights_sum_100(self) -> bool:
        return abs(sum(self._get_weights_from_ui()) - 100.0) <= 1e-6

    def _check_weights_sum(self):
        self.weights = self._get_weights_from_ui()
        total = sum(self.weights)
        if abs(total - 100.0) <= 1e-6:
            self.weights_status_label.configure(text=f"Total weight : {total:.2f} % (OK)", style="StatusGoodCompact.TLabel")
        else:
            self.weights_status_label.configure(text=f"Total weight : {total:.2f} % (must be 100%)", style="StatusBadCompact.TLabel")

        if self.send_button is not None:
            if self.matrix is not None and not self.matrix.empty:
                self.send_button.configure(state=tk.NORMAL)
            else:
                self.send_button.configure(state=tk.DISABLED)

    def _configure_new_matrix(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("New matrix")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.palette["bg"])
        apply_excel_style(self.current_mode)

        frm = ttk.Frame(dialog, padding=12, style="Dialog.TFrame")
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Nombre d'alternatives :", style="DialogLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        alt_var = tk.IntVar(value=DEFAULT_ALTERNATIVES)
        tk.Spinbox(frm, from_=1, to=100, textvariable=alt_var, width=5).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Nombre de critères :", style="DialogLabel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        crit_var = tk.IntVar(value=DEFAULT_CRITERIA)
        tk.Spinbox(frm, from_=1, to=50, textvariable=crit_var, width=5).grid(row=1, column=1, sticky="w", pady=4)

        names_frame = ttk.LabelFrame(frm, text="Noms des critères", padding=8, style="Card.TLabelframe")
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
                ttk.Label(names_frame, text=f"Critère {i_local + 1} :", style="DialogLabel.TLabel").grid(row=i_local, column=0, sticky="w", padx=(0, 8), pady=2)
                v_local = tk.StringVar(value=f"Critère {i_local + 1}")
                crit_vars.append(v_local)
                ttk.Entry(names_frame, textvariable=v_local, width=25, style="Modern.TEntry").grid(row=i_local, column=1, sticky="ew", pady=2)
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
                names.append(crit_vars[i_local].get().strip() or f"Critère {i_local + 1}")
            result["value"] = (n_alt, n_crit, names)
            dialog.destroy()

        ttk.Button(frm, text="OK", command=on_ok, style="Accent.TButton").grid(row=3, column=1, pady=(10, 0), sticky="e")
        self.root.wait_window(dialog)
        return result["value"]

    def _file_new(self):
        config = self._configure_new_matrix()
        if not config:
            return
        n_alt, n_crit, col_names = config
        self.matrix = pd.DataFrame([[1.0] * n_crit for _ in range(n_alt)], index=[f"Alternative {i + 1}" for i in range(n_alt)], columns=col_names)
        self.matrix_structure = self.matrix.copy(deep=True)
        self._refresh_tree()
        self._check_weights_sum()

    def _file_open(self):
        path = filedialog.askopenfilename(filetypes=[("Fichiers Excel (*.xlsx)", "*.xlsx")], defaultextension=".xlsx")
        if not path:
            return
        try:
            self.matrix = pd.read_excel(path, index_col=0)
            self.matrix_structure = self.matrix.copy(deep=True)
            self._refresh_tree()
            self._check_weights_sum()
        except Exception as exc:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier.\n{exc}")

    def _file_save(self):
        if self.matrix is None:
            messagebox.showwarning("Enregistrer", "Aucune matrice chargée.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("Tous les fichiers", "*.*")])
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
            if self.send_button is not None:
                self.send_button.configure(state=tk.DISABLED)
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
        if self.send_button is not None:
            self.send_button.configure(state=tk.NORMAL)

    def _matrix_from_tree(self):
        self.matrix = self.matrix.apply(pd.to_numeric, errors="coerce").fillna(0)
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

    def _send_matrix(self):
        if self.matrix is None or self.matrix.empty:
            messagebox.showwarning("Error", "No matrix loaded.")
            return

        if not self._validate_weights_sum_100():
            messagebox.showwarning("Weights", "The total weight must be exactly 100%.")
            return

        self._matrix_from_tree()
        self.matrix = self.matrix.apply(pd.to_numeric, errors="coerce").fillna(0)

        self.weights = self._get_weights_from_ui()
        self.weights = [w / 100 for w in self.weights]

        self.sent_matrix = self.matrix.copy(deep=True)

        model = Promethee(self.matrix, self.weights)
        pref_matrix = model.build_preference_matrix()
        result = model.compute_flows(pref_matrix)
        self.promethee_result = result

        # Ouvre la fenêtre de résultats PROMETHEE
        self._show_promethee_results(pref_matrix, result)

    def _show_promethee_results(self, pref_matrix: pd.DataFrame, result: pd.DataFrame):
        """Affiche la matrice π(a,b) — Phase 1 — et les flux Φ — Phase 2."""
        win = tk.Toplevel(self.root)
        win.title("Résultats PROMETHEE")
        win.configure(bg=self.palette["bg"])
        win.geometry("1050x650")
        win.minsize(800, 500)
        apply_excel_style(self.current_mode)

        shell = ttk.Frame(win, padding=14, style="App.TFrame")
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)
        shell.rowconfigure(3, weight=1)

        # ── PHASE 1 : matrice π(a,b) ─────────────────────────────────────
        hdr1 = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(10, 6))
        hdr1.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(hdr1, text="Phase 1 — Matrice de préférence agrégée π(a,b)",
                  style="HeroCompactTitle.TLabel").pack(anchor="w")
        ttk.Label(hdr1,
                  text="π(a,b) = Σ wj · Pj(a,b)   —   chaque cellule mesure le degré de préférence de a sur b",
                  style="HeroCompactSubtitle.TLabel").pack(anchor="w", pady=(2, 0))

        border1 = tk.Frame(shell, bg=self.palette["border"], padx=1, pady=1)
        border1.grid(row=1, column=0, sticky="nsew")
        c1 = ttk.Frame(border1, style="TableWrap.TFrame")
        c1.pack(fill=tk.BOTH, expand=True)
        c1.columnconfigure(0, weight=1)
        c1.rowconfigure(0, weight=1)

        alternatives = list(pref_matrix.index)
        cols1 = ["π(a,b)"] + [str(a) for a in alternatives]
        t1 = ttk.Treeview(c1, columns=cols1, show="headings")
        vsb1 = ttk.Scrollbar(c1, orient=tk.VERTICAL, command=t1.yview)
        hsb1 = ttk.Scrollbar(c1, orient=tk.HORIZONTAL, command=t1.xview)
        t1.configure(yscrollcommand=vsb1.set, xscrollcommand=hsb1.set)
        t1.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")
        hsb1.grid(row=1, column=0, sticky="ew")

        t1.heading("π(a,b)", text="a \\ b")
        t1.column("π(a,b)", width=110, anchor="w")
        for a in alternatives:
            t1.heading(str(a), text=str(a))
            t1.column(str(a), width=80, anchor="center")

        for a in alternatives:
            row_vals = [str(a)] + [
                "—" if a == b else f"{pref_matrix.loc[a, b]:.4f}"
                for b in alternatives
            ]
            t1.insert("", tk.END, values=row_vals)

        # ── PHASE 2 : flux Φ+, Φ-, Φnet ─────────────────────────────────
        hdr2 = ttk.Frame(shell, style="HeroCompact.TFrame", padding=(10, 6))
        hdr2.grid(row=2, column=0, sticky="ew", pady=(10, 4))
        ttk.Label(hdr2, text="Phase 2 — Flux PROMETHEE et classement final",
                  style="HeroCompactTitle.TLabel").pack(anchor="w")
        ttk.Label(hdr2,
                  text="Φ+(a) = flux sortant   |   Φ-(a) = flux entrant   |   Φnet = Φ+ − Φ−  (plus grand = meilleur)",
                  style="HeroCompactSubtitle.TLabel").pack(anchor="w", pady=(2, 0))

        border2 = tk.Frame(shell, bg=self.palette["border"], padx=1, pady=1)
        border2.grid(row=3, column=0, sticky="nsew")
        c2 = ttk.Frame(border2, style="TableWrap.TFrame")
        c2.pack(fill=tk.BOTH, expand=True)
        c2.columnconfigure(0, weight=1)
        c2.rowconfigure(0, weight=1)

        cols2 = ["Rang", "Alternative", "Φ+ (sortant)", "Φ- (entrant)", "Φ net"]
        t2 = ttk.Treeview(c2, columns=cols2, show="headings")
        vsb2 = ttk.Scrollbar(c2, orient=tk.VERTICAL, command=t2.yview)
        t2.configure(yscrollcommand=vsb2.set)
        t2.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")

        for col, w in [("Rang", 60), ("Alternative", 140),
                       ("Φ+ (sortant)", 130), ("Φ- (entrant)", 130), ("Φ net", 120)]:
            t2.heading(col, text=col)
            t2.column(col, width=w, anchor="center")

        for rank, alt in enumerate(result.index, start=1):
            t2.insert("", tk.END, values=(
                rank,
                str(alt),
                f"{result.loc[alt, 'Phi+']:.4f}",
                f"{result.loc[alt, 'Phi-']:.4f}",
                f"{result.loc[alt, 'Phi net']:.4f}",
            ))

        ttk.Button(shell, text="Fermer", command=win.destroy,
                   style="Secondary.TButton").grid(row=4, column=0, sticky="e", pady=(10, 0))

    def _open_decision_makers(self):
        win = tk.Toplevel(self.root)
        win.title("Decision makers list")
        win.transient(self.root)
        win.configure(bg=self.palette["bg"])
        apply_excel_style(self.current_mode)

        frame = ttk.Frame(win, padding=12, style="Dialog.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Decision makers", style="DialogTitle.TLabel").pack(anchor="w", pady=(0, 8))

        listbox = tk.Listbox(
            frame,
            bg="#FFFFFF",
            fg="#0F172A",
            selectbackground=self.palette["primary"],
            selectforeground="#FFFFFF",
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.palette["border_soft"],
            font=("Segoe UI", 10),
        )
        listbox.pack(fill=tk.BOTH, expand=True)

        for name in self.decision_maker_names[:self.num_decision_makers]:
            listbox.insert(tk.END, name)

        def open_selected(event=None):
            idx = listbox.curselection()
            if not idx:
                return
            name = listbox.get(idx[0])
            if name not in self.decision_windows or not self.decision_windows[name].window.winfo_exists():
                self.decision_windows[name] = DecisionMakerWindow(self.root, name, self._get_weight_for_decision_maker(name), mode=self.current_mode)
            window = self.decision_windows[name]
            if self.sent_matrix is not None:
                window.receive_matrix(self.sent_matrix)
            window.update_weight(self._get_weight_for_decision_maker(name))
            window.apply_mode(self.current_mode)
            window.window.lift()
            window.window.focus_force()

        listbox.bind("<Double-1>", open_selected)
        ttk.Button(frame, text="Open selected", command=open_selected, style="Accent.TButton").pack(pady=(8, 0), anchor="e")

    def _get_weight_for_decision_maker(self, name: str) -> float:
        weights = self._get_weights_from_ui()
        if name in self.decision_maker_names:
            i = self.decision_maker_names.index(name)
            return weights[i] if i < len(weights) else 0.0
        return 0.0

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    CoordinatorApp().run()