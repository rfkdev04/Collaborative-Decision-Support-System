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
from promethee import aggregate_decision_maker_results


class CoordinatorApp:
    def __init__(self):
        self.sent_matrix = None
        self.current_mode = "dark"
        self.palette = None
        self.root = tk.Tk()
        self.root.title("DSS - Interface Coordinateur")
        self.root.geometry("1180x820")
        self.root.minsize(900, 700)

        self.matrix: Optional[pd.DataFrame] = None
        self.matrix_structure: Optional[pd.DataFrame] = None
        self.weights: List[float] = []
        self.num_decision_makers = 4
        self.expected_decisions = 4
        self.decision_maker_names = ["Politician", "Economist", "Environment representative", "Public representative"]

        self.decision_windows: Dict[str, DecisionMakerWindow] = {}
        self.weight_vars: List[tk.StringVar] = []
        self.weight_spinboxes: List[tk.Spinbox] = []
        self.legend_rows: List[tuple] = []
        self._building_weights = False

        self._weights_pie_colors = ["#054A91", "#3E7CB1", "#81A4CD", "#DBE4EE"]
        self._weights_pie_diameter = 80
        self._weights_pie_padding = 10
        self.weights_pie_canvas = None
        self.mode_button = None
        self.weight_canvas = None
        self.weight_scrollbar = None
        self.weight_scroll_inner = None
        self.weight_canvas_window = None
        self.table_border = None
        self.send_button = None
        self.aggregate_button = None
        self.final_results = None

        # ── Nombre de décideurs requis pour le consensus (choisi par le coordinateur)
        self.consensus_requis_var = tk.IntVar(value=3)

        self._build_ui()
        self._apply_mode(self.current_mode)

    # ─────────────────────────────────────────────
    #  UI BUILD
    # ─────────────────────────────────────────────

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
        ttk.Label(title_box, text="Coordinator workspace for matrix editing, aggregation, and exploitation.", style="HeroCompactSubtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 0))

        self.mode_button = ttk.Button(header, text="Mode clair", style="Secondary.TButton", command=self._toggle_mode)
        self.mode_button.grid(row=0, column=1, rowspan=2, sticky="e")

        content = ttk.Frame(shell, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=0)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

        # ── Left panel scrollable ────────────────
        left_outer = ttk.Frame(content, style="Card.TFrame")
        left_outer.grid(row=0, column=0, sticky="nsw", padx=(0, 10))

        left_canvas = tk.Canvas(left_outer, width=330, highlightthickness=0, bd=0)
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_vsb = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=left_canvas.yview)
        left_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        left_canvas.configure(yscrollcommand=left_vsb.set)

        left_panel = ttk.Frame(left_canvas, style="Card.TFrame", padding=10)
        left_win = left_canvas.create_window((0, 0), window=left_panel, anchor="nw")

        def _sync_left(event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        def _resize_left(event):
            left_canvas.itemconfigure(left_win, width=event.width)
        left_panel.bind("<Configure>", _sync_left)
        left_canvas.bind("<Configure>", _resize_left)

        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        right_panel = ttk.Frame(content, style="Card.TFrame", padding=10)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(0, weight=1)

        # ── Left panel ──────────────────────────
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

        self.weight_canvas = tk.Canvas(weights_scroll_wrap, width=300, height=120, highlightthickness=0, bd=0)
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

        # ── NOUVEAU : seuil d'acceptation du coordinateur ───────────────────
        consensus_card = ttk.LabelFrame(left_panel, text="Seuil d'acceptation", padding=10, style="Card.TLabelframe")
        consensus_card.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(
            consensus_card,
            text="Nombre minimum de décideurs\ndevant accepter une alternative :",
            style="WeightHintCompact.TLabel",
        ).pack(anchor="w", pady=(0, 6))

        consensus_row = ttk.Frame(consensus_card, style="CardInner.TFrame")
        consensus_row.pack(fill=tk.X)

        tk.Spinbox(
            consensus_row,
            from_=1, to=self.num_decision_makers,
            textvariable=self.consensus_requis_var,
            width=4,
            relief="flat",
            highlightthickness=1,
            bg="#FFFFFF", fg="#0F172A",
            buttonbackground="#E8EEFF",
            highlightbackground="#CFD8EA",
            font=("Segoe UI", 13, "bold"),
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(
            consensus_row,
            text=f"décideur(s) sur {self.num_decision_makers}",
            style="WeightNameCompact.TLabel",
        ).pack(side=tk.LEFT)
        # ────────────────────────────────────────────────────────────────────

        # ── Panneau statut classements
        status_card = ttk.LabelFrame(left_panel, text="Classements reçus", padding=10, style="Card.TLabelframe")
        status_card.pack(fill=tk.X, pady=(10, 0))
        self.dm_status_labels: Dict[str, tuple] = {}
        for name in self.decision_maker_names:
            row = ttk.Frame(status_card, style="CardInner.TFrame")
            row.pack(fill=tk.X, pady=2)
            dot = tk.Canvas(row, width=12, height=12, highlightthickness=0, bd=0, bg=self.palette.get("panel_alt", "#f0f0f0"))
            dot.pack(side=tk.LEFT, padx=(0, 6))
            dot.create_oval(2, 2, 10, 10, fill="#CBD5E1", outline="", tags="dot")
            lbl = ttk.Label(row, text=name, style="WeightHintCompact.TLabel")
            lbl.pack(side=tk.LEFT)
            self.dm_status_labels[name] = (dot, lbl)

        # ── Right panel ──────────────────────────
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
        self.send_button.pack(side=tk.LEFT, padx=(0, 6))
        self.aggregate_button = ttk.Button(button_bar, text="Agrégation + Exploitation", command=self._aggregate_and_exploit, state=tk.DISABLED, style="Secondary.TButton")
        self.aggregate_button.pack(side=tk.LEFT)

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

    # ─────────────────────────────────────────────
    #  SCROLLABLE WEIGHTS
    # ─────────────────────────────────────────────

    def _sync_weight_scrollregion(self, event=None):
        self.weight_canvas.update_idletasks()
        self.weight_canvas.configure(scrollregion=self.weight_canvas.bbox("all"))

    def _resize_weight_canvas_window(self, event):
        if self.weight_canvas_window is not None:
            self.weight_canvas.itemconfigure(self.weight_canvas_window, width=event.width)
        self._sync_weight_scrollregion()

    # ─────────────────────────────────────────────
    #  MENUBAR
    # ─────────────────────────────────────────────

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

    # ─────────────────────────────────────────────
    #  MODE
    # ─────────────────────────────────────────────

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

    # ─────────────────────────────────────────────
    #  WEIGHT FIELDS
    # ─────────────────────────────────────────────

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
            sb = tk.Spinbox(row, from_=0, to=100, textvariable=var, width=6, relief="flat", highlightthickness=1,
                            bg="#FFFFFF", fg="#0F172A", buttonbackground="#E8EEFF",
                            highlightbackground="#CFD8EA", highlightcolor=self.palette["primary"],
                            font=("Segoe UI", 10, "bold"),
                            command=lambda idx=i, v=var: self._on_weight_edited(idx, v))
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
            canvas.create_text(self._weights_pie_diameter / 2, self._weights_pie_diameter / 2,
                               text="0%", fill=self.palette["text"], font=("Segoe UI", 11, "bold"))
            return
        start_angle = 90.0
        for i, w in enumerate(self.weights):
            if w <= 0.0:
                continue
            extent = -(w / total) * 360.0
            color = self._weights_pie_colors[i % len(self._weights_pie_colors)]
            canvas.create_arc(x0, y0, x1, y1, start=start_angle, extent=extent,
                              fill=color, outline=self.palette["panel_alt"], width=2)
            start_angle += extent
        canvas.create_text(self._weights_pie_diameter / 2, self._weights_pie_diameter / 2,
                           text=f"{total:.0f}%", fill=self.palette["text"], font=("Segoe UI", 11, "bold"))

    @staticmethod
    def _parse_weight(s: str) -> float:
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
            self.weights_status_label.configure(
                text=f"Total weight : {total:.2f} % (OK)", style="StatusGoodCompact.TLabel")
        else:
            self.weights_status_label.configure(
                text=f"Total weight : {total:.2f} % (must be 100%)", style="StatusBadCompact.TLabel")
        if self.send_button is not None:
            self.send_button.configure(
                state=tk.NORMAL if self.matrix is not None and not self.matrix.empty else tk.DISABLED)

    # ─────────────────────────────────────────────
    #  FILE MENU
    # ─────────────────────────────────────────────

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
                ttk.Label(names_frame, text=f"Critère {i_local + 1} :", style="DialogLabel.TLabel").grid(
                    row=i_local, column=0, sticky="w", padx=(0, 8), pady=2)
                v_local = tk.StringVar(value=f"Critère {i_local + 1}")
                crit_vars.append(v_local)
                ttk.Entry(names_frame, textvariable=v_local, width=25, style="Modern.TEntry").grid(
                    row=i_local, column=1, sticky="ew", pady=2)
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
            names = [crit_vars[i_local].get().strip() or f"Critère {i_local + 1}" for i_local in range(n_crit)]
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
        self.matrix = pd.DataFrame(
            [[1.0] * n_crit for _ in range(n_alt)],
            index=[f"Alternative {i + 1}" for i in range(n_alt)],
            columns=col_names,
        )
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
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("Tous les fichiers", "*.*")])
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

    # ─────────────────────────────────────────────
    #  MATRIX TREE
    # ─────────────────────────────────────────────

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
        entry = ttk.Entry(self.tree, width=max(6, width // 8), style="Modern.TEntry")
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

    # ─────────────────────────────────────────────
    #  SEND MATRIX
    # ─────────────────────────────────────────────

    def _send_matrix(self):
        if self.matrix is None or self.matrix.empty:
            messagebox.showwarning("Error", "No matrix loaded.")
            return
        if not self._validate_weights_sum_100():
            messagebox.showwarning("Weights", "The total weight must be exactly 100%.")
            return
        self._matrix_from_tree()
        self.weights = self._get_weights_from_ui()
        self.sent_matrix = self.matrix.copy(deep=True)
        for name, window in list(self.decision_windows.items()):
            try:
                window.receive_matrix(self.sent_matrix)
                window.update_weight(self._get_weight_for_decision_maker(name))
                window.apply_mode(self.current_mode)
            except Exception:
                pass
        messagebox.showinfo("Success", "Matrix ready. Open a decision maker from the list to view the interface.")

    # ─────────────────────────────────────────────
    #  DECISION MAKERS LIST
    # ─────────────────────────────────────────────

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
            frame, bg="#FFFFFF", fg="#0F172A",
            selectbackground=self.palette["primary"], selectforeground="#FFFFFF",
            relief="flat", highlightthickness=1,
            highlightbackground=self.palette["border_soft"], font=("Segoe UI", 10),
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
                self.decision_windows[name] = DecisionMakerWindow(
                    self.root, name,
                    self._get_weight_for_decision_maker(name),
                    mode=self.current_mode,
                    on_result_ready=self._on_decision_maker_result_ready,
                )
            window = self.decision_windows[name]
            if self.sent_matrix is not None:
                window.receive_matrix(self.sent_matrix)
            window.update_weight(self._get_weight_for_decision_maker(name))
            window.apply_mode(self.current_mode)
            window.window.lift()
            window.window.focus_force()

        listbox.bind("<Double-1>", open_selected)
        ttk.Button(frame, text="Open selected", command=open_selected,
                   style="Accent.TButton").pack(pady=(8, 0), anchor="e")

    def _get_weight_for_decision_maker(self, name: str) -> float:
        weights = self._get_weights_from_ui()
        if name in self.decision_maker_names:
            i = self.decision_maker_names.index(name)
            return weights[i] if i < len(weights) else 0.0
        return 0.0

    # ─────────────────────────────────────────────
    #  CALLBACK : décideur a soumis son classement
    # ─────────────────────────────────────────────

    def _on_decision_maker_result_ready(self, name: str):
        self._refresh_dm_status_panel()

        ready = self._get_ready_decision_makers()
        missing = [n for n in self.decision_maker_names[:self.num_decision_makers] if n not in ready]
        total = self.num_decision_makers

        if len(ready) == total:
            self.aggregate_button.configure(state=tk.NORMAL)
            messagebox.showinfo(
                "Agrégation disponible",
                f"Les {total} décideurs ont soumis leur classement.\nVous pouvez lancer l'agrégation.",
                parent=self.root,
            )
        else:
            self.aggregate_button.configure(state=tk.DISABLED)
            missing_str = "\n• ".join(missing)
            messagebox.showinfo(
                "Classement reçu",
                f"Classement reçu de : {name}\n\nEn attente de :\n• {missing_str}",
                parent=self.root,
            )

    def _get_ready_decision_makers(self) -> List[str]:
        return [
            n for n in self.decision_maker_names[:self.num_decision_makers]
            if n in self.decision_windows
            and self.decision_windows[n].get_promethee_results() is not None
        ]

    def _refresh_dm_status_panel(self):
        ready = self._get_ready_decision_makers()
        for name, (dot, _lbl) in self.dm_status_labels.items():
            color = "#22C55E" if name in ready else "#CBD5E1"
            dot.delete("dot")
            dot.create_oval(2, 2, 10, 10, fill=color, outline="", tags="dot")

    # ─────────────────────────────────────────────
    #  AGRÉGATION
    # ─────────────────────────────────────────────

    def _aggregate_and_exploit(self):
        ready = self._get_ready_decision_makers()
        missing = [n for n in self.decision_maker_names[:self.num_decision_makers] if n not in ready]

        if missing:
            missing_str = "\n• ".join(missing)
            messagebox.showerror(
                "Agrégation impossible",
                f"Classement manquant pour :\n• {missing_str}\n\n"
                "Demandez à ces décideurs de lancer leur calcul avant de continuer.",
                parent=self.root,
            )
            return

        if not self._validate_weights_sum_100():
            messagebox.showwarning("Agrégation", "Les poids des décideurs doivent totaliser 100%.")
            return

        dm_results = []
        for name in self.decision_maker_names[:self.num_decision_makers]:
            window = self.decision_windows[name]
            result = window.get_promethee_results()
            dm_results.append((name, self._get_weight_for_decision_maker(name), result))

        try:
            self.final_results = aggregate_decision_maker_results(dm_results)
            self._show_final_results_window()
        except Exception as exc:
            messagebox.showerror("Agrégation", str(exc))

    def _show_final_results_window(self):
        if self.final_results is None or self.final_results.empty:
            return

        win = tk.Toplevel(self.root)
        win.title("Tableau final — Agrégation")
        win.geometry("700x460")
        win.configure(bg=self.palette["bg"])
        apply_excel_style(self.current_mode)

        frame = ttk.Frame(win, padding=12, style="Dialog.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Tableau final agrégé — ϕ+, ϕ-, ϕ, Rang",
                  style="DialogTitle.TLabel").pack(anchor="w", pady=(0, 8))

        tree = ttk.Treeview(frame, columns=("Alternative", "ϕ+", "ϕ-", "ϕ", "Rang"), show="headings")
        for c in ("Alternative", "ϕ+", "ϕ-", "ϕ", "Rang"):
            tree.heading(c, text=c)
            tree.column(c, width=120 if c != "Alternative" else 200)
        for _, row in self.final_results.iterrows():
            tree.insert("", tk.END, values=(
                row["Alternative"],
                round(float(row["ϕ+"]), 6),
                round(float(row["ϕ-"]), 6),
                round(float(row["ϕ"]), 6),
                int(row["Rang"]),
            ))
        tree.pack(fill=tk.BOTH, expand=True)

        btn_row = ttk.Frame(frame, style="Dialog.TFrame")
        btn_row.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(btn_row, text="Exporter Excel", style="Secondary.TButton",
                   command=self._export_final_results).pack(side=tk.LEFT)

        ttk.Button(btn_row, text="Lancer l'exploitation →", style="Accent.TButton",
                   command=lambda: [win.destroy(), self._run_exploitation()]).pack(side=tk.RIGHT)

    # ─────────────────────────────────────────────
    #  EXPLOITATION ITÉRATIVE
    # ─────────────────────────────────────────────

    def _run_exploitation(self):
        """
        Propose chaque alternative du classement agrégé aux décideurs.
        - Chaque décideur utilise son propre seuil (seuil_pct) défini dans son interface.
        - Le coordinateur choisit le nombre d'acceptations requises via consensus_requis_var.
        - Max 18 essais (ou nombre d'alternatives si inférieur à 18).
        """
        if self.final_results is None or self.final_results.empty:
            messagebox.showwarning("Exploitation", "Lancez d'abord l'agrégation.", parent=self.root)
            return

        # Lire le seuil du coordinateur
        try:
            consensus_requis = max(1, min(self.num_decision_makers, int(self.consensus_requis_var.get())))
        except (ValueError, TypeError):
            consensus_requis = 3

        alternatives = list(self.final_results["Alternative"])
        n_alts = len(alternatives)
        max_essais = min(18, n_alts)

        history = []  # [(alternative, acceptations, refus, seuils_info)]

        for essai, alt_proposee in enumerate(alternatives[:max_essais], start=1):
            acceptations = []
            refus = []
            seuils_info = {}

            for name in self.decision_maker_names[:self.num_decision_makers]:
                window = self.decision_windows.get(name)
                if window is None:
                    refus.append(name)
                    seuils_info[name] = "—"
                    continue

                dm_results = window.get_promethee_results()
                if dm_results is None:
                    refus.append(name)
                    seuils_info[name] = "—"
                    continue

                # Seuil individuel du décideur (en %)
                seuil_pct = window.get_seuil_pct()
                seuils_info[name] = seuil_pct
                seuil_rang = max(1, round(n_alts * seuil_pct / 100))

                dm_alts = list(dm_results["Alternative"])
                if alt_proposee in dm_alts:
                    rang_dm = dm_alts.index(alt_proposee) + 1
                    if rang_dm <= seuil_rang:
                        acceptations.append(name)
                    else:
                        refus.append(name)
                else:
                    refus.append(name)

            history.append((alt_proposee, list(acceptations), list(refus), seuils_info))

            if len(acceptations) >= consensus_requis:
                self._show_exploitation_result(
                    alternative=alt_proposee,
                    essai=essai,
                    acceptations=acceptations,
                    refus=refus,
                    history=history,
                    success=True,
                    n_alts=n_alts,
                    consensus_requis=consensus_requis,
                )
                return

        # Aucun consensus
        self._show_exploitation_result(
            alternative=None,
            essai=max_essais,
            acceptations=[],
            refus=[],
            history=history,
            success=False,
            n_alts=n_alts,
            consensus_requis=consensus_requis,
        )

    def _show_exploitation_result(self, alternative, essai, acceptations, refus,
                                   history, success, n_alts, consensus_requis=3):
        win = tk.Toplevel(self.root)
        win.title("Résultat — Exploitation itérative")
        win.geometry("780x580")
        win.configure(bg=self.palette["bg"])
        apply_excel_style(self.current_mode)

        frame = ttk.Frame(win, padding=16, style="Dialog.TFrame")
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        # ── Titre
        title_text = "✅  Consensus atteint !" if success else "❌  Aucun consensus trouvé"
        ttk.Label(frame, text=title_text, style="DialogTitle.TLabel").grid(row=0, column=0, sticky="w")

        # ── Résumé
        info_frame = ttk.Frame(frame, style="CardInner.TFrame", padding=(0, 8))
        info_frame.grid(row=1, column=0, sticky="ew", pady=(6, 10))

        if success:
            ttk.Label(info_frame, text=f"Alternative retenue : {alternative}",
                      style="MetricCompact.TLabel").pack(anchor="w")
            ttk.Label(info_frame,
                      text=f"Trouvée à l'essai n°{essai}   |   Consensus requis : {consensus_requis} décideur(s) sur {self.num_decision_makers}",
                      style="SectionHintCompact.TLabel").pack(anchor="w", pady=(2, 8))
            ttk.Label(info_frame, text="Accepté par :", style="MutedCapsCompact.TLabel").pack(anchor="w")
            for name in acceptations:
                seuil = self.decision_windows[name].get_seuil_pct() if name in self.decision_windows else "?"
                ttk.Label(info_frame, text=f"   ✅  {name}  (seuil : top {seuil}%)",
                          style="WeightHintCompact.TLabel").pack(anchor="w")
            if refus:
                ttk.Label(info_frame, text="Refusé par :", style="MutedCapsCompact.TLabel").pack(anchor="w", pady=(6, 0))
                for name in refus:
                    seuil = self.decision_windows[name].get_seuil_pct() if name in self.decision_windows else "?"
                    ttk.Label(info_frame, text=f"   ❌  {name}  (seuil : top {seuil}%)",
                              style="WeightHintCompact.TLabel").pack(anchor="w")
        else:
            ttk.Label(info_frame,
                      text=f"Après {essai} essais, aucune alternative n'a obtenu {consensus_requis} acceptation(s).",
                      style="SectionHintCompact.TLabel").pack(anchor="w", pady=(0, 4))
            ttk.Label(info_frame, text="Seuils individuels des décideurs :",
                      style="MutedCapsCompact.TLabel").pack(anchor="w", pady=(4, 2))
            for name in self.decision_maker_names[:self.num_decision_makers]:
                seuil = self.decision_windows[name].get_seuil_pct() if name in self.decision_windows else "?"
                ttk.Label(info_frame, text=f"   • {name} : top {seuil}%",
                          style="WeightHintCompact.TLabel").pack(anchor="w")

        # ── Historique
        ttk.Label(frame, text="Détail des essais :",
                  style="MutedCapsCompact.TLabel").grid(row=2, column=0, sticky="nw", pady=(0, 4))

        hist_container = ttk.Frame(frame, style="CardInner.TFrame")
        hist_container.grid(row=3, column=0, sticky="nsew")
        hist_container.columnconfigure(0, weight=1)
        hist_container.rowconfigure(0, weight=1)

        cols = ("Essai", "Alternative proposée", "Accepté par", "Refusé par", "Résultat")
        hist_tree = ttk.Treeview(hist_container, columns=cols, show="headings", height=min(8, len(history)))
        widths = {"Essai": 45, "Alternative proposée": 170, "Accepté par": 170, "Refusé par": 170, "Résultat": 90}
        for col in cols:
            hist_tree.heading(col, text=col)
            hist_tree.column(col, width=widths.get(col, 120))

        for i, entry in enumerate(history, start=1):
            alt, acc, ref = entry[0], entry[1], entry[2]
            ok = len(acc) >= consensus_requis
            hist_tree.insert("", tk.END, values=(
                i, alt,
                ", ".join(acc) if acc else "—",
                ", ".join(ref) if ref else "—",
                "✅ Consensus" if ok else "❌ Rejeté",
            ))

        vsb = ttk.Scrollbar(hist_container, orient=tk.VERTICAL, command=hist_tree.yview)
        hist_tree.configure(yscrollcommand=vsb.set)
        hist_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # ── Boutons
        btn_row = ttk.Frame(frame, style="Dialog.TFrame")
        btn_row.grid(row=4, column=0, sticky="e", pady=(10, 0))
        ttk.Button(btn_row, text="Exporter Excel", style="Accent.TButton",
                   command=lambda: self._export_exploitation(history, success, alternative, consensus_requis)).pack(side=tk.RIGHT)

    def _export_exploitation(self, history, success, alternative_retenue, consensus_requis=3):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], parent=self.root)
        if not path:
            return
        try:
            rows = []
            for i, entry in enumerate(history, start=1):
                alt, acc, ref = entry[0], entry[1], entry[2]
                seuils_info = entry[3] if len(entry) > 3 else {}
                rows.append({
                    "Essai": i,
                    "Alternative proposée": alt,
                    "Accepté par": ", ".join(acc) if acc else "—",
                    "Refusé par": ", ".join(ref) if ref else "—",
                    "Nb acceptations": len(acc),
                    "Résultat": "Consensus ✅" if len(acc) >= consensus_requis else "Rejeté ❌",
                    "Seuils décideurs": "; ".join(f"{n}:{v}%" for n, v in seuils_info.items()),
                })
            df_hist = pd.DataFrame(rows)
            df_summary = pd.DataFrame([{
                "Consensus atteint": "Oui" if success else "Non",
                "Alternative retenue": alternative_retenue or "—",
                "Nombre d'essais": len(history),
                "Acceptations requises": consensus_requis,
            }])
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_hist.to_excel(writer, index=False, sheet_name="Historique essais")
                df_summary.to_excel(writer, index=False, sheet_name="Résumé")
            messagebox.showinfo("Export", "Résultats d'exploitation exportés avec succès.")
        except Exception as exc:
            messagebox.showerror("Export", str(exc))

    # ─────────────────────────────────────────────
    #  EXPORT AGRÉGATION
    # ─────────────────────────────────────────────

    def _export_final_results(self):
        if self.final_results is None or self.final_results.empty:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            self.final_results.to_excel(path, index=False)
            messagebox.showinfo("Export", "Tableau final exporté avec succès.")
        except Exception as exc:
            messagebox.showerror("Export", str(exc))

    # ─────────────────────────────────────────────
    #  RUN
    # ─────────────────────────────────────────────

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    CoordinatorApp().run()