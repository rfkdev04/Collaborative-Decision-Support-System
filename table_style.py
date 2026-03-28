

import tkinter as tk
from tkinter import ttk


def apply_excel_style() -> None:
    """
    Apply a consistent UI theme across the project:
    - Treeview (Excel-like grid row height + heading font + selection color)
    - Labelframe and buttons (shared colors)
    """
    style = ttk.Style()

    # Use a more customizable theme (Windows defaults may ignore many options otherwise)
    try:
        style.theme_use("clam")
    except tk.TclError:
        # Fallback to whatever is available
        pass

    # Common panel look (used by ttk.LabelFrame)
    bg_panel = "#E9F1FF"
    fg_panel = "#000000"
    border_panel = "#B8C9E6"

    style.configure("TLabelframe", background=bg_panel, foreground=fg_panel, bordercolor=border_panel)
    style.configure("TLabelframe.Label", background=bg_panel, foreground=fg_panel)

    # Buttons (ttk themed)
    style.configure(
        "TButton",
        padding=(10, 5),
        font=("Segoe UI", 9),
    )
    style.map(
        "TButton",
        background=[("active", "#D9E8FF"), ("disabled", "#D9D9D9")],
        foreground=[("active", "#000000"), ("disabled", "#666666")],
    )

    # Treeview (Excel-like)
    style.configure(
        "Treeview",
        rowheight=26,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Treeview.Heading",
        font=("Segoe UI", 9, "bold"),
    )
    style.map("Treeview", background=[("selected", "#CCE5FF")])
