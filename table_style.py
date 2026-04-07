import tkinter as tk
from tkinter import ttk

def apply_excel_style() -> None:
    style = ttk.Style()

    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    # Fonts
    font_main = ("Segoe UI", 10)

    # LabelFrame
    style.configure(
        "TLabelframe",
        background="#F5F7FA",
        bordercolor="#D6DCE5"
    )
    style.configure(
        "TLabelframe.Label",
        background="#F5F7FA",
        foreground="#2F80ED",
        font=("Segoe UI", 10, "bold")
    )

    # Buttons
    style.configure(
        "TButton",
        padding=8,
        font=font_main
    )

    style.map(
        "TButton",
        background=[("active", "#DCEBFF"), ("disabled", "#E0E0E0")]
    )

    # Treeview
    style.configure(
        "Treeview",
        rowheight=28,
        font=("Segoe UI", 10),
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground="#1F2933"
    )

    style.configure(
        "Treeview.Heading",
        font=("Segoe UI", 10, "bold"),
        background="#E9F1FF",
        foreground="#2F80ED"
    )

    style.map(
        "Treeview",
        background=[("selected", "#DCEBFF")]
    )