

from tkinter import ttk


def apply_excel_style() -> None:
    """Apply Excel-like style to all Treeviews: row height, font, headings."""
    style = ttk.Style()
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
