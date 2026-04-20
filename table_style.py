import tkinter as tk
from tkinter import ttk

LIGHT_PALETTE = {
    "bg": "#dbe4ee",
    "panel": "#f5f8fc",
    "panel_alt": "#edf3f9",
    "card": "#ffffff",
    "surface": "#ffffff",
    "border": "#b8cbe0",
    "border_soft": "#cdd9e6",
    "text": "#12304d",
    "text_dark": "#10273d",
    "muted": "#557391",
    "primary": "#054a91",
    "primary_hover": "#043a72",
    "accent": "#3e7cb1",
    "danger": "#c14953",
    "heading_bg": "#dbe4ee",
    "selected": "#c9d9ea",
    "hero": "#edf3f9",
}

DARK_PALETTE = {
    "bg": "#081827",
    "panel": "#0d2237",
    "panel_alt": "#102a43",
    "card": "#0c2033",
    "surface": "#f8fbff",
    "border": "#2c5378",
    "border_soft": "#b7c9dc",
    "text": "#e7f0f8",
    "text_dark": "#10273d",
    "muted": "#9bb8d3",
    "primary": "#81a4cd",
    "primary_hover": "#6d93bf",
    "accent": "#3e7cb1",
    "danger": "#ff7a7a",
    "heading_bg": "#dbe4ee",
    "selected": "#c8d9ec",
    "hero": "#0d2237",
}


def _apply_palette(style: ttk.Style, palette: dict) -> None:
    font_main = ("Segoe UI", 10)
    font_bold = ("Segoe UI", 10, "bold")

    style.configure(".", background=palette["bg"], foreground=palette["text"], font=font_main)

    style.configure("App.TFrame", background=palette["bg"])
    style.configure("HeroCompact.TFrame", background=palette["hero"])
    style.configure("Card.TFrame", background=palette["card"])
    style.configure("CardInner.TFrame", background=palette["card"])
    style.configure("ChartCard.TFrame", background=palette["panel_alt"])
    style.configure("TableWrap.TFrame", background=palette["surface"])
    style.configure("Dialog.TFrame", background=palette["panel"])
    style.configure("WeightRowCompact.TFrame", background=palette["panel_alt"])

    style.configure("TLabelframe", background=palette["card"], bordercolor=palette["border"], relief="solid")
    style.configure("Card.TLabelframe", background=palette["card"], bordercolor=palette["border"], relief="solid")
    style.configure("TLabelframe.Label", background=palette["card"], foreground=palette["text"], font=("Segoe UI Semibold", 11))
    style.configure("Card.TLabelframe.Label", background=palette["card"], foreground=palette["text"], font=("Segoe UI Semibold", 11))

    style.configure("TLabel", background=palette["bg"], foreground=palette["text"], font=font_main)
    style.configure("HeroCompactTitle.TLabel", background=palette["hero"], foreground=palette["text"], font=("Segoe UI Semibold", 14))
    style.configure("HeroCompactSubtitle.TLabel", background=palette["hero"], foreground=palette["muted"], font=("Segoe UI", 9))
    style.configure("SectionTitleCompact.TLabel", background=palette["card"], foreground=palette["text"], font=("Segoe UI Semibold", 10))
    style.configure("SectionHintCompact.TLabel", background=palette["card"], foreground=palette["muted"], font=("Segoe UI", 9))
    style.configure("MutedCapsCompact.TLabel", background=palette["card"], foreground=palette["muted"], font=("Segoe UI", 8, "bold"))
    style.configure("MetricCompact.TLabel", background=palette["card"], foreground=palette["text"], font=("Segoe UI Semibold", 15))
    style.configure("ChartTitleCompact.TLabel", background=palette["panel_alt"], foreground=palette["text"], font=("Segoe UI Semibold", 9))
    style.configure("WeightNameCompact.TLabel", background=palette["panel_alt"], foreground=palette["text"], font=("Segoe UI Semibold", 9))
    style.configure("WeightHintCompact.TLabel", background=palette["panel_alt"], foreground=palette["muted"], font=("Segoe UI", 8))
    style.configure("DialogLabel.TLabel", background=palette["panel"], foreground=palette["text"], font=("Segoe UI", 9))
    style.configure("DialogTitle.TLabel", background=palette["panel"], foreground=palette["text"], font=("Segoe UI Semibold", 12))
    style.configure("StatusGoodCompact.TLabel", background=palette["card"], foreground=palette["accent"], font=("Segoe UI", 9, "bold"))
    style.configure("StatusBadCompact.TLabel", background=palette["card"], foreground=palette["danger"], font=("Segoe UI", 9, "bold"))

    style.configure("TButton", padding=(12, 8), font=font_bold, borderwidth=0, relief="flat")
    style.configure("Accent.TButton", background=palette["primary"], foreground="#FFFFFF")
    style.map(
        "Accent.TButton",
        background=[("active", palette["primary_hover"]), ("disabled", "#aebdcb")],
        foreground=[("disabled", "#eef3f8")],
    )

    style.configure("Secondary.TButton", background=palette["panel_alt"], foreground=palette["text"])
    style.map(
        "Secondary.TButton",
        background=[("active", palette["heading_bg"]), ("disabled", "#c4d1dc")],
        foreground=[("disabled", "#64809a")],
    )

    style.configure(
        "Treeview",
        rowheight=30,
        font=("Segoe UI", 9),
        background=palette["surface"],
        fieldbackground=palette["surface"],
        foreground=palette["text_dark"],
        bordercolor=palette["border_soft"],
        lightcolor=palette["surface"],
        darkcolor=palette["surface"],
        relief="flat",
    )
    style.configure(
        "Treeview.Heading",
        font=("Segoe UI Semibold", 9),
        background=palette["heading_bg"],
        foreground=palette["primary"],
        relief="flat",
        padding=(8, 8),
    )
    style.map(
        "Treeview",
        background=[("selected", palette["selected"])],
        foreground=[("selected", palette["text_dark"])],
    )

    style.configure(
        "TEntry",
        fieldbackground="#FFFFFF",
        foreground=palette["text_dark"],
        bordercolor=palette["border_soft"],
        insertcolor=palette["text_dark"],
        padding=6,
    )
    style.configure(
        "Modern.TEntry",
        fieldbackground="#FFFFFF",
        foreground=palette["text_dark"],
        bordercolor=palette["border_soft"],
        insertcolor=palette["text_dark"],
        padding=6,
    )
    style.map(
        "Modern.TEntry",
        bordercolor=[("focus", palette["primary"])],
        lightcolor=[("focus", palette["primary"])],
        darkcolor=[("focus", palette["primary"])],
    )


def apply_excel_style(mode: str = "dark") -> dict:
    style = ttk.Style()

    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    palette = DARK_PALETTE if mode == "dark" else LIGHT_PALETTE
    _apply_palette(style, palette)
    return palette
