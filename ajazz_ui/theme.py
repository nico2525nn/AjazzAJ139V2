from tkinter import font as tkfont, ttk


PALETTES = {
    "light": {
        "window": "#f3f4f6",
        "sidebar": "#ffffff",
        "sidebar_panel": "#f9fafb",
        "panel": "#ffffff",
        "panel_alt": "#f9fafb",
        "input": "#f3f4f6",
        "border": "#d1d5db",
        "border_soft": "#e5e7eb",
        "text": "#111827",
        "text_muted": "#6b7280",
        "text_on_dark": "#ffffff",
        "accent": "#0d9488",
        "accent_hover": "#0f766e",
        "accent_soft": "#ccfbf1",
        "warning": "#f59e0b",
        "warning_soft": "#fef3c7",
        "danger": "#ef4444",
        "success": "#10b981",
        "table_heading": "#f3f4f6",
    },
    "dark": {
        "window": "#07111f",
        "sidebar": "#091221",
        "sidebar_panel": "#0e1a2d",
        "panel": "#101b2e",
        "panel_alt": "#13213a",
        "input": "#0b1628",
        "border": "#20314c",
        "border_soft": "#182741",
        "text": "#edf5ff",
        "text_muted": "#91a5c4",
        "text_on_dark": "#edf5ff",
        "accent": "#2dd4bf",
        "accent_hover": "#21baa7",
        "accent_soft": "#103a38",
        "warning": "#f59e0b",
        "warning_soft": "#40280b",
        "danger": "#f87171",
        "success": "#34d399",
        "table_heading": "#16233b",
    },
}


def get_palette(theme_name: str) -> dict:
    return PALETTES.get(theme_name, PALETTES["dark"])


def font_family(lang: str) -> str:
    available = set(tkfont.families())
    candidates = (
        ["Noto Sans JP", "Noto Sans CJK JP", "Noto Sans", "Yu Gothic UI", "Meiryo UI", "Meiryo"]
        if lang == "ja"
        else ["Segoe UI Variable Display", "Segoe UI Variable Text", "Bahnschrift", "Segoe UI"]
    )
    for name in candidates:
        if name in available:
            return name
    return candidates[-1]


def apply_ttk_treeview_style(style: ttk.Style, palette: dict, base_font: str | None = None):
    body_font = (base_font, 11) if base_font else None
    heading_font = (base_font, 11, "bold") if base_font else None
    style.theme_use("clam")
    if body_font:
        style.configure(".", font=body_font)
        style.configure("TCombobox", font=body_font)
        style.configure("TEntry", font=body_font)
        style.configure("TSpinbox", font=body_font)
        style.configure("TLabel", font=body_font)
        style.configure("TButton", font=body_font)
    style.configure(
        "Ajazz.Treeview",
        background=palette["input"],
        foreground=palette["text"],
        fieldbackground=palette["input"],
        bordercolor=palette["border"],
        lightcolor=palette["border"],
        darkcolor=palette["border"],
        rowheight=30,
        relief="flat",
        font=body_font,
    )
    style.configure(
        "Ajazz.Treeview.Heading",
        background=palette["table_heading"],
        foreground=palette["text"],
        bordercolor=palette["border"],
        relief="flat",
        padding=(8, 8),
        font=heading_font,
    )
    style.map(
        "Ajazz.Treeview",
        background=[("selected", palette["accent"])],
        foreground=[("selected", palette["window"])],
    )
