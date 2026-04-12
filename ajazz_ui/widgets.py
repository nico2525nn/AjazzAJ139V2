import customtkinter as ctk


class SectionCard(ctk.CTkFrame):
    def __init__(self, master, palette, title="", subtitle="", **kwargs):
        super().__init__(
            master,
            fg_color=palette["panel"],
            border_color=palette["border"],
            border_width=1,
            corner_radius=20,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)
        self.body = self
        if title or subtitle:
            header = ctk.CTkFrame(self, fg_color="transparent")
            self.header_frame = header
            header.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
            header.grid_columnconfigure(0, weight=1)
            if title:
                self.title_label = ctk.CTkLabel(
                    header,
                    text=title,
                    font=ctk.CTkFont(size=18, weight="bold"),
                    anchor="w",
                )
                self.title_label.grid(row=0, column=0, sticky="w")
            if subtitle:
                self.subtitle_label = ctk.CTkLabel(
                    header,
                    text=subtitle,
                    text_color=palette["text_muted"],
                    anchor="w",
                )
                self.subtitle_label.grid(row=1, column=0, sticky="w", pady=(4, 0))


class MetricCard(ctk.CTkFrame):
    def __init__(self, master, palette, label="", value="", accent=False, **kwargs):
        fg_color = palette["accent_soft"] if accent else palette["panel"]
        border_color = palette["accent"] if accent else palette["border"]
        super().__init__(
            master,
            fg_color=fg_color,
            border_color=border_color,
            border_width=1,
            corner_radius=18,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)
        self.label = ctk.CTkLabel(self, text=label, text_color=palette["text_muted"], anchor="w")
        self.label.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 2))
        self.value = ctk.CTkLabel(self, text=value, font=ctk.CTkFont(size=20, weight="bold"), anchor="w")
        self.value.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))

    def set_text(self, label=None, value=None):
        if label is not None:
            self.label.configure(text=label)
        if value is not None:
            self.value.configure(text=value)
