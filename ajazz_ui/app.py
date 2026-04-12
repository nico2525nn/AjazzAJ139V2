import ctypes
import json
import os
import threading
import time
import tkinter as tk
from copy import deepcopy
from tkinter import font as tkfont, messagebox, ttk

import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw, ImageFont

from ajazz_mouse import (
    ACTION_PRESS,
    ACTION_RELEASE,
    EVENT_TYPE_KEYBOARD,
    EVENT_TYPE_MOUSE,
    AjazzMouse,
    MacroEvent,
    MacroProfile,
    MouseKeyBinding,
    decode_macro_profiles,
    encode_macro_profiles,
)
from .i18n import translate, ui_text
from .models import (
    BUTTON_SLOT_NAMES,
    HID_KEY_NAMES,
    KEYBOARD_EVENT_CHOICES,
    KEYBOARD_NAME_TO_CODE,
    KEYBOARD_PRESETS,
    MACRO_MODE_OPTIONS,
    MEDIA_PRESETS,
    MOUSE_CODE_NAMES,
    MOUSE_EVENT_CHOICES,
    MOUSE_NAME_TO_CODE,
    MOUSE_PRESETS,
    clone_binding,
    copy_default_bindings,
    modifier_names,
    resolve_binding_name,
    resolve_macro_event_name,
)
from .theme import apply_ttk_treeview_style, font_family, get_palette
from .widgets import MetricCard, SectionCard

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class AjazzApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.mouse = AjazzMouse(log_callback=self._append_log)
        self.current_config = None
        self.lang = "en"
        self.theme_name = "dark"
        self.mouse_keys = copy_default_bindings()
        self.macro_profiles = [MacroProfile(slot=index, name=f"Macro {index + 1}") for index in range(32)]
        self.selected_button_index = 0
        self.selected_macro_slot = 0
        self.selected_macro_event_index = None
        self.is_recording = False
        self._pressed_keys = set()
        self._pressed_mouse = set()
        self._last_record_time = None
        self._macro_metadata = {"names": {}}
        self._busy_count = 0
        self._status_refresh_in_progress = False
        self._macro_profiles_loaded = False
        self._macro_profiles_loading = False
        self._macro_status_key = "idle"
        self._macro_progress_current = 0
        self._macro_progress_total = 0
        self._saved_mouse_keys = copy_default_bindings()
        self._saved_macro_profiles = deepcopy(self.macro_profiles)
        self._macro_editor_widget = None
        self._macro_editor_info = None
        self._active_page = "perf"
        self._page_frames = {}
        self._nav_buttons = {}
        self._macro_profile_buttons = []
        self._device_summary = {"status": "st_disconnected", "version": "--", "battery": "--"}

        self.config_file = "settings.json"
        self.macro_meta_file = "macro_metadata.json"
        self._load_app_settings()
        self._load_macro_metadata()

        self._theme_palette = get_palette(self.theme_name)
        ctk.set_appearance_mode(self.theme_name)

        self.title(self._t("title"))
        self.geometry("1440x980")
        self.minsize(1260, 860)
        self.configure(fg_color=self._theme_palette["window"])

        self.tray_icon = None
        self._last_tray_battery = -1
        self.protocol("WM_DELETE_WINDOW", self._hide_window)

        self._build_ui()
        self._refresh_all_views()
        self.after(50, self._refresh_status)

        self._bg_thread = threading.Thread(target=self._tray_monitor_loop, daemon=True)
        self._bg_thread.start()

    def _t(self, key, *args):
        return translate(self.lang, key, *args)

    def _ui_text(self, key):
        return ui_text(self.lang, key)

    def _build_ui(self):
        palette = self._theme_palette
        base_font = font_family(self.lang)
        self._apply_global_font(base_font)
        self.style = ttk.Style(self)
        apply_ttk_treeview_style(self.style, palette, base_font=base_font)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=250, fg_color=palette["sidebar"], corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        self.sidebar.grid_rowconfigure(4, weight=1)
        self.sidebar.grid_propagate(False)

        self.brand_panel = ctk.CTkFrame(
            self.sidebar,
            fg_color=palette["sidebar_panel"],
            border_color=palette["border"],
            border_width=1,
            corner_radius=22,
        )
        self.brand_panel.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 14))
        self.brand_title = ctk.CTkLabel(
            self.brand_panel,
            text="AJ139 V2",
            text_color=palette["text_on_dark"],
            font=ctk.CTkFont(family=base_font, size=28, weight="bold"),
        )
        self.brand_title.pack(anchor="w", padx=16, pady=(16, 4))
        self.brand_subtitle = ctk.CTkLabel(
            self.brand_panel,
            text=self._t("tagline"),
            text_color=palette["text_muted"],
            wraplength=200,
            justify="left",
        )
        self.brand_subtitle.pack(anchor="w", padx=16, pady=(0, 16))


        nav_items = [
            ("perf", "nav_perf"),
            ("sys", "nav_sys"),
            ("keys", "nav_keys"),
            ("macro", "nav_macro"),
            ("debug", "nav_debug"),
        ]
        self.nav_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.nav_frame.grid(row=2, column=0, sticky="ew", padx=18, pady=(18, 0))
        self.nav_frame.grid_columnconfigure(0, weight=1)
        for row, (page, label_key) in enumerate(nav_items):
            button = ctk.CTkButton(
                self.nav_frame,
                text=self._t(label_key),
                command=lambda page=page: self._switch_page(page),
                anchor="w",
                height=40,
                corner_radius=20,
                border_width=1,
            )
            button.grid(row=row, column=0, sticky="ew", pady=4)
            self._nav_buttons[page] = button

        self.metrics_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.metrics_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=(18, 0))
        self.metrics_frame.grid_columnconfigure(0, weight=1)

        self.metric_connection = MetricCard(self.metrics_frame, palette, label=self._t("hero_connection"), value="--", accent=True)
        self.metric_connection.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.metric_firmware = MetricCard(self.metrics_frame, palette, label=self._t("hero_firmware"), value="--")
        self.metric_firmware.grid(row=1, column=0, sticky="ew", pady=6)
        self.metric_battery = MetricCard(self.metrics_frame, palette, label=self._t("hero_battery"), value="--")
        self.metric_battery.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        self.sidebar_tools = ctk.CTkFrame(
            self.sidebar,
            fg_color=palette["sidebar_panel"],
            border_color=palette["border"],
            border_width=1,
            corner_radius=22,
        )
        self.sidebar_tools.grid(row=5, column=0, sticky="sew", padx=18, pady=18)
        self.sidebar_tools.grid_columnconfigure(0, weight=1)

        self.lang_label = ctk.CTkLabel(self.sidebar_tools, text=self._t("lang_label"), text_color=palette["text_on_dark"], anchor="w")
        self.lang_label.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 4))
        self.lang_var = tk.StringVar(value=self.lang)
        self.lang_combo = ctk.CTkOptionMenu(self.sidebar_tools, variable=self.lang_var, values=["en", "ja"], command=self.change_language)
        self.lang_combo.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 14))

        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=1, sticky="nsew", padx=(0, 18), pady=18)
        self.main.grid_rowconfigure(0, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        self.page_host = ctk.CTkFrame(self.main, fg_color="transparent")
        self.page_host.grid(row=0, column=0, sticky="nsew")
        self.page_host.grid_rowconfigure(0, weight=1)
        self.page_host.grid_columnconfigure(0, weight=1)

        self._build_perf_page()
        self._build_sys_page()
        self._build_keys_page()
        self._build_macro_page()
        self._build_debug_page()

        self.footer = ctk.CTkFrame(
            self.main,
            fg_color=palette["panel"],
            border_color=palette["border"],
            border_width=1,
            corner_radius=18,
        )
        self.footer.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        self.footer.grid_columnconfigure(0, weight=1)
        self.page_status_var = tk.StringVar(value="")
        self.page_status_label = ctk.CTkLabel(self.footer, textvariable=self.page_status_var, anchor="w")
        self.page_status_label.grid(row=0, column=0, sticky="ew", padx=16, pady=12)
        self.btn_refresh = ctk.CTkButton(self.footer, text=self._t("btn_refresh"), width=150, height=36, corner_radius=18, command=self._refresh_status)
        self.btn_refresh.grid(row=0, column=1, padx=8, pady=12)
        self.btn_save_page = ctk.CTkButton(self.footer, text=self._t("btn_save"), width=150, height=36, corner_radius=18, command=self._save_current_page)
        self.btn_save_page.grid(row=0, column=2, padx=(0, 12), pady=12)

        self._switch_page(self._active_page, allow_prompt=False)
        self._apply_font_to_widgets(self, base_font)

    def _apply_global_font(self, family: str):
        self.option_add("*Font", (family, 11))
        for name in (
            "TkDefaultFont",
            "TkTextFont",
            "TkMenuFont",
            "TkHeadingFont",
            "TkIconFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkTooltipFont",
        ):
            try:
                tkfont.nametofont(name).configure(family=family)
            except tk.TclError:
                pass

    def _build_page_shell(self, key, title_key, subtitle_key):
        page = ctk.CTkFrame(self.page_host, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        page.grid_columnconfigure(0, weight=1)
        self._page_frames[key] = page
        hero = SectionCard(page, self._theme_palette, title=self._t(title_key), subtitle=self._t(subtitle_key))
        hero.grid(row=0, column=0, sticky="ew")
        return page, hero

    def _apply_font_to_widgets(self, root, family: str):
        for child in root.winfo_children():
            if isinstance(
                child,
                (
                    ctk.CTkLabel,
                    ctk.CTkButton,
                    ctk.CTkEntry,
                    ctk.CTkTextbox,
                    ctk.CTkOptionMenu,
                    ctk.CTkComboBox,
                    ctk.CTkCheckBox,
                    ctk.CTkRadioButton,
                    ctk.CTkSegmentedButton,
                ),
            ):
                try:
                    current_font = child.cget("font")
                    if isinstance(current_font, ctk.CTkFont):
                        child.configure(font=ctk.CTkFont(family=family, size=current_font.cget("size"), weight=current_font.cget("weight")))
                    else:
                        child.configure(font=ctk.CTkFont(family=family, size=11))
                except Exception:
                    pass
            
            if isinstance(child, ctk.CTkTabview):
                try:
                    child._segmented_button.configure(font=ctk.CTkFont(family=family, size=11))
                except Exception:
                    pass

            self._apply_font_to_widgets(child, family)

    def _build_perf_page(self):
        page, hero = self._build_page_shell("perf", "nav_perf", "status_subtitle")
        form = ctk.CTkFrame(hero, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 18))
        form.grid_columnconfigure(1, weight=1)

        self.poll_var = tk.StringVar(value="1000")
        self.debounce_var = tk.StringVar(value="4")
        self.lod_var = tk.StringVar(value="1mm")
        self.dpi_idx_var = tk.StringVar(value="1")

        self.lbl_poll, self.poll_combo = self._create_form_row(form, 0, self._t("poll_rate"), self.poll_var, values=["125", "250", "500", "1000"])
        self.lbl_debounce, self.debounce_entry = self._create_form_row(form, 1, self._t("debounce"), self.debounce_var)
        self.lbl_lod, self.lod_combo = self._create_form_row(form, 2, self._t("lod"), self.lod_var, values=["1mm", "2mm"])
        self.lbl_dpi_idx, self.dpi_idx_combo = self._create_form_row(form, 3, self._t("dpi_active"), self.dpi_idx_var, values=["1", "2", "3", "4", "5", "6"])

        self.dpi_frame = SectionCard(page, self._theme_palette, title=self._t("dpi_frame"))
        self.dpi_frame.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        self.dpi_frame.grid_columnconfigure((0, 1), weight=1)
        self.dpi_lbls = []
        self.dpi_vars = []
        self.dpi_entries = []
        for index in range(6):
            card = ctk.CTkFrame(self.dpi_frame, fg_color=self._theme_palette["panel_alt"], corner_radius=16)
            card.grid(row=1 + index // 2, column=index % 2, sticky="ew", padx=18 if index % 2 == 0 else (8, 18), pady=8)
            card.grid_columnconfigure(0, weight=1)
            label = ctk.CTkLabel(card, text=self._t("dpi_label", index + 1), anchor="w")
            label.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
            value = tk.StringVar(value="400")
            entry = ctk.CTkEntry(card, textvariable=value)
            entry.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 12))
            self.dpi_lbls.append(label)
            self.dpi_vars.append(value)
            self.dpi_entries.append(entry)

    def _build_sys_page(self):
        page, hero = self._build_page_shell("sys", "nav_sys", "status_subtitle")
        form = ctk.CTkFrame(hero, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 18))
        form.grid_columnconfigure(1, weight=1)
        self.light_var = tk.StringVar(value="0")
        self.sleep_var = tk.StringVar(value="10")
        self.lbl_light, self.light_combo = self._create_form_row(form, 0, self._t("light_mode"), self.light_var, values=[str(index) for index in range(8)])
        self.lbl_sleep, self.sleep_entry = self._create_form_row(form, 1, self._t("sleep_time"), self.sleep_var)

    def _build_keys_page(self):
        page, hero = self._build_page_shell("keys", "keys_title", "keys_subtitle")
        
        actions_frame = ctk.CTkFrame(hero.header_frame, fg_color="transparent")
        actions_frame.grid(row=0, column=1, rowspan=2, sticky="e", padx=(10, 0))
        self.reset_keys_btn = ctk.CTkButton(actions_frame, text=self._t("keys_reset"), height=36, corner_radius=18, fg_color=self._theme_palette["danger"], hover_color="#dc2626", command=self._reset_key_mapping_on_device)
        self.reset_keys_btn.pack(side="right")

        page.grid_rowconfigure(0, weight=1)
        hero.grid(row=0, column=0, sticky="nsew")
        hero.grid_rowconfigure(1, weight=1)

        layout = ctk.CTkFrame(hero, fg_color="transparent")
        layout.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        layout.grid_rowconfigure(0, weight=1)
        layout.grid_columnconfigure(1, weight=1)
        layout.grid_columnconfigure(2, weight=1)

        self.keys_slot_frame = SectionCard(layout, self._theme_palette, title=self._t("keys_select"))
        self.keys_slot_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.keys_slot_frame.grid_rowconfigure(1, weight=1)
        self.keys_slot_scroll = ctk.CTkScrollableFrame(self.keys_slot_frame, fg_color="transparent", width=250)
        self.keys_slot_scroll.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.key_slot_buttons = []
        for index, slot_name in enumerate(BUTTON_SLOT_NAMES):
            button = ctk.CTkButton(
                self.keys_slot_scroll,
                text=f"{slot_name}: -",
                anchor="w",
                height=40,
                corner_radius=20,
                command=lambda idx=index: self._select_button_slot(idx),
            )
            button.pack(fill="x", pady=4)
            self.key_slot_buttons.append(button)

        center = ctk.CTkFrame(layout, fg_color="transparent")
        center.grid(row=0, column=1, sticky="nsew", padx=(0, 12))
        center.grid_rowconfigure(1, weight=1)
        center.grid_columnconfigure(0, weight=1)

        self.keys_current_frame = SectionCard(center, self._theme_palette, title=self._t("keys_current"))
        self.keys_current_frame.grid(row=0, column=0, sticky="ew")
        self.current_key_var = tk.StringVar(value="-")
        self.current_key_label = ctk.CTkLabel(
            self.keys_current_frame,
            textvariable=self.current_key_var,
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        )
        self.current_key_label.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 6))
        self.keys_hint_label = ctk.CTkLabel(
            self.keys_current_frame,
            text=self._ui_text("keys_hint"),
            text_color=self._theme_palette["text_muted"],
            justify="left",
            wraplength=560,
        )
        self.keys_hint_label.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 18))

        self.keys_presets_frame = SectionCard(center, self._theme_palette, title=self._t("keys_presets"))
        self.keys_presets_frame.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        self.keys_presets_frame.grid_rowconfigure(1, weight=1)
        self.keys_presets_frame.grid_columnconfigure(0, weight=1)
        self.keys_preset_notebook = ctk.CTkTabview(self.keys_presets_frame, segmented_button_selected_color=self._theme_palette["accent"])
        self.keys_preset_notebook.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self.keys_keyboard_tab = self.keys_preset_notebook.add(self._t("keys_keyboard"))
        self.keys_media_tab = self.keys_preset_notebook.add(self._t("keys_media"))
        self.keys_mouse_tab = self.keys_preset_notebook.add(self._t("keys_mouse"))
        self.keyboard_preset_list = self._build_preset_list(self.keys_keyboard_tab, KEYBOARD_PRESETS)
        self.media_preset_list = self._build_preset_list(self.keys_media_tab, MEDIA_PRESETS)
        self.mouse_preset_list = self._build_preset_list(self.keys_mouse_tab, MOUSE_PRESETS)
        for listbox in (self.keyboard_preset_list, self.media_preset_list, self.mouse_preset_list):
            listbox.bind("<Double-Button-1>", lambda _event: self._apply_selected_preset())
            listbox.bind("<Return>", lambda _event: self._apply_selected_preset())

        right = ctk.CTkFrame(layout, fg_color="transparent")
        right.grid(row=0, column=2, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)

        self.keys_modifier_frame = SectionCard(right, self._theme_palette, title=self._t("keys_modifier"))
        self.keys_modifier_frame.grid(row=0, column=0, sticky="ew")
        toggles = ctk.CTkFrame(self.keys_modifier_frame, fg_color="transparent")
        toggles.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        self.modifier_ctrl = tk.BooleanVar()
        self.modifier_shift = tk.BooleanVar()
        self.modifier_alt = tk.BooleanVar()
        self.modifier_win = tk.BooleanVar()
        for column, (text, variable) in enumerate((("Ctrl", self.modifier_ctrl), ("Shift", self.modifier_shift), ("Alt", self.modifier_alt), ("Win", self.modifier_win))):
            checkbox = ctk.CTkCheckBox(toggles, text=text, variable=variable)
            checkbox.grid(row=0, column=column, padx=(0, 8), pady=0, sticky="w")
        self.capture_hint_label = ctk.CTkLabel(self.keys_modifier_frame, text=self._t("keys_capture"), anchor="w")
        self.capture_hint_label.grid(row=2, column=0, sticky="ew", padx=18)
        self.modifier_capture_var = tk.StringVar(value="-")
        self.modifier_capture_entry = ctk.CTkEntry(self.keys_modifier_frame, textvariable=self.modifier_capture_var)
        self.modifier_capture_entry.grid(row=3, column=0, sticky="ew", padx=18, pady=(6, 18))
        self.modifier_capture_entry.bind("<KeyPress>", self._capture_modifier_key)
        self.captured_modifier_hid = None
        for variable in (self.modifier_ctrl, self.modifier_shift, self.modifier_alt, self.modifier_win):
            variable.trace_add("write", lambda *_args: self._apply_modifier_combo())

        self.keys_macro_frame = SectionCard(right, self._theme_palette, title=self._t("keys_macro"))
        self.keys_macro_frame.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        self.keys_macro_frame.grid_columnconfigure(1, weight=1)
        self.key_macro_profile_label = ctk.CTkLabel(self.keys_macro_frame, text=self._t("keys_macro_profile"), anchor="w")
        self.key_macro_profile_label.grid(row=1, column=0, sticky="w", padx=18)
        self.key_macro_profile_var = tk.StringVar(value="")
        self.key_macro_profile_values = [profile.name for profile in self.macro_profiles]
        self.key_macro_profile_combo = ctk.CTkComboBox(self.keys_macro_frame, variable=self.key_macro_profile_var, values=self.key_macro_profile_values, command=lambda _value: self._apply_macro_assignment())
        self.key_macro_profile_combo.grid(row=1, column=1, sticky="ew", padx=(8, 18), pady=(0, 10))
        self.key_macro_mode_label = ctk.CTkLabel(self.keys_macro_frame, text=self._t("keys_macro_mode"), anchor="w")
        self.key_macro_mode_label.grid(row=2, column=0, sticky="w", padx=18)
        self.key_macro_mode_var = tk.StringVar()
        self.key_macro_mode_code = tk.IntVar(value=0)
        self.macro_mode_display = {}
        self.key_macro_mode_combo = ctk.CTkComboBox(self.keys_macro_frame, variable=self.key_macro_mode_var, values=[], command=lambda _value: (self._sync_macro_mode_selection(), self._apply_macro_assignment()))
        self.key_macro_mode_combo.grid(row=2, column=1, sticky="ew", padx=(8, 18), pady=(0, 10))
        self.key_macro_repeat_label = ctk.CTkLabel(self.keys_macro_frame, text=self._t("keys_macro_repeat"), anchor="w")
        self.key_macro_repeat_label.grid(row=3, column=0, sticky="w", padx=18)
        self.key_macro_repeat_var = tk.StringVar(value="1")
        self.key_macro_repeat_spinbox = ctk.CTkEntry(self.keys_macro_frame, textvariable=self.key_macro_repeat_var)
        self.key_macro_repeat_spinbox.grid(row=3, column=1, sticky="ew", padx=(8, 18), pady=(0, 18))
        self.key_macro_repeat_spinbox.bind("<Return>", self._apply_macro_assignment)
        self.key_macro_repeat_spinbox.bind("<FocusOut>", self._apply_macro_assignment)

    def _build_macro_page(self):
        page, hero = self._build_page_shell("macro", "macro_title", "macro_subtitle")
        
        actions_frame = ctk.CTkFrame(hero.header_frame, fg_color="transparent")
        actions_frame.grid(row=0, column=1, rowspan=2, sticky="e", padx=(10, 0))
        self.reset_macro_btn = ctk.CTkButton(actions_frame, text=self._t("macro_reset"), height=36, corner_radius=18, fg_color=self._theme_palette["danger"], hover_color="#dc2626", command=self._reset_macro_data_on_device)
        self.reset_macro_btn.pack(side="right")

        self.macro_hint_label = ctk.CTkLabel(hero, text=self._ui_text("macro_hint"), text_color=self._theme_palette["text_muted"], justify="left")
        self.macro_hint_label.grid(row=1, column=0, sticky="ew", padx=18)
        self.macro_status_var = tk.StringVar(value="")
        self.macro_status_label = ctk.CTkLabel(hero, textvariable=self.macro_status_var, anchor="w")
        self.macro_status_label.grid(row=2, column=0, sticky="ew", padx=18, pady=(8, 0))
        self.macro_progress_row = ctk.CTkFrame(hero, fg_color="transparent")
        self.macro_progress_row.grid(row=3, column=0, sticky="ew", padx=18, pady=(8, 18))
        self.macro_progress_row.grid_columnconfigure(0, weight=1)
        self.macro_progress = ctk.CTkProgressBar(self.macro_progress_row)
        self.macro_progress.grid(row=0, column=0, sticky="ew")
        self.macro_progress_label_var = tk.StringVar(value="")
        self.macro_progress_label = ctk.CTkLabel(self.macro_progress_row, textvariable=self.macro_progress_label_var, width=100, anchor="e")
        self.macro_progress_label.grid(row=0, column=1, padx=(10, 0))

        page.grid_rowconfigure(1, weight=1)
        body = ctk.CTkFrame(page, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_columnconfigure(2, weight=1)

        self.macro_profiles_frame = SectionCard(body, self._theme_palette, title=self._t("macro_profiles"))
        self.macro_profiles_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self.macro_profiles_frame.grid_rowconfigure(1, weight=1)
        self.macro_profiles_scroll = ctk.CTkScrollableFrame(self.macro_profiles_frame, fg_color="transparent", width=250)
        self.macro_profiles_scroll.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        for index in range(32):
            button = ctk.CTkButton(
                self.macro_profiles_scroll,
                text=f"{index + 1:02d}. Macro {index + 1}",
                anchor="w",
                height=40,
                corner_radius=20,
                command=lambda idx=index: self._on_macro_profile_selected(idx),
            )
            button.pack(fill="x", pady=4)
            self._macro_profile_buttons.append(button)

        self.macro_events_frame = SectionCard(body, self._theme_palette, title=self._t("macro_events"))
        self.macro_events_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 12))
        self.macro_events_frame.grid_rowconfigure(1, weight=1)
        self.macro_tree = ttk.Treeview(self.macro_events_frame, columns=("type", "name", "action", "delay"), show="headings", style="Ajazz.Treeview", height=15)
        for column, width in (("type", 100), ("name", 220), ("action", 100), ("delay", 90)):
            self.macro_tree.column(column, width=width, anchor=tk.CENTER if column != "name" else tk.W)
            self.macro_tree.heading(column, text="")
        self.macro_tree.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 8))
        self.macro_tree.bind("<<TreeviewSelect>>", self._on_macro_event_selected)
        self.macro_tree.bind("<Double-1>", self._on_macro_tree_double_click)
        self.macro_edit_hint_label = ctk.CTkLabel(self.macro_events_frame, text=self._ui_text("macro_edit_hint"), text_color=self._theme_palette["text_muted"], justify="left")
        self.macro_edit_hint_label.grid(row=2, column=0, sticky="ew", padx=18)
        event_actions = ctk.CTkFrame(self.macro_events_frame, fg_color="transparent")
        event_actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(10, 10))
        for column in range(4):
            event_actions.grid_columnconfigure(column, weight=1)
        self.macro_move_up_btn = ctk.CTkButton(event_actions, text=self._t("macro_move_up"), height=36, corner_radius=18, command=self._move_macro_event_up)
        self.macro_move_up_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.macro_move_down_btn = ctk.CTkButton(event_actions, text=self._t("macro_move_down"), height=36, corner_radius=18, command=self._move_macro_event_down)
        self.macro_move_down_btn.grid(row=0, column=1, sticky="ew", padx=6)
        self.macro_delete_btn = ctk.CTkButton(event_actions, text=self._t("macro_delete"), height=36, corner_radius=18, fg_color=self._theme_palette["warning"], hover_color=self._theme_palette["accent_hover"], command=self._delete_macro_event)
        self.macro_delete_btn.grid(row=0, column=2, sticky="ew", padx=6)
        self.macro_clear_btn = ctk.CTkButton(event_actions, text=self._t("macro_clear"), height=36, corner_radius=18, fg_color=self._theme_palette["danger"], hover_color="#dc2626", command=self._clear_macro_profile)
        self.macro_clear_btn.grid(row=0, column=3, sticky="ew", padx=(6, 0))
        delay_row = ctk.CTkFrame(self.macro_events_frame, fg_color="transparent")
        delay_row.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 18))
        self.macro_delay_label = ctk.CTkLabel(delay_row, text=self._t("macro_delay"), anchor="w")
        self.macro_delay_label.pack(side="left")
        self.macro_delay_var = tk.StringVar(value="10")
        self.macro_delay_spinbox = ctk.CTkEntry(delay_row, textvariable=self.macro_delay_var, width=120)
        self.macro_delay_spinbox.pack(side="left", padx=8)

        self.macro_controls_frame = SectionCard(body, self._theme_palette, title=self._t("macro_controls"))
        self.macro_controls_frame.grid(row=0, column=2, sticky="nsew")
        self.macro_controls_frame.grid_columnconfigure(0, weight=1)
        self.macro_name_label = ctk.CTkLabel(self.macro_controls_frame, text=self._t("macro_name"), anchor="w")
        self.macro_name_label.grid(row=1, column=0, sticky="ew", padx=18)
        self.macro_name_var = tk.StringVar(value="")
        self.macro_name_entry = ctk.CTkEntry(self.macro_controls_frame, textvariable=self.macro_name_var)
        self.macro_name_entry.grid(row=2, column=0, sticky="ew", padx=18, pady=(6, 12))
        self.record_btn = ctk.CTkButton(self.macro_controls_frame, text=self._t("macro_record_start"), height=40, corner_radius=20, command=self._toggle_recording)
        self.record_btn.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 16))

        self.delay_mode_frame = SectionCard(self.macro_controls_frame, self._theme_palette, title=self._t("macro_delay_mode"))
        self.delay_mode_frame.grid(row=4, column=0, sticky="ew", padx=18)
        self.record_delay_mode = tk.StringVar(value="exact")
        self.delay_exact_radio = ctk.CTkRadioButton(self.delay_mode_frame, text=self._t("macro_delay_exact"), variable=self.record_delay_mode, value="exact")
        self.delay_exact_radio.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 6))
        self.delay_none_radio = ctk.CTkRadioButton(self.delay_mode_frame, text=self._t("macro_delay_none"), variable=self.record_delay_mode, value="none")
        self.delay_none_radio.grid(row=2, column=0, sticky="w", padx=18, pady=(0, 6))
        fixed_row = ctk.CTkFrame(self.delay_mode_frame, fg_color="transparent")
        fixed_row.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 18))
        self.delay_fixed_radio = ctk.CTkRadioButton(fixed_row, text=self._t("macro_delay_fixed"), variable=self.record_delay_mode, value="fixed")
        self.delay_fixed_radio.pack(side="left")
        self.fixed_delay_label = ctk.CTkLabel(fixed_row, text=self._t("macro_fixed_ms"))
        self.fixed_delay_label.pack(side="left", padx=(8, 6))
        self.fixed_delay_var = tk.StringVar(value="10")
        self.fixed_delay_entry = ctk.CTkEntry(fixed_row, textvariable=self.fixed_delay_var, width=90)
        self.fixed_delay_entry.pack(side="left")

        self.manual_key_label = ctk.CTkLabel(self.macro_controls_frame, text=self._t("macro_manual_key"), anchor="w")
        self.manual_key_label.grid(row=5, column=0, sticky="ew", padx=18, pady=(16, 4))
        self.manual_key_map = dict(KEYBOARD_NAME_TO_CODE)
        self.manual_key_var = tk.StringVar(value=KEYBOARD_PRESETS[0]["name"])
        self.manual_key_combo = ctk.CTkComboBox(self.macro_controls_frame, variable=self.manual_key_var, values=list(self.manual_key_map.keys()))
        self.manual_key_combo.grid(row=6, column=0, sticky="ew", padx=18)
        self.add_key_pair_btn = ctk.CTkButton(self.macro_controls_frame, text=self._t("macro_add_key"), height=36, corner_radius=18, command=self._add_manual_key_pair)
        self.add_key_pair_btn.grid(row=7, column=0, sticky="ew", padx=18, pady=(8, 12))
        self.manual_mouse_label = ctk.CTkLabel(self.macro_controls_frame, text=self._t("macro_manual_mouse"), anchor="w")
        self.manual_mouse_label.grid(row=8, column=0, sticky="ew", padx=18, pady=(0, 4))
        self.manual_mouse_map = dict(MOUSE_NAME_TO_CODE)
        self.manual_mouse_var = tk.StringVar(value="Mouse L")
        self.manual_mouse_combo = ctk.CTkComboBox(self.macro_controls_frame, variable=self.manual_mouse_var, values=list(self.manual_mouse_map.keys()))
        self.manual_mouse_combo.grid(row=9, column=0, sticky="ew", padx=18)
        self.add_mouse_pair_btn = ctk.CTkButton(self.macro_controls_frame, text=self._t("macro_add_mouse"), height=36, corner_radius=18, command=self._add_manual_mouse_pair)
        self.add_mouse_pair_btn.grid(row=10, column=0, sticky="ew", padx=18, pady=(8, 16))

        self._set_macro_status("idle")
        self._set_macro_progress(0, 0)
        self._hide_macro_progress()
        self.macro_name_var.trace_add("write", lambda *_args: self._on_macro_name_changed())
        self.macro_delay_spinbox.bind("<Return>", lambda _event: self._update_selected_macro_delay())
        self.macro_delay_spinbox.bind("<FocusOut>", lambda _event: self._update_selected_macro_delay())

    def _build_debug_page(self):
        page, hero = self._build_page_shell("debug", "debug_title", "debug_subtitle")
        page.grid_rowconfigure(0, weight=1)
        hero.grid(row=0, column=0, sticky="nsew")
        hero.grid_rowconfigure(1, weight=1)
        self.log_txt = ctk.CTkTextbox(hero, corner_radius=16, fg_color=self._theme_palette["input"], border_color=self._theme_palette["border"], border_width=1)
        self.log_txt.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self.log_txt.configure(state="disabled")

    def _create_form_row(self, parent, row, label_text, variable, values=None):
        label = ctk.CTkLabel(parent, text=label_text, anchor="w")
        label.grid(row=row, column=0, sticky="w", pady=8)
        if values:
            widget = ctk.CTkComboBox(parent, variable=variable, values=values)
        else:
            widget = ctk.CTkEntry(parent, textvariable=variable)
        widget.grid(row=row, column=1, sticky="ew", pady=8)
        return label, widget

    def _build_preset_list(self, parent, presets):
        frame = tk.Frame(parent, bg=self._theme_palette["input"], highlightbackground=self._theme_palette["border"], highlightthickness=1, bd=0)
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        listbox = tk.Listbox(
            frame,
            height=12,
            exportselection=False,
            bg=self._theme_palette["input"],
            fg=self._theme_palette["text"],
            selectbackground=self._theme_palette["accent"],
            selectforeground=self._theme_palette["window"],
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
        )
        scrollbar = ctk.CTkScrollbar(frame, orientation="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        for preset in presets:
            listbox.insert(tk.END, preset["name"])
        return listbox

    def _capture_ui_snapshot(self):
        snapshot = {"active_page": self._active_page}
        if hasattr(self, "poll_var"):
            snapshot["perf"] = {
                "poll": self.poll_var.get(),
                "debounce": self.debounce_var.get(),
                "lod": self.lod_var.get(),
                "dpi_index": self.dpi_idx_var.get(),
                "dpis": [var.get() for var in self.dpi_vars],
            }
        if hasattr(self, "light_var"):
            snapshot["sys"] = {"light": self.light_var.get(), "sleep": self.sleep_var.get()}
        return snapshot

    def _restore_ui_snapshot(self, snapshot):
        if snapshot.get("perf") and hasattr(self, "poll_var"):
            perf = snapshot["perf"]
            self.poll_var.set(perf["poll"])
            self.debounce_var.set(perf["debounce"])
            self.lod_var.set(perf["lod"])
            self.dpi_idx_var.set(perf["dpi_index"])
            for index, value in enumerate(perf["dpis"]):
                if index < len(self.dpi_vars):
                    self.dpi_vars[index].set(value)
        elif self.current_config:
            self._load_config_to_ui()
        if snapshot.get("sys") and hasattr(self, "light_var"):
            self.light_var.set(snapshot["sys"]["light"])
            self.sleep_var.set(snapshot["sys"]["sleep"])
        elif self.current_config:
            self._load_config_to_ui()
        self._active_page = snapshot.get("active_page", self._active_page)

    def _rebuild_ui(self):
        snapshot = self._capture_ui_snapshot()
        for child in list(self.winfo_children()):
            child.destroy()
        self._page_frames = {}
        self._nav_buttons = {}
        self._macro_profile_buttons = []
        self._build_ui()
        self._restore_ui_snapshot(snapshot)
        self._refresh_all_views()

    def _refresh_all_views(self):
        self.title(self._t("title"))
        self._sync_status_display()
        self._sync_page_header()
        if self.current_config:
            self._load_config_to_ui()
        self._sync_macro_mode_labels()
        self._refresh_key_mapping_ui()
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()
        self.record_btn.configure(text=self._t("macro_record_stop") if self.is_recording else self._t("macro_record_start"))

    def _sync_status_display(self):
        self.metric_connection.set_text(label=self._t("hero_connection"), value=self._t(self._device_summary["status"]))
        self.metric_firmware.set_text(label=self._t("hero_firmware"), value=self._device_summary["version"])
        self.metric_battery.set_text(label=self._t("hero_battery"), value=self._device_summary["battery"])

    def _sync_page_header(self):
        label_map = {
            "perf": self._t("nav_perf"),
            "sys": self._t("nav_sys"),
            "keys": self._t("nav_keys"),
            "macro": self._t("nav_macro"),
            "debug": self._t("nav_debug"),
        }
        _ = label_map.get(self._active_page, self._active_page.title())

    def _sync_macro_mode_labels(self):
        self.macro_mode_display = {code: self._t(label_key) for code, label_key in MACRO_MODE_OPTIONS}
        self.key_macro_mode_combo.configure(values=list(self.macro_mode_display.values()))
        self.key_macro_mode_var.set(self.macro_mode_display.get(self.key_macro_mode_code.get(), self._t("macro_mode_counted")))
        self.macro_tree.heading("type", text=self._t("macro_type"))
        self.macro_tree.heading("name", text=self._t("macro_name_col"))
        self.macro_tree.heading("action", text=self._t("macro_action"))
        self.macro_tree.heading("delay", text=self._t("macro_delay_col"))

    def _set_macro_status(self, state_key):
        self._macro_status_key = state_key
        if hasattr(self, "macro_status_var"):
            self.macro_status_var.set(f"{self._t('macro_status_label')} {self._t(f'macro_status_{state_key}')}")

    def _set_macro_progress(self, current, total):
        self._macro_progress_current = max(0, int(current))
        self._macro_progress_total = max(0, int(total))
        if not hasattr(self, "macro_progress"):
            return
        maximum = self._macro_progress_total or 1
        current = min(self._macro_progress_current, maximum)
        self.macro_progress.set(0 if maximum <= 0 else current / maximum)
        percent = 0 if self._macro_progress_total <= 0 else round((current / self._macro_progress_total) * 100)
        self.macro_progress_label_var.set(f"{percent}% ({current}/{self._macro_progress_total})" if self._macro_progress_total > 0 else "")

    def _show_macro_progress(self):
        if hasattr(self, "macro_progress_row"):
            self.macro_progress_row.grid()

    def _hide_macro_progress(self):
        if hasattr(self, "macro_progress_row"):
            self.macro_progress_row.grid_remove()

    def _on_theme_selected(self):
        self._apply_theme(self.theme_var.get())

    def _apply_theme(self, theme_name, persist=True):
        self.theme_name = theme_name if theme_name in {"light", "dark"} else "dark"
        self._theme_palette = get_palette(self.theme_name)
        ctk.set_appearance_mode(self.theme_name)
        if persist:
            self._save_app_settings()
        self._rebuild_ui()

    def _load_app_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as file:
                    cfg = json.load(file)
                self.lang = cfg.get("lang", "en")
                self.theme_name = "dark"
            except Exception:
                self.lang = "en"
                self.theme_name = "dark"

    def _save_app_settings(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as file:
                json.dump({"lang": self.lang, "theme": "dark"}, file, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_macro_metadata(self):
        if os.path.exists(self.macro_meta_file):
            try:
                with open(self.macro_meta_file, "r", encoding="utf-8") as file:
                    self._macro_metadata = json.load(file)
            except Exception:
                self._macro_metadata = {"names": {}}

    def _save_macro_metadata(self):
        try:
            with open(self.macro_meta_file, "w", encoding="utf-8") as file:
                json.dump(self._macro_metadata, file, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _hide_window(self):
        self.withdraw()

    def _show_window(self):
        self.deiconify()
        self.lift()
        self.focus_force()

    def _exit_app(self, icon, item):
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)

    def _create_tray_menu(self):
        return pystray.Menu(
            pystray.MenuItem(self._t("tray_open"), lambda: self.after(0, self._show_window), default=True),
            pystray.MenuItem(self._t("tray_exit"), self._exit_app),
        )

    def _create_image(self, battery_perc):
        width, height = 64, 64
        if battery_perc < 0:
            bg_color, fg_color = "#475569", "#ffffff"
        elif battery_perc <= 20:
            bg_color, fg_color = "#111827", "#ef4444"
        else:
            bg_color, fg_color = "#111827", "#ffffff"
        image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle([2, 2, width - 2, height - 2], radius=18, fill=bg_color, outline="#ffffff")
        text = "--" if battery_perc < 0 else str(battery_perc)
        try:
            font = ImageFont.truetype("arial.ttf", 32)
        except Exception:
            font = ImageFont.load_default()
        try:
            left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
            text_width = right - left
            text_height = bottom - top
        except AttributeError:
            text_width, text_height = draw.textsize(text, font=font)
        draw.text(((width - text_width) // 2, (height - text_height) // 2 - 2), text, fill=fg_color, font=font)
        return image

    def _tray_monitor_loop(self):
        self.tray_icon = pystray.Icon("AjazzAJ139", self._create_image(-1), "AJ139 V2", self._create_tray_menu())

        def setup(icon):
            icon.visible = True
            while icon.visible:
                try:
                    if self._busy_count or self._macro_profiles_loading or self._status_refresh_in_progress:
                        time.sleep(1)
                        continue
                    if self.mouse.device and self.mouse.is_online():
                        battery = self.mouse.get_battery_info()["battery"]
                        if battery != self._last_tray_battery:
                            self._last_tray_battery = battery
                            icon.icon = self._create_image(battery)
                    elif self._last_tray_battery != -1:
                        self._last_tray_battery = -1
                        icon.icon = self._create_image(-1)
                except Exception:
                    pass
                time.sleep(10)

        self.tray_icon.run(setup)

    def _append_log(self, text):
        if hasattr(self, "log_txt") and self.log_txt.winfo_exists():
            self.log_txt.configure(state="normal")
            self.log_txt.insert("end", text + "\n")
            self.log_txt.see("end")
            self.log_txt.configure(state="disabled")

    def _set_busy(self, busy: bool):
        self._busy_count = max(0, self._busy_count + (1 if busy else -1))
        state = "disabled" if self._busy_count else "normal"
        widget_names = (
            "btn_save_page",
            "btn_refresh",
            "reset_keys_btn",
            "record_btn",
            "add_key_pair_btn",
            "add_mouse_pair_btn",
            "macro_move_up_btn",
            "macro_move_down_btn",
            "macro_delete_btn",
            "macro_clear_btn",
            "reload_macro_btn",
            "reset_macro_btn",
        )
        for name in widget_names:
            widget = getattr(self, name, None)
            if widget is not None:
                try:
                    widget.configure(state=state)
                except tk.TclError:
                    pass
        self.configure(cursor="watch" if self._busy_count else "")

    def _run_in_background(self, worker, on_success=None, on_error=None, busy=True):
        if busy:
            self._set_busy(True)

        def finish_success(result):
            if busy:
                self._set_busy(False)
            if on_success:
                on_success(result)

        def finish_error(exc):
            if busy:
                self._set_busy(False)
            if on_error:
                on_error(exc)
            else:
                messagebox.showerror(self._t("error"), str(exc))

        def runner():
            try:
                result = worker()
            except Exception as exc:
                self.after(0, lambda exc=exc: finish_error(exc))
            else:
                self.after(0, lambda result=result: finish_success(result))

        threading.Thread(target=runner, daemon=True).start()

    def _switch_page(self, page, allow_prompt=True):
        if allow_prompt:
            previous_page = self._active_page
            if previous_page and previous_page != page and self._page_is_dirty(previous_page):
                answer = messagebox.askyesnocancel(self._t("unsaved_changes"), self._t("unsaved_prompt"))
                if answer is None:
                    return
                if answer:
                    self._save_page(previous_page)
                else:
                    self._discard_page_changes(previous_page)
        self._active_page = page
        for name, frame in self._page_frames.items():
            if name == page:
                frame.grid()
            else:
                frame.grid_remove()
        if page in self._page_frames:
            self._page_frames[page].tkraise()
        self._refresh_nav_buttons()
        self._sync_page_header()
        self._refresh_dirty_state()
        if page == "macro":
            self._load_macro_profiles_async()

    def _refresh_nav_buttons(self):
        for page, button in self._nav_buttons.items():
            dirty = self._page_is_dirty(page)
            is_active = page == self._active_page
            if is_active:
                button.configure(
                    fg_color=self._theme_palette["accent"],
                    hover_color=self._theme_palette["accent_hover"],
                    text_color=self._theme_palette["window"],
                    border_color=self._theme_palette["accent"],
                )
            elif dirty:
                button.configure(
                    fg_color=self._theme_palette["warning_soft"],
                    hover_color=self._theme_palette["warning"],
                    text_color=self._theme_palette["text"],
                    border_color=self._theme_palette["warning"],
                )
            else:
                button.configure(
                    fg_color=self._theme_palette["sidebar_panel"],
                    hover_color=self._theme_palette["panel_alt"],
                    text_color=self._theme_palette["text_on_dark"],
                    border_color=self._theme_palette["border"],
                )

    def change_language(self, new_lang, skip_ui_update=False):
        self.lang = new_lang
        self._save_app_settings()
        if not skip_ui_update:
            self._rebuild_ui()

    def _load_config_to_ui(self):
        cfg = self.current_config
        if not cfg:
            return
        self.poll_var.set({0: "125", 1: "250", 2: "500", 3: "1000"}.get(cfg["report_rate_idx"], "1000"))
        self.debounce_var.set(str(cfg["key_respond"]))
        self.lod_var.set({1: "1mm", 2: "2mm"}.get(cfg["lod_value"], "1mm"))
        self.light_var.set(str(cfg["light_mode"]))
        self.sleep_var.set(str(cfg["sleep_light"]))
        self.dpi_idx_var.set(str(cfg["dpi_index"] + 1))
        for index, value in enumerate(cfg["dpis"]):
            self.dpi_vars[index].set(str(value))

    def _saved_perf_snapshot(self):
        if not self.current_config:
            return None
        return {
            "poll": {0: "125", 1: "250", 2: "500", 3: "1000"}.get(self.current_config["report_rate_idx"], "1000"),
            "debounce": str(self.current_config["key_respond"]),
            "lod": {1: "1mm", 2: "2mm"}.get(self.current_config["lod_value"], "1mm"),
            "dpi_index": str(self.current_config["dpi_index"] + 1),
            "dpis": [str(value) for value in self.current_config["dpis"]],
        }

    def _current_perf_snapshot(self):
        return {
            "poll": self.poll_var.get(),
            "debounce": self.debounce_var.get(),
            "lod": self.lod_var.get(),
            "dpi_index": self.dpi_idx_var.get(),
            "dpis": [var.get() for var in self.dpi_vars],
        }

    def _saved_sys_snapshot(self):
        if not self.current_config:
            return None
        return {"light": str(self.current_config["light_mode"]), "sleep": str(self.current_config["sleep_light"])}

    def _current_sys_snapshot(self):
        return {"light": self.light_var.get(), "sleep": self.sleep_var.get()}

    def _page_is_dirty(self, page):
        if page == "perf":
            saved = self._saved_perf_snapshot()
            return bool(saved and saved != self._current_perf_snapshot())
        if page == "sys":
            saved = self._saved_sys_snapshot()
            return bool(saved and saved != self._current_sys_snapshot())
        if page == "keys":
            return self.mouse_keys != self._saved_mouse_keys
        if page == "macro":
            return self.macro_profiles != self._saved_macro_profiles
        return False

    def _set_field_dirty_style(self, widget, dirty, widget_type):
        if widget is None:
            return
        if widget_type in {"entry", "combo"}:
            widget.configure(border_color=self._theme_palette["warning"] if dirty else self._theme_palette["border"])
        elif widget_type == "button":
            widget.configure(
                fg_color=self._theme_palette["warning_soft"] if dirty else self._theme_palette["panel_alt"],
                text_color=self._theme_palette["text"],
                border_color=self._theme_palette["warning"] if dirty else self._theme_palette["border"],
                border_width=1,
            )
        elif widget_type == "label":
            widget.configure(text_color=self._theme_palette["warning"] if dirty else self._theme_palette["text_muted"])

    def _refresh_dirty_state(self):
        perf_saved = self._saved_perf_snapshot()
        perf_current = self._current_perf_snapshot() if hasattr(self, "poll_var") else None
        if perf_saved and perf_current:
            self._set_field_dirty_style(self.poll_combo, perf_current["poll"] != perf_saved["poll"], "combo")
            self._set_field_dirty_style(self.debounce_entry, perf_current["debounce"] != perf_saved["debounce"], "entry")
            self._set_field_dirty_style(self.lod_combo, perf_current["lod"] != perf_saved["lod"], "combo")
            self._set_field_dirty_style(self.dpi_idx_combo, perf_current["dpi_index"] != perf_saved["dpi_index"], "combo")
            for index, entry in enumerate(self.dpi_entries):
                self._set_field_dirty_style(entry, perf_current["dpis"][index] != perf_saved["dpis"][index], "entry")

        sys_saved = self._saved_sys_snapshot()
        sys_current = self._current_sys_snapshot() if hasattr(self, "light_var") else None
        if sys_saved and sys_current:
            self._set_field_dirty_style(self.light_combo, sys_current["light"] != sys_saved["light"], "combo")
            self._set_field_dirty_style(self.sleep_entry, sys_current["sleep"] != sys_saved["sleep"], "entry")

        dirty_slots = {index for index, (current, saved) in enumerate(zip(self.mouse_keys, self._saved_mouse_keys)) if current != saved}
        for index, button in enumerate(self.key_slot_buttons):
            selected = index == self.selected_button_index
            if selected:
                button.configure(fg_color=self._theme_palette["accent"], hover_color=self._theme_palette["accent_hover"], text_color=self._theme_palette["window"])
            elif index in dirty_slots:
                self._set_field_dirty_style(button, True, "button")
            else:
                button.configure(
                    fg_color=self._theme_palette["panel_alt"],
                    hover_color=self._theme_palette["panel"],
                    text_color=self._theme_palette["text"],
                    border_color=self._theme_palette["border"],
                    border_width=1,
                )

        for index, button in enumerate(self._macro_profile_buttons):
            changed = index < len(self._saved_macro_profiles) and self.macro_profiles[index] != self._saved_macro_profiles[index]
            selected = index == self.selected_macro_slot
            if selected:
                button.configure(fg_color=self._theme_palette["accent"], hover_color=self._theme_palette["accent_hover"], text_color=self._theme_palette["window"])
            elif changed:
                button.configure(fg_color=self._theme_palette["warning_soft"], hover_color=self._theme_palette["warning"], text_color=self._theme_palette["text"])
            else:
                button.configure(fg_color=self._theme_palette["panel_alt"], hover_color=self._theme_palette["panel"], text_color=self._theme_palette["text"])

        if self._page_is_dirty(self._active_page):
            self.page_status_var.set(self._t("dirty_hint"))
            self._set_field_dirty_style(self.page_status_label, True, "label")
        else:
            self.page_status_var.set(self._t("clean_hint"))
            self._set_field_dirty_style(self.page_status_label, False, "label")

        save_key = {"perf": "save_perf", "sys": "save_sys", "keys": "save_keys", "macro": "save_macro"}.get(self._active_page, "save_page")
        self.btn_save_page.configure(text=self._t(save_key), state="normal" if self._active_page in {"perf", "sys", "keys", "macro"} else "disabled")
        self._refresh_nav_buttons()

    def _discard_page_changes(self, page):
        if page in {"perf", "sys"} and self.current_config:
            self._load_config_to_ui()
        elif page == "keys":
            self.mouse_keys = [clone_binding(binding) for binding in self._saved_mouse_keys]
            self._refresh_key_mapping_ui()
        elif page == "macro":
            self.macro_profiles = deepcopy(self._saved_macro_profiles)
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _save_current_page(self):
        self._save_page(self._active_page)

    def _save_page(self, page):
        if page == "perf":
            self._save_perf_page()
        elif page == "sys":
            self._save_sys_page()
        elif page == "keys":
            self._write_key_mapping_to_device()
        elif page == "macro":
            self._write_macros_to_device()

    def _merge_macro_names(self, profiles: list[MacroProfile]):
        names = self._macro_metadata.get("names", {})
        merged = []
        for profile in profiles:
            override = names.get(str(profile.slot))
            merged.append(MacroProfile(slot=profile.slot, name=override or profile.name, trigger_mode=profile.trigger_mode, repeat_count=profile.repeat_count, list=deepcopy(profile.list)))
        return merged

    def _reset_current_page(self):
        if self._active_page == "keys":
            self._reset_key_mapping_on_device()
        elif self._active_page == "macro":
            self._reset_macro_data_on_device()

    def _set_buttons_loading_state(self, is_loading):
        buttons = []
        if hasattr(self, "reset_macro_btn"): buttons.append(self.reset_macro_btn)
        if hasattr(self, "reset_keys_btn"): buttons.append(self.reset_keys_btn)
        
        if hasattr(self, "macro_move_up_btn"): buttons.append(self.macro_move_up_btn)
        if hasattr(self, "macro_move_down_btn"): buttons.append(self.macro_move_down_btn)
        if hasattr(self, "macro_delete_btn"): buttons.append(self.macro_delete_btn)
        if hasattr(self, "macro_clear_btn"): buttons.append(self.macro_clear_btn)
        if hasattr(self, "record_btn"): buttons.append(self.record_btn)
        if hasattr(self, "add_key_pair_btn"): buttons.append(self.add_key_pair_btn)
        if hasattr(self, "add_mouse_pair_btn"): buttons.append(self.add_mouse_pair_btn)

        if is_loading:
            for btn in buttons:
                if not hasattr(btn, "_saved_fg"):
                    setattr(btn, "_saved_fg", btn.cget("fg_color"))
                btn.configure(
                    state="disabled",
                    fg_color=self._theme_palette["border"],
                    text_color_disabled=self._theme_palette["text"]
                )
        else:
            for btn in buttons:
                if hasattr(btn, "_saved_fg"):
                    btn.configure(state="normal", fg_color=btn._saved_fg)

    def _refresh_status(self):
        if self._status_refresh_in_progress:
            return
        self._status_refresh_in_progress = True
        self._set_buttons_loading_state(True)
        self.log_txt.configure(state="normal")
        self.log_txt.delete("1.0", "end")
        self.log_txt.configure(state="disabled")
        self._append_log(self._t("log_sys_search"))
        self._device_summary = {"status": "st_disconnected", "version": "--", "battery": "--"}
        self._sync_status_display()

        def worker():
            if not self.mouse.device and not self.mouse.connect():
                return {"connected": False}
            result = {
                "connected": True,
                "version": self.mouse.get_version(),
                "online": self.mouse.is_online(),
                "config": self.mouse.get_config(),
                "keys": self.mouse.get_mouse_keys(),
            }
            if result["online"]:
                result["battery_info"] = self.mouse.get_battery_info()
            return result

        def on_success(result):
            self._status_refresh_in_progress = False
            self._set_buttons_loading_state(False)
            self._macro_profiles_loaded = False
            self._macro_profiles_loading = False
            self._set_macro_status("idle")
            if not result.get("connected"):
                self._device_summary = {"status": "st_disconnected", "version": "--", "battery": "--"}
                self._sync_status_display()
                self._append_log(self._t("log_sys_fail"))
                return

            self._append_log(self._t("log_sys_conn"))
            battery_text = self._t("bat_offline")
            if result["online"]:
                battery_info = result["battery_info"]
                suffix = f" {self._t('bat_charging')}" if battery_info["is_charging"] else ""
                battery_text = f"{battery_info['battery']}%{suffix}"
                self._append_log(self._t("log_sys_bat", battery_info["battery"]))
                self._last_tray_battery = battery_info["battery"]
                if self.tray_icon:
                    self.tray_icon.icon = self._create_image(self._last_tray_battery)
            else:
                self._append_log(self._t("log_sys_off"))
                self._last_tray_battery = -1
                if self.tray_icon:
                    self.tray_icon.icon = self._create_image(-1)
            self._device_summary = {"status": "st_connected", "version": result["version"], "battery": battery_text}
            self._sync_status_display()

            self.current_config = result["config"]
            if self.current_config:
                self._load_config_to_ui()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in result["keys"]]
            self._saved_mouse_keys = [clone_binding(binding) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()

            if self._active_page == "macro":
                self._load_macro_profiles_async()

        def on_error(exc):
            self._status_refresh_in_progress = False
            self._set_buttons_loading_state(False)
            self._set_macro_status("error")
            self._device_summary = {"status": "st_disconnected", "version": "--", "battery": "--"}
            self._sync_status_display()
            self._append_log(f"ERROR: {exc}")

        self._run_in_background(worker, on_success=on_success, on_error=on_error)

    def _load_macro_profiles_async(self, force=False, log=True):
        if self._macro_profiles_loading or (self._macro_profiles_loaded and not force):
            return
        if not self.mouse.device:
            return
        self._macro_profiles_loading = True
        self._set_buttons_loading_state(True)
        self._set_macro_status("loading")
        self._set_macro_progress(0, 0)
        self._show_macro_progress()
        if log:
            self._append_log("System: Loading macro profiles from device...")

        def worker():
            return self._merge_macro_names(
                decode_macro_profiles(
                    self.mouse.get_macro_data(progress_callback=lambda current, total: self.after(0, lambda current=current, total=total: self._set_macro_progress(current, total))),
                    resolve_macro_event_name,
                )
            )

        def on_success(profiles):
            self._macro_profiles_loading = False
            self._set_buttons_loading_state(False)
            self._macro_profiles_loaded = True
            self._set_macro_status("ready")
            self._set_macro_progress(self._macro_progress_total, self._macro_progress_total)
            self._hide_macro_progress()
            self.macro_profiles = profiles
            self._saved_macro_profiles = deepcopy(profiles)
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()
            if log:
                self._append_log("System: Macro profiles loaded from device.")

        def on_error(exc):
            self._macro_profiles_loading = False
            self._set_buttons_loading_state(False)
            self._set_macro_status("error")
            self._hide_macro_progress()
            if log:
                self._append_log(f"ERROR: Failed to load macro data: {exc}")

        self._run_in_background(worker, on_success=on_success, on_error=on_error)

    def _load_macro_profiles_from_device(self, log=True):
        try:
            self.macro_profiles = self._merge_macro_names(decode_macro_profiles(self.mouse.get_macro_data(), resolve_macro_event_name))
            self._macro_profiles_loaded = True
            self._saved_macro_profiles = deepcopy(self.macro_profiles)
            self._set_macro_status("ready")
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self._refresh_dirty_state()
            if log:
                self._append_log("System: Macro profiles loaded from device.")
        except Exception as exc:
            if log:
                self._append_log(f"ERROR: Failed to load macro data: {exc}")

    def _load_key_mapping_from_device(self, log=True):
        try:
            loaded = self.mouse.get_mouse_keys()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in loaded]
            self._refresh_key_mapping_ui()
            if log:
                self._append_log("System: Key mapping loaded from device.")
        except Exception as exc:
            if log:
                self._append_log(f"ERROR: Failed to load key mapping: {exc}")

    def _select_button_slot(self, index):
        self.selected_button_index = index
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _selected_binding(self):
        return self.mouse_keys[self.selected_button_index]

    def _refresh_key_mapping_ui(self):
        if not hasattr(self, "key_slot_buttons"):
            return
        for index, button in enumerate(self.key_slot_buttons):
            button.configure(text=f"{BUTTON_SLOT_NAMES[index]}: {self.mouse_keys[index].name}")
        selected = self._selected_binding()
        self.current_key_var.set(f"{BUTTON_SLOT_NAMES[selected.index]} -> {selected.name}")
        if selected.type == 16:
            self.modifier_ctrl.set(bool(selected.code1 & 0x01 or selected.code1 & 0x10))
            self.modifier_shift.set(bool(selected.code1 & 0x02 or selected.code1 & 0x20))
            self.modifier_alt.set(bool(selected.code1 & 0x04 or selected.code1 & 0x40))
            self.modifier_win.set(bool(selected.code1 & 0x08 or selected.code1 & 0x80))
            self.captured_modifier_hid = selected.code2
            self.modifier_capture_var.set(HID_KEY_NAMES.get(selected.code2, "-"))
        self.key_macro_profile_values = [profile.name for profile in self.macro_profiles]
        self.key_macro_profile_combo.configure(values=self.key_macro_profile_values or [""])
        if selected.type == 112 and self.macro_profiles:
            slot = min(max(selected.code1, 0), len(self.macro_profiles) - 1)
            self.key_macro_profile_var.set(self.macro_profiles[slot].name)
            self.key_macro_repeat_var.set(str(max(1, selected.code2 or 1)))
            self.key_macro_mode_code.set(selected.code3)
            self.key_macro_mode_var.set(self.macro_mode_display.get(selected.code3, self._t("macro_mode_counted")))
        elif self.macro_profiles:
            self.key_macro_profile_var.set(self.macro_profiles[0].name)
            self.key_macro_repeat_var.set("1")
            self.key_macro_mode_code.set(0)
            self.key_macro_mode_var.set(self.macro_mode_display.get(0, self._t("macro_mode_counted")))
        self._refresh_dirty_state()

    def _selected_preset(self):
        current_tab = self.keys_preset_notebook.get()
        source = {
            self._t("keys_keyboard"): (self.keyboard_preset_list, KEYBOARD_PRESETS),
            self._t("keys_media"): (self.media_preset_list, MEDIA_PRESETS),
            self._t("keys_mouse"): (self.mouse_preset_list, MOUSE_PRESETS),
        }[current_tab]
        selection = source[0].curselection()
        return source[1][selection[0]] if selection else None

    def _apply_selected_preset(self):
        preset = self._selected_preset()
        if not preset:
            return
        self.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=preset["name"],
            type=preset["type"],
            code1=preset["code1"],
            code2=preset["code2"],
            code3=preset["code3"],
            lang=preset["lang"],
        )
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _capture_modifier_key(self, event):
        keysym = event.keysym
        hid = None
        if keysym in ("Escape", "BackSpace", "Tab", "space", "Return"):
            hid = {"Escape": 41, "BackSpace": 42, "Tab": 43, "space": 44, "Return": 40}[keysym]
        elif len(keysym) == 1 and keysym.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
            hid = 4 + ord(keysym.upper()) - ord("A") if keysym.isalpha() else 30 + "1234567890".index(keysym)
        hid = hid or next((code for code, name in HID_KEY_NAMES.items() if name == keysym), None)
        if hid is None:
            return "break"
        self.captured_modifier_hid = hid
        self.modifier_capture_var.set(HID_KEY_NAMES.get(hid, f"Key {hid}"))
        self._apply_modifier_combo()
        return "break"

    def _apply_modifier_combo(self):
        if self.captured_modifier_hid is None:
            return
        code1 = (
            (0x01 if self.modifier_ctrl.get() else 0)
            | (0x02 if self.modifier_shift.get() else 0)
            | (0x04 if self.modifier_alt.get() else 0)
            | (0x08 if self.modifier_win.get() else 0)
        )
        name = "+".join(modifier_names(code1) + [HID_KEY_NAMES.get(self.captured_modifier_hid, f"Key {self.captured_modifier_hid}")])
        self.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=name,
            type=16,
            code1=code1,
            code2=self.captured_modifier_hid,
            code3=0,
        )
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _sync_macro_mode_selection(self, event=None):
        reverse_map = {label: code for code, label in self.macro_mode_display.items()}
        self.key_macro_mode_code.set(reverse_map.get(self.key_macro_mode_var.get(), 0))

    def _apply_macro_assignment(self, event=None):
        slot = self.key_macro_profile_values.index(self.key_macro_profile_var.get()) if self.key_macro_profile_var.get() in self.key_macro_profile_values else 0
        profile = self.macro_profiles[slot]
        try:
            repeat_count = max(1, int(self.key_macro_repeat_var.get() or "1"))
        except ValueError:
            repeat_count = 1
            self.key_macro_repeat_var.set("1")
        self.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=profile.name,
            type=112,
            code1=slot,
            code2=repeat_count,
            code3=int(self.key_macro_mode_code.get()),
        )
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _write_key_mapping_to_device(self):
        if not self.mouse.device:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        self._append_log("System: Writing key mapping...")

        def worker():
            self.mouse.set_mouse_keys(self.mouse_keys)
            return None

        def on_success(_):
            self._saved_mouse_keys = [clone_binding(binding) for binding in self.mouse_keys]
            self._refresh_dirty_state()
            self._append_log("System: Key mapping written.")
            messagebox.showinfo(self._t("ok"), self._t("keys_write_ok"))

        self._run_in_background(worker, on_success=on_success)

    def _reset_key_mapping_on_device(self):
        if not self.mouse.device:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        if not messagebox.askyesno(self._t("confirm"), self._t("reset_confirm")):
            return
        self._append_log("System: Resetting key mapping...")

        def worker():
            self.mouse.reset_mouse_keys()
            return self.mouse.get_mouse_keys()

        def on_success(bindings):
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in bindings]
            self._saved_mouse_keys = [clone_binding(binding) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()
            self._append_log("System: Key mapping loaded from device.")

        self._run_in_background(worker, on_success=on_success)

    def _refresh_macro_profile_list(self):
        if not self._macro_profile_buttons:
            return
        for index, button in enumerate(self._macro_profile_buttons):
            profile = self.macro_profiles[index]
            suffix = f" ({len(profile.list)})" if profile.list else ""
            button.configure(text=f"{profile.slot + 1:02d}. {profile.name}{suffix}")
        self.macro_name_var.set(self.macro_profiles[self.selected_macro_slot].name)
        self._refresh_key_mapping_ui()

    def _refresh_macro_events(self):
        self._close_macro_editor()
        for item_id in self.macro_tree.get_children():
            self.macro_tree.delete(item_id)
        profile = self.macro_profiles[self.selected_macro_slot]
        if not profile.list:
            self.macro_tree.insert("", tk.END, values=("", self._t("macro_empty"), "", ""))
            self.selected_macro_event_index = None
            return
        for index, event in enumerate(profile.list):
            event_type = self._t("macro_event_mouse") if event.type == EVENT_TYPE_MOUSE else self._t("macro_event_keyboard")
            action = self._t("macro_action_press") if event.action == ACTION_PRESS else self._t("macro_action_release")
            self.macro_tree.insert("", tk.END, iid=str(index), values=(event_type, event.name, action, event.delay))
        if self.selected_macro_event_index is not None and self.selected_macro_event_index < len(profile.list):
            self.macro_tree.selection_set(str(self.selected_macro_event_index))

    def _on_macro_profile_selected(self, index=None, event=None):
        self._close_macro_editor()
        if index is not None:
            self.selected_macro_slot = index
            self.selected_macro_event_index = None
            self.macro_name_var.set(self.macro_profiles[self.selected_macro_slot].name)
            self._refresh_macro_events()
            self._refresh_dirty_state()

    def _on_macro_event_selected(self, event=None):
        selection = self.macro_tree.selection()
        if selection:
            try:
                self.selected_macro_event_index = int(selection[0])
                self.macro_delay_var.set(str(self.macro_profiles[self.selected_macro_slot].list[self.selected_macro_event_index].delay))
            except Exception:
                self.selected_macro_event_index = None

    def _close_macro_editor(self):
        if self._macro_editor_widget is not None:
            try:
                self._macro_editor_widget.destroy()
            except tk.TclError:
                pass
        self._macro_editor_widget = None
        self._macro_editor_info = None

    def _commit_macro_editor(self):
        info = self._macro_editor_info
        if not info:
            return
        profile = self._selected_profile()
        index = info["index"]
        if index >= len(profile.list):
            self._close_macro_editor()
            return
        macro_event = profile.list[index]
        value = info["var"].get().strip()
        column = info["column"]
        if column == "#2":
            code = KEYBOARD_NAME_TO_CODE.get(value) if macro_event.type == EVENT_TYPE_KEYBOARD else MOUSE_NAME_TO_CODE.get(value)
            if code is not None:
                macro_event.name = value
                macro_event.code = code
        elif column == "#3":
            macro_event.action = ACTION_PRESS if value == self._t("macro_action_press") else ACTION_RELEASE
        elif column == "#4":
            try:
                macro_event.delay = max(0, int(value or "0"))
            except ValueError:
                pass
        self.selected_macro_event_index = index
        self._close_macro_editor()
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _on_macro_tree_double_click(self, event):
        self._close_macro_editor()
        row_id = self.macro_tree.identify_row(event.y)
        column = self.macro_tree.identify_column(event.x)
        if not row_id or column not in {"#2", "#3", "#4"}:
            return
        try:
            index = int(row_id)
        except ValueError:
            return
        profile = self._selected_profile()
        if index >= len(profile.list):
            return
        bbox = self.macro_tree.bbox(row_id, column)
        if not bbox:
            return
        x, y, width, height = bbox
        macro_event = profile.list[index]
        if column == "#2":
            values = KEYBOARD_EVENT_CHOICES if macro_event.type == EVENT_TYPE_KEYBOARD else MOUSE_EVENT_CHOICES if macro_event.type == EVENT_TYPE_MOUSE else None
            if values is None:
                return
            variable = tk.StringVar(value=macro_event.name)
            editor = ttk.Combobox(self.macro_tree, textvariable=variable, values=values, state="readonly")
            editor.bind("<<ComboboxSelected>>", lambda _event: self._commit_macro_editor())
        elif column == "#3":
            variable = tk.StringVar(value=self._t("macro_action_press") if macro_event.action == ACTION_PRESS else self._t("macro_action_release"))
            editor = ttk.Combobox(self.macro_tree, textvariable=variable, values=[self._t("macro_action_press"), self._t("macro_action_release")], state="readonly")
            editor.bind("<<ComboboxSelected>>", lambda _event: self._commit_macro_editor())
        else:
            variable = tk.StringVar(value=str(macro_event.delay))
            editor = tk.Spinbox(self.macro_tree, from_=0, to=60000, textvariable=variable)
        self._macro_editor_widget = editor
        self._macro_editor_info = {"index": index, "column": column, "var": variable}
        editor.place(x=x, y=y, width=width, height=height)
        editor.bind("<Return>", lambda _event: self._commit_macro_editor())
        editor.bind("<Escape>", lambda _event: self._close_macro_editor())
        editor.bind("<FocusOut>", lambda _event: self.after(0, self._commit_macro_editor))
        editor.focus_set()

    def _on_macro_name_changed(self):
        name = self.macro_name_var.get().strip() or f"Macro {self.selected_macro_slot + 1}"
        self.macro_profiles[self.selected_macro_slot].name = name
        self._macro_metadata.setdefault("names", {})[str(self.selected_macro_slot)] = name
        self._refresh_macro_profile_list()
        self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in self.mouse_keys]
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _selected_profile(self):
        return self.macro_profiles[self.selected_macro_slot]

    def _append_event_to_profile(self, event: MacroEvent):
        self._selected_profile().list.append(event)
        self.selected_macro_event_index = len(self._selected_profile().list) - 1
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _add_event_pair(self, name, code, event_type):
        self._append_event_to_profile(MacroEvent(name=name, code=code, type=event_type, action=ACTION_PRESS, delay=10))
        self._append_event_to_profile(MacroEvent(name=name, code=code, type=event_type, action=ACTION_RELEASE, delay=10))

    def _add_manual_key_pair(self):
        key_name = self.manual_key_var.get()
        if key_name in self.manual_key_map:
            self._add_event_pair(key_name, self.manual_key_map[key_name], EVENT_TYPE_KEYBOARD)

    def _add_manual_mouse_pair(self):
        mouse_name = self.manual_mouse_var.get()
        if mouse_name in self.manual_mouse_map:
            self._add_event_pair(mouse_name, self.manual_mouse_map[mouse_name], EVENT_TYPE_MOUSE)

    def _move_macro_event_up(self):
        profile = self._selected_profile()
        index = self.selected_macro_event_index
        if index is None or index <= 0:
            return
        profile.list[index - 1], profile.list[index] = profile.list[index], profile.list[index - 1]
        self.selected_macro_event_index -= 1
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _move_macro_event_down(self):
        profile = self._selected_profile()
        index = self.selected_macro_event_index
        if index is None or index >= len(profile.list) - 1:
            return
        profile.list[index + 1], profile.list[index] = profile.list[index], profile.list[index + 1]
        self.selected_macro_event_index += 1
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _delete_macro_event(self):
        profile = self._selected_profile()
        index = self.selected_macro_event_index
        if index is None or index >= len(profile.list):
            return
        profile.list.pop(index)
        self.selected_macro_event_index = None if not profile.list else min(index, len(profile.list) - 1)
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _clear_macro_profile(self):
        self._selected_profile().list = []
        self.selected_macro_event_index = None
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _update_selected_macro_delay(self):
        index = self.selected_macro_event_index
        profile = self._selected_profile()
        if index is not None and index < len(profile.list):
            try:
                profile.list[index].delay = max(0, int(self.macro_delay_var.get() or "0"))
            except ValueError:
                return
            self._refresh_macro_events()
            self._refresh_dirty_state()

    def _record_delay_value(self, now_ms):
        if self.record_delay_mode.get() == "none":
            return 0
        if self.record_delay_mode.get() == "fixed":
            return max(10, int(self.fixed_delay_var.get() or "10"))
        return 0 if self._last_record_time is None else max(0, now_ms - self._last_record_time)

    def _append_recorded_event(self, name, code, event_type, action):
        profile = self._selected_profile()
        now_ms = int(time.time() * 1000)
        if profile.list:
            profile.list[-1].delay = self._record_delay_value(now_ms)
        profile.list.append(MacroEvent(name=name, code=code, type=event_type, action=action, delay=0))
        self._last_record_time = now_ms
        self.selected_macro_event_index = len(profile.list) - 1
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()

    def _event_target_allows_record(self, widget):
        return widget.winfo_class() not in {"TButton", "Button", "CTkButton", "TCombobox", "Combobox", "CTkComboBox", "Treeview", "Listbox"}

    def _hid_from_key_event(self, event):
        keysym = event.keysym
        if keysym in ("Escape", "BackSpace", "Tab", "space", "Return"):
            return {"Escape": 41, "BackSpace": 42, "Tab": 43, "space": 44, "Return": 40}[keysym]
        if len(keysym) == 1 and keysym.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
            return 4 + ord(keysym.upper()) - ord("A") if keysym.isalpha() else 30 + "1234567890".index(keysym)
        return next((code for code, name in HID_KEY_NAMES.items() if name == keysym), None)

    def _handle_record_key_press(self, event):
        if not self.is_recording:
            return
        hid = self._hid_from_key_event(event)
        if hid is None or hid in self._pressed_keys:
            return "break"
        self._pressed_keys.add(hid)
        self._append_recorded_event(HID_KEY_NAMES.get(hid, f"Key {hid}"), hid, EVENT_TYPE_KEYBOARD, ACTION_PRESS)
        return "break"

    def _handle_record_key_release(self, event):
        if not self.is_recording:
            return
        hid = self._hid_from_key_event(event)
        if hid is None:
            return "break"
        self._pressed_keys.discard(hid)
        self._append_recorded_event(HID_KEY_NAMES.get(hid, f"Key {hid}"), hid, EVENT_TYPE_KEYBOARD, ACTION_RELEASE)
        return "break"

    def _handle_record_mouse_press(self, event):
        if not self.is_recording or not self._event_target_allows_record(event.widget):
            return
        code = {1: 1, 2: 4, 3: 2}.get(event.num)
        if code is None or code in self._pressed_mouse:
            return "break"
        self._pressed_mouse.add(code)
        self._append_recorded_event(MOUSE_CODE_NAMES.get(code, f"Mouse {code}"), code, EVENT_TYPE_MOUSE, ACTION_PRESS)
        return "break"

    def _handle_record_mouse_release(self, event):
        if not self.is_recording or not self._event_target_allows_record(event.widget):
            return
        code = {1: 1, 2: 4, 3: 2}.get(event.num)
        if code is None:
            return "break"
        self._pressed_mouse.discard(code)
        self._append_recorded_event(MOUSE_CODE_NAMES.get(code, f"Mouse {code}"), code, EVENT_TYPE_MOUSE, ACTION_RELEASE)
        return "break"

    def _toggle_recording(self):
        self.is_recording = not self.is_recording
        self._pressed_keys.clear()
        self._pressed_mouse.clear()
        self._last_record_time = None
        self.record_btn.configure(text=self._t("macro_record_stop") if self.is_recording else self._t("macro_record_start"))
        if self.is_recording:
            self.bind_all("<KeyPress>", self._handle_record_key_press)
            self.bind_all("<KeyRelease>", self._handle_record_key_release)
            for sequence in ("<ButtonPress-1>", "<ButtonPress-2>", "<ButtonPress-3>"):
                self.bind_all(sequence, self._handle_record_mouse_press)
            for sequence in ("<ButtonRelease-1>", "<ButtonRelease-2>", "<ButtonRelease-3>"):
                self.bind_all(sequence, self._handle_record_mouse_release)
        else:
            for sequence in (
                "<KeyPress>",
                "<KeyRelease>",
                "<ButtonPress-1>",
                "<ButtonPress-2>",
                "<ButtonPress-3>",
                "<ButtonRelease-1>",
                "<ButtonRelease-2>",
                "<ButtonRelease-3>",
            ):
                self.unbind_all(sequence)

    def _write_macros_to_device(self):
        if not self.mouse.device:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        self._append_log("System: Writing macro data...")
        self._set_macro_status("writing")

        def worker():
            self.mouse.set_macro_data(encode_macro_profiles(self.macro_profiles))
            return None

        def on_success(_):
            self._save_macro_metadata()
            self._macro_profiles_loaded = True
            self._saved_macro_profiles = deepcopy(self.macro_profiles)
            self._set_macro_status("ready")
            self._refresh_dirty_state()
            self._append_log("System: Macro data written.")
            messagebox.showinfo(self._t("ok"), self._t("macro_write_ok"))

        self._run_in_background(worker, on_success=on_success)

    def _reset_macro_data_on_device(self):
        if not self.mouse.device:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        if not messagebox.askyesno(self._t("confirm"), self._t("reset_confirm")):
            return
        self._append_log("System: Resetting macro memory...")
        self._set_macro_status("resetting")

        def worker():
            self.mouse.reset_macro_data()
            return self._merge_macro_names(decode_macro_profiles(self.mouse.get_macro_data(), resolve_macro_event_name))

        def on_success(profiles):
            self.macro_profiles = profiles
            self._macro_profiles_loaded = True
            self._saved_macro_profiles = deepcopy(profiles)
            self._set_macro_status("ready")
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()
            self._append_log("System: Macro profiles loaded from device.")

        self._run_in_background(worker, on_success=on_success)

    def _save_perf_page(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        cfg = deepcopy(self.current_config)
        cfg["report_rate_idx"] = {"125": 0, "250": 1, "500": 2, "1000": 3}.get(self.poll_var.get(), 3)
        cfg["key_respond"] = int(self.debounce_var.get() or "4")
        cfg["lod_value"] = 1 if self.lod_var.get() == "1mm" else 2
        cfg["dpi_index"] = int(self.dpi_idx_var.get() or "1") - 1
        cfg["dpis"] = [max(50, int(var.get() or "400")) for var in self.dpi_vars]
        self._append_log(self._t("log_sys_writing"))

        def on_success(_):
            self.current_config = cfg
            self._refresh_dirty_state()
            messagebox.showinfo(self._t("ok"), self._t("msg_success"))

        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=on_success)

    def _save_sys_page(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror(self._t("error"), self._t("msg_err_conn"))
            return
        cfg = deepcopy(self.current_config)
        cfg["sleep_light"] = int(self.sleep_var.get() or "10")
        cfg["light_mode"] = int(self.light_var.get() or "0")
        self._append_log(self._t("log_sys_writing"))

        def on_success(_):
            self.current_config = cfg
            self._refresh_dirty_state()
            messagebox.showinfo(self._t("ok"), self._t("msg_success"))

        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=on_success)
