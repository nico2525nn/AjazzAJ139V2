"""
ui_app.py – AjazzApp メインアプリ (CustomTkinter ネイティブUI化)

よりモダンでゲーミング向けな "サイドバーナビゲーション" デザイン。
すべての主要ウィジェットを CustomTkinter (CTk) に移行し、フォントも Noto Sans を指定。
一部の特殊ウィジェット(Listbox, Treeview)のみ tk/ttk カスタムスタイルでハイブリッド使用。
"""

import ctypes
import json
import os
import threading
import time
import tkinter as tk
from copy import deepcopy
from tkinter import font as tkfont
from tkinter import messagebox, ttk

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw, ImageFont

from ajazz_mouse import (
    AjazzMouse,
    MacroProfile,
    decode_macro_profiles,
    encode_macro_profiles,
)
from constants import (
    BUTTON_SLOT_NAMES,
    HID_KEY_NAMES,
    LANG_DICT,
    MACRO_MODE_OPTIONS,
    PALETTES,
    UI_TEXT,
)
from ui_helpers import (
    clone_binding,
    copy_default_bindings,
    resolve_binding_name,
    resolve_macro_event_name,
)
from ui_keymapping import KeyMappingTab
from ui_macro import MacroTab

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class AjazzApp(ctk.CTk):
    HID_KEY_NAMES = HID_KEY_NAMES

    def __init__(self):
        super().__init__()

        self.mouse = AjazzMouse(log_callback=self._append_log)
        self.current_config = None
        self.lang = "ja"
        self.theme_name = "dark"
        self.mouse_keys = copy_default_bindings()
        self.macro_profiles: list[MacroProfile] = [
            MacroProfile(slot=i, name=f"Macro {i + 1}") for i in range(32)
        ]
        self._macro_metadata: dict = {"names": {}}
        self._busy_count = 0
        self._status_refresh_in_progress = False
        self._macro_profiles_loaded = False
        self._macro_profiles_loading = False
        self._macro_status_key = "idle"
        self._macro_progress_current = 0
        self._macro_progress_total = 0
        self._saved_mouse_keys = copy_default_bindings()
        self._saved_macro_profiles = deepcopy(self.macro_profiles)
        self._tab_change_guard = False
        self._current_tab_id = "status"
        self._theme_palette: dict = PALETTES["dark"]

        self.config_file = "settings.json"
        self.macro_meta_file = "macro_metadata.json"
        self._load_app_settings()
        self._load_macro_metadata()

        self.title(self._t("title"))
        self.geometry("1300x900")
        self.minsize(1300, 900)

        self.tray_icon = None
        self._last_tray_battery = -1
        self.protocol("WM_DELETE_WINDOW", self._hide_window)

        # ── フォント設定 ──
        # Noto Sans JP をメインに指定。なければフォールバック。
        avail_fonts = tkfont.families()
        def get_best_font(choices):
            for c in choices:
                if c in avail_fonts:
                    return c
            return "MS Gothic" if os.name == "nt" else choices[-1]
            
        self.base_font_family = get_best_font(["Noto Sans JP", "Noto Sans JP Regular", "Noto Sans CJK JP", "Meiryo UI", "Yu Gothic UI", "Segoe UI"])
        base_font_family = self.base_font_family
        
        self.font_main = ctk.CTkFont(family=base_font_family, size=13)
        self.font_bold = ctk.CTkFont(family=base_font_family, size=13, weight="bold")
        self.font_heading = ctk.CTkFont(family=base_font_family, size=16, weight="bold")
        self.font_title = ctk.CTkFont(family=base_font_family, size=24, weight="bold")
        self.font_button = ctk.CTkFont(family=base_font_family, size=14, weight="bold")
        self.font_small = ctk.CTkFont(family=base_font_family, size=12)

        self.log_txt = None
        self._build_ttk_styles(base_font_family)
        self._build_ui()

        for var in (self.poll_var, self.debounce_var, self.lod_var, self.dpi_idx_var, self.light_var, self.sleep_var):
            var.trace_add("write", lambda *_: self._refresh_dirty_state())
        for var in self.dpi_vars:
            var.trace_add("write", lambda *_: self._refresh_dirty_state())

        self._apply_theme(self.theme_name, persist=False)
        self.change_language(self.lang, skip_ui_update=True)
        self.after(50, self._refresh_status)
        self._select_tab("status")

        self._bg_thread = threading.Thread(target=self._tray_monitor_loop, daemon=True)
        self._bg_thread.start()

    # ══════════════════════════════════════════════════════════════════
    # 翻訳 & 設定
    # ══════════════════════════════════════════════════════════════════
    def _t(self, key: str, *args) -> str:
        text = LANG_DICT[self.lang].get(key, key)
        return text.format(*args) if args else text

    def _ui_text(self, key: str) -> str:
        entry = UI_TEXT.get(key, {})
        return entry.get(self.lang, entry.get("en", key))

    def _load_app_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                self.lang = cfg.get("lang", "ja")
                self.theme_name = cfg.get("theme", "dark")
            except Exception:
                self.lang = "ja"
                self.theme_name = "dark"

    def _save_app_settings(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump({"lang": self.lang, "theme": self.theme_name}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_macro_metadata(self):
        if os.path.exists(self.macro_meta_file):
            try:
                with open(self.macro_meta_file, "r", encoding="utf-8") as f:
                    self._macro_metadata = json.load(f)
            except Exception:
                self._macro_metadata = {"names": {}}

    def _save_macro_metadata(self):
        try:
            with open(self.macro_meta_file, "w", encoding="utf-8") as f:
                json.dump(self._macro_metadata, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════
    # テーマ・スタイリング
    # ══════════════════════════════════════════════════════════════════
    def _build_ttk_styles(self, base_font_family):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.base_font_family = base_font_family

    def _apply_theme(self, theme_name: str, persist=True):
        palette = PALETTES.get(theme_name, PALETTES["dark"])
        self.theme_name = theme_name if theme_name in PALETTES else "dark"
        self._theme_palette = palette

        ctk.set_appearance_mode(self.theme_name)
        # CTk ベースカラー調整用（サイドバーや背景）
        self.configure(fg_color=palette["bg"])
        self.sidebar_frame.configure(fg_color=palette["panel"])

        # ttk ハイブリッド用 (Treeview/Listboxなど特殊ウィジェットのみ)
        font = (self.base_font_family, 10)
        font_bold = (self.base_font_family, 10, "bold")
        s = self.style
        s.configure("Treeview",
            background=palette["field"], fieldbackground=palette["field"], foreground=palette["fg"],
            bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"],
            rowheight=28, font=font, relief="flat", borderwidth=1,
        )
        s.map("Treeview",
            background=[("selected", palette["select_bg"])],
            foreground=[("selected", palette["select_fg"])],
        )
        s.configure("Treeview.Heading",
            background=palette["panel"], foreground=palette["fg"],
            bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"],
            font=font_bold, padding=(6, 4), relief="flat", borderwidth=1,
        )
        s.configure("TScrollbar",
            background=palette["panel"], troughcolor=palette["bg"],
            borderwidth=0, arrowsize=12,
        )

        listbox_opts = dict(
            bg=palette["field"], fg=palette["fg"],
            selectbackground=palette["select_bg"], selectforeground=palette["select_fg"],
            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
            font=(self.font_main.cget("family"), 10)
        )
        if self.log_txt:
            mono = ctk.CTkFont(family="Cascadia Code" if self.lang != "ja" else "Consolas", size=10)
            self.log_txt.configure(
                fg_color=palette["field"],
                text_color=palette["muted"],
                font=mono
            )
        if hasattr(self, "km_tab"):
            self.km_tab.apply_listbox_style(listbox_opts)
        if hasattr(self, "macro_tab"):
            self.macro_tab.apply_listbox_style(listbox_opts)

        self._apply_window_chrome(self.theme_name == "dark")
        if persist:
            self._save_app_settings()

    def _apply_window_chrome(self, dark: bool):
        if os.name != "nt":
            return
        try:
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            value = ctypes.c_int(1 if dark else 0)
            for attr in (20, 19):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value))
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════════
    # UI構築
    # ══════════════════════════════════════════════════════════════════
    def _build_ui(self):
        # ── 全体レイアウト ──
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ── サイドバー ──
        self.sidebar_frame = ctk.CTkFrame(self, corner_radius=0, width=220)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="AJ139 V2", font=self.font_title, text_color=PALETTES["dark"]["accent"])
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 30))

        self.sidebar_btns = {}
        tabs_config = [
            ("status", "tab_status"),
            ("perf", "tab_perf"),
            ("sys", "tab_sys"),
            ("keys", "tab_keys"),
            ("macro", "tab_macro"),
            ("debug", "tab_debug"),
        ]
        for i, (tab_id, lang_key) in enumerate(tabs_config):
            btn = ctk.CTkButton(
                self.sidebar_frame, text="",
                fg_color="transparent", text_color=["gray10", "gray90"], hover_color=["gray70", "gray30"],
                anchor="w", font=self.font_button,
                command=lambda tid=tab_id: self._select_tab(tid)
            )
            btn.grid(row=i+1, column=0, padx=15, pady=8, sticky="ew")
            self.sidebar_btns[tab_id] = {"btn": btn, "lang_key": lang_key}

        # ── サイドバー下部 (言語・テーマ設定) ──
        bottom_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        bottom_frame.grid(row=8, column=0, padx=15, pady=(20, 20), sticky="ew")

        self.lbl_lang_label = ctk.CTkLabel(bottom_frame, font=self.font_small, text_color="gray50")
        self.lbl_lang_label.pack(anchor="w")
        self.lang_var = ctk.StringVar(value=self.lang)
        self.lang_combo = ctk.CTkOptionMenu(
            bottom_frame, variable=self.lang_var, values=["en", "ja"],
            font=self.font_small, command=self._on_lang_selected
        )
        self.lang_combo.pack(fill="x", pady=(0, 10))

        self.lbl_theme_label = ctk.CTkLabel(bottom_frame, font=self.font_small, text_color="gray50")
        self.lbl_theme_label.pack(anchor="w")
        self.theme_var = ctk.StringVar(value=self.theme_name)
        self.theme_combo = ctk.CTkOptionMenu(
            bottom_frame, variable=self.theme_var, values=["light", "dark"],
            font=self.font_small, command=self._on_theme_selected
        )
        self.theme_combo.pack(fill="x")

        # ── メインコンテナ ──
        self.main_container = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_container.grid(row=0, column=1, sticky="nsew")
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # ── 各タブのコンテナ ──
        self.tabs = {
            "status": ctk.CTkFrame(self.main_container, fg_color="transparent"),
            "perf": ctk.CTkFrame(self.main_container, fg_color="transparent"),
            "sys": ctk.CTkFrame(self.main_container, fg_color="transparent"),
            "keys": ctk.CTkFrame(self.main_container, fg_color="transparent"),
            "macro": ctk.CTkFrame(self.main_container, fg_color="transparent"),
            "debug": ctk.CTkFrame(self.main_container, fg_color="transparent"),
        }

        self._build_status_tab()
        self._build_perf_tab()
        self._build_sys_tab()
        self._build_keys_tab()
        self._build_macro_tab()
        self._build_debug_tab()

        # ── フッタ (保存・更新など) ──
        self.footer_frame = ctk.CTkFrame(self.main_container, corner_radius=0, fg_color="transparent")
        self.footer_frame.grid(row=1, column=0, sticky="ew", padx=30, pady=20)
        
        self.page_status_var = tk.StringVar(value="")
        self.page_status_label = ctk.CTkLabel(self.footer_frame, textvariable=self.page_status_var, font=self.font_bold)
        self.page_status_label.pack(side=tk.LEFT)
        
        self.btn_save_page = ctk.CTkButton(
            self.footer_frame, text="", font=self.font_button,
            fg_color=PALETTES["dark"]["accent"], hover_color=PALETTES["dark"]["accent_light"], text_color=PALETTES["dark"]["accent_fg"],
            command=self._save_current_page, height=36
        )
        self.btn_save_page.pack(side=tk.RIGHT)

        self.btn_refresh = ctk.CTkButton(self.footer_frame, text="", font=self.font_main, command=self._refresh_status, height=36)
        self.btn_refresh.pack(side=tk.RIGHT, padx=10)

    def _select_tab(self, tab_id: str):
        # Guard changes
        if not self._tab_change_guard and self._current_tab_id != tab_id:
            if self._page_is_dirty(self._current_tab_id):
                answer = messagebox.askyesnocancel(self._t("unsaved_changes"), self._t("unsaved_prompt"))
                if answer is None:
                    self._tab_change_guard = True
                    # Re-select previous
                    self._select_tab(self._current_tab_id)
                    self._tab_change_guard = False
                    return
                if answer:
                    self._save_page(self._current_tab_id)
                else:
                    self._discard_page_changes(self._current_tab_id)
            self._current_tab_id = tab_id

        # Update button colors
        for t_id, data in self.sidebar_btns.items():
            if t_id == tab_id:
                data["btn"].configure(fg_color=["gray70", "gray25"])
            else:
                data["btn"].configure(fg_color="transparent")

        # Show selected frame
        for t_id, frame in self.tabs.items():
            if t_id == tab_id:
                frame.grid(row=0, column=0, sticky="nsew", padx=30, pady=(30, 0))
            else:
                frame.grid_forget()
        
        if tab_id == "macro":
            self.after(50, self._load_macro_profiles_async)
            
        self._refresh_dirty_state()


    # ── 各タブのUI定義 (CTk化) ───────────────────────────────────────────

    def _build_status_tab(self):
        f = self.tabs["status"]
        self.status_var = tk.StringVar()
        self.version_var = tk.StringVar()
        self.battery_var = tk.StringVar()

        self.status_frame_inner = ctk.CTkFrame(f, corner_radius=15)
        self.status_frame_inner.pack(fill="x", padx=20, pady=20)
        
        self.status_frame_title = ctk.CTkLabel(self.status_frame_inner, text="USB / Dongle Status", font=self.font_heading, text_color="#2dd4bf")
        self.status_frame_title.pack(anchor="w", padx=20, pady=(15, 15))

        ctk.CTkLabel(self.status_frame_inner, textvariable=self.status_var, font=self.font_main).pack(anchor="w", padx=20)
        ctk.CTkLabel(self.status_frame_inner, textvariable=self.version_var, font=self.font_main).pack(anchor="w", padx=20, pady=(8, 0))
        ctk.CTkLabel(self.status_frame_inner, textvariable=self.battery_var, font=self.font_main).pack(anchor="w", padx=20, pady=(8, 20))

    def _build_perf_tab(self):
        f = self.tabs["perf"]
        def row(parent):
            cf = ctk.CTkFrame(parent, fg_color="transparent")
            cf.pack(fill="x", pady=6)
            return cf

        self.perf_frame = ctk.CTkFrame(f)
        self.perf_frame.pack(fill="both", expand=True, padx=20, pady=20)

        r1 = row(self.perf_frame)
        self.lbl_poll = ctk.CTkLabel(r1, text="Polling Rate", font=self.font_main, width=180, anchor="w")
        self.lbl_poll.pack(side="left", padx=(20, 0))
        self.poll_var = ctk.StringVar(value="1000")
        self.poll_combo = ctk.CTkOptionMenu(r1, variable=self.poll_var, values=["125", "250", "500", "1000"], font=self.font_main)
        self.poll_combo.pack(side="left")

        r2 = row(self.perf_frame)
        self.lbl_debounce = ctk.CTkLabel(r2, text="Debounce Time", font=self.font_main, width=180, anchor="w")
        self.lbl_debounce.pack(side="left", padx=(20, 0))
        self.debounce_var = ctk.StringVar(value="4")
        self.debounce_entry = ctk.CTkEntry(r2, textvariable=self.debounce_var, width=140, font=self.font_main)
        self.debounce_entry.pack(side="left")

        r3 = row(self.perf_frame)
        self.lbl_lod = ctk.CTkLabel(r3, text="LOD", font=self.font_main, width=180, anchor="w")
        self.lbl_lod.pack(side="left", padx=(20, 0))
        self.lod_var = ctk.StringVar(value="1mm")
        self.lod_combo = ctk.CTkOptionMenu(r3, variable=self.lod_var, values=["1mm", "2mm"], font=self.font_main)
        self.lod_combo.pack(side="left")

        # DPI 枠
        self.dpi_label_title = ctk.CTkLabel(self.perf_frame, text="DPI Levels", font=self.font_heading, text_color="#2dd4bf")
        self.dpi_label_title.pack(anchor="w", padx=20, pady=(20, 10))
        
        self.dpi_lbls = []
        self.dpi_vars = []
        self.dpi_entries = []
        for i in range(6):
            fr = row(self.perf_frame)
            lbl = ctk.CTkLabel(fr, font=self.font_main, width=180, anchor="w")
            lbl.pack(side="left", padx=(20, 0))
            var = ctk.StringVar(value="400")
            ent = ctk.CTkEntry(fr, textvariable=var, width=140, font=self.font_main)
            ent.pack(side="left")
            self.dpi_lbls.append(lbl)
            self.dpi_vars.append(var)
            self.dpi_entries.append(ent)

        idx_row = row(self.perf_frame)
        self.lbl_dpi_idx = ctk.CTkLabel(idx_row, font=self.font_main, width=180, anchor="w")
        self.lbl_dpi_idx.pack(side="left", padx=(20, 0))
        self.dpi_idx_var = ctk.StringVar(value="1")
        self.dpi_idx_combo = ctk.CTkOptionMenu(idx_row, variable=self.dpi_idx_var, values=[str(i+1) for i in range(6)], font=self.font_main, width=80)
        self.dpi_idx_combo.pack(side="left")

    def _build_sys_tab(self):
        f = self.tabs["sys"]
        def row(parent):
            cf = ctk.CTkFrame(parent, fg_color="transparent")
            cf.pack(fill="x", pady=10)
            return cf

        self.sys_frame = ctk.CTkFrame(f)
        self.sys_frame.pack(fill="both", expand=True, padx=20, pady=20)

        r1 = row(self.sys_frame)
        self.lbl_light = ctk.CTkLabel(r1, font=self.font_main, width=180, anchor="w")
        self.lbl_light.pack(side="left", padx=(20, 0))
        self.light_var = ctk.StringVar(value="0")
        self.light_combo = ctk.CTkOptionMenu(r1, variable=self.light_var, values=[str(i) for i in range(8)], font=self.font_main)
        self.light_combo.pack(side="left")

        r2 = row(self.sys_frame)
        self.lbl_sleep = ctk.CTkLabel(r2, font=self.font_main, width=180, anchor="w")
        self.lbl_sleep.pack(side="left", padx=(20, 0))
        self.sleep_var = ctk.StringVar(value="10")
        self.sleep_entry = ctk.CTkEntry(r2, textvariable=self.sleep_var, width=140, font=self.font_main)
        self.sleep_entry.pack(side="left")

    def _build_keys_tab(self):
        self.km_tab = KeyMappingTab(self.tabs["keys"], self)

    def _build_macro_tab(self):
        self.macro_tab = MacroTab(self.tabs["macro"], self)

    def _build_debug_tab(self):
        f = self.tabs["debug"]
        self.debug_frame = ctk.CTkFrame(f)
        self.debug_frame.pack(fill="both", expand=True)
        # CTkTextbox を使用することでスクロール可能なモダンテキストエリアに。
        self.log_txt = ctk.CTkTextbox(self.debug_frame, font=self.font_main, state="disabled")
        self.log_txt.pack(fill="both", expand=True, padx=10, pady=10)

    # ══════════════════════════════════════════════════════════════════
    # 言語・テーマ設定
    # ══════════════════════════════════════════════════════════════════
    def _on_lang_selected(self, val):
        self.change_language(val)

    def _on_theme_selected(self, val):
        reverse = {
            self._t("theme_light"): "light", 
            self._t("theme_dark"): "dark",
            "light": "light",
            "dark": "dark"
        }
        self._apply_theme(reverse.get(val, "dark"))

    def change_language(self, new_lang: str, skip_ui_update=False):
        self.lang = new_lang
        self._save_app_settings()
        self.title(self._t("title"))

        for tab_id, data in self.sidebar_btns.items():
            data["btn"].configure(text=self._t(data["lang_key"]))

        self.lbl_lang_label.configure(text=self._t("lang_label"))
        self.lbl_theme_label.configure(text=self._t("theme_label"))
        
        thm_options = ["light", "dark"] if self.lang == "en" else [self._t("theme_light"), self._t("theme_dark")]
        self.theme_combo.configure(values=thm_options)
        self.theme_var.set(self.theme_name if self.lang == "en" else self._t(f"theme_{self.theme_name}"))

        self.btn_refresh.configure(text=self._t("btn_refresh"))
        self.status_frame_title.configure(text=self._t("status_frame"))
        self.lbl_poll.configure(text=self._t("poll_rate"))
        self.lbl_debounce.configure(text=self._t("debounce"))
        self.lbl_lod.configure(text=self._t("lod"))
        self.dpi_label_title.configure(text=self._t("dpi_frame"))
        self.lbl_dpi_idx.configure(text=self._t("dpi_active"))
        self.lbl_light.configure(text=self._t("light_mode"))
        self.lbl_sleep.configure(text=self._t("sleep_time"))
        for i in range(6):
            self.dpi_lbls[i].configure(text=self._t("dpi_label", i + 1))

        self.km_tab.apply_lang(self._t, self._ui_text)
        self.macro_tab.apply_lang(self._t, self._ui_text)

        if self.tray_icon:
            self.tray_icon.menu = self._create_tray_menu()

        self._refresh_key_mapping_ui()
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()
        if not skip_ui_update:
            self._refresh_status()

    # ══════════════════════════════════════════════════════════════════
    # Dirty State & Widget Styling
    # ══════════════════════════════════════════════════════════════════
    def _set_field_dirty_style(self, widget, dirty: bool, widget_type: str):
        if widget is None:
            return
        # CTk ウィジェットの色を動的に変更
        # border_color 等で変更を表現
        dirty_color = PALETTES["dark"]["dirty_fg"] if self.theme_name == "dark" else PALETTES["light"]["dirty_fg"]
        normal_border = "gray50"
        
        try:
            if widget_type in ("entry", "combo"):
                widget.configure(border_color=dirty_color if dirty else normal_border, border_width=2 if dirty else 1)
            elif widget_type == "button":
                widget.configure(border_color=dirty_color if dirty else normal_border, border_width=2 if dirty else 1)
            elif widget_type == "label":
                widget.configure(text_color=dirty_color if dirty else ("gray10", "gray90"))
        except Exception:
            pass

    def _refresh_dirty_state(self):
        perf_saved = self._saved_perf_snapshot()
        perf_cur = self._current_perf_snapshot()
        if perf_saved:
            self._set_field_dirty_style(self.poll_combo, perf_cur["poll"] != perf_saved["poll"], "combo")
            self._set_field_dirty_style(self.debounce_entry, perf_cur["debounce"] != perf_saved["debounce"], "entry")
            self._set_field_dirty_style(self.lod_combo, perf_cur["lod"] != perf_saved["lod"], "combo")
            self._set_field_dirty_style(self.dpi_idx_combo, perf_cur["dpi_index"] != perf_saved["dpi_index"], "combo")
            for i, ent in enumerate(self.dpi_entries):
                self._set_field_dirty_style(ent, perf_cur["dpis"][i] != perf_saved["dpis"][i], "entry")

        sys_saved = self._saved_sys_snapshot()
        sys_cur = self._current_sys_snapshot()
        if sys_saved:
            self._set_field_dirty_style(self.light_combo, sys_cur["light"] != sys_saved["light"], "combo")
            self._set_field_dirty_style(self.sleep_entry, sys_cur["sleep"] != sys_saved["sleep"], "entry")

        dirty_slots = {
            i for i, (cur, sav) in enumerate(zip(self.mouse_keys, self._saved_mouse_keys)) if cur != sav
        }
        self.km_tab.refresh_dirty(dirty_slots)
        palette = self._theme_palette
        self.macro_tab.refresh_dirty_profiles(self.macro_profiles, self._saved_macro_profiles, palette)

        cur_page = self._current_tab_id
        is_dirty = self._page_is_dirty(cur_page)
        self.page_status_var.set(self._t("dirty_hint") if is_dirty else self._t("clean_hint"))
        self._set_field_dirty_style(self.page_status_label, is_dirty, "label")

        save_key = {"perf": "save_perf", "sys": "save_sys", "keys": "save_keys", "macro": "save_macro"}.get(cur_page, "save_page")
        self.btn_save_page.configure(text=self._t(save_key))
        self.btn_save_page.configure(state="normal" if cur_page in {"perf", "sys", "keys", "macro"} else "disabled")

    # ══════════════════════════════════════════════════════════════════
    # (既存のスナップショット・保存ロジック等は変更なし、以下UI互換ラッパー)
    # ══════════════════════════════════════════════════════════════════
    def _saved_perf_snapshot(self) -> dict | None:
        if not self.current_config: return None
        return {
            "poll": {0: "125", 1: "250", 2: "500", 3: "1000"}.get(self.current_config["report_rate_idx"], "1000"),
            "debounce": str(self.current_config["key_respond"]),
            "lod": {1: "1mm", 2: "2mm"}.get(self.current_config["lod_value"], "1mm"),
            "dpi_index": str(self.current_config["dpi_index"] + 1),
            "dpis": [str(v) for v in self.current_config["dpis"]],
        }

    def _current_perf_snapshot(self) -> dict:
        return {
            "poll": self.poll_var.get(),
            "debounce": self.debounce_var.get(),
            "lod": self.lod_var.get(),
            "dpi_index": self.dpi_idx_var.get(),
            "dpis": [v.get() for v in self.dpi_vars],
        }

    def _saved_sys_snapshot(self) -> dict | None:
        if not self.current_config: return None
        return {"light": str(self.current_config["light_mode"]), "sleep": str(self.current_config["sleep_light"])}

    def _current_sys_snapshot(self) -> dict:
        return {"light": self.light_var.get(), "sleep": self.sleep_var.get()}

    def _page_is_dirty(self, page: str) -> bool:
        if page == "perf": return bool(self._saved_perf_snapshot() and self._saved_perf_snapshot() != self._current_perf_snapshot())
        if page == "sys": return bool(self._saved_sys_snapshot() and self._saved_sys_snapshot() != self._current_sys_snapshot())
        if page == "keys": return self.mouse_keys != self._saved_mouse_keys
        if page == "macro": return self.macro_profiles != self._saved_macro_profiles
        return False

    def _save_current_page(self):
        self._save_page(self._current_tab_id)

    def _save_page(self, page: str | None):
        if page == "perf": self._save_perf_page()
        elif page == "sys": self._save_sys_page()
        elif page == "keys": self._write_key_mapping_to_device()
        elif page == "macro": self._write_macros_to_device()

    def _discard_page_changes(self, page: str):
        if page in ("perf", "sys") and self.current_config: self._load_config_to_ui()
        elif page == "keys":
            self.mouse_keys = [clone_binding(b) for b in self._saved_mouse_keys]
            self._refresh_key_mapping_ui()
        elif page == "macro":
            self.macro_profiles = deepcopy(self._saved_macro_profiles)
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self.mouse_keys = [clone_binding(b, name=resolve_binding_name(b, self.macro_profiles)) for b in self.mouse_keys]
            self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _save_perf_page(self):
        if not self.mouse.device or not self.current_config: return
        cfg = deepcopy(self.current_config)
        cfg["report_rate_idx"] = {"125": 0, "250": 1, "500": 2, "1000": 3}.get(self.poll_var.get(), 3)
        cfg["key_respond"] = int(self.debounce_var.get() or "4")
        cfg["lod_value"] = 1 if self.lod_var.get() == "1mm" else 2
        cfg["dpi_index"] = int(self.dpi_idx_var.get() or "1") - 1
        cfg["dpis"] = [max(50, int(v.get() or "400")) for v in self.dpi_vars]
        self._append_log(self._t("log_sys_writing"))
        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=lambda _: self._on_config_saved(cfg))
        
    def _save_sys_page(self):
        if not self.mouse.device or not self.current_config: return
        cfg = deepcopy(self.current_config)
        cfg["sleep_light"] = int(self.sleep_var.get() or "10")
        cfg["light_mode"] = int(self.light_var.get() or "0")
        self._append_log(self._t("log_sys_writing"))
        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=lambda _: self._on_config_saved(cfg))

    def _on_config_saved(self, cfg):
        self.current_config = cfg
        self._refresh_dirty_state()

    def _write_key_mapping_to_device(self):
        if not self.mouse.device: return
        self._append_log("System: Writing key mapping...")
        def on_success(_):
            self._saved_mouse_keys = [clone_binding(b) for b in self.mouse_keys]
            self._refresh_dirty_state()
            self._append_log("System: Key mapping written.")
        self._run_in_background(lambda: self.mouse.set_mouse_keys(self.mouse_keys), on_success=on_success)

    def _write_macros_to_device(self):
        if not self.mouse.device: return
        self._append_log("System: Writing macro data...")
        self._set_macro_status("writing")
        def on_success(_):
            self._save_macro_metadata()
            self._saved_macro_profiles = deepcopy(self.macro_profiles)
            self._set_macro_status("ready")
            self._refresh_dirty_state()
            self._append_log("System: Macro data written.")
        self._run_in_background(lambda: self.mouse.set_macro_data(encode_macro_profiles(self.macro_profiles)), on_success=on_success)

    def _refresh_status(self):
        if self._status_refresh_in_progress: return
        self._status_refresh_in_progress = True
        self.log_txt.configure(state="normal")
        self.log_txt.delete("1.0", tk.END)
        self.log_txt.configure(state="disabled")
        self._append_log(self._t("log_sys_search"))
        self.status_var.set(self._t("status_val", "..."))
        self.version_var.set(self._t("fw_val", "--"))
        self.battery_var.set(self._t("bat_val", "--"))

        def worker():
            if not self.mouse.device and not self.mouse.connect(): return {"connected": False}
            result = {"connected": True, "version": self.mouse.get_version(), "online": self.mouse.is_online(), "config": self.mouse.get_config(), "keys": self.mouse.get_mouse_keys()}
            if result["online"]: result["battery_info"] = self.mouse.get_battery_info()
            return result

        def on_success(result):
            self._status_refresh_in_progress = False
            self._macro_profiles_loaded = False
            self._macro_profiles_loading = False
            self._set_macro_status("idle")
            if not result.get("connected"):
                self.status_var.set(self._t("status_val", self._t("st_disconnected")))
                self.version_var.set(self._t("fw_val", "--"))
                self.battery_var.set(self._t("bat_val", "--"))
                self._append_log(self._t("log_sys_fail"))
                return
            self.status_var.set(self._t("status_val", self._t("st_connected")))
            self._append_log(self._t("log_sys_conn"))
            self.version_var.set(self._t("fw_val", result["version"]))
            if result["online"]:
                bat = result["battery_info"]
                self.battery_var.set(self._t("bat_val", f"{bat['battery']}%" + (f" {self._t('bat_charging')}" if bat["is_charging"] else "")))
                self._append_log(self._t("log_sys_bat", bat["battery"]))
                self._last_tray_battery = bat["battery"]
                if self.tray_icon: self.tray_icon.icon = self._create_image(bat["battery"])
            else:
                self.battery_var.set(self._t("bat_val", self._t("bat_offline")))
                self._append_log(self._t("log_sys_off"))
                self._last_tray_battery = -1
                if self.tray_icon: self.tray_icon.icon = self._create_image(-1)
            self.current_config = result["config"]
            if self.current_config: self._load_config_to_ui()
            self.mouse_keys = [clone_binding(b, name=resolve_binding_name(b, self.macro_profiles)) for b in result["keys"]]
            self._saved_mouse_keys = [clone_binding(b) for b in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()
            if self._current_tab_id == "macro": self._load_macro_profiles_async()

        def on_error(exc):
            self._status_refresh_in_progress = False
            self._set_macro_status("error")
            self.status_var.set(self._t("status_val", self._t("st_disconnected")))
            self._append_log(f"ERROR: {exc}")

        self._run_in_background(worker, on_success=on_success, on_error=on_error)

    def _load_config_to_ui(self):
        cfg = self.current_config
        self.poll_var.set({0: "125", 1: "250", 2: "500", 3: "1000"}.get(cfg["report_rate_idx"], "1000"))
        self.debounce_var.set(str(cfg["key_respond"]))
        self.lod_var.set({1: "1mm", 2: "2mm"}.get(cfg["lod_value"], "1mm"))
        self.light_var.set(str(cfg["light_mode"]))
        self.sleep_var.set(str(cfg["sleep_light"]))
        self.dpi_idx_var.set(str(cfg["dpi_index"] + 1))
        for i, v in enumerate(cfg["dpis"]): self.dpi_vars[i].set(str(v))

    def _load_macro_profiles_async(self, force=False, log=True):
        if self._macro_profiles_loading or (self._macro_profiles_loaded and not force): return
        if not self.mouse.device: return
        self._macro_profiles_loading = True
        self._set_macro_status("loading")
        self._set_macro_progress(0, 0)
        self._show_macro_progress()
        if log: self._append_log("System: Loading macro profiles from device...")

        def worker():
            return self._merge_macro_names(decode_macro_profiles(
                    self.mouse.get_macro_data(progress_callback=lambda c, t: self.after(0, lambda c=c, t=t: self._set_macro_progress(c, t))),
                    resolve_macro_event_name))

        def on_success(profiles):
            self._macro_profiles_loading = False
            self._macro_profiles_loaded = True
            self._set_macro_status("ready")
            self._hide_macro_progress()
            self.macro_profiles = profiles
            self._saved_macro_profiles = deepcopy(profiles)
            self._refresh_macro_profile_list()
            self._refresh_macro_events()
            self.mouse_keys = [clone_binding(b, name=resolve_binding_name(b, self.macro_profiles)) for b in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()
            if log: self._append_log("System: Macro profiles loaded from device.")

        def on_error(exc):
            self._macro_profiles_loading = False
            self._set_macro_status("error")
            self._hide_macro_progress()
            if log: self._append_log(f"ERROR: Failed to load macro data: {exc}")

        self._run_in_background(worker, on_success=on_success, on_error=on_error)

    def _merge_macro_names(self, profiles: list[MacroProfile]) -> list[MacroProfile]:
        names = self._macro_metadata.get("names", {})
        return [MacroProfile(slot=p.slot, name=names.get(str(p.slot), p.name), trigger_mode=p.trigger_mode, repeat_count=p.repeat_count, list=deepcopy(p.list)) for p in profiles]

    def _set_macro_status(self, state_key: str):
        self._macro_status_key = state_key
        label = self._ui_text("macro_status_label")
        value = self._ui_text(f"macro_status_{state_key}")
        if hasattr(self, "macro_tab"): self.macro_tab.macro_status_var.set(f"{label} {value}")

    def _set_macro_progress(self, current: int, total: int):
        self._macro_progress_current = max(0, int(current))
        self._macro_progress_total = max(0, int(total))
        if hasattr(self, "macro_tab"): self.macro_tab.set_progress(self._macro_progress_current, self._macro_progress_total)

    def _show_macro_progress(self):
        if hasattr(self, "macro_tab"): self.macro_tab.show_progress()

    def _hide_macro_progress(self):
        if hasattr(self, "macro_tab"): self.macro_tab.hide_progress()

    def _append_log(self, text: str):
        if hasattr(self, "log_txt") and self.log_txt.winfo_exists():
            self.log_txt.configure(state="normal")
            self.log_txt.insert(tk.END, text + "\n")
            self.log_txt.see(tk.END)
            self.log_txt.configure(state="disabled")

    def _hide_window(self): self.withdraw()
    def _show_window(self): self.deiconify(); self.lift(); self.focus_force()
    def _exit_app(self, icon, item):
        if self.tray_icon: self.tray_icon.stop()
        os._exit(0)
    def _create_tray_menu(self): return pystray.Menu(pystray.MenuItem(self._t("tray_open"), lambda: self.after(0, self._show_window), default=True), pystray.MenuItem(self._t("tray_exit"), self._exit_app))
    def _create_image(self, battery_perc: int):
        w, h = 64, 64
        bg, fg = ("gray", "white") if battery_perc < 0 else (("black", "red") if battery_perc <= 20 else ("black", "white"))
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([2, 2, w - 2, h - 2], fill=bg, outline="white")
        text = "--" if battery_perc < 0 else str(battery_perc)
        try: font = ImageFont.truetype("arial.ttf", 36)
        except Exception: font = ImageFont.load_default()
        try: l, t, r, b = draw.textbbox((0, 0), text, font=font); tw, th = r - l, b - t
        except AttributeError: tw, th = draw.textsize(text, font=font)
        draw.text(((w - tw) // 2, (h - th) // 2 - 2), text, fill=fg, font=font)
        return img
    def _tray_monitor_loop(self):
        self.tray_icon = pystray.Icon("AjazzAJ139", self._create_image(-1), "AJ139 V2", self._create_tray_menu())
        def setup(icon):
            icon.visible = True
            while icon.visible:
                try:
                    if self._busy_count or self._macro_profiles_loading or self._status_refresh_in_progress:
                        time.sleep(1); continue
                    if self.mouse.device and self.mouse.is_online():
                        bat = self.mouse.get_battery_info()["battery"]
                        if bat != self._last_tray_battery:
                            self._last_tray_battery = bat
                            icon.icon = self._create_image(bat)
                    elif self._last_tray_battery != -1:
                        self._last_tray_battery = -1
                        icon.icon = self._create_image(-1)
                except Exception: pass
                time.sleep(10)
        self.tray_icon.run(setup)

    def _set_busy(self, busy: bool):
        self._busy_count = max(0, self._busy_count + (1 if busy else -1))
        state = "disabled" if self._busy_count else "normal"
        for w in (self.btn_save_page, self.btn_refresh):
            if w: w.configure(state=state)
        for btn_name in ("reset_keys_btn", "record_btn", "add_key_pair_btn", "add_mouse_pair_btn", "macro_move_up_btn", "macro_move_down_btn", "macro_delete_btn", "macro_clear_btn", "reload_macro_btn", "reset_macro_btn"):
            btn = getattr(self.macro_tab, btn_name, None) or getattr(self.km_tab, btn_name, None)
            if btn: btn.configure(state=state)
        self.configure(cursor="watch" if self._busy_count else "")

    def _run_in_background(self, worker, on_success=None, on_error=None, busy=True):
        if busy: self._set_busy(True)
        def runner():
            try: result = worker()
            except Exception as exc: self.after(0, lambda: (self._set_busy(False) if busy else None, on_error(exc) if on_error else messagebox.showerror("Error", str(exc))))
            else: self.after(0, lambda: (self._set_busy(False) if busy else None, on_success(result) if on_success else None))
        threading.Thread(target=runner, daemon=True).start()

    def _refresh_key_mapping_ui(self):
        self.km_tab.refresh(self.mouse_keys, self.macro_profiles)
        self._refresh_dirty_state()
    def _refresh_macro_profile_list(self):
        self.macro_tab.refresh_profile_list(self.macro_profiles)
        self._refresh_key_mapping_ui()
    def _refresh_macro_events(self):
        self.macro_tab.refresh_events(self.macro_profiles, self._t)
