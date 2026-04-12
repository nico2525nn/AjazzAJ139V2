"""
ui_macro.py – Macro タブのウィジェットとロジック (CustomTkinter化)
"""

import time
import tkinter as tk
from copy import deepcopy
from tkinter import messagebox, ttk
import customtkinter as ctk

from ajazz_mouse import (
    ACTION_PRESS,
    ACTION_RELEASE,
    EVENT_TYPE_KEYBOARD,
    EVENT_TYPE_MOUSE,
    MacroEvent,
    MacroProfile,
    decode_macro_profiles,
    encode_macro_profiles,
)
from constants import (
    KEYBOARD_EVENT_CHOICES,
    KEYBOARD_NAME_TO_CODE,
    KEYBOARD_PRESETS,
    MOUSE_EVENT_CHOICES,
    MOUSE_NAME_TO_CODE,
)
from ui_helpers import clone_binding, resolve_binding_name, resolve_macro_event_name


class MacroTab:
    def __init__(self, parent_frame: ctk.CTkFrame, app):
        self.app = app
        self.frame = parent_frame

        self.selected_macro_slot = 0
        self.selected_macro_event_index: int | None = None
        self.is_recording = False
        self._pressed_keys: set[int] = set()
        self._pressed_mouse: set[int] = set()
        self._last_record_time: int | None = None
        self._macro_editor_widget = None
        self._macro_editor_info: dict | None = None

        self.macro_status_var = ctk.StringVar(value="")
        self.macro_progress_var = tk.DoubleVar(value=0)
        self.macro_progress_label_var = ctk.StringVar(value="")
        self.macro_name_var = ctk.StringVar()
        self.record_delay_mode = ctk.StringVar(value="exact")
        self.fixed_delay_var = ctk.StringVar(value="10")
        self.manual_key_var = ctk.StringVar(value=KEYBOARD_PRESETS[0]["name"])
        self.manual_mouse_var = ctk.StringVar(value="Mouse L")
        self.macro_delay_var = ctk.StringVar(value="10")

        self.manual_key_map = dict(KEYBOARD_NAME_TO_CODE)
        self.manual_mouse_map = dict(MOUSE_NAME_TO_CODE)

        self._build()

        self.macro_name_var.trace_add("write", lambda *_: self._on_macro_name_changed())

    # ──────────────────────────────── UI 構築 ────────────────────────────────

    def _build(self):
        # ── ヘッダ ──
        top = ctk.CTkFrame(self.frame, fg_color="transparent")
        top.pack(fill="x", pady=(0, 20))
        
        self.macro_title_label = ctk.CTkLabel(top, text="Macro Config", font=self.app.font_title, text_color=("#0ea5a5", "#2dd4bf"))
        self.macro_title_label.pack(anchor="w")

        self.macro_hint_label = ctk.CTkLabel(top, text="", justify="left", font=self.app.font_small, text_color="gray50")
        self.macro_hint_label.pack(anchor="w")
        
        self.macro_status_label = ctk.CTkLabel(top, textvariable=self.macro_status_var, font=self.app.font_bold)
        self.macro_status_label.pack(anchor="w", pady=(5, 0))

        self.macro_progress_row = ctk.CTkFrame(top, fg_color="transparent")
        self.macro_progress = ctk.CTkProgressBar(self.macro_progress_row, variable=self.macro_progress_var, progress_color="#2dd4bf")
        self.macro_progress.pack(side="left", fill="x", expand=True)
        self.macro_progress_label = ctk.CTkLabel(self.macro_progress_row, textvariable=self.macro_progress_label_var, width=120, anchor="e", font=self.app.font_main)
        self.macro_progress_label.pack(side="left", padx=(10, 0))

        # ── 3 カラムレイアウト ──
        content = ctk.CTkFrame(self.frame, fg_color="transparent")
        content.pack(fill="both", expand=True)
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)

        left = ctk.CTkFrame(content, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 15))
        
        center = ctk.CTkFrame(content, fg_color="transparent")
        center.grid(row=0, column=1, sticky="nsew", padx=(0, 15))
        
        right = ctk.CTkFrame(content, fg_color="transparent")
        right.grid(row=0, column=2, sticky="nsew")

        # ── 左: プロファイルリスト ──
        self.macro_profiles_frame = ctk.CTkFrame(left, corner_radius=16, fg_color=("#ffffff", "#21242f"), border_width=1, border_color=("#d1d5db", "#2e3241"))
        self.macro_profiles_frame.pack(fill="both", expand=True)
        
        self.macro_profiles_title = ctk.CTkLabel(self.macro_profiles_frame, text="Macro Slots", font=self.app.font_heading, text_color=("#0ea5a5", "#2dd4bf"))
        self.macro_profiles_title.pack(anchor="w", padx=15, pady=(15, 5))
        
        # TK Listbox inside CTkFrame
        list_container = ctk.CTkFrame(self.macro_profiles_frame, fg_color="transparent")
        list_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        f_name = self.app.font_main.cget("family")
        self.macro_profile_listbox = tk.Listbox(
            list_container, width=22, exportselection=False, 
            borderwidth=0, highlightthickness=0,
            font=(f_name, 10)
        )
        self.macro_profile_listbox.pack(side="left", fill="both", expand=True)
        sb_prof = ctk.CTkScrollbar(list_container, orientation="vertical", command=self.macro_profile_listbox.yview, fg_color="black", button_color="#333333", button_hover_color="#444444")
        self.macro_profile_listbox.configure(yscrollcommand=sb_prof.set)
        sb_prof.pack(side="right", fill="y")
        self.macro_profile_listbox.bind("<<ListboxSelect>>", self._on_macro_profile_selected)

        # ── 中央: イベントツリー ──
        self.macro_events_frame = ctk.CTkFrame(center, corner_radius=16, fg_color=("#ffffff", "#21242f"), border_width=1, border_color=("#d1d5db", "#2e3241"))
        self.macro_events_frame.pack(fill="both", expand=True)
        
        self.macro_events_title = ctk.CTkLabel(self.macro_events_frame, text="Events", font=self.app.font_heading, text_color=("#0ea5a5", "#2dd4bf"))
        self.macro_events_title.pack(anchor="w", padx=15, pady=(15, 5))

        tree_container = ctk.CTkFrame(self.macro_events_frame, fg_color="transparent")
        tree_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.macro_tree = ttk.Treeview(
            tree_container,
            columns=("type", "name", "action", "delay"),
            show="headings",
            height=13,
        )
        for col, width in (("type", 90), ("name", 160), ("action", 90), ("delay", 70)):
            self.macro_tree.column(col, width=width, anchor=tk.CENTER if col != "name" else tk.W)
            self.macro_tree.heading(col, text="")
        self.macro_tree.pack(side="left", fill="both", expand=True)
        
        sb = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.macro_tree.yview, fg_color="black", button_color="#333333", button_hover_color="#444444")
        self.macro_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(2, 0))
        
        self.macro_tree.bind("<<TreeviewSelect>>", self._on_macro_event_selected)
        self.macro_tree.bind("<Double-1>", self._on_macro_tree_double_click)

        self.macro_edit_hint_label = ctk.CTkLabel(self.macro_events_frame, text="", justify="left", font=self.app.font_small, text_color="gray50")
        self.macro_edit_hint_label.pack(anchor="w", padx=10, pady=(0, 5))

        # イベント操作ボタン
        event_actions = ctk.CTkFrame(self.macro_events_frame, fg_color="transparent")
        event_actions.pack(fill="x", padx=10, pady=(0, 10))
        
        self.macro_move_up_btn = ctk.CTkButton(event_actions, text="▲ Up", font=self.app.font_main, corner_radius=16, width=70, command=self._move_event_up)
        self.macro_move_up_btn.pack(side="left")
        self.macro_move_down_btn = ctk.CTkButton(event_actions, text="▼ Down", font=self.app.font_main, corner_radius=16, width=70, command=self._move_event_down)
        self.macro_move_down_btn.pack(side="left", padx=5)
        self.macro_delete_btn = ctk.CTkButton(event_actions, text="Delete", font=self.app.font_main, corner_radius=16, width=70, command=self._delete_event, fg_color="#b91c1c", hover_color="#991b1b")
        self.macro_delete_btn.pack(side="left", padx=5)
        self.macro_clear_btn = ctk.CTkButton(event_actions, text="Clear", font=self.app.font_main, corner_radius=16, width=70, command=self._clear_profile, fg_color=("#d1d5db", "gray30"), hover_color=("#f3f4f6", "gray40"), text_color=("black", "white"))
        self.macro_clear_btn.pack(side="left", padx=5)

        # 遅延更新行
        delay_row = ctk.CTkFrame(self.macro_events_frame, fg_color="transparent")
        delay_row.pack(fill="x", padx=10, pady=(0, 10))
        self.macro_delay_label = ctk.CTkLabel(delay_row, text="Delay (ms):", font=self.app.font_main)
        self.macro_delay_label.pack(side="left")
        self.macro_delay_entry = ctk.CTkEntry(delay_row, textvariable=self.macro_delay_var, width=80, font=self.app.font_main)
        self.macro_delay_entry.pack(side="left", padx=10)
        self.macro_delay_entry.bind("<Return>", lambda _e: self._update_selected_macro_delay())
        self.macro_delay_entry.bind("<FocusOut>", lambda _e: self._update_selected_macro_delay())

        # ── 右: コントロールパネル ──
        self.macro_controls_frame = ctk.CTkFrame(right, width=280, corner_radius=16, fg_color=("#ffffff", "#21242f"), border_width=1, border_color=("#d1d5db", "#2e3241"))
        self.macro_controls_frame.pack(fill="both", expand=True)
        
        self.macro_controls_title = ctk.CTkLabel(self.macro_controls_frame, text="Recorder / Editor", font=self.app.font_heading, text_color=("#0ea5a5", "#2dd4bf"))
        self.macro_controls_title.pack(anchor="w", padx=15, pady=(15, 5))

        # 名前行
        name_row = ctk.CTkFrame(self.macro_controls_frame, fg_color="transparent")
        name_row.pack(fill="x", padx=10)
        self.macro_name_label = ctk.CTkLabel(name_row, text="Name:", font=self.app.font_main)
        self.macro_name_label.pack(side="left")
        ctk.CTkEntry(name_row, textvariable=self.macro_name_var, width=150, font=self.app.font_main).pack(side="left", padx=(10, 0))

        # 録画ボタン
        self.record_btn = ctk.CTkButton(
            self.macro_controls_frame,
            text="",
            font=self.app.font_button, corner_radius=20, height=40,
            fg_color="#0ea5a5", hover_color="#0d8a8a", text_color="white",
            command=self._toggle_recording,
        )
        self.record_btn.pack(fill="x", padx=10, pady=15)

        # 遅延モード
        self.delay_mode_frame = ctk.CTkFrame(self.macro_controls_frame, border_width=1, border_color="gray30")
        self.delay_mode_frame.pack(fill="x", padx=10, pady=(0, 15))
        
        self.delay_mode_title = ctk.CTkLabel(self.delay_mode_frame, text="Delay Handling", font=self.app.font_bold)
        self.delay_mode_title.pack(anchor="w", padx=10, pady=(5, 5))

        self.delay_exact_radio = ctk.CTkRadioButton(self.delay_mode_frame, text="Use actual delays", variable=self.record_delay_mode, value="exact", font=self.app.font_main)
        self.delay_exact_radio.pack(anchor="w", padx=10, pady=5)
        self.delay_none_radio = ctk.CTkRadioButton(self.delay_mode_frame, text="Zero delay", variable=self.record_delay_mode, value="none", font=self.app.font_main)
        self.delay_none_radio.pack(anchor="w", padx=10, pady=5)
        
        fixed_row = ctk.CTkFrame(self.delay_mode_frame, fg_color="transparent")
        fixed_row.pack(fill="x", padx=10, pady=(5, 10))
        self.delay_fixed_radio = ctk.CTkRadioButton(fixed_row, text="Fixed:", variable=self.record_delay_mode, value="fixed", font=self.app.font_main)
        self.delay_fixed_radio.pack(side="left")
        ctk.CTkEntry(fixed_row, textvariable=self.fixed_delay_var, width=60, font=self.app.font_main).pack(side="left", padx=(10, 0))

        # 手動追加
        man_frame = ctk.CTkFrame(self.macro_controls_frame, fg_color="transparent")
        man_frame.pack(fill="x", padx=10)
        
        self.manual_key_label = ctk.CTkLabel(man_frame, text="Manual key pair:", font=self.app.font_main)
        self.manual_key_label.pack(anchor="w")
        self.manual_key_combo = ctk.CTkOptionMenu(man_frame, variable=self.manual_key_var, values=list(self.manual_key_map.keys()), font=self.app.font_main)
        self.manual_key_combo.pack(fill="x", pady=(2, 5))
        self.add_key_pair_btn = ctk.CTkButton(man_frame, text="Add Key Pair", font=self.app.font_main, corner_radius=16, command=self._add_manual_key_pair, fg_color=("#d1d5db", "gray30"), hover_color=("#f3f4f6", "gray40"), text_color=("black", "white"))
        self.add_key_pair_btn.pack(fill="x")
        
        self.manual_mouse_label = ctk.CTkLabel(man_frame, text="Manual mouse pair:", font=self.app.font_main)
        self.manual_mouse_label.pack(anchor="w", pady=(15, 0))
        self.manual_mouse_combo = ctk.CTkOptionMenu(man_frame, variable=self.manual_mouse_var, values=list(self.manual_mouse_map.keys()), font=self.app.font_main)
        self.manual_mouse_combo.pack(fill="x", pady=(2, 5))
        self.add_mouse_pair_btn = ctk.CTkButton(man_frame, text="Add Mouse Pair", font=self.app.font_main, corner_radius=16, command=self._add_manual_mouse_pair, fg_color=("#d1d5db", "gray30"), hover_color=("#f3f4f6", "gray40"), text_color=("black", "white"))
        self.add_mouse_pair_btn.pack(fill="x")

        # デバイス操作
        dev_frame = ctk.CTkFrame(self.macro_controls_frame, fg_color="transparent")
        dev_frame.pack(fill="x", padx=10, pady=(20, 10))
        
        self.reload_macro_btn = ctk.CTkButton(dev_frame, text="Reload From Device", font=self.app.font_button, corner_radius=20, height=36, fg_color=("#0ea5a5", "#2dd4bf"), hover_color=("#0d8a8a", "#5eead4"), text_color="white", command=lambda: self.app._load_macro_profiles_async(force=True))
        self.reload_macro_btn.pack(fill="x")
        self.reset_macro_btn = ctk.CTkButton(dev_frame, text="Reset Macro Memory", font=self.app.font_button, corner_radius=20, height=36, fg_color="#b91c1c", hover_color="#991b1b", text_color="white", command=self._reset_macro_data_on_device)
        self.reset_macro_btn.pack(fill="x", pady=(10, 0))

        self._hide_progress()

    # ──────────────────────────────── 言語更新 ────────────────────────────────

    def apply_lang(self, t, ui_text):
        self.macro_hint_label.configure(text=ui_text("macro_hint"))
        self.macro_edit_hint_label.configure(text=ui_text("macro_edit_hint"))
        self.macro_profiles_title.configure(text=t("macro_profiles"))
        self.macro_events_title.configure(text=t("macro_events"))
        self.macro_controls_title.configure(text=t("macro_controls"))
        self.macro_name_label.configure(text=t("macro_name"))
        self.delay_mode_title.configure(text=t("macro_delay_mode"))
        self.delay_exact_radio.configure(text=t("macro_delay_exact"))
        self.delay_none_radio.configure(text=t("macro_delay_none"))
        self.delay_fixed_radio.configure(text=t("macro_delay_fixed").replace("(ms):", ""))
        self.manual_key_label.configure(text=t("macro_manual_key"))
        self.manual_mouse_label.configure(text=t("macro_manual_mouse"))
        self.add_key_pair_btn.configure(text=t("macro_add_key"))
        self.add_mouse_pair_btn.configure(text=t("macro_add_mouse"))
        
        # Adjusting move/delete text if not explicitly translated or fall back
        self.macro_move_up_btn.configure(text=t("macro_move_up").replace("▲ ", ""))
        self.macro_move_down_btn.configure(text=t("macro_move_down").replace("▼ ", ""))
        self.macro_delete_btn.configure(text=t("macro_delete"))
        self.macro_clear_btn.configure(text=t("macro_clear"))
        
        self.macro_delay_label.configure(text=t("macro_delay"))
        self.reload_macro_btn.configure(text=t("macro_reload"))
        self.reset_macro_btn.configure(text=t("macro_reset"))
        self.record_btn.configure(
            text=t("macro_record_stop") if self.is_recording else t("macro_record_start")
        )
        self.macro_tree.heading("type", text=t("macro_type"))
        self.macro_tree.heading("name", text=t("macro_name_col"))
        self.macro_tree.heading("action", text=t("macro_action"))
        self.macro_tree.heading("delay", text=t("macro_delay_col"))

    def apply_listbox_style(self, opts: dict):
        self.macro_profile_listbox.configure(**opts)

    # ──────────────────────────────── プログレス ──────────────────────────────

    def set_status(self, ui_text, status_label: str, status_value: str):
        self.macro_status_var.set(f"{status_label} {status_value}")

    def set_progress(self, current: int, total: int):
        self.macro_progress_var.set(min(current, max(total, 1)))
        if total > 0:
            self.macro_progress.set(min(current, max(total, 1)) / max(total, 1))
        percent = 0 if total <= 0 else round(min(current, total) / total * 100)
        self.macro_progress_label_var.set(
            f"{percent}% ({min(current, total)}/{total})" if total > 0 else ""
        )

    def show_progress(self):
        if not self.macro_progress_row.winfo_ismapped():
            self.macro_progress_row.pack(fill="x", pady=(4, 0))

    def _hide_progress(self):
        if self.macro_progress_row.winfo_ismapped():
            self.macro_progress_row.pack_forget()

    def hide_progress(self):
        self._hide_progress()

    # ──────────────────────────────── プロファイルリスト ──────────────────────

    def refresh_profile_list(self, macro_profiles: list[MacroProfile]):
        self.macro_profile_listbox.delete(0, tk.END)
        for p in macro_profiles:
            suffix = f" ({len(p.list)})" if p.list else ""
            self.macro_profile_listbox.insert(tk.END, f"{p.slot + 1:02d}. {p.name}{suffix}")
        self.macro_profile_listbox.selection_clear(0, tk.END)
        self.macro_profile_listbox.selection_set(self.selected_macro_slot)
        self.macro_name_var.set(macro_profiles[self.selected_macro_slot].name)

    def refresh_dirty_profiles(self, macro_profiles, saved_profiles, palette: dict):
        self.macro_profile_listbox.selection_clear(0, tk.END)
        for i, _ in enumerate(macro_profiles):
            if i < self.macro_profile_listbox.size():
                changed = i < len(saved_profiles) and macro_profiles[i] != saved_profiles[i]
                self.macro_profile_listbox.itemconfig(
                    i,
                    bg=palette["dirty"] if changed else palette["field"],
                    fg=palette["fg"],
                )
        if macro_profiles:
            self.macro_profile_listbox.selection_set(self.selected_macro_slot)

    # ──────────────────────────────── イベントツリー ──────────────────────────

    def refresh_events(self, macro_profiles: list[MacroProfile], t):
        self._close_editor()
        for iid in self.macro_tree.get_children():
            self.macro_tree.delete(iid)
        profile = macro_profiles[self.selected_macro_slot]
        if not profile.list:
            self.macro_tree.insert("", tk.END, values=("", t("macro_empty"), "", ""))
            self.selected_macro_event_index = None
            return
        for i, event in enumerate(profile.list):
            event_type = t("macro_event_mouse") if event.type == EVENT_TYPE_MOUSE else t("macro_event_keyboard")
            action = t("macro_action_press") if event.action == ACTION_PRESS else t("macro_action_release")
            self.macro_tree.insert("", tk.END, iid=str(i), values=(event_type, event.name, action, event.delay))
        if (
            self.selected_macro_event_index is not None
            and self.selected_macro_event_index < len(profile.list)
        ):
            self.macro_tree.selection_set(str(self.selected_macro_event_index))

    # ──────────────────────────────── イベントハンドラ ────────────────────────

    def _on_macro_profile_selected(self, event=None):
        self._close_editor()
        sel = self.macro_profile_listbox.curselection()
        if sel:
            self.selected_macro_slot = sel[0]
            self.selected_macro_event_index = None
            self.macro_name_var.set(self.app.macro_profiles[self.selected_macro_slot].name)
            self.app._refresh_macro_events()

    def _on_macro_event_selected(self, event=None):
        sel = self.macro_tree.selection()
        if sel:
            try:
                self.selected_macro_event_index = int(sel[0])
                self.macro_delay_var.set(
                    str(self.app.macro_profiles[self.selected_macro_slot].list[self.selected_macro_event_index].delay)
                )
            except Exception:
                self.selected_macro_event_index = None

    # ──────────────────────────────── インラインエディタ ──────────────────────

    def _close_editor(self):
        if self._macro_editor_widget is not None:
            try:
                self._macro_editor_widget.destroy()
            except tk.TclError:
                pass
        self._macro_editor_widget = None
        self._macro_editor_info = None

    def _commit_editor(self):
        info = self._macro_editor_info
        if not info:
            return
        profile = self.app.macro_profiles[self.selected_macro_slot]
        index = info["index"]
        if index >= len(profile.list):
            self._close_editor()
            return
        macro_event = profile.list[index]
        value = info["var"].get().strip()
        column = info["column"]
        t = self.app._t

        if column == "#2":
            code = (KEYBOARD_NAME_TO_CODE if macro_event.type == EVENT_TYPE_KEYBOARD else MOUSE_NAME_TO_CODE).get(value)
            if code is not None:
                macro_event.name = value
                macro_event.code = code
        elif column == "#3":
            macro_event.action = ACTION_PRESS if value == t("macro_action_press") else ACTION_RELEASE
        elif column == "#4":
            try:
                macro_event.delay = max(0, int(value or "0"))
            except ValueError:
                pass

        self.selected_macro_event_index = index
        self._close_editor()
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _on_macro_tree_double_click(self, event):
        self._close_editor()
        row_id = self.macro_tree.identify_row(event.y)
        column = self.macro_tree.identify_column(event.x)
        if not row_id or column not in {"#2", "#3", "#4"}:
            return
        try:
            index = int(row_id)
        except ValueError:
            return
        profile = self.app.macro_profiles[self.selected_macro_slot]
        if index >= len(profile.list):
            return
        bbox = self.macro_tree.bbox(row_id, column)
        if not bbox:
            return
        x, y, width, height = bbox
        macro_event = profile.list[index]
        t = self.app._t

        # 編集用ウィジェットのみ ttk のまま使用（Treeviewの上に被せるため）
        if column == "#2":
            choices = (
                KEYBOARD_EVENT_CHOICES if macro_event.type == EVENT_TYPE_KEYBOARD
                else MOUSE_EVENT_CHOICES if macro_event.type == EVENT_TYPE_MOUSE
                else None
            )
            if choices is None:
                return
            var = tk.StringVar(value=macro_event.name)
            editor = ttk.Combobox(self.macro_tree, textvariable=var, values=choices, state="readonly")
            editor.bind("<<ComboboxSelected>>", lambda _e: self._commit_editor())
        elif column == "#3":
            var = tk.StringVar(
                value=t("macro_action_press") if macro_event.action == ACTION_PRESS else t("macro_action_release")
            )
            editor = ttk.Combobox(
                self.macro_tree,
                textvariable=var,
                values=[t("macro_action_press"), t("macro_action_release")],
                state="readonly",
            )
            editor.bind("<<ComboboxSelected>>", lambda _e: self._commit_editor())
        else:
            var = tk.StringVar(value=str(macro_event.delay))
            editor = tk.Spinbox(self.macro_tree, from_=0, to=60000, textvariable=var)

        self._macro_editor_widget = editor
        self._macro_editor_info = {"index": index, "column": column, "var": var}
        editor.place(x=x, y=y, width=width, height=height)
        editor.bind("<Return>", lambda _e: self._commit_editor())
        editor.bind("<Escape>", lambda _e: self._close_editor())
        editor.bind("<FocusOut>", lambda _e: self.frame.after(100, self._commit_editor))
        editor.focus_set()

    # ──────────────────────────────── マクロ編集 ──────────────────────────────

    def _on_macro_name_changed(self):
        name = self.macro_name_var.get().strip() or f"Macro {self.selected_macro_slot + 1}"
        self.app.macro_profiles[self.selected_macro_slot].name = name
        self.app._macro_metadata.setdefault("names", {})[str(self.selected_macro_slot)] = name
        self.app._refresh_macro_profile_list()
        self.app.mouse_keys = [
            clone_binding(b, name=resolve_binding_name(b, self.app.macro_profiles))
            for b in self.app.mouse_keys
        ]
        self.app._refresh_key_mapping_ui()
        self.app._refresh_dirty_state()

    def _append_event(self, event: MacroEvent):
        self.app.macro_profiles[self.selected_macro_slot].list.append(event)
        self.selected_macro_event_index = len(self.app.macro_profiles[self.selected_macro_slot].list) - 1
        self.app._refresh_macro_profile_list()
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _add_event_pair(self, name: str, code: int, event_type: int):
        self._append_event(MacroEvent(name=name, code=code, type=event_type, action=ACTION_PRESS, delay=10))
        self._append_event(MacroEvent(name=name, code=code, type=event_type, action=ACTION_RELEASE, delay=10))

    def _add_manual_key_pair(self):
        key = self.manual_key_var.get()
        if key in self.manual_key_map:
            self._add_event_pair(key, self.manual_key_map[key], EVENT_TYPE_KEYBOARD)

    def _add_manual_mouse_pair(self):
        mouse = self.manual_mouse_var.get()
        if mouse in self.manual_mouse_map:
            self._add_event_pair(mouse, self.manual_mouse_map[mouse], EVENT_TYPE_MOUSE)

    def _move_event_up(self):
        profile = self.app.macro_profiles[self.selected_macro_slot]
        i = self.selected_macro_event_index
        if i is None or i <= 0:
            return
        profile.list[i - 1], profile.list[i] = profile.list[i], profile.list[i - 1]
        self.selected_macro_event_index -= 1
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _move_event_down(self):
        profile = self.app.macro_profiles[self.selected_macro_slot]
        i = self.selected_macro_event_index
        if i is None or i >= len(profile.list) - 1:
            return
        profile.list[i + 1], profile.list[i] = profile.list[i], profile.list[i + 1]
        self.selected_macro_event_index += 1
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _delete_event(self):
        profile = self.app.macro_profiles[self.selected_macro_slot]
        i = self.selected_macro_event_index
        if i is None or i >= len(profile.list):
            return
        profile.list.pop(i)
        self.selected_macro_event_index = (
            None if not profile.list else min(i, len(profile.list) - 1)
        )
        self.app._refresh_macro_profile_list()
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _clear_profile(self):
        self.app.macro_profiles[self.selected_macro_slot].list = []
        self.selected_macro_event_index = None
        self.app._refresh_macro_profile_list()
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _update_selected_macro_delay(self):
        i = self.selected_macro_event_index
        profile = self.app.macro_profiles[self.selected_macro_slot]
        if i is not None and i < len(profile.list):
            try:
                profile.list[i].delay = max(0, int(self.macro_delay_var.get() or "0"))
            except ValueError:
                pass
            self.app._refresh_macro_events()
            self.app._refresh_dirty_state()

    # ──────────────────────────────── 録画 ────────────────────────────────────

    def _record_delay(self, now_ms: int) -> int:
        mode = self.record_delay_mode.get()
        if mode == "none":
            return 0
        if mode == "fixed":
            return max(10, int(self.fixed_delay_var.get() or "10"))
        return 0 if self._last_record_time is None else max(0, now_ms - self._last_record_time)

    def _append_recorded_event(self, name: str, code: int, event_type: int, action: int):
        profile = self.app.macro_profiles[self.selected_macro_slot]
        now_ms = int(time.time() * 1000)
        if profile.list:
            profile.list[-1].delay = self._record_delay(now_ms)
        profile.list.append(MacroEvent(name=name, code=code, type=event_type, action=action, delay=0))
        self._last_record_time = now_ms
        self.selected_macro_event_index = len(profile.list) - 1
        self.app._refresh_macro_profile_list()
        self.app._refresh_macro_events()
        self.app._refresh_dirty_state()

    def _event_target_allows_record(self, widget) -> bool:
        cls_name = widget.winfo_class()
        return "Button" not in cls_name and "Entry" not in cls_name and "Listbox" not in cls_name and "Treeview" not in cls_name and "Combobox" not in cls_name

    def _hid_from_key_event(self, event) -> int | None:
        keysym = event.keysym
        simple = {"Escape": 41, "BackSpace": 42, "Tab": 43, "space": 44, "Return": 40}
        if keysym in simple:
            return simple[keysym]
        if len(keysym) == 1 and keysym.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
            return (4 + ord(keysym.upper()) - ord("A")) if keysym.isalpha() else (30 + "1234567890".index(keysym))
        from constants import HID_KEY_NAMES as _HKN
        return next((c for c, n in _HKN.items() if n == keysym), None)

    def handle_record_key_press(self, event):
        if not self.is_recording:
            return
        hid = self._hid_from_key_event(event)
        if hid is None or hid in self._pressed_keys:
            return "break"
        self._pressed_keys.add(hid)
        from constants import HID_KEY_NAMES
        self._append_recorded_event(HID_KEY_NAMES.get(hid, f"Key {hid}"), hid, EVENT_TYPE_KEYBOARD, ACTION_PRESS)
        return "break"

    def handle_record_key_release(self, event):
        if not self.is_recording:
            return
        hid = self._hid_from_key_event(event)
        if hid is None:
            return "break"
        self._pressed_keys.discard(hid)
        from constants import HID_KEY_NAMES
        self._append_recorded_event(HID_KEY_NAMES.get(hid, f"Key {hid}"), hid, EVENT_TYPE_KEYBOARD, ACTION_RELEASE)
        return "break"

    def handle_record_mouse_press(self, event):
        if not self.is_recording or not self._event_target_allows_record(event.widget):
            return
        code = {1: 1, 2: 4, 3: 2}.get(event.num)
        if code is None or code in self._pressed_mouse:
            return "break"
        self._pressed_mouse.add(code)
        from constants import MOUSE_CODE_NAMES
        self._append_recorded_event(MOUSE_CODE_NAMES.get(code, f"Mouse {code}"), code, EVENT_TYPE_MOUSE, ACTION_PRESS)
        return "break"

    def handle_record_mouse_release(self, event):
        if not self.is_recording or not self._event_target_allows_record(event.widget):
            return
        code = {1: 1, 2: 4, 3: 2}.get(event.num)
        if code is None:
            return "break"
        self._pressed_mouse.discard(code)
        from constants import MOUSE_CODE_NAMES
        self._append_recorded_event(MOUSE_CODE_NAMES.get(code, f"Mouse {code}"), code, EVENT_TYPE_MOUSE, ACTION_RELEASE)
        return "break"

    def _toggle_recording(self):
        self.is_recording = not self.is_recording
        self._pressed_keys.clear()
        self._pressed_mouse.clear()
        self._last_record_time = None
        t = self.app._t
        self.record_btn.configure(
            text=t("macro_record_stop") if self.is_recording else t("macro_record_start")
        )
        if self.is_recording:
            self.record_btn.configure(fg_color="#b91c1c", hover_color="#991b1b")
            self.app.bind_all("<KeyPress>", self.handle_record_key_press)
            self.app.bind_all("<KeyRelease>", self.handle_record_key_release)
            for seq in ("<ButtonPress-1>", "<ButtonPress-2>", "<ButtonPress-3>"):
                self.app.bind_all(seq, self.handle_record_mouse_press)
            for seq in ("<ButtonRelease-1>", "<ButtonRelease-2>", "<ButtonRelease-3>"):
                self.app.bind_all(seq, self.handle_record_mouse_release)
        else:
            self.record_btn.configure(fg_color="#0ea5a5", hover_color="#0d8a8a")
            for seq in (
                "<KeyPress>", "<KeyRelease>",
                "<ButtonPress-1>", "<ButtonPress-2>", "<ButtonPress-3>",
                "<ButtonRelease-1>", "<ButtonRelease-2>", "<ButtonRelease-3>",
            ):
                self.app.unbind_all(seq)

    # ──────────────────────────────── デバイス操作 ────────────────────────────

    def _reset_macro_data_on_device(self):
        t = self.app._t
        if not self.app.mouse.device:
            messagebox.showerror("Error", t("msg_err_conn"))
            return
        if not messagebox.askyesno("Confirm", t("reset_confirm")):
            return
        self.app._append_log("System: Resetting macro memory...")
        self.app._set_macro_status("resetting")

        def worker():
            self.app.mouse.reset_macro_data()
            return self.app._merge_macro_names(
                decode_macro_profiles(self.app.mouse.get_macro_data(), resolve_macro_event_name)
            )

        def on_success(profiles):
            self.app.macro_profiles = profiles
            self.app._macro_profiles_loaded = True
            self.app._saved_macro_profiles = deepcopy(profiles)
            self.app._set_macro_status("ready")
            self.app._refresh_macro_profile_list()
            self.app._refresh_macro_events()
            self.app.mouse_keys = [
                clone_binding(b, name=resolve_binding_name(b, self.app.macro_profiles))
                for b in self.app.mouse_keys
            ]
            self.app._refresh_key_mapping_ui()
            self.app._refresh_dirty_state()
            self.app._append_log("System: Macro profiles loaded from device.")

        self.app._run_in_background(worker, on_success=on_success)
