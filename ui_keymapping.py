"""
ui_keymapping.py – Key Mapping タブのウィジェットとロジック (CustomTkinter化)
"""

import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from ajazz_mouse import MacroProfile, MouseKeyBinding
from constants import (
    BUTTON_SLOT_NAMES,
    HID_KEY_NAMES,
    KEYBOARD_PRESETS,
    MACRO_MODE_OPTIONS,
    MEDIA_PRESETS,
    MOUSE_PRESETS,
)
from ui_helpers import clone_binding, modifier_names, resolve_binding_name


class KeyMappingTab:
    def __init__(self, parent_frame: ctk.CTkFrame, app):
        self.app = app
        self.frame = parent_frame

        self.selected_button_index = 0
        self.captured_modifier_hid: int | None = None
        self.modifier_ctrl = ctk.BooleanVar()
        self.modifier_shift = ctk.BooleanVar()
        self.modifier_alt = ctk.BooleanVar()
        self.modifier_win = ctk.BooleanVar()
        self.modifier_capture_var = ctk.StringVar(value="-")
        
        self.key_slot_vars: list[ctk.StringVar] = []
        self.key_slot_buttons: list[ctk.CTkButton] = []
        
        self.current_key_var = ctk.StringVar(value="-")
        
        self.key_macro_profile_var = ctk.StringVar()
        self.key_macro_mode_var = ctk.StringVar()
        self.key_macro_mode_code = ctk.IntVar(value=0)
        self.key_macro_repeat_var = ctk.StringVar(value="1")
        self.macro_mode_display: dict[int, str] = {}

        self._build()
        for v in (self.modifier_ctrl, self.modifier_shift, self.modifier_alt, self.modifier_win):
            v.trace_add("write", lambda *_: self._apply_modifier_combo())

    # ──────────────────────────────── UI 構築 ────────────────────────────────

    def _build(self):
        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)

        left = ctk.CTkFrame(self.frame, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        
        right = ctk.CTkFrame(self.frame, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")

        # ── 左: ボタンスロット ──
        self.keys_slot_frame_title = ctk.CTkLabel(left, text="Button Slots", font=self.app.font_heading, text_color="#2dd4bf")
        self.keys_slot_frame_title.pack(anchor="w", pady=(0, 10))

        self.keys_slot_frame = ctk.CTkFrame(left)
        self.keys_slot_frame.pack(fill="y", expand=True)

        for index, slot_name in enumerate(BUTTON_SLOT_NAMES):
            var = ctk.StringVar(value=f"{slot_name}: -")
            btn = ctk.CTkButton(
                self.keys_slot_frame,
                textvariable=var,
                width=220,
                height=36,
                font=self.app.font_main,
                fg_color="transparent",
                border_width=1,
                border_color="gray50",
                text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                command=lambda idx=index: self._select_slot(idx),
            )
            btn.pack(fill="x", padx=10, pady=6)
            self.key_slot_vars.append(var)
            self.key_slot_buttons.append(btn)

        # ── 右上: 現在の割り当て ──
        self.keys_current_frame = ctk.CTkFrame(right)
        self.keys_current_frame.pack(fill="x", pady=(0, 15))
        
        inner_curr = ctk.CTkFrame(self.keys_current_frame, fg_color="transparent")
        inner_curr.pack(padx=20, pady=15, fill="x")
        
        self.keys_current_title = ctk.CTkLabel(inner_curr, text="Current Assignment", font=self.app.font_heading, text_color="#2dd4bf")
        self.keys_current_title.pack(anchor="w")
        
        ctk.CTkLabel(
            inner_curr,
            textvariable=self.current_key_var,
            font=self.app.font_title,
        ).pack(anchor="w", pady=(5, 5))
        
        self.keys_hint_label = ctk.CTkLabel(
            inner_curr,
            text="",
            font=self.app.font_small,
            text_color="gray50"
        )
        self.keys_hint_label.pack(anchor="w")

        # ── 右中: プリセット ──
        self.keys_presets_frame = ctk.CTkFrame(right)
        self.keys_presets_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        inner_preset = ctk.CTkFrame(self.keys_presets_frame, fg_color="transparent")
        inner_preset.pack(fill="both", expand=True, padx=20, pady=15)
        
        self.keys_presets_title = ctk.CTkLabel(inner_preset, text="Preset Assignments", font=self.app.font_heading, text_color="#2dd4bf")
        self.keys_presets_title.pack(anchor="w", pady=(0, 10))

        self.keys_preset_tabview = ctk.CTkTabview(inner_preset, corner_radius=10)
        try:
            # 内部のセグメントボタンに対してフォントを適用
            self.keys_preset_tabview._segmented_button.configure(font=self.app.font_main)
        except Exception:
            pass
        self.keys_preset_tabview.pack(fill="both", expand=True)

        self.keys_keyboard_tab = self.keys_preset_tabview.add("Keyboard")
        self.keys_media_tab = self.keys_preset_tabview.add("Media")
        self.keys_mouse_tab = self.keys_preset_tabview.add("Mouse")

        self.keyboard_preset_list = self._build_preset_list(self.keys_keyboard_tab, KEYBOARD_PRESETS)
        self.media_preset_list = self._build_preset_list(self.keys_media_tab, MEDIA_PRESETS)
        self.mouse_preset_list = self._build_preset_list(self.keys_mouse_tab, MOUSE_PRESETS)
        
        for lb in (self.keyboard_preset_list, self.media_preset_list, self.mouse_preset_list):
            lb.bind("<Double-Button-1>", lambda _e: self._apply_selected_preset())
            lb.bind("<Return>", lambda _e: self._apply_selected_preset())

        # ── 右下: 修飾キー / マクロアサイン ──
        lower = ctk.CTkFrame(right, fg_color="transparent")
        lower.pack(fill="x")
        lower.grid_columnconfigure((0, 1), weight=1)

        # 修飾キーフレーム
        self.keys_modifier_frame = ctk.CTkFrame(lower)
        self.keys_modifier_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        inner_mod = ctk.CTkFrame(self.keys_modifier_frame, fg_color="transparent")
        inner_mod.pack(fill="both", expand=True, padx=15, pady=15)
        
        self.keys_modifier_title = ctk.CTkLabel(inner_mod, text="Modifier Combo", font=self.app.font_heading, text_color="#2dd4bf")
        self.keys_modifier_title.pack(anchor="w", pady=(0, 10))
        
        mod_row = ctk.CTkFrame(inner_mod, fg_color="transparent")
        mod_row.pack(fill="x")
        for text, var in [
            ("Ctrl", self.modifier_ctrl), ("Shift", self.modifier_shift),
            ("Alt", self.modifier_alt), ("Win", self.modifier_win),
        ]:
            ctk.CTkCheckBox(mod_row, text=text, variable=var, font=self.app.font_main).pack(side="left", padx=(0, 10))
            
        self.capture_hint_label = ctk.CTkLabel(inner_mod, text="Press a key in the capture box", font=self.app.font_small, text_color="gray50")
        self.capture_hint_label.pack(anchor="w", pady=(10, 2))
        
        self.modifier_capture_entry = ctk.CTkEntry(
            inner_mod,
            textvariable=self.modifier_capture_var,
            width=200,
            font=self.app.font_main
        )
        self.modifier_capture_entry.pack(anchor="w")
        self.modifier_capture_entry.bind("<KeyPress>", self._capture_modifier_key)

        # マクロアサインフレーム
        self.keys_macro_frame = ctk.CTkFrame(lower)
        self.keys_macro_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        inner_mac = ctk.CTkFrame(self.keys_macro_frame, fg_color="transparent")
        inner_mac.pack(fill="both", expand=True, padx=15, pady=15)
        inner_mac.grid_columnconfigure(1, weight=1)

        self.keys_macro_title = ctk.CTkLabel(inner_mac, text="Macro Assignment", font=self.app.font_heading, text_color="#2dd4bf")
        self.keys_macro_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.key_macro_profile_label = ctk.CTkLabel(inner_mac, text="Macro Profile:", font=self.app.font_main)
        self.key_macro_profile_label.grid(row=1, column=0, sticky="w")
        self.key_macro_profile_combo = ctk.CTkOptionMenu(
            inner_mac,
            variable=self.key_macro_profile_var,
            font=self.app.font_main,
            command=self._apply_macro_assignment
        )
        self.key_macro_profile_combo.grid(row=1, column=1, sticky="ew", padx=(10, 0))

        self.key_macro_mode_label = ctk.CTkLabel(inner_mac, text="Run Mode:", font=self.app.font_main)
        self.key_macro_mode_label.grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.key_macro_mode_combo = ctk.CTkOptionMenu(
            inner_mac,
            variable=self.key_macro_mode_var,
            font=self.app.font_main,
            command=self._apply_macro_mode_assignment
        )
        self.key_macro_mode_combo.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=(10, 0))

        self.key_macro_repeat_label = ctk.CTkLabel(inner_mac, text="Repeat Count:", font=self.app.font_main)
        self.key_macro_repeat_label.grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.key_macro_repeat_entry = ctk.CTkEntry(
            inner_mac,
            textvariable=self.key_macro_repeat_var,
            width=100,
            font=self.app.font_main
        )
        self.key_macro_repeat_entry.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        self.key_macro_repeat_entry.bind("<Return>", lambda _e: self._apply_macro_assignment())
        self.key_macro_repeat_entry.bind("<FocusOut>", lambda _e: self._apply_macro_assignment())

        # デバイス操作フレーム
        actions_frame = ctk.CTkFrame(right, fg_color="transparent")
        actions_frame.pack(fill="x", pady=(15, 0))
        self.reset_keys_btn = ctk.CTkButton(
            actions_frame, 
            text="Reset Key Mapping", 
            font=self.app.font_button, 
            fg_color="gray30", hover_color="gray40", text_color="white",
            command=self._reset_key_mapping_on_device
        )
        self.reset_keys_btn.pack(side="right")

    def _build_preset_list(self, parent, presets) -> tk.Listbox:
        # リストの外側に少し余白を設けることで、親の丸角が見えるようにする
        frame = ctk.CTkFrame(parent, fg_color="black", corner_radius=8)
        frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        f_name = self.app.font_main.cget("family")
        # Listbox 自体は角丸にできないため、親フレームで包んで余白をあける
        lb = tk.Listbox(
            frame, height=12, exportselection=False, 
            borderwidth=0, highlightthickness=0,
            bg="black", fg="white", # 黒背景に統一
            font=(f_name, 11)
        )
        sb = ctk.CTkScrollbar(frame, orientation="vertical", command=lb.yview, fg_color="black", button_color="#333333", button_hover_color="#444444")
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        sb.pack(side="right", fill="y", padx=(0, 5), pady=5)
        
        for preset in presets:
            lb.insert(tk.END, preset["name"])
        return lb
        
    def _apply_macro_mode_assignment(self, val):
        self._sync_macro_mode_selection()
        self._apply_macro_assignment()

    # ──────────────────────────────── 言語更新 ────────────────────────────────

    def apply_lang(self, t, ui_text):
        self.keys_slot_frame_title.configure(text=t("keys_select"))
        self.keys_current_title.configure(text=t("keys_current"))
        self.keys_hint_label.configure(text=ui_text("keys_hint"))
        self.keys_presets_title.configure(text=t("keys_presets"))
        self.keys_modifier_title.configure(text=t("keys_modifier"))
        self.keys_macro_title.configure(text=t("keys_macro"))
        
        # CTkTabview doesn't have an easy way to rename tabs once created without recreating or accessing inner dict
        # So we keep it English or try internal accesses. Skipping Tabview rename for simplicity, or we can clear and add.
        # It's an acceptable compromise to leave it English / setup once, but let's try internal rename if possible.
        try:
            self.keys_preset_tabview._segmented_button._buttons_dict["Keyboard"].configure(text=t("keys_keyboard"))
            self.keys_preset_tabview._segmented_button._buttons_dict["Media"].configure(text=t("keys_media"))
            self.keys_preset_tabview._segmented_button._buttons_dict["Mouse"].configure(text=t("keys_mouse"))
        except Exception:
            pass

        self.capture_hint_label.configure(text=t("keys_capture"))
        self.key_macro_profile_label.configure(text=t("keys_macro_profile"))
        self.key_macro_mode_label.configure(text=t("keys_macro_mode"))
        self.key_macro_repeat_label.configure(text=t("keys_macro_repeat"))
        self.reset_keys_btn.configure(text=t("keys_reset"))

        self.macro_mode_display = {v: t(k) for v, k in MACRO_MODE_OPTIONS}
        self.key_macro_mode_combo.configure(values=list(self.macro_mode_display.values()))
        self.key_macro_mode_var.set(
            self.macro_mode_display.get(self.key_macro_mode_code.get(), t("macro_mode_counted"))
        )

    def apply_listbox_style(self, opts: dict):
        for lb in (self.keyboard_preset_list, self.media_preset_list, self.mouse_preset_list):
            lb.configure(**opts)

    # ──────────────────────────────── UI 更新 ────────────────────────────────

    def refresh(self, mouse_keys: list[MouseKeyBinding], macro_profiles: list[MacroProfile]):
        for i, var in enumerate(self.key_slot_vars):
            var.set(f"{BUTTON_SLOT_NAMES[i]}: {mouse_keys[i].name}")
            
        selected = mouse_keys[self.selected_button_index]
        self.current_key_var.set(f"{BUTTON_SLOT_NAMES[selected.index]} → {selected.name}")
        
        if selected.type == 16:
            self.modifier_ctrl.set(bool(selected.code1 & 0x01 or selected.code1 & 0x10))
            self.modifier_shift.set(bool(selected.code1 & 0x02 or selected.code1 & 0x20))
            self.modifier_alt.set(bool(selected.code1 & 0x04 or selected.code1 & 0x40))
            self.modifier_win.set(bool(selected.code1 & 0x08 or selected.code1 & 0x80))
            self.captured_modifier_hid = selected.code2
            self.modifier_capture_var.set(HID_KEY_NAMES.get(selected.code2, "-"))
            
        if selected.type == 112 and macro_profiles:
            idx = min(max(selected.code1, 0), len(macro_profiles) - 1)
            self.key_macro_profile_var.set(macro_profiles[idx].name)
            self.key_macro_repeat_var.set(str(max(1, selected.code2 or 1)))
            self.key_macro_mode_code.set(selected.code3)
            self.key_macro_mode_var.set(
                self.macro_mode_display.get(selected.code3, list(self.macro_mode_display.values())[0] if self.macro_mode_display else "")
            )
            
        profiles_list = [p.name for p in macro_profiles]
        if profiles_list:
            self.key_macro_profile_combo.configure(values=profiles_list)
            if self.key_macro_profile_var.get() not in profiles_list:
                self.key_macro_profile_var.set(profiles_list[0])

    def refresh_dirty(self, dirty_slots: set[int]):
        for i, btn in enumerate(self.key_slot_buttons):
            self.app._set_field_dirty_style(btn, i in dirty_slots, "button")

    # ──────────────────────────────── スロット選択 ────────────────────────────

    def _select_slot(self, index: int):
        self.selected_button_index = index
        # Change active button style
        for i, btn in enumerate(self.key_slot_buttons):
            if i == index:
                btn.configure(fg_color="#2dd4bf", text_color=("gray10", "gray10"), hover_color="#5eead4")
            else:
                btn.configure(fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"))
        
        self.app._refresh_key_mapping_ui()
        self.app._refresh_dirty_state()

    # ──────────────────────────────── プリセット ──────────────────────────────

    def _selected_preset(self) -> dict | None:
        tab_name = self.keys_preset_tabview.get()
        mapping = {
            "Keyboard": (self.keyboard_preset_list, KEYBOARD_PRESETS),
            "Media": (self.media_preset_list, MEDIA_PRESETS),
            "Mouse": (self.mouse_preset_list, MOUSE_PRESETS)
        }
        if tab_name not in mapping:
            return None
        lb, presets = mapping[tab_name]
        sel = lb.curselection()
        return presets[sel[0]] if sel else None

    def _apply_selected_preset(self):
        preset = self._selected_preset()
        if not preset:
            return
        self.app.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=preset["name"],
            type=preset["type"],
            code1=preset["code1"],
            code2=preset["code2"],
            code3=preset["code3"],
            lang=preset["lang"],
        )
        self.app._refresh_key_mapping_ui()
        self.app._refresh_dirty_state()

    # ──────────────────────────────── 修飾キー ────────────────────────────────

    def _capture_modifier_key(self, event):
        keysym = event.keysym
        hid: int | None = None
        simple_map = {"Escape": 41, "BackSpace": 42, "Tab": 43, "space": 44, "Return": 40}
        if keysym in simple_map:
            hid = simple_map[keysym]
        elif len(keysym) == 1 and keysym.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789":
            hid = (4 + ord(keysym.upper()) - ord("A")) if keysym.isalpha() else (30 + "1234567890".index(keysym))
        hid = hid or next((c for c, n in HID_KEY_NAMES.items() if n == keysym), None)
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
        name = "+".join(
            modifier_names(code1) + [HID_KEY_NAMES.get(self.captured_modifier_hid, f"Key {self.captured_modifier_hid}")]
        )
        self.app.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=name,
            type=16,
            code1=code1,
            code2=self.captured_modifier_hid,
            code3=0,
        )
        self.app._refresh_key_mapping_ui()
        self.app._refresh_dirty_state()

    # ──────────────────────────────── マクロアサイン ──────────────────────────

    def _sync_macro_mode_selection(self):
        reverse = {label: code for code, label in self.macro_mode_display.items()}
        self.key_macro_mode_code.set(reverse.get(self.key_macro_mode_var.get(), 0))

    def _apply_macro_assignment(self, _val=None):
        if not self.app.macro_profiles: return
        p_name = self.key_macro_profile_var.get()
        slot = next((i for i, p in enumerate(self.app.macro_profiles) if p.name == p_name), 0)
        
        profile = self.app.macro_profiles[slot]
        try:
            repeat = max(1, int(self.key_macro_repeat_var.get() or "1"))
        except ValueError:
            repeat = 1
            
        self.app.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=profile.name,
            type=112,
            code1=slot,
            code2=repeat,
            code3=int(self.key_macro_mode_code.get()),
        )
        self.app._refresh_key_mapping_ui()
        self.app._refresh_dirty_state()

    # ──────────────────────────────── デバイス操作 ────────────────────────────

    def _reset_key_mapping_on_device(self):
        t = self.app._t
        if not self.app.mouse.device:
            messagebox.showerror("Error", t("msg_err_conn"))
            return
        if not messagebox.askyesno("Confirm", t("reset_confirm")):
            return
        self.app._append_log("System: Resetting key mapping...")

        def worker():
            self.app.mouse.reset_mouse_keys()
            return self.app.mouse.get_mouse_keys()

        def on_success(bindings):
            self.app.mouse_keys = [
                clone_binding(b, name=resolve_binding_name(b, self.app.macro_profiles))
                for b in bindings
            ]
            self.app._saved_mouse_keys = [clone_binding(b) for b in self.app.mouse_keys]
            self.app._refresh_key_mapping_ui()
            self.app._refresh_dirty_state()
            self.app._append_log("System: Key mapping loaded from device.")

        self.app._run_in_background(worker, on_success=on_success)
