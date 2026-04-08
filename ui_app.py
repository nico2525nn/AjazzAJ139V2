import json
import ctypes
import os
import threading
import time
import tkinter as tk
from tkinter import font as tkfont

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
from copy import deepcopy
from tkinter import messagebox, ttk

import pystray
from PIL import Image, ImageDraw, ImageFont

from ajazz_mouse import (
    ACTION_PRESS,
    ACTION_RELEASE,
    DEFAULT_MOUSE_KEY_BINDINGS,
    EVENT_TYPE_KEYBOARD,
    EVENT_TYPE_MOUSE,
    AjazzMouse,
    MacroEvent,
    MacroProfile,
    MouseKeyBinding,
    decode_macro_profiles,
    encode_macro_profiles,
)


LANG_DICT = {
    "en": {
        "title": "Ajazz AJ139 V2 Control",
        "tab_status": "Status",
        "tab_perf": "Performance",
        "tab_sys": "System / Lighting",
        "tab_keys": "Key Mapping",
        "tab_macro": "Macro",
        "tab_debug": "Debug Logs",
        "btn_write": "Write Settings",
        "btn_refresh": "Refresh Status",
        "lang_label": "Language",
        "status_frame": "USB / Dongle Status",
        "status_val": "Status: {0}",
        "fw_val": "Firmware: {0}",
        "bat_val": "Battery: {0}",
        "st_disconnected": "Disconnected (Device not found)",
        "st_connected": "Connected (USB / Dongle)",
        "bat_offline": "Offline (Sleeping or disconnected)",
        "bat_charging": "(Charging)",
        "poll_rate": "Polling Rate (Hz):",
        "debounce": "Debounce Time (ms):",
        "lod": "LOD (Lift Off Distance):",
        "dpi_frame": "DPI Levels",
        "dpi_label": "DPI {0}:",
        "dpi_active": "Active DPI Index:",
        "light_mode": "Lighting Mode:",
        "sleep_time": "LED Auto Sleep (min):",
        "log_sys_search": "System: Searching for USB device...",
        "log_sys_conn": "System: Connected successfully.",
        "log_sys_fail": "System: Connection failed. Check the cable or dongle.",
        "log_sys_bat": "System: Online detected. Battery={0}%",
        "log_sys_off": "System: Device is offline or sleeping.",
        "log_sys_writing": "System: ==== Writing settings... ====",
        "msg_success": "Settings written successfully.",
        "msg_err_conn": "Device not connected or configuration not loaded.",
        "theme_label": "Theme",
        "theme_light": "Light",
        "theme_dark": "Dark",
        "save_page": "Save This Page",
        "save_disabled": "Nothing to save on this page",
        "save_perf": "Save Performance",
        "save_sys": "Save Lighting",
        "save_keys": "Save Key Mapping",
        "save_macro": "Save Macros",
        "unsaved_changes": "Unsaved changes",
        "unsaved_prompt": "There are unsaved changes on this page. Save them before leaving?",
        "dirty_hint": "Unsaved changes",
        "clean_hint": "Saved",
        "tray_open": "Open Settings",
        "tray_exit": "Exit",
        "keys_select": "Button Slots",
        "keys_current": "Current Assignment",
        "keys_presets": "Preset Assignments",
        "keys_keyboard": "Keyboard",
        "keys_media": "Media",
        "keys_mouse": "Mouse",
        "keys_apply_preset": "Apply Preset",
        "keys_modifier": "Modifier Combo",
        "keys_capture": "Press a key in the capture box",
        "keys_apply_modifier": "Apply Combo",
        "keys_macro": "Macro Assignment",
        "keys_macro_profile": "Macro Profile:",
        "keys_macro_mode": "Run Mode:",
        "keys_macro_repeat": "Repeat Count:",
        "keys_apply_macro": "Apply Macro",
        "keys_write": "Write Key Mapping",
        "keys_reset": "Reset Key Mapping",
        "macro_profiles": "Macro Slots",
        "macro_events": "Events",
        "macro_controls": "Recorder / Editor",
        "macro_name": "Display Name:",
        "macro_save_name": "Save Name",
        "macro_record_start": "Start Recording",
        "macro_record_stop": "Stop Recording",
        "macro_delay_mode": "Delay Handling",
        "macro_delay_exact": "Use actual delays",
        "macro_delay_none": "Zero delay",
        "macro_delay_fixed": "Fixed delay",
        "macro_fixed_ms": "Fixed delay (ms):",
        "macro_manual_key": "Manual key pair:",
        "macro_manual_mouse": "Manual mouse pair:",
        "macro_add_key": "Add Key Pair",
        "macro_add_mouse": "Add Mouse Pair",
        "macro_move_up": "Move Up",
        "macro_move_down": "Move Down",
        "macro_delete": "Delete Event",
        "macro_clear": "Clear Profile",
        "macro_update_delay": "Update Delay",
        "macro_delay": "Delay (ms):",
        "macro_write": "Write Macros",
        "macro_reload": "Reload From Device",
        "macro_reset": "Reset Macro Memory",
        "macro_empty": "No events in this profile.",
        "macro_type": "Type",
        "macro_name_col": "Name",
        "macro_action": "Action",
        "macro_delay_col": "Delay",
        "macro_action_press": "Press",
        "macro_action_release": "Release",
        "macro_event_mouse": "Mouse",
        "macro_event_keyboard": "Keyboard",
        "macro_mode_counted": "Counted",
        "macro_mode_vendor2": "Vendor Mode 2",
        "macro_mode_vendor3": "Vendor Mode 3",
        "macro_write_ok": "Macro data written successfully.",
        "keys_write_ok": "Key mapping written successfully.",
        "reset_confirm": "Reset the selected data on the device?",
    },
    "ja": {
        "title": "Ajazz AJ139 V2 コントロール",
        "tab_status": "状態",
        "tab_perf": "パフォーマンス",
        "tab_sys": "システム / ライティング",
        "tab_keys": "キーマッピング",
        "tab_macro": "マクロ",
        "tab_debug": "デバッグログ",
        "btn_write": "設定を書き込む",
        "btn_refresh": "状態を更新",
        "lang_label": "言語",
        "status_frame": "USB / ドングル状態",
        "status_val": "状態: {0}",
        "fw_val": "ファームウェア: {0}",
        "bat_val": "バッテリー: {0}",
        "st_disconnected": "未接続 (デバイスが見つかりません)",
        "st_connected": "接続中 (USB / ドングル)",
        "bat_offline": "オフライン (スリープ中または未接続)",
        "bat_charging": "(充電中)",
        "poll_rate": "ポーリングレート (Hz):",
        "debounce": "デバウンス時間 (ms):",
        "lod": "LOD (Lift Off Distance):",
        "dpi_frame": "DPI レベル",
        "dpi_label": "DPI {0}:",
        "dpi_active": "アクティブ DPI:",
        "light_mode": "ライティングモード:",
        "sleep_time": "LED 自動スリープ (分):",
        "log_sys_search": "System: USB デバイスを検索中...",
        "log_sys_conn": "System: 接続しました。",
        "log_sys_fail": "System: 接続に失敗しました。ケーブルまたはドングルを確認してください。",
        "log_sys_bat": "System: オンライン検出。バッテリー={0}%",
        "log_sys_off": "System: デバイスはオフラインまたはスリープ中です。",
        "log_sys_writing": "System: ==== 設定を書き込み中... ====",
        "msg_success": "設定を書き込みました。",
        "msg_err_conn": "デバイス未接続、または設定が読み込まれていません。",
        "tray_open": "設定を開く",
        "tray_exit": "終了",
        "keys_select": "ボタンスロット",
        "keys_current": "現在の割り当て",
        "keys_presets": "プリセット割り当て",
        "keys_keyboard": "キーボード",
        "keys_media": "メディア",
        "keys_mouse": "マウス",
        "keys_apply_preset": "プリセットを適用",
        "keys_modifier": "修飾キーコンボ",
        "keys_capture": "下の入力欄でキーを押してください",
        "keys_apply_modifier": "コンボを適用",
        "keys_macro": "マクロ割り当て",
        "keys_macro_profile": "マクロプロファイル:",
        "keys_macro_mode": "実行モード:",
        "keys_macro_repeat": "繰り返し回数:",
        "keys_apply_macro": "マクロを適用",
        "keys_write": "キーマッピングを書き込む",
        "keys_reset": "キーマッピングをリセット",
        "macro_profiles": "マクロスロット",
        "macro_events": "イベント",
        "macro_controls": "録画 / 編集",
        "macro_name": "表示名:",
        "macro_save_name": "名前を保存",
        "macro_record_start": "録画開始",
        "macro_record_stop": "録画停止",
        "macro_delay_mode": "遅延の扱い",
        "macro_delay_exact": "実際の遅延を使う",
        "macro_delay_none": "遅延なし",
        "macro_delay_fixed": "固定遅延",
        "macro_fixed_ms": "固定遅延 (ms):",
        "macro_manual_key": "手動キー追加:",
        "macro_manual_mouse": "手動マウス追加:",
        "macro_add_key": "キー対を追加",
        "macro_add_mouse": "マウス対を追加",
        "macro_move_up": "上へ",
        "macro_move_down": "下へ",
        "macro_delete": "イベント削除",
        "macro_clear": "プロファイルをクリア",
        "macro_update_delay": "遅延を更新",
        "macro_delay": "遅延 (ms):",
        "macro_write": "マクロを書き込む",
        "macro_reload": "デバイスから再読込",
        "macro_reset": "マクロ領域を初期化",
        "macro_empty": "このプロファイルにイベントはありません。",
        "macro_type": "種類",
        "macro_name_col": "名前",
        "macro_action": "動作",
        "macro_delay_col": "遅延",
        "macro_action_press": "押下",
        "macro_action_release": "解放",
        "macro_event_mouse": "マウス",
        "macro_event_keyboard": "キーボード",
        "macro_mode_counted": "回数指定",
        "macro_mode_vendor2": "Vendor Mode 2",
        "macro_mode_vendor3": "Vendor Mode 3",
        "macro_write_ok": "マクロを書き込みました。",
        "keys_write_ok": "キーマッピングを書き込みました。",
        "reset_confirm": "デバイス上のデータをリセットしますか？",
        "theme_label": "テーマ",
        "theme_light": "ライト",
        "theme_dark": "ダーク",
        "save_page": "このページを保存",
        "save_disabled": "このページに保存する変更はありません",
        "save_perf": "パフォーマンスを保存",
        "save_sys": "ライティングを保存",
        "save_keys": "キーマッピングを保存",
        "save_macro": "マクロを保存",
        "unsaved_changes": "未保存の変更",
        "unsaved_prompt": "このページに未保存の変更があります。タブを切り替える前に保存しますか？",
        "dirty_hint": "未保存の変更あり",
        "clean_hint": "保存済み",
    },
}


def _build_hid_name_map():
    mapping = {}
    for offset, char in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ", start=4):
        mapping[offset] = char
    for offset, char in enumerate("1234567890", start=30):
        mapping[offset] = char
    mapping.update(
        {
            40: "Enter",
            41: "Esc",
            42: "Backspace",
            43: "Tab",
            44: "Space",
            45: "-",
            46: "=",
            47: "[",
            48: "]",
            49: "\\",
            51: ";",
            52: "'",
            53: "`",
            54: ",",
            55: ".",
            56: "/",
            57: "CapsLock",
            58: "F1",
            59: "F2",
            60: "F3",
            61: "F4",
            62: "F5",
            63: "F6",
            64: "F7",
            65: "F8",
            66: "F9",
            67: "F10",
            68: "F11",
            69: "F12",
            104: "F13",
            105: "F14",
            106: "F15",
            107: "F16",
            108: "F17",
            109: "F18",
            110: "F19",
            111: "F20",
            112: "F21",
            113: "F22",
            114: "F23",
            115: "F24",
            70: "PrintScreen",
            71: "ScrollLock",
            72: "Pause",
            73: "Insert",
            74: "Home",
            75: "PageUp",
            76: "Delete",
            77: "End",
            78: "PageDown",
            79: "Right",
            80: "Left",
            81: "Down",
            82: "Up",
            83: "NumLock",
            84: "Num /",
            85: "Num *",
            86: "Num -",
            87: "Num +",
            88: "Num Enter",
            89: "Num 1",
            90: "Num 2",
            91: "Num 3",
            92: "Num 4",
            93: "Num 5",
            94: "Num 6",
            95: "Num 7",
            96: "Num 8",
            97: "Num 9",
            98: "Num 0",
            99: "Num .",
            101: "Menu",
            168: "Mute",
            169: "Volume Up",
            170: "Volume Down",
        }
    )
    return mapping


HID_KEY_NAMES = _build_hid_name_map()
MOUSE_CODE_NAMES = {1: "Mouse L", 2: "Mouse R", 4: "Mouse M", 8: "Mouse Backward", 16: "Mouse Forward"}
KEYBOARD_NAME_TO_CODE = {name: code for code, name in sorted(HID_KEY_NAMES.items())}
KEYBOARD_EVENT_CHOICES = list(KEYBOARD_NAME_TO_CODE.keys())
MOUSE_NAME_TO_CODE = {name: code for code, name in MOUSE_CODE_NAMES.items()}
MOUSE_EVENT_CHOICES = list(MOUSE_NAME_TO_CODE.keys())
BUTTON_SLOT_NAMES = [
    "Left Button",
    "Right Button",
    "Middle Button",
    "Backward Button",
    "Forward Button",
    "DPI Button",
    "Wheel Up",
    "Wheel Down",
]
MACRO_MODE_OPTIONS = [(0, "macro_mode_counted"), (2, "macro_mode_vendor2"), (3, "macro_mode_vendor3")]


def _make_binding(name, binding_type, code1, code2, code3=0, lang=""):
    return {"name": name, "type": binding_type, "code1": code1, "code2": code2, "code3": code3, "lang": lang}


KEYBOARD_PRESETS = [_make_binding(name, 16, 0, code, 0) for code, name in sorted(HID_KEY_NAMES.items())]
MEDIA_PRESETS = [
    _make_binding("Volume +", 48, 233, 0),
    _make_binding("Volume -", 48, 234, 0),
    _make_binding("Mute", 48, 226, 0),
    _make_binding("Play / Pause", 48, 205, 0),
    _make_binding("Stop", 48, 183, 0),
    _make_binding("Prev Track", 48, 182, 0),
    _make_binding("Next Track", 48, 181, 0),
    _make_binding("Homepage", 48, 35, 2),
    _make_binding("Web Refresh", 48, 39, 2),
]
MOUSE_PRESETS = [
    _make_binding("Left Click", 32, 1, 0),
    _make_binding("Right Click", 32, 2, 0),
    _make_binding("Middle Click", 32, 4, 0),
    _make_binding("Backward", 32, 8, 0),
    _make_binding("Forward", 32, 16, 0),
    _make_binding("Scroll Up", 33, 56, 1),
    _make_binding("Scroll Down", 33, 56, 255),
    _make_binding("DPI Loop +", 33, 85, 0),
    _make_binding("Disable", 32, 0, 0),
]


def clone_binding(binding: MouseKeyBinding, **updates) -> MouseKeyBinding:
    data = {
        "index": binding.index,
        "name": binding.name,
        "type": binding.type,
        "code1": binding.code1,
        "code2": binding.code2,
        "code3": binding.code3,
        "lang": binding.lang,
    }
    data.update(updates)
    return MouseKeyBinding(**data)


def copy_default_bindings():
    return [clone_binding(binding) for binding in DEFAULT_MOUSE_KEY_BINDINGS]


def modifier_names(code1: int):
    names = []
    if code1 & 0x01 or code1 & 0x10:
        names.append("CTRL")
    if code1 & 0x08 or code1 & 0x80:
        names.append("WIN")
    if code1 & 0x04 or code1 & 0x40:
        names.append("ALT")
    if code1 & 0x02 or code1 & 0x20:
        names.append("SHIFT")
    return names


def resolve_binding_name(binding: MouseKeyBinding, macro_profiles: list[MacroProfile]) -> str:
    if binding.type == 112:
        slot = min(max(binding.code1, 0), len(macro_profiles) - 1) if macro_profiles else 0
        return macro_profiles[slot].name if macro_profiles else f"Macro {slot + 1}"
    if binding.type == 16:
        key_name = HID_KEY_NAMES.get(binding.code2, "Unassigned")
        mods = modifier_names(binding.code1)
        return "+".join(mods + [key_name]) if mods else key_name
    for preset in MOUSE_PRESETS + MEDIA_PRESETS:
        if (
            preset["type"] == binding.type
            and preset["code1"] == binding.code1
            and preset["code2"] == binding.code2
            and preset["code3"] == binding.code3
        ):
            return preset["name"]
    return binding.name or "Unassigned"


def resolve_macro_event_name(event_type: int, code: int, action: int) -> str:
    if event_type == EVENT_TYPE_MOUSE:
        return MOUSE_CODE_NAMES.get(code, f"Mouse {code}")
    if event_type == EVENT_TYPE_KEYBOARD:
        return HID_KEY_NAMES.get(code, f"Key {code}")
    return f"Event {code}"


class AjazzApp(tk.Tk):
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
        self._tab_change_guard = False
        self._current_tab_id = None
        self._macro_editor_widget = None
        self._macro_editor_info = None

        self.config_file = "settings.json"
        self.macro_meta_file = "macro_metadata.json"
        self._load_app_settings()
        self._load_macro_metadata()

        self.title(self._t("title"))
        self.geometry("1180x1050")
        self.minsize(1080, 1050)

        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.tray_icon = None
        self._last_tray_battery = -1
        self.protocol("WM_DELETE_WINDOW", self._hide_window)

        self._build_ui()
        self._current_tab_id = self.notebook.select()
        for variable in (self.poll_var, self.debounce_var, self.lod_var, self.dpi_idx_var, self.light_var, self.sleep_var):
            variable.trace_add("write", lambda *_args: self._refresh_dirty_state())
        for variable in self.dpi_vars:
            variable.trace_add("write", lambda *_args: self._refresh_dirty_state())
        self._apply_theme(self.theme_name, persist=False)
        self.change_language(self.lang, skip_ui_update=True)
        self.after(50, self._refresh_status)

        self._bg_thread = threading.Thread(target=self._tray_monitor_loop, daemon=True)
        self._bg_thread.start()

    def _t(self, key, *args):
        text = LANG_DICT[self.lang].get(key, key)
        if args:
            return text.format(*args)
        return text

    def _ui_text(self, key):
        texts = {
            "keys_hint": {
                "en": "1. Select a button. 2. Choose or edit its assignment. 3. Write the result to the device.",
                "ja": "1. ボタンを選ぶ 2. 割り当てを選ぶまたは編集する 3. デバイスへ書き込む",
            },
            "keys_device_actions": {
                "en": "Device Actions",
                "ja": "デバイス操作",
            },
            "macro_hint": {
                "en": "Load a slot, edit its events, then write changes to the device.",
                "ja": "スロットを読み込んでイベントを編集し、最後にデバイスへ書き込みます。",
            },
            "macro_status_label": {
                "en": "Status:",
                "ja": "状態:",
            },
            "macro_status_idle": {
                "en": "Macro data will load when this tab opens.",
                "ja": "このタブを開くとマクロデータを読み込みます。",
            },
            "macro_status_loading": {
                "en": "Loading macro data from device...",
                "ja": "デバイスからマクロデータを読み込み中...",
            },
            "macro_status_ready": {
                "en": "Macro data loaded.",
                "ja": "マクロデータを読み込みました。",
            },
            "macro_status_writing": {
                "en": "Writing macro data to device...",
                "ja": "デバイスへマクロデータを書き込み中...",
            },
            "macro_status_resetting": {
                "en": "Resetting macro memory on device...",
                "ja": "デバイス上のマクロ領域を初期化中...",
            },
            "macro_status_error": {
                "en": "Failed to load macro data.",
                "ja": "マクロデータの読み込みに失敗しました。",
            },
            "macro_edit_hint": {
                "en": "Double-click Name, Action, or Delay to edit directly.",
                "ja": "Name / Action / Delay はダブルクリックで直接編集できます。",
            },
            "macro_device_actions": {
                "en": "Device Actions",
                "ja": "デバイス操作",
            },
        }
        entry = texts.get(key, {})
        return entry.get(self.lang, entry.get("en", key))

    def _set_macro_status(self, state_key):
        self._macro_status_key = state_key
        if hasattr(self, "macro_status_var"):
            self.macro_status_var.set(f"{self._ui_text('macro_status_label')} {self._ui_text(f'macro_status_{state_key}')}")

    def _set_macro_progress(self, current, total):
        self._macro_progress_current = max(0, int(current))
        self._macro_progress_total = max(0, int(total))
        if not hasattr(self, "macro_progress_var"):
            return
        maximum = self._macro_progress_total or 1
        self.macro_progress.configure(mode="determinate", maximum=maximum)
        self.macro_progress_var.set(min(self._macro_progress_current, maximum))
        percent = 0 if self._macro_progress_total <= 0 else round((min(self._macro_progress_current, self._macro_progress_total) / self._macro_progress_total) * 100)
        self.macro_progress_label_var.set(f"{percent}% ({min(self._macro_progress_current, self._macro_progress_total)}/{self._macro_progress_total})" if self._macro_progress_total > 0 else "")

    def _show_macro_progress(self):
        if hasattr(self, "macro_progress_row") and not self.macro_progress_row.winfo_ismapped():
            self.macro_progress_row.pack(fill=tk.X, pady=(4, 0))

    def _hide_macro_progress(self):
        if hasattr(self, "macro_progress_row") and self.macro_progress_row.winfo_ismapped():
            self.macro_progress_row.pack_forget()

    def _apply_window_chrome(self, dark):
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

    def _style_combobox_popdown(self, widget, palette):
        try:
            popdown = self.tk.call("ttk::combobox::PopdownWindow", str(widget))
            listbox_path = f"{popdown}.f.l"
            self.tk.call(
                listbox_path,
                "configure",
                "-background",
                palette["field"],
                "-foreground",
                palette["fg"],
                "-selectbackground",
                palette["accent"],
                "-selectforeground",
                palette["fg"],
                "-highlightthickness",
                0,
                "-borderwidth",
                0,
            )
        except Exception:
            pass

    def _apply_theme(self, theme_name, persist=True):
        palettes = {
            "light": {
                "bg": "#f0f2f5",
                "panel": "#e4e7ec",
                "field": "#ffffff",
                "fg": "#1a1d23",
                "muted": "#6b7280",
                "accent": "#0ea5a5",
                "accent_light": "#2dd4bf",
                "accent_dark": "#0d8a8a",
                "accent_fg": "#ffffff",
                "dirty": "#fef3c7",
                "dirty_fg": "#b45309",
                "border": "#d1d5db",
                "trough": "#e5e7eb",
                "tab_selected": "#ffffff",
                "heading": "#f3f4f6",
                "select_bg": "#0ea5a5",
                "select_fg": "#ffffff",
            },
            "dark": {
                "bg": "#0f1117",
                "panel": "#1a1d27",
                "field": "#21242f",
                "fg": "#e8eaed",
                "muted": "#8b8fa3",
                "accent": "#2dd4bf",
                "accent_light": "#5eead4",
                "accent_dark": "#14b8a6",
                "accent_fg": "#0f1117",
                "dirty": "#422006",
                "dirty_fg": "#fbbf24",
                "border": "#2e3241",
                "trough": "#151821",
                "tab_selected": "#21242f",
                "heading": "#1a1d27",
                "select_bg": "#2dd4bf",
                "select_fg": "#0f1117",
            },
        }
        palette = palettes.get(theme_name, palettes["dark"])
        self.theme_name = theme_name if theme_name in palettes else "dark"
        self._theme_palette = palette
        self.configure(bg=palette["bg"])
        
        available_fonts = tkfont.families()
        def get_best_font(choices):
            for f in choices:
                if f in available_fonts:
                    return f
            return choices[-1]

        is_ja = self.lang == "ja"
        font_family = get_best_font(["Noto Sans JP", "Noto Sans CJK JP", "Meiryo", "Meiryo UI", "Yu Gothic UI"]) if is_ja else get_best_font(["Segoe UI Variable Text", "Segoe UI Semibold", "Segoe UI"])
        font_size = 10
        font = (font_family, font_size)
        font_bold = (font_family, font_size, "bold")
        font_small = (font_family, font_size - 1)
        font_heading = (font_family, font_size + 1, "bold")
        self.style.configure(".", background=palette["bg"], foreground=palette["fg"], fieldbackground=palette["field"], font=font, borderwidth=1, bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"])
        self.style.configure("TFrame", background=palette["bg"], relief="flat")
        self.style.configure("TLabel", background=palette["bg"], foreground=palette["fg"], font=font, relief="flat")
        self.style.configure("Heading.TLabel", background=palette["bg"], foreground=palette["accent"], font=font_heading)
        self.style.configure("Muted.TLabel", background=palette["bg"], foreground=palette["muted"], font=font_small)
        self.style.configure("TLabelframe", background=palette["bg"], foreground=palette["accent"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], relief="solid", borderwidth=1)
        self.style.configure("TLabelframe.Label", background=palette["bg"], foreground=palette["accent"], font=font_bold)
        self.style.configure("TButton", background=palette["panel"], foreground=palette["fg"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], focusthickness=0, focuscolor=palette["border"], padding=(12, 6), font=font, relief="flat", borderwidth=1)
        self.style.map("TButton", background=[("active", palette["border"]), ("pressed", palette["accent_dark"])], foreground=[("disabled", palette["muted"]), ("pressed", palette["accent_fg"])], bordercolor=[("active", palette["muted"])])
        self.style.configure("Accent.TButton", background=palette["accent"], foreground=palette["accent_fg"], bordercolor=palette["accent_dark"], lightcolor=palette["accent_dark"], darkcolor=palette["accent_dark"], font=font_bold, padding=(14, 7), relief="flat")
        self.style.map("Accent.TButton", background=[("active", palette["accent_light"]), ("pressed", palette["accent_dark"])], foreground=[("active", palette["accent_fg"]), ("pressed", palette["accent_fg"])])
        self.style.configure("TEntry", fieldbackground=palette["field"], foreground=palette["fg"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], padding=(6, 4), font=font, relief="flat", borderwidth=1)
        self.style.configure("TCombobox", fieldbackground=palette["field"], foreground=palette["fg"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], arrowcolor=palette["muted"], padding=(6, 4), font=font, relief="flat", borderwidth=1)
        self.style.map("TCombobox", fieldbackground=[("readonly", palette["field"])], foreground=[("readonly", palette["fg"])], arrowcolor=[("active", palette["accent"])])
        self.style.configure("TSpinbox", fieldbackground=palette["field"], foreground=palette["fg"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], arrowcolor=palette["muted"], padding=(6, 4), font=font, relief="flat", borderwidth=1)
        self.style.configure("Treeview", background=palette["field"], fieldbackground=palette["field"], foreground=palette["fg"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], rowheight=28, font=font, relief="flat", borderwidth=1)
        self.style.map("Treeview", background=[("selected", palette["select_bg"])], foreground=[("selected", palette["select_fg"])])
        self.style.configure("Treeview.Heading", background=palette["heading"], foreground=palette["muted"], bordercolor=palette["border"], lightcolor=palette["border"], darkcolor=palette["border"], font=font_bold, padding=(6, 4), relief="flat", borderwidth=1)
        self.style.configure("TNotebook", background=palette["bg"], borderwidth=0, lightcolor=palette["bg"], darkcolor=palette["bg"], bordercolor=palette["bg"], tabmargins=(0, 2, 0, 0))
        self.style.configure("TNotebook.Tab", background=palette["panel"], foreground=palette["muted"], bordercolor=palette["panel"], lightcolor=palette["panel"], darkcolor=palette["panel"], focuscolor=palette["tab_selected"], borderwidth=1, padding=(16, 8), font=font, relief="flat")
        self.style.map("TNotebook.Tab", background=[("selected", palette["tab_selected"]), ("active", palette["field"])], foreground=[("selected", palette["accent"]), ("active", palette["fg"])], bordercolor=[("selected", palette["tab_selected"]), ("active", palette["field"])], lightcolor=[("selected", palette["tab_selected"]), ("active", palette["field"])], darkcolor=[("selected", palette["tab_selected"]), ("active", palette["field"])], padding=[("selected", (16, 8))])
        self.style.configure("TScrollbar", background=palette["panel"], troughcolor=palette["bg"], borderwidth=0, arrowsize=12)
        self.style.configure("TCheckbutton", background=palette["bg"], foreground=palette["fg"], font=font)
        self.style.map("TCheckbutton", background=[("active", palette["bg"])])
        self.style.configure("TRadiobutton", background=palette["bg"], foreground=palette["fg"], font=font)
        self.style.map("TRadiobutton", background=[("active", palette["bg"])])
        self.style.configure("TSeparator", background=palette["border"])
        self.style.configure("Macro.Horizontal.TProgressbar", troughcolor=palette["trough"], background=palette["accent"], lightcolor=palette["accent_light"], darkcolor=palette["accent_dark"], bordercolor=palette["border"])
        self.style.configure("PageSave.TButton", background=palette["accent"], foreground=palette["accent_fg"], bordercolor=palette["accent_dark"], font=font_bold, padding=(16, 7))
        self.style.map("PageSave.TButton", background=[("active", palette["accent_light"]), ("pressed", palette["accent_dark"]), ("disabled", palette["panel"])], foreground=[("disabled", palette["muted"]), ("active", palette["accent_fg"])])
        self.style.configure("Dirty.TEntry", fieldbackground=palette["dirty"], foreground=palette["fg"])
        self.style.configure("Dirty.TCombobox", fieldbackground=palette["dirty"], foreground=palette["fg"], arrowcolor=palette["fg"])
        self.style.configure("Dirty.TButton", background=palette["dirty"], foreground=palette["dirty_fg"], bordercolor=palette["dirty_fg"])
        self.style.configure("Dirty.TLabel", background=palette["bg"], foreground=palette["dirty_fg"], font=font_bold)
        self.style.configure("Clean.TLabel", background=palette["bg"], foreground=palette["accent"], font=font)
        listbox_opts = dict(bg=palette["field"], fg=palette["fg"], selectbackground=palette["select_bg"], selectforeground=palette["select_fg"], relief=tk.FLAT, bd=0, highlightthickness=1, highlightbackground=palette["border"], highlightcolor=palette["accent"], font=(font_family, font_size))
        if hasattr(self, "log_txt"):
            self.log_txt.configure(bg=palette["field"], fg=palette["muted"], insertbackground=palette["fg"], font=("Cascadia Code", 10) if not is_ja else ("Consolas", 10))
        if hasattr(self, "debug_frame"):
            self.debug_frame.configure(bg=palette["field"], highlightbackground=palette["border"], highlightcolor=palette["accent"], highlightthickness=1)
        if hasattr(self, "log_scrollbar"):
            self.log_scrollbar.configure(bg=palette["panel"], troughcolor=palette["bg"], activebackground=palette["field"], relief=tk.FLAT, bd=0, elementborderwidth=0, highlightbackground=palette["border"], highlightcolor=palette["border"], highlightthickness=0)
        if hasattr(self, "macro_profile_listbox"):
            self.macro_profile_listbox.configure(**listbox_opts)
        if hasattr(self, "keyboard_preset_list"):
            for listbox in (self.keyboard_preset_list, self.media_preset_list, self.mouse_preset_list):
                listbox.configure(**listbox_opts)
        for combobox_name in (
            "lang_combo", "theme_combo", "poll_combo", "lod_combo", "dpi_idx_combo", "light_combo",
            "key_macro_profile_combo", "key_macro_mode_combo", "manual_key_combo", "manual_mouse_combo",
        ):
            combobox = getattr(self, combobox_name, None)
            if combobox is not None:
                self._style_combobox_popdown(combobox, palette)
        self._apply_window_chrome(self.theme_name == "dark")
        if persist:
            self._save_app_settings()
        self._refresh_dirty_state()

    def _on_theme_selected(self):
        reverse_map = {
            "light": "light",
            "dark": "dark",
            self._t("theme_light"): "light",
            self._t("theme_dark"): "dark",
        }
        self._apply_theme(reverse_map.get(self.theme_var.get(), "light"))

    def _load_app_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as file:
                    cfg = json.load(file)
                self.lang = cfg.get("lang", "en")
                self.theme_name = cfg.get("theme", "light")
            except Exception:
                self.lang = "en"
                self.theme_name = "light"

    def _save_app_settings(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as file:
                json.dump({"lang": self.lang, "theme": self.theme_name}, file, ensure_ascii=False, indent=2)
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
            bg_color, fg_color = "gray", "white"
        elif battery_perc <= 20:
            bg_color, fg_color = "black", "red"
        else:
            bg_color, fg_color = "black", "white"

        image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.ellipse([2, 2, width - 2, height - 2], fill=bg_color, outline="white")
        text = "--" if battery_perc < 0 else str(battery_perc)

        try:
            font = ImageFont.truetype("arial.ttf", 36)
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
            self.log_txt.insert(tk.END, text + "\n")
            self.log_txt.see(tk.END)
            self.log_txt.configure(state="disabled")

    def _set_busy(self, busy: bool):
        self._busy_count = max(0, self._busy_count + (1 if busy else -1))
        state = tk.DISABLED if self._busy_count else tk.NORMAL
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
            if widget is None:
                continue
            try:
                widget.config(state=state)
            except tk.TclError:
                pass
        try:
            self.configure(cursor="watch" if self._busy_count else "")
        except tk.TclError:
            pass

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
                messagebox.showerror("Error", str(exc))

        def runner():
            try:
                result = worker()
            except Exception as exc:
                self.after(0, lambda exc=exc: finish_error(exc))
            else:
                self.after(0, lambda result=result: finish_success(result))

        threading.Thread(target=runner, daemon=True).start()

    def _on_notebook_tab_changed(self, event=None):
        selected = self.notebook.select()
        if self._current_tab_id is None:
            self._current_tab_id = selected
        if not self._tab_change_guard and self._current_tab_id != selected:
            previous_page = self._page_for_tab(self._current_tab_id)
            if previous_page and self._page_is_dirty(previous_page):
                answer = messagebox.askyesnocancel(self._t("unsaved_changes"), self._t("unsaved_prompt"))
                if answer is None:
                    self._tab_change_guard = True
                    self.notebook.select(self._current_tab_id)
                    self._tab_change_guard = False
                    return
                if answer:
                    self._save_page(previous_page)
                else:
                    self._discard_page_changes(previous_page)
            self._current_tab_id = selected
        else:
            self._current_tab_id = selected
        self._refresh_dirty_state()
        if self.notebook.select() == str(self.tab_macro):
            self._load_macro_profiles_async()

    def _build_ui(self):
        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=16, pady=(10, 0))
        self.app_title_label = ttk.Label(header, text="AJ139 V2", style="Heading.TLabel")
        self.app_title_label.pack(side=tk.LEFT, padx=(0, 16))
        self.lbl_lang_label = ttk.Label(header, style="Muted.TLabel")
        self.lbl_lang_label.pack(side=tk.LEFT)
        self.lang_var = tk.StringVar(value=self.lang)
        self.lang_combo = ttk.Combobox(header, textvariable=self.lang_var, values=["en", "ja"], state="readonly", width=6)
        self.lang_combo.pack(side=tk.LEFT, padx=(4, 12))
        self.lang_combo.bind("<<ComboboxSelected>>", lambda event: self.change_language(self.lang_var.get()))
        self.lbl_theme_label = ttk.Label(header, style="Muted.TLabel")
        self.lbl_theme_label.pack(side=tk.LEFT)
        self.theme_var = tk.StringVar(value=self.theme_name)
        self.theme_combo = ttk.Combobox(header, textvariable=self.theme_var, state="readonly", width=10)
        self.theme_combo.pack(side=tk.LEFT, padx=(4, 0))
        self.theme_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_theme_selected())

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 6))
        self.notebook.bind("<<NotebookTabChanged>>", self._on_notebook_tab_changed)
        self.tab_status = ttk.Frame(self.notebook, padding=16)
        self.tab_perf = ttk.Frame(self.notebook, padding=16)
        self.tab_sys = ttk.Frame(self.notebook, padding=16)
        self.tab_keys = ttk.Frame(self.notebook, padding=12)
        self.tab_macro = ttk.Frame(self.notebook, padding=12)
        self.tab_debug = ttk.Frame(self.notebook, padding=12)

        for tab in (self.tab_status, self.tab_perf, self.tab_sys, self.tab_keys, self.tab_macro, self.tab_debug):
            self.notebook.add(tab, text="")

        self._build_status_tab()
        self._build_perf_tab()
        self._build_sys_tab()
        self._build_keys_tab()
        self._build_macro_tab()
        self._build_debug_tab()

        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=16, pady=(4, 12))
        self.page_status_var = tk.StringVar(value="")
        self.page_status_label = ttk.Label(bottom, textvariable=self.page_status_var)
        self.page_status_label.pack(side=tk.LEFT)
        self.btn_save_page = ttk.Button(bottom, command=self._save_current_page, style="PageSave.TButton")
        self.btn_save_page.pack(side=tk.RIGHT)
        self.btn_refresh = ttk.Button(bottom, command=self._refresh_status)
        self.btn_refresh.pack(side=tk.RIGHT, padx=8)

    def _build_status_tab(self):
        self.status_var = tk.StringVar()
        self.version_var = tk.StringVar()
        self.battery_var = tk.StringVar()
        self.status_frame = ttk.LabelFrame(self.tab_status, padding=16)
        self.status_frame.pack(fill=tk.X)
        ttk.Label(self.status_frame, textvariable=self.status_var, style="Heading.TLabel").pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(self.status_frame, textvariable=self.version_var).pack(anchor=tk.W, pady=(4, 0))
        ttk.Label(self.status_frame, textvariable=self.battery_var).pack(anchor=tk.W, pady=(4, 0))

    def _build_perf_tab(self):
        row1 = ttk.Frame(self.tab_perf)
        row1.pack(fill=tk.X, pady=4)
        self.lbl_poll = ttk.Label(row1, width=24)
        self.lbl_poll.pack(side=tk.LEFT)
        self.poll_var = tk.StringVar(value="1000")
        self.poll_combo = ttk.Combobox(row1, textvariable=self.poll_var, values=["125", "250", "500", "1000"], state="readonly")
        self.poll_combo.pack(side=tk.LEFT)

        row2 = ttk.Frame(self.tab_perf)
        row2.pack(fill=tk.X, pady=4)
        self.lbl_debounce = ttk.Label(row2, width=24)
        self.lbl_debounce.pack(side=tk.LEFT)
        self.debounce_var = tk.StringVar(value="4")
        self.debounce_entry = ttk.Entry(row2, textvariable=self.debounce_var, width=12)
        self.debounce_entry.pack(side=tk.LEFT)

        row3 = ttk.Frame(self.tab_perf)
        row3.pack(fill=tk.X, pady=4)
        self.lbl_lod = ttk.Label(row3, width=24)
        self.lbl_lod.pack(side=tk.LEFT)
        self.lod_var = tk.StringVar(value="1mm")
        self.lod_combo = ttk.Combobox(row3, textvariable=self.lod_var, values=["1mm", "2mm"], state="readonly")
        self.lod_combo.pack(side=tk.LEFT)

        self.dpi_frame = ttk.LabelFrame(self.tab_perf, padding=10)
        self.dpi_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.dpi_lbls = []
        self.dpi_vars = []
        self.dpi_entries = []
        for index in range(6):
            frame = ttk.Frame(self.dpi_frame)
            frame.pack(fill=tk.X, pady=3)
            label = ttk.Label(frame, width=16)
            label.pack(side=tk.LEFT)
            value = tk.StringVar(value="400")
            entry = ttk.Entry(frame, textvariable=value, width=12)
            entry.pack(side=tk.LEFT)
            self.dpi_lbls.append(label)
            self.dpi_vars.append(value)
            self.dpi_entries.append(entry)

        dpi_index_row = ttk.Frame(self.dpi_frame)
        dpi_index_row.pack(fill=tk.X, pady=6)
        self.lbl_dpi_idx = ttk.Label(dpi_index_row, width=16)
        self.lbl_dpi_idx.pack(side=tk.LEFT)
        self.dpi_idx_var = tk.StringVar(value="1")
        self.dpi_idx_combo = ttk.Combobox(dpi_index_row, textvariable=self.dpi_idx_var, values=["1", "2", "3", "4", "5", "6"], state="readonly", width=8)
        self.dpi_idx_combo.pack(side=tk.LEFT)

    def _build_sys_tab(self):
        row1 = ttk.Frame(self.tab_sys)
        row1.pack(fill=tk.X, pady=8)
        self.lbl_light = ttk.Label(row1, width=24)
        self.lbl_light.pack(side=tk.LEFT)
        self.light_var = tk.StringVar(value="0")
        self.light_combo = ttk.Combobox(row1, textvariable=self.light_var, values=[str(i) for i in range(8)], state="readonly")
        self.light_combo.pack(side=tk.LEFT)

        row2 = ttk.Frame(self.tab_sys)
        row2.pack(fill=tk.X, pady=8)
        self.lbl_sleep = ttk.Label(row2, width=24)
        self.lbl_sleep.pack(side=tk.LEFT)
        self.sleep_var = tk.StringVar(value="10")
        self.sleep_entry = ttk.Entry(row2, textvariable=self.sleep_var, width=12)
        self.sleep_entry.pack(side=tk.LEFT)

    def _build_keys_tab(self):
        left = ttk.Frame(self.tab_keys)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        right = ttk.Frame(self.tab_keys)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.keys_slot_frame = ttk.LabelFrame(left, padding=8)
        self.keys_slot_frame.pack(fill=tk.Y, expand=True)
        self.key_slot_vars = []
        self.key_slot_buttons = []
        for index, slot_name in enumerate(BUTTON_SLOT_NAMES):
            variable = tk.StringVar(value=f"{slot_name}: -")
            button = ttk.Button(self.keys_slot_frame, textvariable=variable, width=28, command=lambda idx=index: self._select_button_slot(idx))
            button.pack(fill=tk.X, pady=2)
            self.key_slot_vars.append(variable)
            self.key_slot_buttons.append(button)

        self.keys_current_frame = ttk.LabelFrame(right, padding=12)
        self.keys_current_frame.pack(fill=tk.X)
        self.current_key_var = tk.StringVar(value="-")
        ttk.Label(self.keys_current_frame, textvariable=self.current_key_var, style="Heading.TLabel").pack(anchor=tk.W, pady=(0, 4))
        self.keys_hint_label = ttk.Label(self.keys_current_frame, justify=tk.LEFT, wraplength=620, style="Muted.TLabel")
        self.keys_hint_label.pack(anchor=tk.W, pady=(4, 0))

        self.keys_presets_frame = ttk.LabelFrame(right, padding=8)
        self.keys_presets_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.keys_preset_notebook = ttk.Notebook(self.keys_presets_frame)
        self.keys_preset_notebook.pack(fill=tk.BOTH, expand=True)
        self.keys_keyboard_tab = ttk.Frame(self.keys_preset_notebook, padding=4)
        self.keys_media_tab = ttk.Frame(self.keys_preset_notebook, padding=4)
        self.keys_mouse_tab = ttk.Frame(self.keys_preset_notebook, padding=4)
        self.keys_preset_notebook.add(self.keys_keyboard_tab, text="")
        self.keys_preset_notebook.add(self.keys_media_tab, text="")
        self.keys_preset_notebook.add(self.keys_mouse_tab, text="")
        self.keyboard_preset_list = self._build_preset_list(self.keys_keyboard_tab, KEYBOARD_PRESETS)
        self.media_preset_list = self._build_preset_list(self.keys_media_tab, MEDIA_PRESETS)
        self.mouse_preset_list = self._build_preset_list(self.keys_mouse_tab, MOUSE_PRESETS)
        for listbox in (self.keyboard_preset_list, self.media_preset_list, self.mouse_preset_list):
            listbox.bind("<Double-Button-1>", lambda _event: self._apply_selected_preset())
            listbox.bind("<Return>", lambda _event: self._apply_selected_preset())

        lower = ttk.Frame(right)
        lower.pack(fill=tk.X, pady=(10, 0))

        self.keys_modifier_frame = ttk.LabelFrame(lower, padding=8)
        self.keys_modifier_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.modifier_ctrl = tk.BooleanVar()
        self.modifier_shift = tk.BooleanVar()
        self.modifier_alt = tk.BooleanVar()
        self.modifier_win = tk.BooleanVar()
        modifiers_row = ttk.Frame(self.keys_modifier_frame)
        modifiers_row.pack(fill=tk.X)
        ttk.Checkbutton(modifiers_row, text="Ctrl", variable=self.modifier_ctrl).pack(side=tk.LEFT)
        ttk.Checkbutton(modifiers_row, text="Shift", variable=self.modifier_shift).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(modifiers_row, text="Alt", variable=self.modifier_alt).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(modifiers_row, text="Win", variable=self.modifier_win).pack(side=tk.LEFT, padx=4)
        self.capture_hint_label = ttk.Label(self.keys_modifier_frame)
        self.capture_hint_label.pack(anchor=tk.W, pady=(8, 2))
        self.modifier_capture_var = tk.StringVar(value="-")
        self.modifier_capture_entry = ttk.Entry(self.keys_modifier_frame, textvariable=self.modifier_capture_var, width=24)
        self.modifier_capture_entry.pack(anchor=tk.W)
        self.modifier_capture_entry.bind("<KeyPress>", self._capture_modifier_key)
        self.captured_modifier_hid = None
        for variable in (self.modifier_ctrl, self.modifier_shift, self.modifier_alt, self.modifier_win):
            variable.trace_add("write", lambda *_args: self._apply_modifier_combo())

        self.keys_macro_frame = ttk.LabelFrame(lower, padding=8)
        self.keys_macro_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.key_macro_profile_label = ttk.Label(self.keys_macro_frame)
        self.key_macro_profile_label.grid(row=0, column=0, sticky="w")
        self.key_macro_profile_var = tk.StringVar()
        self.key_macro_profile_combo = ttk.Combobox(self.keys_macro_frame, textvariable=self.key_macro_profile_var, state="readonly", width=28)
        self.key_macro_profile_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.key_macro_mode_label = ttk.Label(self.keys_macro_frame)
        self.key_macro_mode_label.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.key_macro_mode_var = tk.StringVar()
        self.key_macro_mode_code = tk.IntVar(value=0)
        self.macro_mode_display = {}
        self.key_macro_mode_combo = ttk.Combobox(self.keys_macro_frame, textvariable=self.key_macro_mode_var, state="readonly", width=28)
        self.key_macro_mode_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))
        self.key_macro_mode_combo.bind("<<ComboboxSelected>>", lambda _event: (self._sync_macro_mode_selection(), self._apply_macro_assignment()))
        self.key_macro_repeat_label = ttk.Label(self.keys_macro_frame)
        self.key_macro_repeat_label.grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.key_macro_repeat_var = tk.StringVar(value="1")
        self.key_macro_repeat_spinbox = ttk.Spinbox(self.keys_macro_frame, from_=1, to=255, textvariable=self.key_macro_repeat_var, width=10)
        self.key_macro_repeat_spinbox.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(6, 0))
        self.keys_macro_frame.columnconfigure(1, weight=1)
        self.key_macro_profile_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_macro_assignment())
        self.key_macro_repeat_spinbox.bind("<Return>", self._apply_macro_assignment)
        self.key_macro_repeat_spinbox.bind("<FocusOut>", self._apply_macro_assignment)

        self.keys_device_frame = ttk.LabelFrame(right, padding=8)
        self.keys_device_frame.pack(fill=tk.X, pady=(10, 0))
        key_actions = ttk.Frame(self.keys_device_frame)
        key_actions.pack(anchor=tk.E)
        self.reset_keys_btn = ttk.Button(key_actions, command=self._reset_key_mapping_on_device)
        self.reset_keys_btn.pack(side=tk.RIGHT)

    def _build_preset_list(self, parent, presets):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        listbox = tk.Listbox(frame, height=12, exportselection=False)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        for preset in presets:
            listbox.insert(tk.END, preset["name"])
        return listbox

    def _build_macro_tab(self):
        top = ttk.Frame(self.tab_macro)
        top.pack(fill=tk.X, pady=(0, 8))
        self.macro_hint_label = ttk.Label(top, justify=tk.LEFT, style="Muted.TLabel")
        self.macro_hint_label.pack(anchor=tk.W)
        self.macro_status_var = tk.StringVar(value="")
        self.macro_status_label = ttk.Label(top, textvariable=self.macro_status_var)
        self.macro_status_label.pack(anchor=tk.W, pady=(4, 0))
        self.macro_progress_row = ttk.Frame(top)
        self.macro_progress_var = tk.DoubleVar(value=0)
        self.macro_progress = ttk.Progressbar(self.macro_progress_row, variable=self.macro_progress_var, mode="determinate", maximum=1, style="Macro.Horizontal.TProgressbar")
        self.macro_progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.macro_progress_label_var = tk.StringVar(value="")
        self.macro_progress_label = ttk.Label(self.macro_progress_row, textvariable=self.macro_progress_label_var, width=12, anchor=tk.E)
        self.macro_progress_label.pack(side=tk.LEFT, padx=(8, 0))

        left = ttk.Frame(self.tab_macro)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        center = ttk.Frame(self.tab_macro)
        center.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        right = ttk.Frame(self.tab_macro)
        right.pack(side=tk.LEFT, fill=tk.Y)

        self.macro_profiles_frame = ttk.LabelFrame(left, padding=8)
        self.macro_profiles_frame.pack(fill=tk.BOTH, expand=True)
        self.macro_profile_listbox = tk.Listbox(self.macro_profiles_frame, width=26, exportselection=False)
        self.macro_profile_listbox.pack(fill=tk.BOTH, expand=True)
        self.macro_profile_listbox.bind("<<ListboxSelect>>", self._on_macro_profile_selected)

        self.macro_events_frame = ttk.LabelFrame(center, padding=8)
        self.macro_events_frame.pack(fill=tk.BOTH, expand=True)
        self.macro_tree = ttk.Treeview(self.macro_events_frame, columns=("type", "name", "action", "delay"), show="headings", height=13)
        for column, width in (("type", 90), ("name", 180), ("action", 90), ("delay", 80)):
            self.macro_tree.column(column, width=width, anchor=tk.CENTER if column != "name" else tk.W)
            self.macro_tree.heading(column, text="")
        self.macro_tree.pack(fill=tk.BOTH, expand=True)
        self.macro_tree.bind("<<TreeviewSelect>>", self._on_macro_event_selected)
        self.macro_tree.bind("<Double-1>", self._on_macro_tree_double_click)
        self.macro_edit_hint_label = ttk.Label(self.macro_events_frame, justify=tk.LEFT, style="Muted.TLabel")
        self.macro_edit_hint_label.pack(anchor=tk.W, pady=(6, 0))

        event_actions = ttk.Frame(self.macro_events_frame)
        event_actions.pack(fill=tk.X, pady=(8, 0))
        self.macro_move_up_btn = ttk.Button(event_actions, command=self._move_macro_event_up)
        self.macro_move_up_btn.pack(side=tk.LEFT)
        self.macro_move_down_btn = ttk.Button(event_actions, command=self._move_macro_event_down)
        self.macro_move_down_btn.pack(side=tk.LEFT, padx=4)
        self.macro_delete_btn = ttk.Button(event_actions, command=self._delete_macro_event)
        self.macro_delete_btn.pack(side=tk.LEFT, padx=4)
        self.macro_clear_btn = ttk.Button(event_actions, command=self._clear_macro_profile)
        self.macro_clear_btn.pack(side=tk.LEFT, padx=4)

        delay_row = ttk.Frame(self.macro_events_frame)
        delay_row.pack(fill=tk.X, pady=(8, 0))
        self.macro_delay_label = ttk.Label(delay_row)
        self.macro_delay_label.pack(side=tk.LEFT)
        self.macro_delay_var = tk.StringVar(value="10")
        self.macro_delay_spinbox = ttk.Spinbox(delay_row, from_=0, to=60000, textvariable=self.macro_delay_var, width=10)
        self.macro_delay_spinbox.pack(side=tk.LEFT, padx=6)

        self.macro_controls_frame = ttk.LabelFrame(right, padding=8)
        self.macro_controls_frame.pack(fill=tk.BOTH, expand=True)
        name_row = ttk.Frame(self.macro_controls_frame)
        name_row.pack(fill=tk.X)
        self.macro_name_label = ttk.Label(name_row)
        self.macro_name_label.pack(side=tk.LEFT)
        self.macro_name_var = tk.StringVar()
        self.macro_name_entry = ttk.Entry(name_row, textvariable=self.macro_name_var, width=18)
        self.macro_name_entry.pack(side=tk.LEFT, padx=6)
        self.record_btn = ttk.Button(self.macro_controls_frame, command=self._toggle_recording, style="Accent.TButton")
        self.record_btn.pack(fill=tk.X, pady=(10, 0))

        self.delay_mode_frame = ttk.LabelFrame(self.macro_controls_frame, padding=6)
        self.delay_mode_frame.pack(fill=tk.X, pady=(10, 0))
        self.record_delay_mode = tk.StringVar(value="exact")
        self.delay_exact_radio = ttk.Radiobutton(self.delay_mode_frame, variable=self.record_delay_mode, value="exact")
        self.delay_exact_radio.pack(anchor=tk.W)
        self.delay_none_radio = ttk.Radiobutton(self.delay_mode_frame, variable=self.record_delay_mode, value="none")
        self.delay_none_radio.pack(anchor=tk.W)
        fixed_row = ttk.Frame(self.delay_mode_frame)
        fixed_row.pack(fill=tk.X)
        self.delay_fixed_radio = ttk.Radiobutton(fixed_row, variable=self.record_delay_mode, value="fixed")
        self.delay_fixed_radio.pack(side=tk.LEFT)
        self.fixed_delay_label = ttk.Label(fixed_row)
        self.fixed_delay_label.pack(side=tk.LEFT, padx=(4, 0))
        self.fixed_delay_var = tk.StringVar(value="10")
        ttk.Spinbox(fixed_row, from_=10, to=1000, textvariable=self.fixed_delay_var, width=8).pack(side=tk.LEFT, padx=6)

        ttk.Separator(self.macro_controls_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        self.manual_key_label = ttk.Label(self.macro_controls_frame)
        self.manual_key_label.pack(anchor=tk.W)
        self.manual_key_map = dict(KEYBOARD_NAME_TO_CODE)
        self.manual_key_var = tk.StringVar(value=KEYBOARD_PRESETS[0]["name"])
        self.manual_key_combo = ttk.Combobox(self.macro_controls_frame, textvariable=self.manual_key_var, values=list(self.manual_key_map.keys()), state="readonly", width=20)
        self.manual_key_combo.pack(fill=tk.X, pady=(2, 4))
        self.add_key_pair_btn = ttk.Button(self.macro_controls_frame, command=self._add_manual_key_pair)
        self.add_key_pair_btn.pack(fill=tk.X)
        self.manual_mouse_label = ttk.Label(self.macro_controls_frame)
        self.manual_mouse_label.pack(anchor=tk.W, pady=(10, 0))
        self.manual_mouse_map = dict(MOUSE_NAME_TO_CODE)
        self.manual_mouse_var = tk.StringVar(value="Mouse L")
        self.manual_mouse_combo = ttk.Combobox(self.macro_controls_frame, textvariable=self.manual_mouse_var, values=list(self.manual_mouse_map.keys()), state="readonly", width=20)
        self.manual_mouse_combo.pack(fill=tk.X, pady=(2, 4))
        self.add_mouse_pair_btn = ttk.Button(self.macro_controls_frame, command=self._add_manual_mouse_pair)
        self.add_mouse_pair_btn.pack(fill=tk.X)

        ttk.Separator(self.macro_controls_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        self.macro_device_frame = ttk.LabelFrame(self.macro_controls_frame, padding=6)
        self.macro_device_frame.pack(fill=tk.X)
        self.reload_macro_btn = ttk.Button(self.macro_device_frame, command=lambda: self._load_macro_profiles_async(force=True))
        self.reload_macro_btn.pack(fill=tk.X)
        self.reset_macro_btn = ttk.Button(self.macro_device_frame, command=self._reset_macro_data_on_device)
        self.reset_macro_btn.pack(fill=tk.X, pady=(6, 0))

        self._set_macro_status("idle")
        self._set_macro_progress(0, 0)
        self._hide_macro_progress()
        self.macro_name_var.trace_add("write", lambda *_args: self._on_macro_name_changed())
        self.macro_delay_spinbox.bind("<Return>", lambda _event: self._update_selected_macro_delay())
        self.macro_delay_spinbox.bind("<FocusOut>", lambda _event: self._update_selected_macro_delay())

    def _build_debug_tab(self):
        self.debug_frame = tk.Frame(self.tab_debug, bd=0, highlightthickness=1)
        self.debug_frame.pack(fill=tk.BOTH, expand=True)
        self.log_scrollbar = tk.Scrollbar(self.debug_frame, orient=tk.VERTICAL)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_txt = tk.Text(
            self.debug_frame,
            wrap=tk.WORD,
            state="disabled",
            font=("Consolas", 10),
            yscrollcommand=self.log_scrollbar.set,
            bd=0,
            highlightthickness=0,
            relief=tk.FLAT,
        )
        self.log_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_scrollbar.configure(command=self.log_txt.yview)

    def change_language(self, new_lang, skip_ui_update=False):
        self.lang = new_lang
        self.lang_var.set(new_lang)
        self._save_app_settings()
        labels = [
            self.tab_status,
            self.tab_perf,
            self.tab_sys,
            self.tab_keys,
            self.tab_macro,
            self.tab_debug,
        ]
        keys = ["tab_status", "tab_perf", "tab_sys", "tab_keys", "tab_macro", "tab_debug"]
        self.title(self._t("title"))
        for tab, key in zip(labels, keys):
            self.notebook.tab(tab, text=self._t(key))
        self.lbl_lang_label.config(text=self._t("lang_label"))
        self.lbl_theme_label.config(text=self._t("theme_label"))
        self.theme_combo["values"] = ["light", "dark"] if self.lang == "en" else [self._t("theme_light"), self._t("theme_dark")]
        self.theme_var.set(self.theme_name if self.lang == "en" else self._t(f"theme_{self.theme_name}"))
        self.btn_refresh.config(text=self._t("btn_refresh"))
        self.status_frame.config(text=self._t("status_frame"))
        self.lbl_poll.config(text=self._t("poll_rate"))
        self.lbl_debounce.config(text=self._t("debounce"))
        self.lbl_lod.config(text=self._t("lod"))
        self.dpi_frame.config(text=self._t("dpi_frame"))
        self.lbl_dpi_idx.config(text=self._t("dpi_active"))
        self.lbl_light.config(text=self._t("light_mode"))
        self.lbl_sleep.config(text=self._t("sleep_time"))
        for index in range(6):
            self.dpi_lbls[index].config(text=self._t("dpi_label", index + 1))

        self.keys_slot_frame.config(text=self._t("keys_select"))
        self.keys_current_frame.config(text=self._t("keys_current"))
        self.keys_hint_label.config(text=self._ui_text("keys_hint"))
        self.keys_presets_frame.config(text=self._t("keys_presets"))
        self.keys_modifier_frame.config(text=self._t("keys_modifier"))
        self.keys_macro_frame.config(text=self._t("keys_macro"))
        self.keys_device_frame.config(text=self._ui_text("keys_device_actions"))
        self.keys_preset_notebook.tab(self.keys_keyboard_tab, text=self._t("keys_keyboard"))
        self.keys_preset_notebook.tab(self.keys_media_tab, text=self._t("keys_media"))
        self.keys_preset_notebook.tab(self.keys_mouse_tab, text=self._t("keys_mouse"))
        self.capture_hint_label.config(text=self._t("keys_capture"))
        self.key_macro_profile_label.config(text=self._t("keys_macro_profile"))
        self.key_macro_mode_label.config(text=self._t("keys_macro_mode"))
        self.key_macro_repeat_label.config(text=self._t("keys_macro_repeat"))
        self.reset_keys_btn.config(text=self._t("keys_reset"))

        self.macro_profiles_frame.config(text=self._t("macro_profiles"))
        self.macro_events_frame.config(text=self._t("macro_events"))
        self.macro_controls_frame.config(text=self._t("macro_controls"))
        self.macro_hint_label.config(text=self._ui_text("macro_hint"))
        self.macro_edit_hint_label.config(text=self._ui_text("macro_edit_hint"))
        self.macro_device_frame.config(text=self._ui_text("macro_device_actions"))
        self.macro_name_label.config(text=self._t("macro_name"))
        self.delay_mode_frame.config(text=self._t("macro_delay_mode"))
        self.delay_exact_radio.config(text=self._t("macro_delay_exact"))
        self.delay_none_radio.config(text=self._t("macro_delay_none"))
        self.delay_fixed_radio.config(text=self._t("macro_delay_fixed"))
        self.fixed_delay_label.config(text=self._t("macro_fixed_ms"))
        self.manual_key_label.config(text=self._t("macro_manual_key"))
        self.manual_mouse_label.config(text=self._t("macro_manual_mouse"))
        self.add_key_pair_btn.config(text=self._t("macro_add_key"))
        self.add_mouse_pair_btn.config(text=self._t("macro_add_mouse"))
        self.macro_move_up_btn.config(text=self._t("macro_move_up"))
        self.macro_move_down_btn.config(text=self._t("macro_move_down"))
        self.macro_delete_btn.config(text=self._t("macro_delete"))
        self.macro_clear_btn.config(text=self._t("macro_clear"))
        self.macro_delay_label.config(text=self._t("macro_delay"))
        self.reload_macro_btn.config(text=self._t("macro_reload"))
        self.reset_macro_btn.config(text=self._t("macro_reset"))
        self._set_macro_status(self._macro_status_key)

        self.macro_mode_display = {value: self._t(key) for value, key in MACRO_MODE_OPTIONS}
        self.key_macro_mode_combo["values"] = list(self.macro_mode_display.values())
        self.key_macro_mode_var.set(self.macro_mode_display.get(self.key_macro_mode_code.get(), self._t("macro_mode_counted")))
        self.record_btn.config(text=self._t("macro_record_stop") if self.is_recording else self._t("macro_record_start"))
        self.macro_tree.heading("type", text=self._t("macro_type"))
        self.macro_tree.heading("name", text=self._t("macro_name_col"))
        self.macro_tree.heading("action", text=self._t("macro_action"))
        self.macro_tree.heading("delay", text=self._t("macro_delay_col"))

        if self.tray_icon:
            self.tray_icon.menu = self._create_tray_menu()

        self._refresh_key_mapping_ui()
        self._refresh_macro_profile_list()
        self._refresh_macro_events()
        self._refresh_dirty_state()
        if not skip_ui_update:
            self._refresh_status()

    def _load_config_to_ui(self):
        cfg = self.current_config
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
        return {
            "light": str(self.current_config["light_mode"]),
            "sleep": str(self.current_config["sleep_light"]),
        }

    def _current_sys_snapshot(self):
        return {
            "light": self.light_var.get(),
            "sleep": self.sleep_var.get(),
        }

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

    def _page_for_tab(self, tab_id):
        mapping = {
            str(self.tab_perf): "perf",
            str(self.tab_sys): "sys",
            str(self.tab_keys): "keys",
            str(self.tab_macro): "macro",
        }
        return mapping.get(tab_id)

    def _set_field_dirty_style(self, widget, dirty, widget_type):
        if widget is None:
            return
        styles = {
            "entry": ("TEntry", "Dirty.TEntry"),
            "combo": ("TCombobox", "Dirty.TCombobox"),
            "button": ("TButton", "Dirty.TButton"),
            "label": ("Clean.TLabel", "Dirty.TLabel"),
        }
        normal_style, dirty_style = styles[widget_type]
        widget.configure(style=dirty_style if dirty else normal_style)

    def _refresh_dirty_state(self):
        perf_saved = self._saved_perf_snapshot()
        perf_current = self._current_perf_snapshot()
        if perf_saved:
            self._set_field_dirty_style(self.poll_combo, perf_current["poll"] != perf_saved["poll"], "combo")
            self._set_field_dirty_style(self.debounce_entry, perf_current["debounce"] != perf_saved["debounce"], "entry")
            self._set_field_dirty_style(self.lod_combo, perf_current["lod"] != perf_saved["lod"], "combo")
            self._set_field_dirty_style(self.dpi_idx_combo, perf_current["dpi_index"] != perf_saved["dpi_index"], "combo")
            for index, entry in enumerate(self.dpi_entries):
                self._set_field_dirty_style(entry, perf_current["dpis"][index] != perf_saved["dpis"][index], "entry")

        sys_saved = self._saved_sys_snapshot()
        sys_current = self._current_sys_snapshot()
        if sys_saved:
            self._set_field_dirty_style(self.light_combo, sys_current["light"] != sys_saved["light"], "combo")
            self._set_field_dirty_style(self.sleep_entry, sys_current["sleep"] != sys_saved["sleep"], "entry")

        dirty_slots = {index for index, (current, saved) in enumerate(zip(self.mouse_keys, self._saved_mouse_keys)) if current != saved}
        for index, button in enumerate(self.key_slot_buttons):
            self._set_field_dirty_style(button, index in dirty_slots, "button")

        self.macro_profile_listbox.selection_clear(0, tk.END)
        for index, _profile in enumerate(self.macro_profiles):
            if index < self.macro_profile_listbox.size():
                changed = index < len(self._saved_macro_profiles) and self.macro_profiles[index] != self._saved_macro_profiles[index]
                palette = getattr(self, "_theme_palette", {"dirty": "#fff4cc", "field": "#ffffff", "fg": "#1d1d1f"})
                self.macro_profile_listbox.itemconfig(index, bg=palette["dirty"] if changed else palette["field"], fg=palette["fg"])
        if self.macro_profiles:
            self.macro_profile_listbox.selection_set(self.selected_macro_slot)

        current_page = self._page_for_tab(self.notebook.select())
        if current_page and self._page_is_dirty(current_page):
            self.page_status_var.set(self._t("dirty_hint"))
            self._set_field_dirty_style(self.page_status_label, True, "label")
        else:
            self.page_status_var.set(self._t("clean_hint"))
            self._set_field_dirty_style(self.page_status_label, False, "label")

        save_key = {
            "perf": "save_perf",
            "sys": "save_sys",
            "keys": "save_keys",
            "macro": "save_macro",
        }.get(current_page, "save_page")
        self.btn_save_page.config(text=self._t(save_key))
        if current_page in {"perf", "sys", "keys", "macro"}:
            self.btn_save_page.config(state=tk.NORMAL)
        else:
            self.btn_save_page.config(state=tk.DISABLED)

    def _discard_page_changes(self, page):
        if page == "perf" and self.current_config:
            self._load_config_to_ui()
        elif page == "sys" and self.current_config:
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
        page = self._page_for_tab(self.notebook.select())
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

    def _refresh_status(self):
        if self._status_refresh_in_progress:
            return
        self._status_refresh_in_progress = True
        self.log_txt.configure(state="normal")
        self.log_txt.delete("1.0", tk.END)
        self.log_txt.configure(state="disabled")
        self._append_log(self._t("log_sys_search"))
        self.status_var.set(self._t("status_val", "..."))
        self.version_var.set(self._t("fw_val", "--"))
        self.battery_var.set(self._t("bat_val", "--"))

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
                battery_info = result["battery_info"]
                suffix = f" {self._t('bat_charging')}" if battery_info["is_charging"] else ""
                self.battery_var.set(self._t("bat_val", f"{battery_info['battery']}%{suffix}"))
                self._append_log(self._t("log_sys_bat", battery_info["battery"]))
                self._last_tray_battery = battery_info["battery"]
                if self.tray_icon:
                    self.tray_icon.icon = self._create_image(self._last_tray_battery)
            else:
                self.battery_var.set(self._t("bat_val", self._t("bat_offline")))
                self._append_log(self._t("log_sys_off"))
                self._last_tray_battery = -1
                if self.tray_icon:
                    self.tray_icon.icon = self._create_image(-1)

            self.current_config = result["config"]
            if self.current_config:
                self._load_config_to_ui()
            self.mouse_keys = [clone_binding(binding, name=resolve_binding_name(binding, self.macro_profiles)) for binding in result["keys"]]
            self._saved_mouse_keys = [clone_binding(binding) for binding in self.mouse_keys]
            self._refresh_key_mapping_ui()
            self._refresh_dirty_state()

            if self.notebook.select() == str(self.tab_macro):
                self._load_macro_profiles_async()

        def on_error(exc):
            self._status_refresh_in_progress = False
            self._set_macro_status("error")
            self.status_var.set(self._t("status_val", self._t("st_disconnected")))
            self.version_var.set(self._t("fw_val", "--"))
            self.battery_var.set(self._t("bat_val", "--"))
            self._append_log(f"ERROR: {exc}")

        self._run_in_background(worker, on_success=on_success, on_error=on_error)

    def _load_macro_profiles_async(self, force=False, log=True):
        if self._macro_profiles_loading or (self._macro_profiles_loaded and not force):
            return
        if not self.mouse.device:
            return
        self._macro_profiles_loading = True
        self._set_macro_status("loading")
        self._set_macro_progress(0, 0)
        self._show_macro_progress()
        if log:
            self._append_log("System: Loading macro profiles from device...")

        def worker():
            return self._merge_macro_names(
                decode_macro_profiles(
                    self.mouse.get_macro_data(
                        progress_callback=lambda current, total: self.after(0, lambda current=current, total=total: self._set_macro_progress(current, total))
                    ),
                    resolve_macro_event_name,
                )
            )

        def on_success(profiles):
            self._macro_profiles_loading = False
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
        for index, variable in enumerate(self.key_slot_vars):
            variable.set(f"{BUTTON_SLOT_NAMES[index]}: {self.mouse_keys[index].name}")
        selected = self._selected_binding()
        self.current_key_var.set(f"{BUTTON_SLOT_NAMES[selected.index]} -> {selected.name}")
        if selected.type == 16:
            self.modifier_ctrl.set(bool(selected.code1 & 0x01 or selected.code1 & 0x10))
            self.modifier_shift.set(bool(selected.code1 & 0x02 or selected.code1 & 0x20))
            self.modifier_alt.set(bool(selected.code1 & 0x04 or selected.code1 & 0x40))
            self.modifier_win.set(bool(selected.code1 & 0x08 or selected.code1 & 0x80))
            self.captured_modifier_hid = selected.code2
            self.modifier_capture_var.set(HID_KEY_NAMES.get(selected.code2, "-"))
        if selected.type == 112 and self.macro_profiles:
            self.key_macro_profile_combo.current(min(max(selected.code1, 0), len(self.macro_profiles) - 1))
            self.key_macro_repeat_var.set(str(max(1, selected.code2 or 1)))
            self.key_macro_mode_code.set(selected.code3)
            self.key_macro_mode_var.set(self.macro_mode_display.get(selected.code3, self._t("macro_mode_counted")))
        self.key_macro_profile_combo["values"] = [profile.name for profile in self.macro_profiles]
        if self.macro_profiles and self.key_macro_profile_combo.current() < 0:
            self.key_macro_profile_combo.current(0)
        self._refresh_dirty_state()

    def _selected_preset(self):
        current_tab = self.keys_preset_notebook.index(self.keys_preset_notebook.select())
        source = {
            0: (self.keyboard_preset_list, KEYBOARD_PRESETS),
            1: (self.media_preset_list, MEDIA_PRESETS),
            2: (self.mouse_preset_list, MOUSE_PRESETS),
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
            if keysym.isalpha():
                hid = 4 + ord(keysym.upper()) - ord("A")
            else:
                hid = 30 + "1234567890".index(keysym)
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
        code1 = (0x01 if self.modifier_ctrl.get() else 0) | (0x02 if self.modifier_shift.get() else 0) | (0x04 if self.modifier_alt.get() else 0) | (0x08 if self.modifier_win.get() else 0)
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
        slot = self.key_macro_profile_combo.current()
        if slot < 0:
            slot = 0
        profile = self.macro_profiles[slot]
        self.mouse_keys[self.selected_button_index] = MouseKeyBinding(
            index=self.selected_button_index,
            name=profile.name,
            type=112,
            code1=slot,
            code2=max(1, int(self.key_macro_repeat_var.get() or "1")),
            code3=int(self.key_macro_mode_code.get()),
        )
        self._refresh_key_mapping_ui()
        self._refresh_dirty_state()

    def _write_key_mapping_to_device(self):
        if not self.mouse.device:
            messagebox.showerror("Error", self._t("msg_err_conn"))
            return
        self._append_log("System: Writing key mapping...")

        def worker():
            self.mouse.set_mouse_keys(self.mouse_keys)
            return None

        def on_success(_):
            self._saved_mouse_keys = [clone_binding(binding) for binding in self.mouse_keys]
            self._refresh_dirty_state()
            self._append_log("System: Key mapping written.")
            messagebox.showinfo("OK", self._t("keys_write_ok"))

        self._run_in_background(worker, on_success=on_success)

    def _reset_key_mapping_on_device(self):
        if not self.mouse.device:
            messagebox.showerror("Error", self._t("msg_err_conn"))
            return
        if not messagebox.askyesno("Confirm", self._t("reset_confirm")):
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
        self.macro_profile_listbox.delete(0, tk.END)
        for profile in self.macro_profiles:
            suffix = f" ({len(profile.list)})" if profile.list else ""
            self.macro_profile_listbox.insert(tk.END, f"{profile.slot + 1:02d}. {profile.name}{suffix}")
        self.macro_profile_listbox.selection_clear(0, tk.END)
        self.macro_profile_listbox.selection_set(self.selected_macro_slot)
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

    def _on_macro_profile_selected(self, event=None):
        self._close_macro_editor()
        selection = self.macro_profile_listbox.curselection()
        if selection:
            self.selected_macro_slot = selection[0]
            self.selected_macro_event_index = None
            self.macro_name_var.set(self.macro_profiles[self.selected_macro_slot].name)
            self._refresh_macro_events()

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
            if macro_event.type == EVENT_TYPE_KEYBOARD:
                code = KEYBOARD_NAME_TO_CODE.get(value)
            else:
                code = MOUSE_NAME_TO_CODE.get(value)
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
            if macro_event.type == EVENT_TYPE_KEYBOARD:
                values = KEYBOARD_EVENT_CHOICES
            elif macro_event.type == EVENT_TYPE_MOUSE:
                values = MOUSE_EVENT_CHOICES
            else:
                return
            variable = tk.StringVar(value=macro_event.name)
            editor = ttk.Combobox(self.macro_tree, textvariable=variable, values=values, state="readonly")
            editor.bind("<<ComboboxSelected>>", lambda _event: self._commit_macro_editor())
        elif column == "#3":
            variable = tk.StringVar(
                value=self._t("macro_action_press") if macro_event.action == ACTION_PRESS else self._t("macro_action_release")
            )
            editor = ttk.Combobox(
                self.macro_tree,
                textvariable=variable,
                values=[self._t("macro_action_press"), self._t("macro_action_release")],
                state="readonly",
            )
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
            profile.list[index].delay = max(0, int(self.macro_delay_var.get() or "0"))
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
        return widget.winfo_class() not in {"TButton", "Button", "TCombobox", "Combobox", "Treeview", "Listbox"}

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
        self.record_btn.config(text=self._t("macro_record_stop") if self.is_recording else self._t("macro_record_start"))
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
            messagebox.showerror("Error", self._t("msg_err_conn"))
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
            messagebox.showinfo("OK", self._t("macro_write_ok"))

        self._run_in_background(worker, on_success=on_success)

    def _reset_macro_data_on_device(self):
        if not self.mouse.device:
            messagebox.showerror("Error", self._t("msg_err_conn"))
            return
        if not messagebox.askyesno("Confirm", self._t("reset_confirm")):
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

    def _save_page(self, page):
        if page == "perf":
            self._save_perf_page()
        elif page == "sys":
            self._save_sys_page()
        elif page == "keys":
            self._write_key_mapping_to_device()
        elif page == "macro":
            self._write_macros_to_device()

    def _save_perf_page(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror("Error", self._t("msg_err_conn"))
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
            messagebox.showinfo("OK", self._t("msg_success"))

        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=on_success)

    def _save_sys_page(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror("Error", self._t("msg_err_conn"))
            return
        cfg = deepcopy(self.current_config)
        cfg["sleep_light"] = int(self.sleep_var.get() or "10")
        cfg["light_mode"] = int(self.light_var.get() or "0")
        self._append_log(self._t("log_sys_writing"))

        def on_success(_):
            self.current_config = cfg
            self._refresh_dirty_state()
            messagebox.showinfo("OK", self._t("msg_success"))

        self._run_in_background(lambda: self.mouse.set_config(cfg), on_success=on_success)
