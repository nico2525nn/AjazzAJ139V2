"""
constants.py – 言語辞書 / HIDキーマップ / プリセット定義
"""

from ajazz_mouse import MouseKeyBinding

# ──────────────────────────────────────────────────────────────
# 言語辞書
# ──────────────────────────────────────────────────────────────
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
        "macro_record_start": "⏺  Start Recording",
        "macro_record_stop": "⏹  Stop Recording",
        "macro_delay_mode": "Delay Handling",
        "macro_delay_exact": "Use actual delays",
        "macro_delay_none": "Zero delay",
        "macro_delay_fixed": "Fixed delay",
        "macro_fixed_ms": "Fixed delay (ms):",
        "macro_manual_key": "Manual key pair:",
        "macro_manual_mouse": "Manual mouse pair:",
        "macro_add_key": "Add Key Pair",
        "macro_add_mouse": "Add Mouse Pair",
        "macro_move_up": "▲ Move Up",
        "macro_move_down": "▼ Move Down",
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
        "macro_record_start": "⏺  録画開始",
        "macro_record_stop": "⏹  録画停止",
        "macro_delay_mode": "遅延の扱い",
        "macro_delay_exact": "実際の遅延を使う",
        "macro_delay_none": "遅延なし",
        "macro_delay_fixed": "固定遅延",
        "macro_fixed_ms": "固定遅延 (ms):",
        "macro_manual_key": "手動キー追加:",
        "macro_manual_mouse": "手動マウス追加:",
        "macro_add_key": "キー対を追加",
        "macro_add_mouse": "マウス対を追加",
        "macro_move_up": "▲ 上へ",
        "macro_move_down": "▼ 下へ",
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

UI_TEXT = {
    "keys_hint": {
        "en": "1. Select a button.  2. Choose or edit its assignment.  3. Write the result to the device.",
        "ja": "1. ボタンを選ぶ  2. 割り当てを選ぶまたは編集する  3. デバイスへ書き込む",
    },
    "keys_device_actions": {"en": "Device Actions", "ja": "デバイス操作"},
    "macro_hint": {
        "en": "Load a slot, edit its events, then write changes to the device.",
        "ja": "スロットを読み込んでイベントを編集し、最後にデバイスへ書き込みます。",
    },
    "macro_status_label": {"en": "Status:", "ja": "状態:"},
    "macro_status_idle": {
        "en": "Macro data will load when this tab opens.",
        "ja": "このタブを開くとマクロデータを読み込みます。",
    },
    "macro_status_loading": {
        "en": "Loading macro data from device...",
        "ja": "デバイスからマクロデータを読み込み中...",
    },
    "macro_status_ready": {"en": "Macro data loaded.", "ja": "マクロデータを読み込みました。"},
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
    "macro_device_actions": {"en": "Device Actions", "ja": "デバイス操作"},
}

# ──────────────────────────────────────────────────────────────
# HID キーマップ
# ──────────────────────────────────────────────────────────────
def _build_hid_name_map() -> dict[int, str]:
    mapping: dict[int, str] = {}
    for offset, char in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ", start=4):
        mapping[offset] = char
    for offset, char in enumerate("1234567890", start=30):
        mapping[offset] = char
    mapping.update(
        {
            40: "Enter", 41: "Esc", 42: "Backspace", 43: "Tab", 44: "Space",
            45: "-", 46: "=", 47: "[", 48: "]", 49: "\\", 51: ";", 52: "'",
            53: "`", 54: ",", 55: ".", 56: "/", 57: "CapsLock",
            58: "F1", 59: "F2", 60: "F3", 61: "F4", 62: "F5", 63: "F6",
            64: "F7", 65: "F8", 66: "F9", 67: "F10", 68: "F11", 69: "F12",
            104: "F13", 105: "F14", 106: "F15", 107: "F16", 108: "F17",
            109: "F18", 110: "F19", 111: "F20", 112: "F21", 113: "F22",
            114: "F23", 115: "F24",
            70: "PrintScreen", 71: "ScrollLock", 72: "Pause",
            73: "Insert", 74: "Home", 75: "PageUp", 76: "Delete",
            77: "End", 78: "PageDown", 79: "Right", 80: "Left",
            81: "Down", 82: "Up", 83: "NumLock",
            84: "Num /", 85: "Num *", 86: "Num -", 87: "Num +",
            88: "Num Enter", 89: "Num 1", 90: "Num 2", 91: "Num 3",
            92: "Num 4", 93: "Num 5", 94: "Num 6", 95: "Num 7",
            96: "Num 8", 97: "Num 9", 98: "Num 0", 99: "Num .",
            101: "Menu", 168: "Mute", 169: "Volume Up", 170: "Volume Down",
        }
    )
    return mapping


HID_KEY_NAMES: dict[int, str] = _build_hid_name_map()
MOUSE_CODE_NAMES: dict[int, str] = {
    1: "Mouse L", 2: "Mouse R", 4: "Mouse M",
    8: "Mouse Backward", 16: "Mouse Forward",
}
KEYBOARD_NAME_TO_CODE: dict[str, int] = {name: code for code, name in sorted(HID_KEY_NAMES.items())}
KEYBOARD_EVENT_CHOICES: list[str] = list(KEYBOARD_NAME_TO_CODE.keys())
MOUSE_NAME_TO_CODE: dict[str, int] = {name: code for code, name in MOUSE_CODE_NAMES.items()}
MOUSE_EVENT_CHOICES: list[str] = list(MOUSE_NAME_TO_CODE.keys())

# ──────────────────────────────────────────────────────────────
# ボタン名
# ──────────────────────────────────────────────────────────────
BUTTON_SLOT_NAMES: list[str] = [
    "Left Button", "Right Button", "Middle Button",
    "Backward Button", "Forward Button", "DPI Button",
    "Wheel Up", "Wheel Down",
]

# ──────────────────────────────────────────────────────────────
# マクロモード
# ──────────────────────────────────────────────────────────────
MACRO_MODE_OPTIONS: list[tuple[int, str]] = [
    (0, "macro_mode_counted"),
    (2, "macro_mode_vendor2"),
    (3, "macro_mode_vendor3"),
]

# ──────────────────────────────────────────────────────────────
# プリセット
# ──────────────────────────────────────────────────────────────
def _make_binding(name, binding_type, code1, code2, code3=0, lang="") -> dict:
    return {"name": name, "type": binding_type, "code1": code1, "code2": code2, "code3": code3, "lang": lang}


KEYBOARD_PRESETS: list[dict] = [
    _make_binding(name, 16, 0, code, 0) for code, name in sorted(HID_KEY_NAMES.items())
]
MEDIA_PRESETS: list[dict] = [
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
MOUSE_PRESETS: list[dict] = [
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

# ──────────────────────────────────────────────────────────────
# カラーパレット
# ──────────────────────────────────────────────────────────────
PALETTES: dict[str, dict[str, str]] = {
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
        "accent_fg": "#ffffff",
        "dirty": "#422006",
        "dirty_fg": "#fbbf24",
        "border": "#2e3241",
        "trough": "#151821",
        "tab_selected": "#21242f",
        "heading": "#1a1d27",
        "select_bg": "#2dd4bf",
        "select_fg": "#ffffff",
    },
}
