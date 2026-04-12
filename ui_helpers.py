"""
ui_helpers.py – バインディング操作 / 名前解決などのユーティリティ
"""

from ajazz_mouse import (
    ACTION_PRESS,
    DEFAULT_MOUSE_KEY_BINDINGS,
    EVENT_TYPE_KEYBOARD,
    EVENT_TYPE_MOUSE,
    MacroEvent,
    MacroProfile,
    MouseKeyBinding,
)
from constants import HID_KEY_NAMES, MOUSE_CODE_NAMES, MOUSE_PRESETS, MEDIA_PRESETS


# ──────────────────────────────────────────────────────────────
# バインディングのクローン / デフォルト生成
# ──────────────────────────────────────────────────────────────

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


def copy_default_bindings() -> list[MouseKeyBinding]:
    return [clone_binding(b) for b in DEFAULT_MOUSE_KEY_BINDINGS]


# ──────────────────────────────────────────────────────────────
# 修飾キー名
# ──────────────────────────────────────────────────────────────

def modifier_names(code1: int) -> list[str]:
    names: list[str] = []
    if code1 & 0x01 or code1 & 0x10:
        names.append("CTRL")
    if code1 & 0x08 or code1 & 0x80:
        names.append("WIN")
    if code1 & 0x04 or code1 & 0x40:
        names.append("ALT")
    if code1 & 0x02 or code1 & 0x20:
        names.append("SHIFT")
    return names


# ──────────────────────────────────────────────────────────────
# 名前解決
# ──────────────────────────────────────────────────────────────

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
