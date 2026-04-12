from copy import deepcopy

from ajazz_mouse import DEFAULT_MOUSE_KEY_BINDINGS, EVENT_TYPE_KEYBOARD, EVENT_TYPE_MOUSE, MacroProfile, MouseKeyBinding


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


def clone_macro_profiles(profiles):
    return [MacroProfile(slot=profile.slot, name=profile.name, trigger_mode=profile.trigger_mode, repeat_count=profile.repeat_count, list=deepcopy(profile.list)) for profile in profiles]
