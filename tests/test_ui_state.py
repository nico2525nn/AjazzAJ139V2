import unittest

from ajazz_mouse import ACTION_PRESS, EVENT_TYPE_KEYBOARD, MacroEvent, MacroProfile, MouseKeyBinding, encode_macro_profiles
import ui_app


class _FakeMouse:
    def __init__(self, mouse_keys=None, macro_data=None):
        self._mouse_keys = mouse_keys or []
        self._macro_data = macro_data or [0] * 4096

    def get_mouse_keys(self):
        return self._mouse_keys

    def get_macro_data(self):
        return self._macro_data


class _FakeApp:
    def __init__(self):
        self._macro_metadata = {"names": {}}
        self.mouse_keys = []
        self.macro_profiles = [MacroProfile(slot=index, name=f"Macro {index + 1}") for index in range(32)]
        self.mouse = _FakeMouse()
        self.logs = []
        self.key_refreshes = 0
        self.macro_profile_refreshes = 0
        self.macro_event_refreshes = 0

    def _append_log(self, text):
        self.logs.append(text)

    def _refresh_key_mapping_ui(self):
        self.key_refreshes += 1

    def _refresh_macro_profile_list(self):
        self.macro_profile_refreshes += 1

    def _refresh_macro_events(self):
        self.macro_event_refreshes += 1

    def _merge_macro_names(self, profiles):
        return ui_app.AjazzApp._merge_macro_names(self, profiles)

    def _set_macro_status(self, state_key):
        self.macro_status = state_key


class UiStateTests(unittest.TestCase):
    def test_keyboard_choices_include_extended_function_keys(self):
        self.assertEqual(ui_app.KEYBOARD_NAME_TO_CODE["F13"], 104)
        self.assertEqual(ui_app.KEYBOARD_NAME_TO_CODE["F24"], 115)
        self.assertIn("F24", ui_app.KEYBOARD_EVENT_CHOICES)

    def test_merge_macro_names_prefers_local_display_name(self):
        app = _FakeApp()
        app._macro_metadata = {"names": {"1": "Burst Fire"}}
        source_profiles = [
            MacroProfile(slot=0, name="Macro 1"),
            MacroProfile(slot=1, name="Macro 2", list=[MacroEvent(name="A", code=4, type=EVENT_TYPE_KEYBOARD, action=ACTION_PRESS, delay=0)]),
        ]

        merged = ui_app.AjazzApp._merge_macro_names(app, source_profiles)

        self.assertEqual(merged[0].name, "Macro 1")
        self.assertEqual(merged[1].name, "Burst Fire")
        merged[1].list.append(MacroEvent(name="B", code=5, type=EVENT_TYPE_KEYBOARD, action=ACTION_PRESS, delay=0))
        self.assertEqual(len(source_profiles[1].list), 1)

    def test_load_key_mapping_uses_macro_profile_names(self):
        app = _FakeApp()
        app.macro_profiles[1].name = "Rapid Tap"
        app.mouse = _FakeMouse(
            mouse_keys=[
                MouseKeyBinding(index=0, name="", type=112, code1=1, code2=3, code3=2),
                MouseKeyBinding(index=1, name="", type=16, code1=0x05, code2=4, code3=0),
            ]
        )

        ui_app.AjazzApp._load_key_mapping_from_device(app, log=False)

        self.assertEqual(app.mouse_keys[0].name, "Rapid Tap")
        self.assertEqual(app.mouse_keys[1].name, "CTRL+ALT+A")
        self.assertEqual(app.key_refreshes, 1)

    def test_load_macro_profiles_from_device_merges_saved_names(self):
        app = _FakeApp()
        app._macro_metadata = {"names": {"0": "Saved Macro"}}
        device_profiles = [MacroProfile(slot=index, name=f"Macro {index + 1}") for index in range(32)]
        device_profiles[0].list = [
            MacroEvent(name="A", code=4, type=EVENT_TYPE_KEYBOARD, action=ACTION_PRESS, delay=0),
        ]
        app.mouse = _FakeMouse(macro_data=encode_macro_profiles(device_profiles))

        ui_app.AjazzApp._load_macro_profiles_from_device(app, log=False)

        self.assertEqual(app.macro_profiles[0].name, "Saved Macro")
        self.assertEqual(len(app.macro_profiles[0].list), 1)
        self.assertEqual(app.macro_profile_refreshes, 1)
        self.assertEqual(app.macro_event_refreshes, 1)


if __name__ == "__main__":
    unittest.main()
