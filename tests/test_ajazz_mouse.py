import unittest

from ajazz_mouse import (
    ACTION_PRESS,
    ACTION_RELEASE,
    DEFAULT_MOUSE_KEY_BINDINGS,
    EVENT_TYPE_KEYBOARD,
    EVENT_TYPE_MOUSE,
    MacroEvent,
    MacroProfile,
    MouseKeyBinding,
    build_get_mouse_keys_request,
    build_macro_read_matcher,
    build_macro_write_packets,
    build_reset_mouse_keys_request,
    build_set_mouse_keys_request,
    decode_macro_profiles,
    encode_macro_profiles,
)


class AjazzMouseProtocolTests(unittest.TestCase):
    def test_get_mouse_keys_packet_matches_web_protocol(self):
        self.assertEqual(build_get_mouse_keys_request(), [0, 85, 8, 165, 11, 32, 0, 0, 0, 0, 0])

    def test_reset_mouse_keys_packet_uses_default_mapping(self):
        packet = build_reset_mouse_keys_request()
        self.assertEqual(packet[:6], [0, 85, 9, 165, 34, 32])
        expected_tail = [
            32, 1, 0, 0,
            32, 2, 0, 0,
            32, 4, 0, 0,
            32, 8, 0, 0,
            32, 16, 0, 0,
            33, 85, 0, 0,
            33, 56, 1, 0,
            33, 56, 255, 0,
        ]
        self.assertEqual(packet[9:41], expected_tail)

    def test_set_mouse_keys_packet_serializes_custom_binding(self):
        bindings = [MouseKeyBinding(index=b.index, name=b.name, type=b.type, code1=b.code1, code2=b.code2, code3=b.code3) for b in DEFAULT_MOUSE_KEY_BINDINGS]
        bindings[3] = MouseKeyBinding(index=3, name="Macro 2", type=112, code1=1, code2=3, code3=2)
        packet = build_set_mouse_keys_request(bindings)
        self.assertEqual(packet[:6], [0, 85, 9, 165, 34, 49])
        self.assertEqual(packet[21:25], [112, 1, 3, 2])

    def test_macro_round_trip_preserves_profiles(self):
        profiles = [
            MacroProfile(
                slot=0,
                name="Macro 1",
                list=[
                    MacroEvent(name="A", code=4, type=EVENT_TYPE_KEYBOARD, action=ACTION_PRESS, delay=0),
                    MacroEvent(name="A", code=4, type=EVENT_TYPE_KEYBOARD, action=ACTION_RELEASE, delay=25),
                ],
            ),
            MacroProfile(
                slot=4,
                name="Macro 5",
                list=[
                    MacroEvent(name="Mouse L", code=1, type=EVENT_TYPE_MOUSE, action=ACTION_PRESS, delay=10),
                    MacroEvent(name="Mouse L", code=1, type=EVENT_TYPE_MOUSE, action=ACTION_RELEASE, delay=15),
                ],
            ),
        ]
        encoded = encode_macro_profiles(profiles)
        decoded = decode_macro_profiles(encoded, lambda event_type, code, action: f"{event_type}:{code}:{action}")

        self.assertEqual(decoded[0].list[0].code, 4)
        self.assertEqual(decoded[0].list[0].action, ACTION_PRESS)
        self.assertEqual(decoded[0].list[1].delay, 25)
        self.assertEqual(decoded[4].list[0].type, EVENT_TYPE_MOUSE)
        self.assertEqual(decoded[4].list[1].action, ACTION_RELEASE)
        self.assertEqual(decoded[4].list[1].delay, 15)

    def test_macro_packets_split_into_56_byte_chunks_and_commit(self):
        payload = list(range(120))
        packets = build_macro_write_packets(payload)
        self.assertEqual(len(packets), 4)
        self.assertEqual(packets[0][:9], [0, 85, 13, 0, 0, 56, 0, 0, 0])
        self.assertEqual(packets[1][:9], [0, 85, 13, 0, 0, 56, 56, 0, 0])
        self.assertEqual(packets[2][:9], [0, 85, 13, 0, 0, 8, 112, 0, 0])
        self.assertEqual(packets[3], [0, 85, 16, 165, 34, 0, 0, 0, 5])

    def test_macro_read_matcher_accepts_vendor_response_header(self):
        matcher = build_macro_read_matcher(56)
        self.assertTrue(matcher([0xAA, 0xFA, 0xA5, 0x2E, 0x38, 0x01, 0x01, 0x00, 0xD0, 0x24]))
        self.assertFalse(matcher([0xAA, 0x30, 0xA5, 0x0B, 0x0A, 0x01, 0x01, 0x00]))

    def test_get_macro_data_reports_chunk_progress(self):
        from ajazz_mouse import AjazzMouse, MACRO_CHUNK_SIZE, MACRO_DATA_SIZE

        mouse = AjazzMouse()
        mouse.device = object()
        progress_updates = []

        def fake_send_command(payload, matcher=None, timeout=1.0, log_timeout=True):
            chunk_length = payload[2]
            response = [0] * 64
            response[8 : 8 + chunk_length] = [1] * chunk_length
            return response

        mouse._send_command = fake_send_command

        data = mouse.get_macro_data(progress_callback=lambda current, total: progress_updates.append((current, total)))

        total_chunks = (MACRO_DATA_SIZE + MACRO_CHUNK_SIZE - 1) // MACRO_CHUNK_SIZE
        self.assertEqual(len(data), MACRO_DATA_SIZE)
        self.assertEqual(progress_updates[0], (0, total_chunks))
        self.assertEqual(progress_updates[-1], (total_chunks, total_chunks))
        self.assertEqual(len(progress_updates), total_chunks + 1)


if __name__ == "__main__":
    unittest.main()
