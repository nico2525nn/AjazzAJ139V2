import threading
import time
from dataclasses import dataclass, field
from typing import Callable, Iterable, Optional, Sequence

import hid


MAX_MACRO_PROFILES = 32
MACRO_DATA_SIZE = 4096
MACRO_HEADER_SIZE = 64
MACRO_CHUNK_SIZE = 56

EVENT_TYPE_MOUSE = 1
EVENT_TYPE_KEYBOARD = 2
EVENT_TYPE_OTHER = 3

ACTION_PRESS = 1
ACTION_RELEASE = 2


@dataclass
class MouseKeyBinding:
    index: int
    name: str
    type: int
    code1: int
    code2: int
    code3: int
    lang: str = ""


@dataclass
class MacroEvent:
    name: str
    code: int
    type: int
    action: int
    delay: int


@dataclass
class MacroProfile:
    slot: int
    name: str
    trigger_mode: int = 0
    repeat_count: int = 1
    list: list[MacroEvent] = field(default_factory=list)


DEFAULT_MOUSE_KEY_BINDINGS = [
    MouseKeyBinding(0, "Left Click", 32, 1, 0, 0),
    MouseKeyBinding(1, "Right Click", 32, 2, 0, 0),
    MouseKeyBinding(2, "Middle Click", 32, 4, 0, 0),
    MouseKeyBinding(3, "Backward", 32, 8, 0, 0),
    MouseKeyBinding(4, "Forward", 32, 16, 0, 0),
    MouseKeyBinding(5, "DPI Loop +", 33, 85, 0, 0),
    MouseKeyBinding(6, "Scroll Up", 33, 56, 1, 0),
    MouseKeyBinding(7, "Scroll Down", 33, 56, 255, 0),
]


def _little_endian(value: int) -> list[int]:
    value = int(value) & 0xFFFF
    return [value & 0xFF, (value >> 8) & 0xFF]


def _coerce_binding(index: int, binding: MouseKeyBinding | dict) -> MouseKeyBinding:
    if isinstance(binding, MouseKeyBinding):
        return MouseKeyBinding(
            index=index,
            name=binding.name,
            type=int(binding.type),
            code1=int(binding.code1),
            code2=int(binding.code2),
            code3=int(binding.code3),
            lang=binding.lang,
        )
    return MouseKeyBinding(
        index=index,
        name=str(binding.get("name", f"Button {index + 1}")),
        type=int(binding.get("type", 0)),
        code1=int(binding.get("code1", 0)),
        code2=int(binding.get("code2", 0)),
        code3=int(binding.get("code3", 0)),
        lang=str(binding.get("lang", "")),
    )


def _coerce_event(event: MacroEvent | dict) -> MacroEvent:
    if isinstance(event, MacroEvent):
        return MacroEvent(
            name=event.name,
            code=int(event.code),
            type=int(event.type),
            action=int(event.action),
            delay=max(0, int(event.delay)),
        )
    return MacroEvent(
        name=str(event.get("name", "")),
        code=int(event.get("code", 0)),
        type=int(event.get("type", EVENT_TYPE_KEYBOARD)),
        action=int(event.get("action", ACTION_PRESS)),
        delay=max(0, int(event.get("delay", 0))),
    )


def build_get_mouse_keys_request() -> list[int]:
    return [0, 85, 8, 165, 11, 32, 0, 0, 0, 0, 0]


def build_set_mouse_keys_request(bindings: Sequence[MouseKeyBinding | dict]) -> list[int]:
    packet = [0] * 64
    packet[0] = 0
    packet[1] = 85
    packet[2] = 9
    packet[3] = 165
    packet[4] = 34
    packet[5] = 49

    for index, binding in enumerate(bindings[:8]):
        normalized = _coerce_binding(index, binding)
        base = 9 + (index * 4)
        packet[base] = normalized.type & 0xFF
        packet[base + 1] = normalized.code1 & 0xFF
        packet[base + 2] = normalized.code2 & 0xFF
        packet[base + 3] = normalized.code3 & 0xFF

    return packet


def build_reset_mouse_keys_request() -> list[int]:
    packet = build_set_mouse_keys_request(DEFAULT_MOUSE_KEY_BINDINGS)
    packet[5] = 32
    return packet


def build_macro_read_request(length: int, offset: int) -> list[int]:
    return [6, 12, int(length) & 0xFF, *_little_endian(offset)]


def build_macro_read_matcher(length: int) -> Callable[[Sequence[int]], bool]:
    expected_length = int(length) & 0xFF

    def matcher(response: Sequence[int]) -> bool:
        return (
            len(response) >= 8
            and response[0] == 0xAA
            and response[1] in (0xFA, 0x0C)
            and response[4] == expected_length
        )

    return matcher


def build_macro_reset_request() -> list[int]:
    return [6, 15, 4]


def build_macro_write_packets(data: Sequence[int]) -> list[list[int]]:
    payload = [int(value) & 0xFF for value in data]
    packets: list[list[int]] = []

    for offset in range(0, len(payload), MACRO_CHUNK_SIZE):
        chunk = payload[offset : offset + MACRO_CHUNK_SIZE]
        packets.append(
            [0, 85, 13, 0, 0, len(chunk), *_little_endian(offset), 0, *chunk]
        )

    packets.append([0, 85, 16, 165, 34, 0, 0, 0, 5])
    return packets


def encode_macro_profiles(profiles: Sequence[MacroProfile | dict]) -> list[int]:
    data = [0] * MACRO_DATA_SIZE
    cursor = MACRO_HEADER_SIZE
    slots: dict[int, MacroProfile] = {}

    for item in profiles:
        if isinstance(item, MacroProfile):
            slot = int(item.slot)
            profile = MacroProfile(
                slot=slot,
                name=item.name,
                trigger_mode=int(item.trigger_mode),
                repeat_count=int(item.repeat_count),
                list=[_coerce_event(event) for event in item.list],
            )
        else:
            slot = int(item.get("slot", 0))
            profile = MacroProfile(
                slot=slot,
                name=str(item.get("name", f"Macro {slot + 1}")),
                trigger_mode=int(item.get("trigger_mode", 0)),
                repeat_count=int(item.get("repeat_count", 1)),
                list=[_coerce_event(event) for event in item.get("list", [])],
            )
        if 0 <= slot < MAX_MACRO_PROFILES:
            slots[slot] = profile

    for slot in range(MAX_MACRO_PROFILES):
        profile = slots.get(slot)
        if not profile or not profile.list:
            continue

        data[slot * 2 : slot * 2 + 2] = _little_endian(cursor)

        for event_index, event in enumerate(profile.list):
            if cursor + 4 > MACRO_DATA_SIZE:
                raise ValueError("Macro data exceeds on-device storage.")

            flags = 0
            if event.action == ACTION_PRESS:
                flags |= 0x40
            if event.type == EVENT_TYPE_MOUSE:
                flags |= 0x03
            elif event.type == EVENT_TYPE_KEYBOARD:
                flags |= 0x02
            elif event.type == EVENT_TYPE_OTHER:
                flags |= 0x04
            if event_index == len(profile.list) - 1:
                flags |= 0x80

            delay_low, delay_high = _little_endian(event.delay)
            data[cursor] = delay_low
            data[cursor + 1] = delay_high
            data[cursor + 2] = flags
            data[cursor + 3] = int(event.code) & 0xFF
            cursor += 4

    return data[:cursor]


def decode_macro_profiles(
    data: Sequence[int],
    name_resolver: Optional[Callable[[int, int, int], str]] = None,
) -> list[MacroProfile]:
    payload = [int(value) & 0xFF for value in data[:MACRO_DATA_SIZE]]
    profiles = [
        MacroProfile(slot=index, name=f"Macro {index + 1}") for index in range(MAX_MACRO_PROFILES)
    ]

    offsets: list[tuple[int, int]] = []
    for slot in range(MAX_MACRO_PROFILES):
        offset = payload[slot * 2] | (payload[slot * 2 + 1] << 8)
        if MACRO_HEADER_SIZE <= offset < len(payload):
            offsets.append((slot, offset))

    offsets.sort(key=lambda item: item[1])
    for position, (slot, offset) in enumerate(offsets):
        limit = len(payload)
        if position + 1 < len(offsets):
            limit = offsets[position + 1][1]

        cursor = offset
        events: list[MacroEvent] = []
        while cursor + 3 < min(limit, len(payload)):
            delay = payload[cursor] | (payload[cursor + 1] << 8)
            flags = payload[cursor + 2]
            code = payload[cursor + 3]
            kind = flags & 0x07

            if kind == 0x03:
                event_type = EVENT_TYPE_MOUSE
            elif kind == 0x02:
                event_type = EVENT_TYPE_KEYBOARD
            elif kind == 0x04:
                event_type = EVENT_TYPE_OTHER
            else:
                break

            action = ACTION_PRESS if flags & 0x40 else ACTION_RELEASE
            name = (
                name_resolver(event_type, code, action)
                if name_resolver
                else f"Event {code}"
            )
            events.append(
                MacroEvent(
                    name=name,
                    code=code,
                    type=event_type,
                    action=action,
                    delay=delay,
                )
            )
            cursor += 4

            if flags & 0x80:
                break

        profiles[slot].list = events

    return profiles


class AjazzMouse:
    VENDOR_IDS = [0xA8A4, 0xA8A5]
    PRODUCT_ID = 0x2255
    USAGE_PAGE = 0xFF01
    USAGE = 0x10

    def __init__(self, log_callback: Optional[Callable[[str], None]] = None):
        self.device = None
        self.log_callback = log_callback
        self.lock = threading.RLock()

    def _log(self, prefix: str, data):
        if not self.log_callback:
            return
        if isinstance(data, (list, tuple, bytes, bytearray)):
            hex_str = " ".join(f"{value:02X}" for value in data)
            self.log_callback(f"{prefix}: {hex_str}")
        else:
            self.log_callback(f"{prefix}: {data}")

    def connect(self) -> bool:
        with self.lock:
            for device_info in hid.enumerate():
                if (
                    device_info["vendor_id"] in self.VENDOR_IDS
                    and device_info["product_id"] == self.PRODUCT_ID
                    and device_info["usage_page"] == self.USAGE_PAGE
                    and device_info["usage"] == self.USAGE
                ):
                    self.device = hid.device()
                    self.device.open_path(device_info["path"])
                    self.device.set_nonblocking(False)
                    return True
            return False

    def close(self):
        with self.lock:
            if self.device:
                self.device.close()
                self.device = None

    def _flush_input_queue(self):
        self.device.set_nonblocking(True)
        try:
            while self.device.read(65):
                pass
        finally:
            self.device.set_nonblocking(False)

    def _normalize_response(self, payload: Sequence[int]) -> list[int]:
        raw = list(payload)
        if len(raw) == 65 and raw[0] == 0:
            return raw[1:]
        return raw

    def _default_matcher(self, payload: Sequence[int]) -> Callable[[Sequence[int]], bool]:
        expected_sub_id = payload[2]

        def matcher(response: Sequence[int]) -> bool:
            return len(response) >= 2 and response[0] == 0xAA and response[1] == expected_sub_id

        return matcher

    def _send_command(
        self,
        payload: Sequence[int],
        matcher: Optional[Callable[[Sequence[int]], bool]] = None,
        timeout: float = 1.0,
        log_timeout: bool = True,
    ) -> list[int]:
        with self.lock:
            if not self.device:
                raise ConnectionError("Device is not connected.")

            self._flush_input_queue()

            buffer = [0] * 65
            for index, value in enumerate(payload[:65]):
                buffer[index] = int(value) & 0xFF

            self._log("SEND", buffer)
            self.device.write(buffer)

            response_matcher = matcher or self._default_matcher(payload)
            start_time = time.time()

            while time.time() - start_time < timeout:
                raw = self.device.read(65, timeout_ms=200)
                if not raw:
                    continue

                parsed = self._normalize_response(raw)
                if response_matcher(parsed):
                    self._log("RECV (MATCHED)", raw)
                    return parsed
                self._log("RECV (IGNORED)", raw)

            if log_timeout:
                self._log("ERROR", "Timed out waiting for response.")
            return [0] * 64

    def get_version(self) -> str:
        response = self._send_command([0, 85, 3])
        if len(response) > 25 and response[0] == 0xAA:
            values = []
            for index in (23, 24, 25):
                values.append(chr(response[index]) if 48 <= response[index] <= 57 else "0")
            return ".".join(values)
        return "0.0.0"

    def get_battery_info(self) -> dict:
        response = self._send_command([0, 85, 48, 165, 11, 46, 1, 1, 0, 0, 0])
        if len(response) > 9 and response[0] == 0xAA:
            return {"battery": response[8], "is_charging": bool(response[9])}
        return {"battery": 0, "is_charging": False}

    def is_online(self) -> bool:
        response = self._send_command([0, 85, 237, 0, 1, 46, 0, 0])
        return len(response) > 8 and response[0] == 0xAA and response[8] == 2

    def get_config(self) -> Optional[dict]:
        response = self._send_command([0, 85, 14, 165, 11, 47, 1, 1, 0, 0, 0])
        if len(response) <= 55 or response[0] != 0xAA:
            return None

        return {
            "light_mode": response[9],
            "report_rate_idx": max(0, response[10] - 1),
            "dpi_count": response[11],
            "dpi_index": max(0, response[12] - 1),
            "dpis": [
                response[13] | (response[14] << 8),
                response[15] | (response[16] << 8),
                response[17] | (response[18] << 8),
                response[19] | (response[20] << 8),
                response[21] | (response[22] << 8),
                response[23] | (response[24] << 8),
            ],
            "scroll_flag": response[48],
            "lod_value": response[49],
            "sensor_flag": response[50],
            "key_respond": response[51],
            "sleep_light": response[52],
            "highspeed_mode": response[53],
            "wakeup_flag": response[55],
        }

    def set_config(self, config: dict):
        request = [0] * 65
        request[1] = 85
        request[2] = 15
        request[3] = 174
        request[4] = 10
        request[5] = 47
        request[6] = 1
        request[7] = 1

        request[10] = int(config.get("light_mode", 0))
        request[11] = int(config.get("report_rate_idx", 3)) + 1
        request[12] = int(config.get("dpi_count", 6))
        request[13] = int(config.get("dpi_index", 0)) + 1

        dpis = list(config.get("dpis", [400, 800, 1200, 1600, 2400, 3200]))
        for index in range(6):
            value = int(dpis[index]) if index < len(dpis) else 400
            request[14 + index * 2] = value & 0xFF
            request[15 + index * 2] = (value >> 8) & 0xFF

        request[49] = int(config.get("scroll_flag", 0))
        request[50] = int(config.get("lod_value", 1))
        request[51] = int(config.get("sensor_flag", 53))
        request[52] = int(config.get("key_respond", 4))
        request[53] = int(config.get("sleep_light", 10))
        request[54] = int(config.get("highspeed_mode", 0))
        request[55] = int(config.get("wakeup_flag", 0))

        self._send_command(request)

    def get_mouse_keys(self) -> list[MouseKeyBinding]:
        response = self._send_command(build_get_mouse_keys_request())
        if len(response) < 40:
            return [
                MouseKeyBinding(
                    index=binding.index,
                    name=binding.name,
                    type=binding.type,
                    code1=binding.code1,
                    code2=binding.code2,
                    code3=binding.code3,
                    lang=binding.lang,
                )
                for binding in DEFAULT_MOUSE_KEY_BINDINGS
            ]

        bindings = []
        payload = response[8:]
        for index in range(8):
            base = index * 4
            bindings.append(
                MouseKeyBinding(
                    index=index,
                    name=f"Button {index + 1}",
                    type=payload[base],
                    code1=payload[base + 1],
                    code2=payload[base + 2],
                    code3=payload[base + 3],
                )
            )
        return bindings

    def set_mouse_keys(self, bindings: Sequence[MouseKeyBinding | dict]):
        self._send_command(build_set_mouse_keys_request(bindings))

    def reset_mouse_keys(self):
        self._send_command(build_reset_mouse_keys_request())

    def get_macro_data(self, progress_callback: Optional[Callable[[int, int], None]] = None) -> list[int]:
        data: list[int] = []
        total_chunks = (MACRO_DATA_SIZE + MACRO_CHUNK_SIZE - 1) // MACRO_CHUNK_SIZE
        if progress_callback:
            progress_callback(0, total_chunks)
        for chunk_index, offset in enumerate(range(0, MACRO_DATA_SIZE, MACRO_CHUNK_SIZE), start=1):
            chunk_length = min(MACRO_CHUNK_SIZE, MACRO_DATA_SIZE - offset)
            response = self._send_command(
                build_macro_read_request(chunk_length, offset),
                matcher=build_macro_read_matcher(chunk_length),
                timeout=0.08,
                log_timeout=False,
            )
            if response and any(response):
                chunk = response[8 : 8 + chunk_length]
                if len(chunk) < chunk_length:
                    chunk = [*chunk, *([0] * (chunk_length - len(chunk)))]
            else:
                chunk = [0] * chunk_length
            data.extend(chunk)
            if progress_callback:
                progress_callback(chunk_index, total_chunks)
        return data[:MACRO_DATA_SIZE]

    def set_macro_data(self, data: Sequence[int]):
        for packet in build_macro_write_packets(data):
            self._send_command(packet)

    def reset_macro_data(self):
        self._send_command(
            build_macro_reset_request(),
            matcher=lambda packet: bool(packet),
        )
