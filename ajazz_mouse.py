import hid

class AjazzMouse:
    VENDOR_IDS = [0xA8A4, 0xA8A5]  # 43172, 43173
    PRODUCT_ID = 0x2255            # 8789
    USAGE_PAGE = 0xFF01
    USAGE = 0x10

    def __init__(self, log_callback=None):
        self.device = None
        self.log_callback = log_callback

    def _log(self, prefix, data):
        """16進数にフォーマットしてコールバックに渡す"""
        if self.log_callback:
            hex_str = " ".join([f"{b:02X}" for b in data])
            self.log_callback(f"{prefix}: {hex_str}")

    def connect(self):
        """マウスを検索して接続する"""
        for d in hid.enumerate():
            if d['vendor_id'] in self.VENDOR_IDS and d['product_id'] == self.PRODUCT_ID:
                if d['usage_page'] == self.USAGE_PAGE and d['usage'] == self.USAGE:
                    self.device = hid.device()
                    self.device.open_path(d['path'])
                    self.device.set_nonblocking(False)
                    return True
        return False

    def close(self):
        """デバイスのクローズ"""
        if self.device:
            self.device.close()
            self.device = None

    def _send_command(self, payload):
        """65バイト(Report ID込)のコマンドを送信し、結果を受け取る"""
        if not self.device:
            raise ConnectionError("デバイスが接続されていません")

        buffer = [0] * 65
        for i in range(len(payload)):
            buffer[i] = payload[i]
        
        self._log("SEND", buffer)
        self.device.write(buffer)
        
        resp = self.device.read(65, timeout_ms=1000)
        resp_list = list(resp)
        self._log("RECV", resp_list)
        
        # JSのパース処理とIndexを一致させるため、レスポンス先頭のReport ID(0)を除去して64バイト配列にします
        if len(resp_list) == 65 and resp_list[0] == 0:
            return resp_list[1:]
        elif len(resp_list) == 64:
            return resp_list
            
        return [0] * 64

    def get_version(self):
        """ファームウェアのバージョンを取得"""
        req = [0, 85, 3]
        resp = self._send_command(req)
        if len(resp) > 25:
            v1 = chr(resp[23]) if 48 <= resp[23] <= 57 else "0"
            v2 = chr(resp[24]) if 48 <= resp[24] <= 57 else "0"
            v3 = chr(resp[25]) if 48 <= resp[25] <= 57 else "0"
            return f"{v1}.{v2}.{v3}"
        return "0.0.0"

    def get_battery_info(self):
        """バッテリー残量の取得"""
        req = [0, 85, 48, 165, 11, 46, 1, 1, 0, 0, 0]
        resp = self._send_command(req)
        if len(resp) > 9:
            return {
                "battery": resp[8],
                "is_charging": bool(resp[9])
            }
        return {"battery": 0, "is_charging": False}

    def is_online(self):
        """無線接続でのオンライン・オフライン判定"""
        req = [0, 85, 237, 0, 1, 46, 0, 0]
        resp = self._send_command(req)
        if len(resp) > 8:
            return resp[8] == 2
        return False

    def get_config(self):
        """各種設定値の完全取得"""
        req = [0, 85, 14, 165, 11, 47, 1, 1, 0, 0, 0]
        resp = self._send_command(req)
        if len(resp) > 55:
            config = {
                "light_mode": resp[9],
                "report_rate_idx": max(0, resp[10] - 1),
                "dpi_count": resp[11],
                "dpi_index": max(0, resp[12] - 1),
                "dpis": [
                    resp[13] | (resp[14] << 8),
                    resp[15] | (resp[16] << 8),
                    resp[17] | (resp[18] << 8),
                    resp[19] | (resp[20] << 8),
                    resp[21] | (resp[22] << 8),
                    resp[23] | (resp[24] << 8),
                ],
                "scroll_flag": resp[48],
                "lod_value": resp[49],
                "sensor_flag": resp[50],
                "key_respond": resp[51],
                "sleep_light": resp[52],
                "highspeed_mode": resp[53],
                "wakeup_flag": resp[55]
            }
            return config
        return None

    def set_config(self, config):
        """各種設定値の書き込み"""
        req = [0] * 65
        req[1] = 85
        req[2] = 15
        req[3] = 174
        req[4] = 10
        req[5] = 47
        req[6] = 1
        req[7] = 1
        
        req[10] = config.get("light_mode", 0)
        req[11] = config.get("report_rate_idx", 3) + 1
        req[12] = config.get("dpi_count", 6)
        req[13] = config.get("dpi_index", 0) + 1
        
        dpis = config.get("dpis", [400, 800, 1200, 1600, 2400, 3200])
        req[14] = dpis[0] & 0xFF
        req[15] = dpis[0] >> 8
        req[16] = dpis[1] & 0xFF
        req[17] = dpis[1] >> 8
        req[18] = dpis[2] & 0xFF
        req[19] = dpis[2] >> 8
        req[20] = dpis[3] & 0xFF
        req[21] = dpis[3] >> 8
        req[22] = dpis[4] & 0xFF
        req[23] = dpis[4] >> 8
        req[24] = dpis[5] & 0xFF
        req[25] = dpis[5] >> 8
        
        req[49] = config.get("scroll_flag", 0)
        req[50] = config.get("lod_value", 1)
        req[51] = config.get("sensor_flag", 53)
        req[52] = config.get("key_respond", 4)
        req[53] = config.get("sleep_light", 10)
        req[54] = config.get("highspeed_mode", 0)
        req[55] = config.get("wakeup_flag", 0)
        
        self._send_command(req)
