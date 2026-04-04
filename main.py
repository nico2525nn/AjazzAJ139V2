import os
import sys
import json
import threading
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

import pystray
from PIL import Image, ImageDraw, ImageFont

from ajazz_mouse import AjazzMouse

# --- 多言語辞書 (i18n) ---
LANG_DICT = {
    "en": {
        "title": "Ajazz AJ139 V2 Control - Extended",
        "tab_status": "Status",
        "tab_perf": "Performance",
        "tab_sys": "System / Lighting",
        "tab_debug": "Debug Logs",
        "btn_write": "Write Settings",
        "btn_refresh": "Refresh Status",
        "lang_label": "Language / 言語:",
        
        "status_frame": "USB / Dongle Status",
        "status_val": "Status: {0}",
        "fw_val": "Firmware: {0}",
        "bat_val": "Battery: {0}",
        
        "st_disconnected": "Disconnected (Device not found)",
        "st_connected": "Connected (USB/Dongle)",
        "bat_offline": "Offline (Sleeping or disconnected)",
        "bat_charging": "(Charging)",
        
        "poll_rate": "Polling Rate (Hz):",
        "debounce": "Debounce Time (ms):",
        "lod": "LOD (Lift Off Distance):",
        "dpi_frame": "DPI Levels (1-6)",
        "dpi_color": "DPI Color {0}:",
        "dpi_active": "Currently Active DPI Index (1-6):",
        
        "light_mode": "Lighting Mode (0-7):",
        "sleep_time": "LED Auto Sleep (min):",
        
        "log_sys_search": "System: Searching for USB device...",
        "log_sys_conn": "System: Connected successfully. \n",
        "log_sys_fail": "System: Connection failed. Is it plugged in?",
        "log_sys_bat": "System: Online detected. Battery={0}%",
        "log_sys_off": "System: is_online = False. Check the RECV logs.",
        "log_sys_writing": "System: ==== Writing settings... ====",
        
        "msg_success": "Settings written successfully!",
        "msg_err_conn": "Error: Device not connected or config not loaded.",
        
        "tray_open": "Open Settings",
        "tray_exit": "Exit",
    },
    "ja": {
        "title": "Ajazz AJ139 V2 コントロール - 拡張版",
        "tab_status": "基本ステータス",
        "tab_perf": "パフォーマンス設定",
        "tab_sys": "ライティング / その他",
        "tab_debug": "通信デバッグログ",
        "btn_write": "設定をマウスに書き込む",
        "btn_refresh": "最新状態を読み直す",
        "lang_label": "Language / 言語:",
        
        "status_frame": "USB / ドングルステータス",
        "status_val": "状態: {0}",
        "fw_val": "ファームウェア: {0}",
        "bat_val": "バッテリー: {0}",
        
        "st_disconnected": "未接続 (デバイスが見つかりません)",
        "st_connected": "接続中 (USB連携)",
        "bat_offline": "オフライン (スリープ中など)",
        "bat_charging": "(充電中)",
        
        "poll_rate": "レポートレート (Hz):",
        "debounce": "デバウンスタイム (ms):",
        "lod": "LOD (Lift Off Distance):",
        "dpi_frame": "DPI レベル (DPI 1 ～ 6 の値)",
        "dpi_color": "DPI 色 {0}:",
        "dpi_active": "現在アクティブなDPIインデックス (1-6):",
        
        "light_mode": "ライティングモード (0~7):",
        "sleep_time": "LEDオートスリープ (分):",
        
        "log_sys_search": "System: USBデバイスを検索中...",
        "log_sys_conn": "System: デバイスと接続しました。\n",
        "log_sys_fail": "System: デバイスが見つかりません。USBに刺さっていますか？",
        "log_sys_bat": "System: オンライン判定。バッテリー={0}%",
        "log_sys_off": "System: is_online=False. 通信ログを確認してください。",
        "log_sys_writing": "System: ==== 設定を書き込みます ====",
        
        "msg_success": "設定を書き込みました！",
        "msg_err_conn": "エラー: デバイスと通信できていないか、設定データがありません。",
        
        "tray_open": "設定画面を開く / Open Settings",
        "tray_exit": "完全終了 (常駐を解除) / Exit",
    }
}

class AjazzApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.mouse = AjazzMouse(log_callback=self._append_log)
        self.current_config = None
        self.lang = "en"
        
        # 設定ファイルの読み込み
        self.config_file = "settings.json"
        self._load_app_settings()

        self.title(self._t("title"))
        self.geometry("650x650")
        
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        # トレイアイコン関連
        self.tray_icon = None
        self._last_tray_battery = -1
        
        # クロスボタン(X)の振る舞いを変える（常駐）
        self.protocol('WM_DELETE_WINDOW', self._hide_window)
        
        # UI構築と初期化
        self._build_ui()
        self.change_language(self.lang)
        self._refresh_status()

        # バッテリー監視用バックグラウンドスレッド (daemon=True にするとメインスレッド終了時に死ぬ)
        self._bg_thread = threading.Thread(target=self._tray_monitor_loop, daemon=True)
        self._bg_thread.start()

    def _t(self, key, *args):
        text = LANG_DICT[self.lang].get(key, key)
        if args:
            return text.format(*args)
        return text

    def _load_app_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    cfg = json.load(f)
                    self.lang = cfg.get("lang", "en")
            except:
                pass

    def _save_app_settings(self):
        try:
            with open(self.config_file, "w") as f:
                json.dump({"lang": self.lang}, f)
        except:
            pass

    def change_language(self, new_lang, skip_ui_update=False):
        self.lang = new_lang
        self.lang_var.set(new_lang)
        self._save_app_settings()
        
        # UIの全テキストを更新
        self.title(self._t("title"))
        self.notebook.tab(self.tab_status, text=self._t("tab_status"))
        self.notebook.tab(self.tab_perf, text=self._t("tab_perf"))
        self.notebook.tab(self.tab_sys, text=self._t("tab_sys"))
        self.notebook.tab(self.tab_debug, text=self._t("tab_debug"))
        
        self.lbl_lang_label.config(text=self._t("lang_label"))
        self.btn_write.config(text=self._t("btn_write"))
        self.btn_refresh.config(text=self._t("btn_refresh"))
        
        self.status_frame.config(text=self._t("status_frame"))
        self.lbl_poll.config(text=self._t("poll_rate"))
        self.lbl_debounce.config(text=self._t("debounce"))
        self.lbl_lod.config(text=self._t("lod"))
        
        self.dpi_frame.config(text=self._t("dpi_frame"))
        for i in range(6):
            self.dpi_lbls[i].config(text=self._t("dpi_color", i+1))
        self.lbl_dpi_idx.config(text=self._t("dpi_active"))
        
        self.lbl_light.config(text=self._t("light_mode"))
        self.lbl_sleep.config(text=self._t("sleep_time"))

        if self.tray_icon:
            self.tray_icon.menu = self._create_tray_menu()
            
        if not skip_ui_update:
            self._refresh_status()

    def _hide_window(self):
        """ウィンドウを隠してトレイ常駐"""
        self.withdraw()
        
    def _show_window(self):
        """ウィンドウを再表示"""
        self.deiconify()
        self.lift()
        self.focus_force()

    def _exit_app(self, icon, item):
        """完全終了"""
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)

    # ------------------
    # System Tray Logic
    # ------------------
    def _create_tray_menu(self):
        return pystray.Menu(
            pystray.MenuItem(self._t("tray_open"), lambda: self.after(0, self._show_window), default=True),
            pystray.MenuItem(self._t("tray_exit"), self._exit_app)
        )

    def _create_image(self, battery_perc):
        """Pillowを使って数字アイコンを動的に生成"""
        # トレイ用の64x64画像
        width, height = 64, 64
        # バッテリー残量に応じて色を変える
        if battery_perc < 0:
            bg_color, fg_color = "gray", "white"
        elif battery_perc <= 20: # 20%以下で赤
            bg_color, fg_color = "black", "red"
        else:
            bg_color, fg_color = "black", "white"
            
        image = Image.new('RGBA', (width, height), (0, 0, 0, 0)) # 透明背景
        dc = ImageDraw.Draw(image)
        
        dc.ellipse([2, 2, width-2, height-2], fill=bg_color, outline="white")
        text = "--" if battery_perc < 0 else str(battery_perc)
        
        try:
            # Arialかシステムの一般的なフォントを使用
            font = ImageFont.truetype("arial.ttf", 36)
        except:
            font = ImageFont.load_default()

        # テキストを中央ぞろえする
        try:
            left, top, right, bottom = dc.textbbox((0, 0), text, font=font)
        except AttributeError:
            w, h = dc.textsize(text, font=font)
            left, top, right, bottom = 0, 0, w, h
            
        tw = right - left
        th = bottom - top
        x = (width - tw) // 2
        y = (height - th) // 2 - 2

        dc.text((x, y), text, fill=fg_color, font=font)
        return image

    def _tray_monitor_loop(self):
        img = self._create_image(-1)
        self.tray_icon = pystray.Icon("AjazzAJ139", img, "AJ139 V2", self._create_tray_menu())
        
        def setup(icon):
            icon.visible = True
            while icon.visible:
                try:
                    if self.mouse.device and self.mouse.is_online():
                        binfo = self.mouse.get_battery_info()
                        bat = binfo["battery"]
                        if bat != self._last_tray_battery:
                            self._last_tray_battery = bat
                            icon.icon = self._create_image(bat)
                    else:
                        if self._last_tray_battery != -1:
                            self._last_tray_battery = -1
                            icon.icon = self._create_image(-1)
                except Exception:
                    pass

                # 10秒ごとにバッテリー確認 (※GUI側操作に影響を与えないよう控えめに)
                time.sleep(10)

        # ブロックして実行
        self.tray_icon.run(setup)

    # ------------------
    # UI Logic
    # ------------------
    def _append_log(self, text):
        if hasattr(self, 'log_txt') and self.log_txt.winfo_exists():
            self.log_txt.configure(state='normal')
            self.log_txt.insert(tk.END, text + "\n")
            self.log_txt.see(tk.END)
            self.log_txt.configure(state='disabled')

    def _build_ui(self):
        # 共通上部ヘッダー部：言語切り替え
        hdr = ttk.Frame(self)
        hdr.pack(fill=tk.X, padx=10, pady=(5,0))
        self.lbl_lang_label = ttk.Label(hdr, text=self._t("lang_label"))
        self.lbl_lang_label.pack(side=tk.LEFT)
        
        self.lang_var = tk.StringVar(value=self.lang)
        lang_cb = ttk.Combobox(hdr, textvariable=self.lang_var, values=["en", "ja"], state="readonly", width=5)
        lang_cb.pack(side=tk.LEFT, padx=5)
        lang_cb.bind("<<ComboboxSelected>>", lambda e: self.change_language(self.lang_var.get(), skip_ui_update=False))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1
        self.tab_status = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_status, text=self._t("tab_status"))
        self._build_status_tab()
        
        # Tab 2
        self.tab_perf = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_perf, text=self._t("tab_perf"))
        self._build_perf_tab()
        
        # Tab 3
        self.tab_sys = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_sys, text=self._t("tab_sys"))
        self._build_sys_tab()
        
        # Tab 4
        self.tab_debug = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_debug, text=self._t("tab_debug"))
        self._build_debug_tab()
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        self.btn_write = ttk.Button(btn_frame, text=self._t("btn_write"), command=self._apply_config)
        self.btn_write.pack(side=tk.RIGHT)
        
        self.btn_refresh = ttk.Button(btn_frame, text=self._t("btn_refresh"), command=self._refresh_status)
        self.btn_refresh.pack(side=tk.RIGHT, padx=10)

    def _build_status_tab(self):
        self.status_var = tk.StringVar()
        self.battery_var = tk.StringVar()
        self.version_var = tk.StringVar()
        
        self.status_frame = ttk.LabelFrame(self.tab_status, text=self._t("status_frame"), padding=10)
        self.status_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(self.status_frame, textvariable=self.status_var, font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
        ttk.Label(self.status_frame, textvariable=self.version_var).pack(anchor=tk.W)
        ttk.Label(self.status_frame, textvariable=self.battery_var).pack(anchor=tk.W)

    def _build_perf_tab(self):
        row1 = ttk.Frame(self.tab_perf)
        row1.pack(fill=tk.X, pady=5)
        self.lbl_poll = ttk.Label(row1, width=25)
        self.lbl_poll.pack(side=tk.LEFT)
        self.poll_var = tk.StringVar()
        ttk.Combobox(row1, textvariable=self.poll_var, values=["125", "250", "500", "1000"], state="readonly").pack(side=tk.LEFT)
        
        row2 = ttk.Frame(self.tab_perf)
        row2.pack(fill=tk.X, pady=5)
        self.lbl_debounce = ttk.Label(row2, width=25)
        self.lbl_debounce.pack(side=tk.LEFT)
        self.debounce_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.debounce_var).pack(side=tk.LEFT)

        row3 = ttk.Frame(self.tab_perf)
        row3.pack(fill=tk.X, pady=5)
        self.lbl_lod = ttk.Label(row3, width=25)
        self.lbl_lod.pack(side=tk.LEFT)
        self.lod_var = tk.StringVar()
        ttk.Combobox(row3, textvariable=self.lod_var, values=["1mm", "2mm"], state="readonly").pack(side=tk.LEFT)
        
        self.dpi_frame = ttk.LabelFrame(self.tab_perf, padding=10)
        self.dpi_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.dpi_lbls = []
        self.dpi_vars = []
        for i in range(6):
            f = ttk.Frame(self.dpi_frame)
            f.pack(fill=tk.X, pady=2)
            lbl = ttk.Label(f, width=15)
            lbl.pack(side=tk.LEFT)
            self.dpi_lbls.append(lbl)
            
            var = tk.StringVar(value="400")
            ttk.Entry(f, textvariable=var, width=15).pack(side=tk.LEFT)
            self.dpi_vars.append(var)
            
        self.dpi_idx_var = tk.StringVar()
        f_idx = ttk.Frame(self.dpi_frame)
        f_idx.pack(fill=tk.X, pady=10)
        self.lbl_dpi_idx = ttk.Label(f_idx)
        self.lbl_dpi_idx.pack(side=tk.LEFT)
        ttk.Combobox(f_idx, textvariable=self.dpi_idx_var, values=["1", "2", "3", "4", "5", "6"], state="readonly", width=5).pack(side=tk.LEFT)

    def _build_sys_tab(self):
        row1 = ttk.Frame(self.tab_sys)
        row1.pack(fill=tk.X, pady=10)
        self.lbl_light = ttk.Label(row1, width=25)
        self.lbl_light.pack(side=tk.LEFT)
        self.light_var = tk.StringVar()
        ttk.Combobox(row1, textvariable=self.light_var, values=["0", "1", "2", "3", "4", "5", "6", "7"], state="readonly").pack(side=tk.LEFT)

        row2 = ttk.Frame(self.tab_sys)
        row2.pack(fill=tk.X, pady=10)
        self.lbl_sleep = ttk.Label(row2, width=25)
        self.lbl_sleep.pack(side=tk.LEFT)
        self.sleep_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.sleep_var).pack(side=tk.LEFT)

    def _build_debug_tab(self):
        self.log_txt = scrolledtext.ScrolledText(self.tab_debug, wrap=tk.WORD, state='disabled', font=("Consolas", 10))
        self.log_txt.pack(fill=tk.BOTH, expand=True, pady=0)

    def _refresh_status(self):
        if hasattr(self, 'log_txt'):
            self.log_txt.configure(state='normal')
            self.log_txt.delete(1.0, tk.END)
            self.log_txt.configure(state='disabled')
        
        self._append_log(self._t("log_sys_search"))
        
        if not self.mouse.device:
            if not self.mouse.connect():
                self.status_var.set(self._t("status_val", self._t("st_disconnected")))
                self._append_log(self._t("log_sys_fail"))
                self.battery_var.set(self._t("bat_val", "--"))
                self.version_var.set(self._t("fw_val", "--"))
                
                if self.tray_icon: 
                    self._last_tray_battery = -1
                    self.tray_icon.icon = self._create_image(-1)
                return

        self.status_var.set(self._t("status_val", self._t("st_connected")))
        self._append_log(self._t("log_sys_conn"))
        
        ver = self.mouse.get_version()
        self.version_var.set(self._t("fw_val", ver))
            
        is_online = self.mouse.is_online()
        if is_online:
            bat_info = self.mouse.get_battery_info()
            charging = self._t("bat_charging") if bat_info["is_charging"] else ""
            bat = bat_info['battery']
            self.battery_var.set(self._t("bat_val", f"{bat}% {charging}"))
            self._append_log(self._t("log_sys_bat", bat))
            
            if self.tray_icon:
                self._last_tray_battery = bat
                self.tray_icon.icon = self._create_image(bat)
        else:
            self.battery_var.set(self._t("bat_val", self._t("bat_offline")))
            self._append_log(self._t("log_sys_off"))
            if self.tray_icon: 
                self._last_tray_battery = -1
                self.tray_icon.icon = self._create_image(-1)
            
        self.current_config = self.mouse.get_config()
        if self.current_config:
            self._load_config_to_ui()

    def _load_config_to_ui(self):
        cfg = self.current_config
        pr_map = {0: "125", 1: "250", 2: "500", 3: "1000"}
        self.poll_var.set(pr_map.get(cfg["report_rate_idx"], "1000"))
        
        self.debounce_var.set(str(cfg["key_respond"]))
        
        lod_map = {1: "1mm", 2: "2mm"}
        self.lod_var.set(lod_map.get(cfg["lod_value"], "1mm"))
        
        for i, val in enumerate(cfg["dpis"]):
            self.dpi_vars[i].set(str(val))
            
        self.dpi_idx_var.set(str(cfg["dpi_index"] + 1))
        
        self.light_var.set(str(cfg["light_mode"]))
        self.sleep_var.set(str(cfg["sleep_light"]))


    def _apply_config(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror("Error", self._t("msg_err_conn"))
            return
            
        try:
            cfg = self.current_config
            
            pr_val = self.poll_var.get()
            pr_remap = {"125": 0, "250": 1, "500": 2, "1000": 3}
            cfg["report_rate_idx"] = pr_remap.get(pr_val, 3)
            
            cfg["key_respond"] = int(self.debounce_var.get() or "4")
            cfg["lod_value"] = 1 if self.lod_var.get() == "1mm" else 2
            
            for i in range(6):
                cfg["dpis"][i] = int(self.dpi_vars[i].get() or "400")
                
            cfg["dpi_index"] = int(self.dpi_idx_var.get() or "1") - 1
            cfg["sleep_light"] = int(self.sleep_var.get() or "10")
            cfg["light_mode"] = int(self.light_var.get() or "0")
            
            self._append_log(self._t("log_sys_writing"))
            self.mouse.set_config(cfg)
            messagebox.showinfo("OK", self._t("msg_success"))
            
        except Exception as e:
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = AjazzApp()
    app.mainloop()

