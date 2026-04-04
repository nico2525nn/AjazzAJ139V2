import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from ajazz_mouse import AjazzMouse

class AjazzApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ajazz AJ139 V2 Control - Advanced")
        self.geometry("650x650")
        
        # マウスモジュール初期化時に、通信があったら _append_log メソッドを呼ぶよう設定
        self.mouse = AjazzMouse(log_callback=self._append_log)
        self.current_config = None
        
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        
        self._build_ui()
        self._refresh_status()

    def _append_log(self, text):
        """デバッグログエリアに通信の16進数データを追記する"""
        if hasattr(self, 'log_txt'):
            self.log_txt.configure(state='normal')
            self.log_txt.insert(tk.END, text + "\n")
            self.log_txt.see(tk.END)
            self.log_txt.configure(state='disabled')

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === タブ1: ステータス ===
        self.tab_status = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_status, text="基本ステータス")
        self._build_status_tab()
        
        # === タブ2: パフォーマンス ===
        self.tab_perf = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_perf, text="パフォーマンス設定")
        self._build_perf_tab()
        
        # === タブ3: システム＆ライティング ===
        self.tab_sys = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_sys, text="ライティング / その他設定")
        self._build_sys_tab()
        
        # === タブ4: デバッグログ ===
        self.tab_debug = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_debug, text="通信デバッグログ")
        self._build_debug_tab()
        
        # 画面下部の共通ボタン群
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_frame, text="設定をマウスに書き込む", command=self._apply_config).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="マウスから最新状態を読み直す", command=self._refresh_status).pack(side=tk.RIGHT, padx=10)

    def _build_status_tab(self):
        self.status_var = tk.StringVar(value="状態: 未接続")
        self.battery_var = tk.StringVar(value="バッテリー: -")
        self.version_var = tk.StringVar(value="ファームウェア: -")
        
        status_frame = ttk.LabelFrame(self.tab_status, text="USB / ドングルステータス", padding=10)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(status_frame, textvariable=self.status_var, font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
        ttk.Label(status_frame, textvariable=self.version_var).pack(anchor=tk.W)
        ttk.Label(status_frame, textvariable=self.battery_var).pack(anchor=tk.W)

    def _build_debug_tab(self):
        log_frame = ttk.LabelFrame(self.tab_debug, text="USB通信の生ログ (16進数)", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(log_frame, text="※不具合発生時や未知の設定を探る際、このSENDとRECVのパケットが役立ちます。").pack(anchor=tk.W)
        self.log_txt = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', font=("Consolas", 10))
        self.log_txt.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

    def _build_perf_tab(self):
        # Polling Rate
        row1 = ttk.Frame(self.tab_perf)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="レポートレート (Hz):", width=25).pack(side=tk.LEFT)
        self.poll_var = tk.StringVar()
        poll_combo = ttk.Combobox(row1, textvariable=self.poll_var, values=["125", "250", "500", "1000"], state="readonly")
        poll_combo.pack(side=tk.LEFT)
        
        # Debounce
        row2 = ttk.Frame(self.tab_perf)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="デバウンスタイム (ms):", width=25).pack(side=tk.LEFT)
        self.debounce_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.debounce_var).pack(side=tk.LEFT)

        # LOD
        row3 = ttk.Frame(self.tab_perf)
        row3.pack(fill=tk.X, pady=5)
        ttk.Label(row3, text="LOD (Lift Off Distance):", width=25).pack(side=tk.LEFT)
        self.lod_var = tk.StringVar()
        lod_combo = ttk.Combobox(row3, textvariable=self.lod_var, values=["1mm", "2mm"], state="readonly")
        lod_combo.pack(side=tk.LEFT)
        
        # DPI Info 
        dpi_frame = ttk.LabelFrame(self.tab_perf, text="DPI レベル (DPI 1 ～ 6 の値)", padding=10)
        dpi_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.dpi_vars = []
        for i in range(6):
            f = ttk.Frame(dpi_frame)
            f.pack(fill=tk.X, pady=2)
            ttk.Label(f, text=f"DPI 色 {i+1}:", width=12).pack(side=tk.LEFT)
            var = tk.StringVar(value="400")
            ttk.Entry(f, textvariable=var, width=15).pack(side=tk.LEFT)
            self.dpi_vars.append(var)
            
        self.dpi_idx_var = tk.StringVar()
        f_idx = ttk.Frame(dpi_frame)
        f_idx.pack(fill=tk.X, pady=10)
        ttk.Label(f_idx, text="現在アクティブなDPIインデックス (1-6):").pack(side=tk.LEFT)
        ttk.Combobox(f_idx, textvariable=self.dpi_idx_var, values=["1", "2", "3", "4", "5", "6"], state="readonly", width=5).pack(side=tk.LEFT)

    def _build_sys_tab(self):
        row1 = ttk.Frame(self.tab_sys)
        row1.pack(fill=tk.X, pady=10)
        ttk.Label(row1, text="ライティングモード:", width=25).pack(side=tk.LEFT)
        self.light_var = tk.StringVar()
        ttk.Combobox(row1, textvariable=self.light_var, values=["オフ", "モード1", "モード2", "モード3", "モード4", "モード5", "モード6", "モード7"], state="readonly").pack(side=tk.LEFT)

        row2 = ttk.Frame(self.tab_sys)
        row2.pack(fill=tk.X, pady=10)
        ttk.Label(row2, text="LEDオートスリープ (分):", width=25).pack(side=tk.LEFT)
        self.sleep_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.sleep_var).pack(side=tk.LEFT)

    def _refresh_status(self):
        # ログをクリア
        if hasattr(self, 'log_txt'):
            self.log_txt.configure(state='normal')
            self.log_txt.delete(1.0, tk.END)
            self.log_txt.configure(state='disabled')
        
        self._append_log("System: USBデバイスを検索中...")
        
        if not self.mouse.device:
            if not self.mouse.connect():
                self.status_var.set("状態: ドングル見つかりません (切断)")
                self._append_log("System: デバイスが見つかりません。USBに刺さっていますか？")
                return

        self.status_var.set("状態: 接続中 (USB連携)")
        self._append_log("System: デバイスと接続しました。\n")
        
        ver = self.mouse.get_version()
        self.version_var.set(f"ファームウェア: {ver}")
            
        is_online = self.mouse.is_online()
        if is_online:
            bat_info = self.mouse.get_battery_info()
            charging = " (充電中)" if bat_info["is_charging"] else ""
            self.battery_var.set(f"バッテリー: {bat_info['battery']}% {charging}")
            self._append_log(f"System: オンライン判定。バッテリー={bat_info['battery']}%")
        else:
            self.battery_var.set("バッテリー: オフライン (スリープ中など) ※ログを確認")
            self._append_log("System: is_onlineの判定がFalseになりました。上記RECVのログにバッテリーぽいものが見えるか確認してください。")
            
        self.current_config = self.mouse.get_config()
        if self.current_config:
            self._load_config_to_ui()

    def _load_config_to_ui(self):
        cfg = self.current_config
        
        # Report Rate
        pr_map = {0: "125", 1: "250", 2: "500", 3: "1000"}
        self.poll_var.set(pr_map.get(cfg["report_rate_idx"], "1000"))
        
        # Debounce
        self.debounce_var.set(str(cfg["key_respond"]))
        
        # LOD
        lod_map = {1: "1mm", 2: "2mm"}
        self.lod_var.set(lod_map.get(cfg["lod_value"], "1mm"))
        
        # DPIs
        for i, val in enumerate(cfg["dpis"]):
            self.dpi_vars[i].set(str(val))
            
        self.dpi_idx_var.set(str(cfg["dpi_index"] + 1))
        
        # Lighting (0=Off, 1~7=Modes)
        idx = cfg["light_mode"]
        if idx == 0:
             self.light_var.set("オフ")
        else:
             self.light_var.set(f"モード{min(idx, 7)}")
             
        # Sleep
        self.sleep_var.set(str(cfg["sleep_light"]))

    def _apply_config(self):
        if not self.mouse.device or not self.current_config:
            messagebox.showerror("エラー", "デバイスと通信できていないか、設定データがありません。")
            return
            
        try:
            cfg = self.current_config
            
            # Polling Rate
            pr_val = self.poll_var.get()
            pr_remap = {"125": 0, "250": 1, "500": 2, "1000": 3}
            cfg["report_rate_idx"] = pr_remap.get(pr_val, 3)
            
            # Debounce
            cfg["key_respond"] = int(self.debounce_var.get() or "4")
            
            # LOD
            cfg["lod_value"] = 1 if self.lod_var.get() == "1mm" else 2
            
            # DPI
            for i in range(6):
                cfg["dpis"][i] = int(self.dpi_vars[i].get() or "400")
                
            cfg["dpi_index"] = int(self.dpi_idx_var.get() or "1") - 1
            
            # Sleep
            cfg["sleep_light"] = int(self.sleep_var.get() or "10")
            
            # Light
            lt_str = self.light_var.get()
            if lt_str == "オフ":
                cfg["light_mode"] = 0
            elif lt_str.startswith("モード"):
                cfg["light_mode"] = int(lt_str.replace("モード", ""))
            
            self._append_log("\nSystem: ==== 設定を書き込みます ====")
            self.mouse.set_config(cfg)
            messagebox.showinfo("成功", "設定を書き込みました！")
            
        except Exception as e:
            messagebox.showerror("書き込みエラー", str(e))

if __name__ == "__main__":
    app = AjazzApp()
    app.mainloop()
