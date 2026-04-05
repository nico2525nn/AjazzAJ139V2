# Ajazz AJ139 V2 コントロールソフト

[English README](./README.md)

[Ajazz AJ139 V2](https://139v2mc.yjx2012.com/) 用の非公式オンデバイス設定ツールです。

ベンダーの Web UI を使わず、`hidapi` と Tkinter ベースのローカルアプリから直接マウス設定を変更できます。タスクトレイ常駐、キーマッピング、マクロ管理まで含めて Python だけで動きます。

## 主な機能
- ファームウェア、オンライン状態、バッテリー情報の表示
- パフォーマンス設定
- DPI 段階設定
- ポーリングレート
- デバウンス時間
- LOD
- ライティング / スリープ設定
- 8 ボタン分のキーマッピング
- マウス機能、メディアキー、通常キー、修飾キーコンボの割り当て
- `F24` までの拡張ファンクションキー対応
- 32 スロットのマクロ管理
- オンデバイスのマクロ読込 / 書込 / 初期化
- マクロ録画と手動イベント追加
- マクロ表の `Name` / `Action` / `Delay` の直接編集
- 英語 / 日本語 UI
- 生の HID 通信ログ表示
- タスクトレイのバッテリー表示

## 動作要件
- Python 3.8 以上
- Windows 推奨
- 必要な Python パッケージ
  - `hidapi`
  - `Pillow`
  - `pystray`
  - `pyinstaller`（exe 化用）

インストール:

```powershell
python -m pip install -r requirements.txt
```

## 起動方法

```powershell
python main.py
```

`pystray` などの不足エラーが出る場合は、起動に使っている Python と同じ環境へ入っているか確認してください。

```powershell
python -m pip install -r requirements.txt
python main.py
```

## ビルド方法

付属の `build.bat` を使います。

```powershell
build.bat
```

内部では次を実行します。

```powershell
python -m PyInstaller --noconfirm --onefile --windowed main.py
```

## 使い方メモ
- ウィンドウを閉じると終了ではなくタスクトレイへ格納されます。
- `Refresh Status` ではまず基本状態を読み込みます。
- マクロデータは起動直後ではなく `Macro` タブを開いたときに遅延読込します。
- 重い HID 操作はバックグラウンドで実行するため、UI が固まりにくくなっています。
- `Macro` タブにはロード中 / 書込中 / 初期化中の状態表示があります。
- このデバイスではマクロ読込チャンクが空で返ることがあり、アプリ側で長時間タイムアウトしないよう対策しています。

## 既知の注意点
- マクロ領域の仕様はベンダー Web アプリを元にしたリバースエンジニアリングです。
- マクロ領域が空に近い個体では、通常設定より読込が遅く見えることがあります。
- 最終的な確認は実機テストが前提です。

## 主なファイル
- [main.py](/D:/学校/python/AjazzAJ139V2/main.py): エントリーポイント
- [ui_app.py](/D:/学校/python/AjazzAJ139V2/ui_app.py): Tkinter UI とバックグラウンド処理
- [ajazz_mouse.py](/D:/学校/python/AjazzAJ139V2/ajazz_mouse.py): HID プロトコル実装
- [build.bat](/D:/学校/python/AjazzAJ139V2/build.bat): Windows 用ビルド補助
- [requirements.txt](/D:/学校/python/AjazzAJ139V2/requirements.txt): 依存関係
