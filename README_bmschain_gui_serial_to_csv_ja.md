# bmschain_gui_serial_to_csv.py ドキュメント

## 1. 目的
`AEK_POW_BMS63CHAIN_app_serialStep_GUI()` のシリアル出力を受け取り、Excelで開けるCSVへ変換します。

## 2. 作成時の計画
1. GUI向けフレーム仕様 (`TOTDEV ... ENDData`) をソースコードから確定
2. ログ入力とCOMライブ受信の両方に対応
3. `ENDData` 単位でフレーム復元し、固定セクション + 可変 `FAULTS` をCSV列に展開
4. `FAULTS` 列名は Cソースから自動抽出（無ければ `fault_001` 形式でフォールバック）

## 3. スクリプト
- ファイル: `BMSCHAIN_serial_tools/bmschain_gui_serial_to_csv.py`

## 4. 前提
- Python 3.10+
- `uv` がインストール済みであること
- 依存パッケージ (`pyserial`) は `uv run` 時に自動解決されます

## 5. 使い方

### 5.1 ログファイルをCSVへ変換
```bash
uv run bmschain_gui_serial_to_csv.py --input capture.txt --output bms_gui_output.csv
```

### 5.2 COMポートからライブ受信してCSVへ保存
```bash
uv run bmschain_gui_serial_to_csv.py --serial-port COM7 --baudrate 115200 --duration-s 20 --output bms_gui_output_live.csv
```

### 5.3 主なオプション
- `--input <path>`: 既存ログを読む
- `--serial-port <COMx>`: COMから直接受信
- `--output <path>`: 出力CSV
- `--baudrate <int>`: 既定 `115200`
- `--duration-s <float>`: ライブ受信時間（省略時は継続）
- `--max-frames <int>`: 最大フレーム数で停止
- `--source-c <path>`: `FAULTS` 名抽出に使う Cファイル
- `--strict`: 不正フレームで即終了（既定はスキップ）
- `--no-progress`: シリアル受信中の進捗表示を無効化

## 6. 出力CSV仕様
- 文字コード: UTF-8 BOM（Excelで開きやすい）
- 1行 = 1フレーム（1デバイス分）
- 代表列:
  - `frame_index`, `total_devices`, `chain_id`, `device_id`
  - `soc_cell1..14`
  - `vcell1_v..vcell14_v`
  - `temp_cell1_raw..temp_cell14_raw`
  - `bal_cell1..bal_cell14`
  - `current_a`, `pack_voltage_v`, `vref_v`, `vuv_threshold_v`, `vov_threshold_v`, `gput_threshold_v`, `gpot_threshold_v`
  - `fault_001_xxx ...`（最大検出数まで）
  - `vtref_v`

## 7. エラーハンドリング
- `ENDData` で復元できるフレームのみ対象
- 不正フレームは既定で警告表示してスキップ
- `--strict` 指定時は最初の不正フレームで終了

## 8. 補足
- フォーマット詳細は `docs/serial_output_format_ja.md` を参照してください。
- `FAULTS` 187項目の完全順序は `docs/faults_order_187_ja.md` を参照してください。
- このREADMEのコマンドは `BMSCHAIN_serial_tools` フォルダ内で実行する前提です。
- `--serial-port` モードでは、受信中に `frames` / `elapsed` などの進捗を表示します。

## 9. 電圧・電流プロット付きExcel生成
- スクリプト: `csv_to_excel_voltage_current_template.py`
- 目的: CSVから電圧(各セル)と電流のグラフを含む `.xlsx` を生成

```bash
uv run csv_to_excel_voltage_current_template.py --input-csv bms_gui_output_live.csv --output-xlsx bms_plot_template.xlsx
```

特定デバイスのみ対象にする場合:

```bash
uv run csv_to_excel_voltage_current_template.py --input-csv bms_gui_output_live.csv --output-xlsx bms_plot_template_dev1.xlsx --chain-id 0 --device-id 1
```
