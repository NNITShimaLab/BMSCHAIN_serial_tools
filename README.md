# BMSCHAIN Serial Tools

`BMSCHAIN_GUI` 向けのシリアル出力を解析し、Excelで扱えるCSVへ変換するツール群です。

## Requirements

- Python 3.10+
- [uv](https://docs.astral.sh/uv/)

## Quick Start (uv)

このフォルダで実行します。

```bash
uv run bmschain_gui_serial_to_csv.py --help
```

ログファイルからCSVを生成:

```bash
uv run bmschain_gui_serial_to_csv.py --input capture.txt --output out.csv
```

COMポートからライブ受信:

```bash
uv run bmschain_gui_serial_to_csv.py --serial-port COM7 --baudrate 115200 --duration 20s --output out_live.csv
```

Duration accepts unit suffixes:
- `20s` (seconds)
- `5m` (minutes)
- `4h` (hours)

Rows are written to CSV incrementally during capture, so long captures do not require buffering all frames in memory.

## Project Files

- `bmschain_gui_serial_to_csv.py`: 変換スクリプト
- `csv_to_excel_voltage_current_template.py`: CSVから電圧・電流グラフ付きExcelを生成
- `README_bmschain_gui_serial_to_csv_ja.md`: 日本語の詳細手順
- `docs/serial_output_format_ja.md`: 送信フォーマット解説
- `docs/faults_order_187_ja.md`: FAULTS項目順

## Excel Plot Template

Generate `.xlsx` with voltage/current charts from CSV:

```bash
uv run csv_to_excel_voltage_current_template.py --input-csv bms_gui_output_live.csv --output-xlsx bms_plot_template.xlsx
```

Charts are configured as straight lines with small point markers (no smoothing curve interpolation).

## Prepare GitHub Repository

このフォルダをそのままGitリポジトリ化できます。

```bash
git init -b main
git add .
git commit -m "Initial commit: BMSCHAIN serial tools"
```

GitHub側で `NNITShimaLab` 配下に空リポジトリを作成後:

```bash
git remote add origin https://github.com/NNITShimaLab/<repository-name>.git
git push -u origin main
```
