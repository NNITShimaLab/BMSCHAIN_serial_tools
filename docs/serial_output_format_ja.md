# SPC58EC BMSCHAIN GUI シリアル出力フォーマット解説

## 1. 対象
- 対象プロジェクト: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery`
- 比較対象: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN application for discovery`（人間可読出力）

## 2. 使用されているシリアルAPI
- 初期化: `sd_lld_start(&SD5, &serial_config_configuration_name)`
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:207`
- 送信: `sd_lld_write(&SD5, (uint8_t*)outputMessage, len)`
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:204`
- 設定 (`serial_config_configuration_name`):
  - Baud: `115200`
  - Format: `8bit / parity none` (8N1 相当)
  - API mode: `SPC5_LIN_API_MODE_SYNCHRONOUS`
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/components/spc58ecxx_low_level_drivers_component_rla/cfg/serial_lld_cfg.c:37`

## 3. 送信周期と実行位置
- `main_core0()` で `AEK_POW_BMS63CHAIN_app_serialStep_GUI(1000)` を繰り返し呼び出し。
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/main.c:25`
- 実際の送信は `osalThreadGetMilliseconds() % timeStamp == 0` の時のみ実施。
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:219`
- GUI向けフォーマット関数: `AEK_POW_BMS63CHAIN_app_serialStep_GUI()`
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:211`

## 4. 1フレームの構造
- 区切り文字: `;`
- フレーム終端: `ENDData`
- フレーム開始の専用マーカはなし（`TOTDEV` から開始）
- 1デバイスごとに1フレーム送信

### 文法（概略）
```text
TOTDEV;<n>;
CHAIN;<chain_id>;DEV;<dev_id>;
SOC;<14値>;
Vcell:;<14値>;
TEMP:;<14値>;
BAL:;<14値>;
Curr:;<1値>;
totV:;<1値>;
Vref:;<1値>;
VUV:;<1値>;
VOV:;<1値>;
GPUT:;<1値>;
GPOT:;<1値>;
FAULTS:;<187値>;
VTREF;<1値>;
ENDData
```

## 5. セクション別フィールド数
- `TOTDEV`: 1（全デバイス数）
- `CHAIN`: 1
- `DEV`: 1
- `SOC`: 14
- `Vcell:`: 14
- `TEMP:`: 14
- `BAL:`: 14
- `Curr:`: 1
- `totV:`: 1
- `Vref:`: 1
- `VUV:`: 1
- `VOV:`: 1
- `GPUT:`: 1
- `GPOT:`: 1
- `FAULTS:`: 187（0/1中心の故障フラグ群）
- `VTREF`: 1

数値フィールド総数: `254`

## 6. `CHAIN application`（人間可読）との違い
- 人間可読側 (`AEK_POW_BMS63CHAIN_app_serialStep`) は改行ベース + カンマ形式:
  - 例: `DEV,1,` / `S1,080,` / `V1,3.95,`
  - 参照: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:211`
- GUI側は `;` 連結 + `ENDData` 終端で、`FAULTS` を含む大規模フレームを送信。

## 7. `FAULTS` の並び順
- `FAULTS` は固定順で 187 項目。
- 並び順は `sendMessage()` の呼び出し順そのまま。
  - 開始付近: `ovrLatch`, `tCycleOvf`, `sensePlusOpen`, ...
  - 終了付近: `curSenseOvcNorm`, `oscFail`, `clkMonEn`, `clkMonInitDone`, `openBistFail`
  - 参照範囲: `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:269` ～ `SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery/source/AEK_POW_BMS63CHAIN_app_mng.c:456`
- 完全な 187 項目順リスト: `BMSCHAIN_serial_tools/docs/faults_order_187_ja.md`

## 8. 実装上の注意
- ログ中に改行が入っても、実体は `ENDData` で1フレーム復元可能。
- `TEMP` は名称上 Temperature だが、実装では NTC計算値を送っているため、生温度そのものとは限らない。
- GUI側送信は `sprintf` を多用しているため、想定外の値幅では文字列破損の可能性がある。解析時は破損フレームをスキップできる実装が安全。
