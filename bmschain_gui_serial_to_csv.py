#!/usr/bin/env python3
"""BMSCHAIN GUIのシリアルストリームをExcel向けCSVに変換する。"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from pathlib import Path
from typing import Callable, Dict, List, Optional


CELL_COUNT = 14
FRAME_END_MARKER = "ENDData"
SOURCE_RELATIVE_PATH = Path("source") / "AEK_POW_BMS63CHAIN_app_mng.c"
PROJECT_DIR_NAME = (
    "SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery"
)


class FrameParseError(Exception):
    """生フレーム1件を解析できない場合の例外。"""


def strip_prefix(value: str, prefix: str) -> str:
    if value.startswith(prefix):
        return value[len(prefix) :]
    return value


def parse_int_token(token: str, label: str) -> int:
    try:
        return int(token)
    except ValueError:
        try:
            float_value = float(token)
        except ValueError as exc:
            raise FrameParseError(f"{label} is not an integer: {token!r}") from exc
        if not float_value.is_integer():
            raise FrameParseError(f"{label} is not an integer: {token!r}")
        return int(float_value)


def parse_float_token(token: str, label: str) -> float:
    try:
        return float(token)
    except ValueError as exc:
        raise FrameParseError(f"{label} is not a float: {token!r}") from exc


def tokenize_frame(raw_frame: str) -> List[str]:
    # 送信側で区切り直後に空白が入る場合があるため、トークン単位で除去する。
    normalized = raw_frame.replace("\ufeff", "")
    return [token.strip() for token in normalized.split(";") if token.strip()]


def expect_label(tokens: List[str], index: int, label: str) -> int:
    if index >= len(tokens):
        raise FrameParseError(f"Expected label {label!r}, but frame ended early")
    if tokens[index] != label:
        raise FrameParseError(
            f"Expected label {label!r} at token {index}, found {tokens[index]!r}"
        )
    return index + 1


def parse_fixed_values(
    tokens: List[str],
    index: int,
    count: int,
    parser: Callable[[str, str], float],
    section_name: str,
) -> tuple[List[float], int]:
    if index + count > len(tokens):
        raise FrameParseError(
            f"Section {section_name!r} is truncated: expected {count} values"
        )
    values: List[float] = []
    for cell_index in range(count):
        label = f"{section_name}[{cell_index + 1}]"
        values.append(parser(tokens[index + cell_index], label))
    return values, index + count


def parse_raw_frame(raw_frame: str) -> Dict[str, object]:
    tokens = tokenize_frame(raw_frame)
    index = 0

    index = expect_label(tokens, index, "TOTDEV")
    total_devices = parse_int_token(tokens[index], "TOTDEV")
    index += 1

    index = expect_label(tokens, index, "CHAIN")
    chain_id = parse_int_token(tokens[index], "CHAIN")
    index += 1

    index = expect_label(tokens, index, "DEV")
    device_id = parse_int_token(tokens[index], "DEV")
    index += 1

    index = expect_label(tokens, index, "SOC")
    soc_values, index = parse_fixed_values(
        tokens, index, CELL_COUNT, parse_int_token, "SOC"
    )

    index = expect_label(tokens, index, "Vcell:")
    vcell_values, index = parse_fixed_values(
        tokens, index, CELL_COUNT, parse_float_token, "Vcell"
    )

    index = expect_label(tokens, index, "TEMP:")
    temp_values, index = parse_fixed_values(
        tokens, index, CELL_COUNT, parse_float_token, "TEMP"
    )

    index = expect_label(tokens, index, "BAL:")
    bal_values, index = parse_fixed_values(
        tokens, index, CELL_COUNT, parse_int_token, "BAL"
    )

    index = expect_label(tokens, index, "Curr:")
    current_a = parse_float_token(tokens[index], "Curr")
    index += 1

    index = expect_label(tokens, index, "totV:")
    pack_voltage_v = parse_float_token(tokens[index], "totV")
    index += 1

    index = expect_label(tokens, index, "Vref:")
    vref_v = parse_float_token(tokens[index], "Vref")
    index += 1

    index = expect_label(tokens, index, "VUV:")
    vuv_threshold_v = parse_float_token(tokens[index], "VUV")
    index += 1

    index = expect_label(tokens, index, "VOV:")
    vov_threshold_v = parse_float_token(tokens[index], "VOV")
    index += 1

    index = expect_label(tokens, index, "GPUT:")
    gput_threshold_v = parse_float_token(tokens[index], "GPUT")
    index += 1

    index = expect_label(tokens, index, "GPOT:")
    gpot_threshold_v = parse_float_token(tokens[index], "GPOT")
    index += 1

    index = expect_label(tokens, index, "FAULTS:")
    fault_values: List[int] = []
    while index < len(tokens) and tokens[index] != "VTREF":
        fault_label = f"FAULTS[{len(fault_values) + 1}]"
        fault_values.append(parse_int_token(tokens[index], fault_label))
        index += 1

    index = expect_label(tokens, index, "VTREF")
    vtref_v = parse_float_token(tokens[index], "VTREF")

    return {
        "total_devices": total_devices,
        "chain_id": chain_id,
        "device_id": device_id,
        "soc_values": soc_values,
        "vcell_values": vcell_values,
        "temp_values": temp_values,
        "bal_values": bal_values,
        "current_a": current_a,
        "pack_voltage_v": pack_voltage_v,
        "vref_v": vref_v,
        "vuv_threshold_v": vuv_threshold_v,
        "vov_threshold_v": vov_threshold_v,
        "gput_threshold_v": gput_threshold_v,
        "gpot_threshold_v": gpot_threshold_v,
        "fault_values": fault_values,
        "vtref_v": vtref_v,
    }


def split_frames_from_text(raw_text: str) -> List[str]:
    normalized = raw_text.replace("\r", "").replace("\n", "")
    chunks = normalized.split(FRAME_END_MARKER)
    return [chunk.strip() for chunk in chunks if chunk.strip()]


def extract_fault_names_from_c(c_source_path: Path) -> List[str]:
    if not c_source_path.exists():
        return []

    source_text = c_source_path.read_text(encoding="utf-8", errors="ignore")

    function_match = re.search(
        r"void\s+AEK_POW_BMS63CHAIN_app_serialStep_GUI\s*\([^)]*\)\s*\{(?P<body>.*?)(?:sendMessage\(\"ENDData\"\))",
        source_text,
        flags=re.S,
    )
    body = function_match.group("body") if function_match else source_text

    names: List[str] = []
    for line in body.splitlines():
        if line.lstrip().startswith("//"):
            continue
        for match in re.finditer(
            r"AEK_POW_BMS63CHAIN_fastDiag\[[^\]]+\]\.([A-Za-z0-9_]+)", line
        ):
            names.append(match.group(1))
    return names


def build_fault_column_names(
    fault_count: int, fault_names_from_source: List[str]
) -> List[str]:
    columns: List[str] = []
    for index in range(fault_count):
        if index < len(fault_names_from_source):
            raw_name = strip_prefix(
                fault_names_from_source[index], "AEK_POW_BMS63CHAIN_"
            )
            columns.append(f"fault_{index + 1:03d}_{raw_name}")
        else:
            columns.append(f"fault_{index + 1:03d}")
    return columns


def write_csv(
    output_path: Path,
    parsed_frames: List[Dict[str, object]],
    fault_columns: List[str],
) -> None:
    headers = [
        "frame_index",
        "total_devices",
        "chain_id",
        "device_id",
    ]
    headers += [f"soc_cell{cell}" for cell in range(1, CELL_COUNT + 1)]
    headers += [f"vcell{cell}_v" for cell in range(1, CELL_COUNT + 1)]
    headers += [f"temp_cell{cell}_raw" for cell in range(1, CELL_COUNT + 1)]
    headers += [f"bal_cell{cell}" for cell in range(1, CELL_COUNT + 1)]
    headers += [
        "current_a",
        "pack_voltage_v",
        "vref_v",
        "vuv_threshold_v",
        "vov_threshold_v",
        "gput_threshold_v",
        "gpot_threshold_v",
    ]
    headers += fault_columns
    headers += ["vtref_v"]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(headers)

        for frame_index, frame in enumerate(parsed_frames, start=1):
            row: List[object] = [
                frame_index,
                frame["total_devices"],
                frame["chain_id"],
                frame["device_id"],
            ]
            row.extend(frame["soc_values"])  # type: ignore[arg-type]
            row.extend(frame["vcell_values"])  # type: ignore[arg-type]
            row.extend(frame["temp_values"])  # type: ignore[arg-type]
            row.extend(frame["bal_values"])  # type: ignore[arg-type]
            row.extend(
                [
                    frame["current_a"],
                    frame["pack_voltage_v"],
                    frame["vref_v"],
                    frame["vuv_threshold_v"],
                    frame["vov_threshold_v"],
                    frame["gput_threshold_v"],
                    frame["gpot_threshold_v"],
                ]
            )

            faults: List[int] = frame["fault_values"]  # type: ignore[assignment]
            row.extend(faults)
            if len(faults) < len(fault_columns):
                row.extend([""] * (len(fault_columns) - len(faults)))

            row.append(frame["vtref_v"])
            writer.writerow(row)


def read_text_file(path: Path, encoding: str) -> str:
    return path.read_text(encoding=encoding, errors="ignore")


def collect_frames_from_serial(
    port: str,
    baudrate: int,
    timeout_s: float,
    duration_s: Optional[float],
    max_frames: Optional[int],
    show_progress: bool,
) -> List[str]:
    try:
        import serial  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pyserial is required for --serial-port mode. Install with: pip install pyserial"
        ) from exc

    serial_port = serial.Serial(port=port, baudrate=baudrate, timeout=timeout_s)
    raw_frames: List[str] = []
    buffer = ""
    start_time = time.monotonic()
    last_progress_ts = 0.0

    def print_progress(force: bool = False) -> None:
        nonlocal last_progress_ts
        if not show_progress:
            return
        now = time.monotonic()
        if not force and (now - last_progress_ts) < 0.5:
            return
        last_progress_ts = now
        elapsed = now - start_time
        parts = [f"frames={len(raw_frames)}", f"elapsed={elapsed:.1f}s"]
        if duration_s is not None:
            remain = max(duration_s - elapsed, 0.0)
            parts.append(f"remaining={remain:.1f}s")
        if max_frames is not None:
            parts.append(f"target={len(raw_frames)}/{max_frames}")
        print(
            "\r[INFO] Capturing... " + ", ".join(parts),
            end="",
            file=sys.stderr,
            flush=True,
        )

    try:
        print_progress(force=True)
        try:
            while True:
                if duration_s is not None and time.monotonic() - start_time >= duration_s:
                    break
                if max_frames is not None and len(raw_frames) >= max_frames:
                    break

                chunk = serial_port.read(serial_port.in_waiting or 1)
                if not chunk:
                    continue

                buffer += chunk.decode("ascii", errors="ignore")
                buffer = buffer.replace("\r", "").replace("\n", "")

                while FRAME_END_MARKER in buffer:
                    frame_raw, buffer = buffer.split(FRAME_END_MARKER, 1)
                    frame_raw = frame_raw.strip()
                    if frame_raw:
                        raw_frames.append(frame_raw)
                        print_progress(force=True)
                    if max_frames is not None and len(raw_frames) >= max_frames:
                        break
                print_progress()
        except KeyboardInterrupt:
            pass
    finally:
        serial_port.close()
        if show_progress:
            print(file=sys.stderr)

    return raw_frames


def parse_frames(
    raw_frames: List[str],
    strict: bool,
) -> List[Dict[str, object]]:
    parsed: List[Dict[str, object]] = []
    for index, raw_frame in enumerate(raw_frames, start=1):
        try:
            parsed.append(parse_raw_frame(raw_frame))
        except FrameParseError as exc:
            message = f"[WARN] Skipped frame #{index}: {exc}"
            if strict:
                raise RuntimeError(message) from exc
            print(message, file=sys.stderr)
    return parsed


def discover_default_source_path(script_path: Path) -> Path:
    candidates = [
        script_path.parent / PROJECT_DIR_NAME / SOURCE_RELATIVE_PATH,
        script_path.parent.parent / PROJECT_DIR_NAME / SOURCE_RELATIVE_PATH,
        script_path.parent.parent / SOURCE_RELATIVE_PATH,
        Path.cwd() / PROJECT_DIR_NAME / SOURCE_RELATIVE_PATH,
        Path.cwd() / SOURCE_RELATIVE_PATH,
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    # 見つからない場合でも、エラーメッセージに使える代表パスを返す。
    return candidates[1]


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert BMSCHAIN GUI serial stream into CSV for Excel."
    )
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        "--input",
        type=Path,
        help="Path to raw serial log text file.",
    )
    input_group.add_argument(
        "--serial-port",
        help="Serial port (example: COM7). Requires pyserial.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Output CSV path.",
    )
    parser.add_argument(
        "--baudrate",
        type=int,
        default=115200,
        help="Baud rate for --serial-port mode. Default: 115200",
    )
    parser.add_argument(
        "--duration-s",
        type=float,
        default=None,
        help="Capture duration in seconds for --serial-port mode. Default: until Ctrl+C",
    )
    parser.add_argument(
        "--max-frames",
        type=int,
        default=None,
        help="Maximum number of frames to capture in --serial-port mode.",
    )
    parser.add_argument(
        "--input-encoding",
        default="utf-8",
        help="Text encoding for --input mode. Default: utf-8",
    )
    parser.add_argument(
        "--source-c",
        type=Path,
        default=None,
        help=(
            "Optional C source path used to extract FAULTS names in order. "
            "Default: auto-detect from project tree."
        ),
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail on the first malformed frame instead of skipping it.",
    )
    parser.add_argument(
        "--no-progress",
        action="store_true",
        help="Disable live progress display during serial capture.",
    )
    return parser


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()

    if args.input is not None:
        raw_text = read_text_file(args.input, args.input_encoding)
        raw_frames = split_frames_from_text(raw_text)
    else:
        try:
            raw_frames = collect_frames_from_serial(
                port=args.serial_port,
                baudrate=args.baudrate,
                timeout_s=0.05,
                duration_s=args.duration_s,
                max_frames=args.max_frames,
                show_progress=(not args.no_progress),
            )
        except RuntimeError as exc:
            print(f"[ERROR] {exc}", file=sys.stderr)
            return 1

    if not raw_frames:
        print("[ERROR] No raw frames found", file=sys.stderr)
        return 1

    parsed_frames = parse_frames(raw_frames, strict=args.strict)
    if not parsed_frames:
        print("[ERROR] No valid frames could be parsed", file=sys.stderr)
        return 1

    max_fault_count = max(len(frame["fault_values"]) for frame in parsed_frames)  # type: ignore[arg-type]
    source_c_path = args.source_c
    if source_c_path is None:
        source_c_path = discover_default_source_path(Path(__file__).resolve())
    fault_names = extract_fault_names_from_c(source_c_path)
    fault_columns = build_fault_column_names(max_fault_count, fault_names)

    write_csv(args.output, parsed_frames, fault_columns)
    print(
        "[INFO] Wrote CSV: "
        f"{args.output} (frames={len(parsed_frames)}, faults_per_frame_max={max_fault_count})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
