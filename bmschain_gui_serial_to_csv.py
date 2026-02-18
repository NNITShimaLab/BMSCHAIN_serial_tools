#!/usr/bin/env python3
"""BMSCHAIN GUIのシリアルストリームをExcel向けCSVに変換する。"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from pathlib import Path
from typing import Callable, Dict, Iterator, List, Optional


CELL_COUNT = 14
FRAME_END_MARKER = "ENDData"
SOURCE_RELATIVE_PATH = Path("source") / "AEK_POW_BMS63CHAIN_app_mng.c"
PROJECT_DIR_NAME = (
    "SPC58EC - AEK_POW_BMS63EN_SOC_Est_SingleAccess_CHAIN_GUI_application for discovery"
)
DEFAULT_FAULT_COUNT = 187


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


def parse_duration_to_seconds(duration_text: str) -> float:
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*([smhSMH]?)\s*", duration_text)
    if not match:
        raise ValueError(
            "Invalid --duration format. Use examples like: 20s, 5m, 4h, 30"
        )
    value = float(match.group(1))
    unit = match.group(2).lower() if match.group(2) else "s"
    if unit == "s":
        return value
    if unit == "m":
        return value * 60.0
    if unit == "h":
        return value * 3600.0
    raise ValueError(f"Unsupported duration unit: {unit}")


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


def iter_frames_from_text_file(
    path: Path,
    encoding: str,
    chunk_chars: int = 65536,
) -> Iterator[str]:
    buffer = ""
    with path.open("r", encoding=encoding, errors="ignore") as fp:
        while True:
            chunk = fp.read(chunk_chars)
            if not chunk:
                break
            buffer += chunk.replace("\r", "").replace("\n", "")
            while FRAME_END_MARKER in buffer:
                frame_raw, buffer = buffer.split(FRAME_END_MARKER, 1)
                frame_raw = frame_raw.strip()
                if frame_raw:
                    yield frame_raw


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


class CsvStreamWriter:
    def __init__(self, output_path: Path, fault_columns: List[str]) -> None:
        self.output_path = output_path
        self.fault_columns = fault_columns
        self._fp = None
        self._writer = None

    def __enter__(self) -> "CsvStreamWriter":
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
        headers += self.fault_columns
        headers += ["vtref_v"]

        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self._fp = self.output_path.open("w", encoding="utf-8-sig", newline="")
        self._writer = csv.writer(self._fp)
        self._writer.writerow(headers)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._fp is not None:
            self._fp.close()
            self._fp = None
            self._writer = None

    def write_frame(self, frame_index: int, frame: Dict[str, object]) -> None:
        if self._writer is None:
            raise RuntimeError("CsvStreamWriter is not opened.")

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
        if len(faults) > len(self.fault_columns):
            faults = faults[: len(self.fault_columns)]
        row.extend(faults)
        if len(faults) < len(self.fault_columns):
            row.extend([""] * (len(self.fault_columns) - len(faults)))

        row.append(frame["vtref_v"])
        self._writer.writerow(row)


def iter_frames_from_serial(
    port: str,
    baudrate: int,
    timeout_s: float,
    duration_seconds: Optional[float],
    max_frames: Optional[int],
    show_progress: bool,
) -> Iterator[str]:
    try:
        import serial  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pyserial is required for --serial-port mode. Install with: pip install pyserial"
        ) from exc

    serial_port = serial.Serial(port=port, baudrate=baudrate, timeout=timeout_s)
    captured_frames = 0
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
        parts = [f"frames={captured_frames}", f"elapsed={elapsed:.1f}s"]
        if duration_seconds is not None:
            remain = max(duration_seconds - elapsed, 0.0)
            parts.append(f"remaining={remain:.1f}s")
        if max_frames is not None:
            parts.append(f"target={captured_frames}/{max_frames}")
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
                if (
                    duration_seconds is not None
                    and time.monotonic() - start_time >= duration_seconds
                ):
                    break
                if max_frames is not None and captured_frames >= max_frames:
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
                        captured_frames += 1
                        print_progress(force=True)
                        yield frame_raw
                    if max_frames is not None and captured_frames >= max_frames:
                        break
                print_progress()
        except KeyboardInterrupt:
            pass
    finally:
        serial_port.close()
        if show_progress:
            print(file=sys.stderr)


def stream_frames_to_csv(
    raw_frames: Iterator[str],
    writer: CsvStreamWriter,
    strict: bool,
) -> tuple[int, int, int]:
    parsed_count = 0
    skipped_count = 0
    max_fault_count = 0

    for source_index, raw_frame in enumerate(raw_frames, start=1):
        try:
            frame = parse_raw_frame(raw_frame)
        except FrameParseError as exc:
            message = f"[WARN] Skipped frame #{source_index}: {exc}"
            if strict:
                raise RuntimeError(message) from exc
            print(message, file=sys.stderr)
            skipped_count += 1
            continue

        parsed_count += 1
        fault_count = len(frame["fault_values"])  # type: ignore[arg-type]
        max_fault_count = max(max_fault_count, fault_count)
        writer.write_frame(parsed_count, frame)

    return parsed_count, skipped_count, max_fault_count if max_fault_count > 0 else 0


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
        "--duration",
        type=str,
        default=None,
        help="Capture duration for --serial-port mode (examples: 20s, 5m, 4h).",
    )
    parser.add_argument(
        "--max-frames",
        type=int,
        default=None,
        help="Maximum number of frames to capture in --serial-port mode.",
    )
    parser.add_argument(
        "--duration-s",
        type=float,
        default=None,
        help=argparse.SUPPRESS,
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
    source_c_path = args.source_c
    if source_c_path is None:
        source_c_path = discover_default_source_path(Path(__file__).resolve())
    fault_names = extract_fault_names_from_c(source_c_path)
    expected_fault_count = len(fault_names) if fault_names else DEFAULT_FAULT_COUNT
    fault_columns = build_fault_column_names(expected_fault_count, fault_names)

    duration_seconds: Optional[float] = None
    if args.duration is not None:
        try:
            duration_seconds = parse_duration_to_seconds(args.duration)
        except ValueError as exc:
            print(f"[ERROR] {exc}", file=sys.stderr)
            return 1
    elif args.duration_s is not None:
        duration_seconds = args.duration_s

    if args.input is not None:
        raw_frames = iter_frames_from_text_file(args.input, args.input_encoding)
    else:
        try:
            raw_frames = iter_frames_from_serial(
                port=args.serial_port,
                baudrate=args.baudrate,
                timeout_s=0.05,
                duration_seconds=duration_seconds,
                max_frames=args.max_frames,
                show_progress=(not args.no_progress),
            )
        except RuntimeError as exc:
            print(f"[ERROR] {exc}", file=sys.stderr)
            return 1

    with CsvStreamWriter(args.output, fault_columns) as writer:
        parsed_count, skipped_count, max_fault_count = stream_frames_to_csv(
            raw_frames=raw_frames,
            writer=writer,
            strict=args.strict,
        )

    if parsed_count == 0:
        print("[ERROR] No valid frames could be parsed", file=sys.stderr)
        return 1

    if max_fault_count > len(fault_columns):
        print(
            "[WARN] Some frames contained more FAULTS than configured columns; extra values were truncated.",
            file=sys.stderr,
        )

    print(
        "[INFO] Wrote CSV: "
        f"{args.output} (frames={parsed_count}, skipped={skipped_count}, faults_per_frame_max={max_fault_count})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
