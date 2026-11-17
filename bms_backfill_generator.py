#!/usr/bin/env python3
"""Automates BMS CSV timeseries extraction (via Mapping Key), COV forward-filling, delete-mode command generation, reporting, and backfill CLI execution."""

from collections import defaultdict
import csv
from datetime import datetime, timezone
import io
import os
from pathlib import Path
import pty
import re
import select
import subprocess
import sys
import time
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

try:
  import openpyxl
  from openpyxl.styles import Font, PatternFill
except ImportError:
  sys.exit(
      "Error: 'openpyxl' is required to read/write .xlsx files. Install it via"
      " 'pip install openpyxl'."
  )

# Seattle, WA timezone (handles PST -08:00 and PDT -07:00 automatically)
SEATTLE_TZ = ZoneInfo("America/Los_Angeles")

# Regular output timestamp format: YYYY-MM-DD HH:MM:SS±HH:MM (no fractional seconds)
VALID_TIMESTAMP_REGEX = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}$"
)

# Regex to extract generated Bigtable delete commands from BackfillTool delete-mode output
BT_DELETE_REGEX = re.compile(r"(bt\s+delete\s+/bigtable/[^\r\n]+)")

POPULATE_STEPS_DESCRIPTION = """
Populate Mode Workflow Steps:
  1. Read BMS CSV standardFieldName mapping key (.xlsx)
  2. Read proxy_id <-> device_num_id mappings (.csv)
  3. Match BMS timeseries data (.csv) [heavy step]
  4. Generate per-proxyId CSVs (with COV forward-fill) and CLI command files
  5. Write Mapping Key Matching & broken timestamps report tabs (backfill_report.xlsx)
  6. Run backfill CLI commands (p4 g4d / g4 sync / blaze run) & record results
"""

DELETE_STEPS_DESCRIPTION = """
Delete Mode Workflow Steps:
  1. Read proxy_id <-> device_num_id mappings (.csv) & prompt for building + start/end timestamps
  2. Generate per-proxyId delete CSVs and delete CLI command files
  3. Run delete CLI commands (p4 g4d / g4 sync / blaze run), extract 'bt delete' commands & write summary/report
"""


class MappingRow:
  """Represents a single mapping row from the BMS CSV standardFieldName mapping key."""

  def __init__(
      self,
      row_number: int,
      unique_col2_value: str,
      proxy_id: str,
      building: str,
      standard_field_name: str,
  ) -> None:
    self.row_number = row_number
    self.unique_col2_value = unique_col2_value
    self.proxy_id = proxy_id
    self.building = building
    self.standard_field_name = standard_field_name
    self.matched_bms_rows: int = 0


def render_progress_bar(
    current: int,
    total: int,
    label: str,
    extra: str = "",
    width: int = 30,
) -> None:
  """Renders an inline terminal progress bar and clears trailing line characters."""
  ratio = min(1.0, current / total) if total > 0 else 1.0
  filled = int(width * ratio)
  bar = "#" * filled + "-" * (width - filled)
  suffix = f" | {extra}" if extra else ""
  sys.stdout.write(f"\r  [{bar}] {ratio * 100:5.1f}% | {label}{suffix}\033[K")
  sys.stdout.flush()


def convert_and_classify_seattle_tz(
    timestamp_str: str,
    cache: Dict[str, Tuple[Optional[str], bool, Optional[str]]],
) -> Tuple[Optional[str], bool, Optional[str]]:
  """Converts an ISO-8601 timestamp to Seattle time and classifies whether it is regular or COV/irregular.

  Returns:
    (formatted_timestamp, is_regular, error_reason)
  """
  raw = timestamp_str.strip()
  cached = cache.get(raw)
  if cached is not None:
    return cached

  result: Tuple[Optional[str], bool, Optional[str]]
  try:
    iso_input = f"{raw[:-1]}+00:00" if raw.endswith(("Z", "z")) else raw
    dt = datetime.fromisoformat(iso_input)
    if dt.tzinfo is None:
      dt = dt.replace(tzinfo=timezone.utc)
    dt_local = dt.astimezone(SEATTLE_TZ)

    if dt_local.microsecond != 0:
      formatted = dt_local.isoformat(sep=" ", timespec="milliseconds")
      result = (
          formatted,
          False,
          "COV / fractional seconds (forward-filled, omitted from CSV)",
      )
    else:
      formatted = dt_local.isoformat(sep=" ", timespec="seconds")
      if not VALID_TIMESTAMP_REGEX.match(formatted):
        result = (
            None,
            False,
            f"Formatted timestamp '{formatted}' failed format validation",
        )
      else:
        result = (formatted, True, None)
  except ValueError as exc:
    result = (None, False, f"Invalid ISO-8601 timestamp ({exc})")

  if len(cache) < 250_000:
    cache[raw] = result
  return result


def parse_and_format_delete_timestamp(raw_input: str) -> str:
  """Parses a user-entered start/end timestamp and formats it in Seattle time ('YYYY-MM-DD HH:MM:SS±HH:MM').

  If the user enters a naive timestamp (without a timezone offset), it is
  interpreted as Seattle local time (America/Los_Angeles).
  """
  cleaned = raw_input.strip()
  iso_input = f"{cleaned[:-1]}+00:00" if cleaned.endswith(("Z", "z")) else cleaned
  dt = datetime.fromisoformat(iso_input)
  if dt.tzinfo is None:
    dt_local = dt.replace(tzinfo=SEATTLE_TZ)
  else:
    dt_local = dt.astimezone(SEATTLE_TZ)

  formatted = dt_local.isoformat(sep=" ", timespec="seconds")
  if not VALID_TIMESTAMP_REGEX.match(formatted):
    raise ValueError(
        f"Timestamp '{raw_input}' did not format to 'YYYY-MM-DD HH:MM:SS±HH:MM'"
        f" (got '{formatted}')."
    )
  return formatted


def prompt_delete_timestamp(prompt_text: str) -> str:
  """Prompts the user for a delete-mode timestamp and normalizes/validates it."""
  while True:
    raw = input(prompt_text).strip()
    if not raw:
      print("Timestamp cannot be blank. Please try again.")
      continue
    try:
      formatted = parse_and_format_delete_timestamp(raw)
      print(f"  -> Normalized Seattle timestamp: {formatted}")
      return formatted
    except ValueError as exc:
      print(
          f"  Invalid timestamp ({exc}). Examples: '2025-03-01 00:00:00-08:00'"
          " or '2025-03-01 00:00:00'."
      )


def read_mapping_key_xlsx(
    xlsx_path: Path, building_name: str
) -> List[MappingRow]:
  """Reads the BMS CSV standardFieldName mapping key (.xlsx) with columns unique_col2_value, proxyId, standardFieldName."""
  wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
  ws = wb.active

  rows_iter = ws.iter_rows(values_only=True)
  try:
    header_row = next(rows_iter)
  except StopIteration:
    wb.close()
    raise ValueError(f"Mapping key file '{xlsx_path}' is empty.")

  headers = [
      str(col).strip() if col is not None else "" for col in header_row
  ]
  headers_lower = {name.lower(): i for i, name in enumerate(headers) if name}

  required_cols = {"unique_col2_value", "proxyid", "standardfieldname"}
  missing = required_cols - set(headers_lower.keys())
  if missing:
    wb.close()
    raise ValueError(
        "Mapping key .xlsx is missing required columns:"
        f" {', '.join(sorted(missing))}"
    )

  col2_idx = headers_lower["unique_col2_value"]
  proxy_idx = headers_lower["proxyid"]
  field_idx = headers_lower["standardfieldname"]

  mappings: List[MappingRow] = []
  for row_num, row in enumerate(rows_iter, start=2):
    if not row:
      continue

    def get_cell(idx: int) -> str:
      if idx < len(row) and row[idx] is not None:
        return str(row[idx]).strip()
      return ""

    unique_col2 = get_cell(col2_idx)
    proxy_id = get_cell(proxy_idx)
    standard_field_name = get_cell(field_idx)

    if not (unique_col2 and proxy_id and standard_field_name):
      continue

    mappings.append(
        MappingRow(
            row_number=row_num,
            unique_col2_value=unique_col2,
            proxy_id=proxy_id,
            building=building_name.strip(),
            standard_field_name=standard_field_name,
        )
    )

  wb.close()
  return mappings


def read_device_num_id_mappings(csv_path: Path) -> Dict[str, str]:
  """Reads the proxy_id <-> external_id mappings CSV into a dict."""
  device_map: Dict[str, str] = {}
  with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
    sample = f.read(4096)
    f.seek(0)
    first_line = sample.splitlines()[0] if sample else ""
    delimiter = "\t" if "\t" in first_line else ","
    reader = csv.DictReader(f, delimiter=delimiter)
    if not reader.fieldnames:
      raise ValueError(f"Mapping CSV '{csv_path}' is empty or missing headers.")

    normalized_fields = {
        name.strip().lower(): name for name in reader.fieldnames if name
    }
    if (
        "proxy_id" not in normalized_fields
        or "device_num_id" not in normalized_fields
    ):
      raise ValueError(
          "Mapping CSV must contain 'proxy_id' and 'device_num_id' columns."
      )

    proxy_col = normalized_fields["proxy_id"]
    device_col = normalized_fields["device_num_id"]

    for row in reader:
      proxy_id = (row.get(proxy_col) or "").strip()
      device_num_id = (row.get(device_col) or "").strip()
      if proxy_id and device_num_id:
        device_map[proxy_id] = device_num_id

  return device_map


def process_bms_csv(
    bms_csv_path: Path, mappings: List[MappingRow]
) -> Tuple[
    Dict[str, Dict[str, Dict[str, str]]],
    Dict[str, List[str]],
    Dict[Tuple[str, str, str, str], int],
]:
  """Streams the headerless BMS CSV, matching directly on Column 2 and ingesting regular + COV timestamps."""
  by_exact_col2: Dict[str, List[MappingRow]] = defaultdict(list)
  for item in mappings:
    by_exact_col2[item.unique_col2_value].append(item)

  timeseries_by_proxy: Dict[str, Dict[str, Dict[str, str]]] = defaultdict(
      lambda: defaultdict(dict)
  )
  matched_fields_by_proxy: Dict[str, List[str]] = defaultdict(list)
  matched_fields_set: Dict[str, set] = defaultdict(set)
  broken_timestamps: Dict[Tuple[str, str, str, str], int] = defaultdict(int)

  ts_cache: Dict[str, Tuple[Optional[str], bool, Optional[str]]] = {}

  total_bytes = bms_csv_path.stat().st_size
  last_update = 0.0
  total_rows = 0
  matched_bms_count = 0

  with bms_csv_path.open("rb") as raw_f:
    text_f = io.TextIOWrapper(raw_f, encoding="utf-8-sig", newline="")
    reader = csv.reader(text_f)
    for row in reader:
      if not row:
        continue
      if len(row) == 1 and "\t" in row[0]:
        row = row[0].split("\t")
      if len(row) < 5:
        continue

      total_rows += 1
      col2_path = row[1].strip()

      if total_rows & 0x3FFF == 0:
        now = time.monotonic()
        if now - last_update >= 0.15:
          mb_done = raw_f.tell() / (1024 * 1024)
          mb_total = total_bytes / (1024 * 1024)
          render_progress_bar(
              raw_f.tell(),
              total_bytes,
              f"{mb_done:.1f}/{mb_total:.1f} MB",
              f"{total_rows:,} rows ({matched_bms_count:,} matched)",
          )
          last_update = now

      matched_rows = by_exact_col2.get(col2_path)
      if not matched_rows:
        continue

      raw_ts = row[3].strip()
      local_ts, is_regular, ts_err = convert_and_classify_seattle_tz(
          raw_ts, ts_cache
      )

      if local_ts is None:
        reason = ts_err or "Invalid timestamp"
        for matched in matched_rows:
          broken_timestamps[
              (matched.proxy_id, matched.standard_field_name, raw_ts, reason)
          ] += 1
        continue

      if not is_regular and ts_err:
        for matched in matched_rows:
          broken_timestamps[
              (matched.proxy_id, matched.standard_field_name, raw_ts, ts_err)
          ] += 1

      col5_value = row[4].strip()
      matched_bms_count += 1
      for matched in matched_rows:
        matched.matched_bms_rows += 1
        p_id = matched.proxy_id
        field = matched.standard_field_name
        timeseries_by_proxy[p_id][local_ts][field] = col5_value
        if field not in matched_fields_set[p_id]:
          matched_fields_set[p_id].add(field)
          matched_fields_by_proxy[p_id].append(field)

  mb_total = total_bytes / (1024 * 1024)
  render_progress_bar(
      total_bytes,
      total_bytes,
      f"{mb_total:.1f}/{mb_total:.1f} MB",
      f"{total_rows:,} rows ({matched_bms_count:,} matched)",
  )
  print()

  return timeseries_by_proxy, matched_fields_by_proxy, broken_timestamps


def build_populate_cli_command(
    bug_number: str,
    data_file_abs: Path,
    device_num_id: str,
) -> str:
  """Formats the backfill populate CLI command string for a proxy_id."""
  normalized_bug = bug_number.strip()
  if normalized_bug and not normalized_bug.startswith("b/"):
    normalized_bug = f"b/{normalized_bug}"

  return (
      f'admin_session --reason="{normalized_bug}" -- \\\n'
      "blaze run \\\n"
      "java/com/google/corp/bizapps/rews/datalake/tools/backfill:backfill_tool"
      " -- \\\n"
      '--mode="populate" \\\n'
      f'--data_file="{data_file_abs}" \\\n'
      f"--device_num_id={device_num_id.strip()} \\\n"
      '--data_field_name="points" \\\n'
      '--present_value_field_name="present_value" \\\n'
      "--environment=prod\n"
  )


def build_delete_cli_command(
    bug_number: str,
    data_file_abs: Path,
    device_num_id: str,
) -> str:
  """Formats the backfill delete CLI command string for a proxy_id."""
  normalized_bug = bug_number.strip()
  if normalized_bug and not normalized_bug.startswith("b/"):
    normalized_bug = f"b/{normalized_bug}"

  return (
      f'admin_session --reason="{normalized_bug}" -- \\\n'
      "blaze run \\\n"
      "java/com/google/corp/bizapps/rews/datalake/tools/backfill:backfill_tool"
      " -- \\\n"
      '--mode="delete" \\\n'
      f'--data_file="{data_file_abs}" \\\n'
      f"--device_num_id={device_num_id.strip()} \\\n"
      "--environment=prod\n"
  )


def write_populate_proxy_outputs(
    root_dir: Path,
    building: str,
    proxy_id: str,
    matched_fields: List[str],
    ts_data: Optional[Dict[str, Dict[str, str]]],
    bug_number: str,
    device_num_id: str,
    write_timeseries_csv: bool = True,
) -> Tuple[Path, Path]:
  """Writes <building>_<proxy_id>.csv (COV forward-filled, irregular timestamps omitted) and _cli_command.txt."""
  proxy_dir = (root_dir / "proxy_ids" / proxy_id).resolve()
  proxy_dir.mkdir(parents=True, exist_ok=True)

  file_stem = f"{building}_{proxy_id}"
  data_csv_path = (proxy_dir / f"{file_stem}.csv").resolve()
  cli_txt_path = (proxy_dir / f"{file_stem}_cli_command.txt").resolve()

  if write_timeseries_csv and ts_data is not None:
    last_seen: Dict[str, str] = {}
    with data_csv_path.open("w", encoding="utf-8", newline="") as f:
      writer = csv.writer(f, quoting=csv.QUOTE_ALL)
      writer.writerow(["timestamp", *matched_fields])
      for local_ts in sorted(ts_data.keys()):
        point_values = ts_data[local_ts]
        for field in matched_fields:
          val = point_values.get(field, "")
          if val != "":
            last_seen[field] = val

        if not VALID_TIMESTAMP_REGEX.match(local_ts):
          continue

        row = [local_ts] + [
            last_seen.get(field, "") for field in matched_fields
        ]
        writer.writerow(row)

  cli_content = build_populate_cli_command(
      bug_number=bug_number,
      data_file_abs=data_csv_path,
      device_num_id=device_num_id,
  )
  cli_txt_path.write_text(cli_content, encoding="utf-8")

  return proxy_dir, cli_txt_path


def write_delete_proxy_outputs(
    root_dir: Path,
    building: str,
    proxy_id: str,
    start_ts: str,
    end_ts: str,
    bug_number: str,
    device_num_id: str,
) -> Tuple[Path, Path]:
  """Writes the 3-line delete <building>_<proxy_id>.csv and delete _cli_command.txt."""
  proxy_dir = (root_dir / "proxy_ids" / proxy_id).resolve()
  proxy_dir.mkdir(parents=True, exist_ok=True)

  file_stem = f"{building}_{proxy_id}"
  data_csv_path = (proxy_dir / f"{file_stem}.csv").resolve()
  cli_txt_path = (proxy_dir / f"{file_stem}_cli_command.txt").resolve()

  with data_csv_path.open("w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Timestamp"])
    writer.writerow([start_ts])
    writer.writerow([end_ts])

  cli_content = build_delete_cli_command(
      bug_number=bug_number,
      data_file_abs=data_csv_path,
      device_num_id=device_num_id,
  )
  cli_txt_path.write_text(cli_content, encoding="utf-8")

  return proxy_dir, cli_txt_path


def extract_bt_delete_commands(output_text: str) -> List[str]:
  """Extracts 'bt delete /bigtable/...' commands from BackfillTool delete output."""
  matches = BT_DELETE_REGEX.findall(output_text)
  cleaned: List[str] = []
  for m in matches:
    cmd = m.strip()
    if cmd not in cleaned:
      cleaned.append(cmd)
  return cleaned


def write_delete_summary_file(
    root_dir: Path, cli_results: List[Dict[str, str]]
) -> Path:
  """Writes all extracted 'bt delete' commands to delete_commands_summary.txt for easy bug pasting."""
  summary_path = (root_dir / "delete_commands_summary.txt").resolve()
  lines: List[str] = [
      "# Carson Backfill Delete Commands Summary",
      f"# Generated: {datetime.now(SEATTLE_TZ).isoformat(sep=' ', timespec='seconds')}",
      "",
  ]
  for res in cli_results:
    proxy_id = res["proxy_id"]
    dev_id = res["device_num_id"]
    bt_cmds = res.get("bt_delete_commands", "").strip()
    lines.append(f"## {proxy_id} (device_num_id: {dev_id}) - {res['status']}")
    if bt_cmds:
      lines.append(bt_cmds)
    else:
      lines.append("# No 'bt delete' commands found in output.")
    lines.append("")

  summary_path.write_text("\n".join(lines), encoding="utf-8")
  return summary_path


def write_matching_and_cli_report(
    report_path: Path,
    mappings: Optional[List[MappingRow]],
    broken_timestamps: Optional[Dict[Tuple[str, str, str, str], int]] = None,
    cli_results: Optional[List[Dict[str, str]]] = None,
) -> None:
  """Writes or updates backfill_report.xlsx with Mapping Key Matching, broken timestamps, and CLI Execution Results tabs."""
  header_font = Font(bold=True, color="FFFFFF")
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  ok_fill = PatternFill(
      start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
  )
  fail_fill = PatternFill(
      start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
  )

  if mappings is None and report_path.is_file():
    wb = openpyxl.load_workbook(report_path)
  else:
    wb = openpyxl.Workbook()

  # Tab 1: Mapping Key Matching
  if mappings is not None:
    if "Mapping Key Matching" in wb.sheetnames:
      ws_match = wb["Mapping Key Matching"]
      ws_match.delete_rows(1, ws_match.max_row)
    else:
      ws_match = wb.active
      ws_match.title = "Mapping Key Matching"

    match_headers = [
        "Row",
        "Status",
        "Matched BMS Row Count",
        "Building",
        "proxyId",
        "unique_col2_value",
        "standardFieldName",
    ]
    ws_match.append(match_headers)
    for cell in ws_match[1]:
      cell.font = header_font
      cell.fill = header_fill

    for m in mappings:
      is_matched = m.matched_bms_rows > 0
      status = "MATCHED" if is_matched else "UNMATCHED"
      ws_match.append([
          m.row_number,
          status,
          m.matched_bms_rows,
          m.building,
          m.proxy_id,
          m.unique_col2_value,
          m.standard_field_name,
      ])
      status_cell = ws_match.cell(row=ws_match.max_row, column=2)
      status_cell.fill = ok_fill if is_matched else fail_fill

  # Tab 2: broken timestamps
  if broken_timestamps is not None:
    if "broken timestamps" in wb.sheetnames:
      ws_broken = wb["broken timestamps"]
      ws_broken.delete_rows(1, ws_broken.max_row)
    else:
      ws_broken = wb.create_sheet(title="broken timestamps")

    broken_headers = [
        "proxyId",
        "standardFieldName",
        "Original Timestamp",
        "Reason",
        "Occurrences",
    ]
    ws_broken.append(broken_headers)
    for cell in ws_broken[1]:
      cell.font = header_font
      cell.fill = header_fill

    for (p_id, field_name, orig_ts, reason), count in sorted(
        broken_timestamps.items()
    ):
      ws_broken.append([p_id, field_name, orig_ts, reason, count])

  # Tab 3: CLI Execution Results
  if "CLI Execution Results" in wb.sheetnames:
    ws_cli = wb["CLI Execution Results"]
    ws_cli.delete_rows(1, ws_cli.max_row)
  elif mappings is None and wb.active.title == "Sheet":
    ws_cli = wb.active
    ws_cli.title = "CLI Execution Results"
  else:
    ws_cli = wb.create_sheet(title="CLI Execution Results")

  cli_headers = [
      "proxyId",
      "device_num_id",
      "Status",
      "Return Code",
      "CLI Command File",
      "Error Log File",
      "Generated bt delete Commands",
      "Command Output / Error Summary",
  ]
  ws_cli.append(cli_headers)
  for cell in ws_cli[1]:
    cell.font = header_font
    cell.fill = header_fill

  if cli_results:
    for res in cli_results:
      ws_cli.append([
          res["proxy_id"],
          res["device_num_id"],
          res["status"],
          res["return_code"],
          res["cli_file"],
          res.get("error_log_file", ""),
          res.get("bt_delete_commands", ""),
          res["output"],
      ])
      status_cell = ws_cli.cell(row=ws_cli.max_row, column=3)
      status_cell.fill = (
          ok_fill if res["status"] == "SUCCEEDED" else fail_fill
      )

  report_path.parent.mkdir(parents=True, exist_ok=True)
  wb.save(report_path)


def discover_existing_cli_files(
    root_dir: Path,
) -> Tuple[List[str], Dict[str, str], Dict[str, Path]]:
  """Discovers existing *_cli_command.txt files under <root>/proxy_ids."""
  proxy_root = root_dir / "proxy_ids"
  if not proxy_root.is_dir():
    sys.exit(
        f"Error: Directory '{proxy_root}' not found. Generate the CLI files"
        " first."
    )

  unique_proxy_ids: List[str] = []
  device_num_ids: Dict[str, str] = {}
  cli_files: Dict[str, Path] = {}

  for proxy_dir in sorted(p for p in proxy_root.iterdir() if p.is_dir()):
    proxy_id = proxy_dir.name
    candidates = sorted(proxy_dir.glob("*_cli_command.txt"))
    if not candidates:
      continue
    cli_path = candidates[-1].resolve()
    cmd_text = cli_path.read_text(encoding="utf-8")
    m = re.search(r"--device_num_id=(\S+)", cmd_text)
    dev_id = m.group(1).strip() if m else "UNKNOWN"

    unique_proxy_ids.append(proxy_id)
    device_num_ids[proxy_id] = dev_id
    cli_files[proxy_id] = cli_path

  if not unique_proxy_ids:
    sys.exit(f"Error: No '*_cli_command.txt' files found under '{proxy_root}'.")
  return unique_proxy_ids, device_num_ids, cli_files


def write_error_log(cli_path: Path, cmd_text: str, full_output: str) -> Path:
  """Writes the full terminal output for a failed CLI execution into the device's folder."""
  base_stem = re.sub(r"_cli_command$", "", cli_path.stem)
  log_path = (cli_path.parent / f"{base_stem}_error_log.txt").resolve()
  log_body = (
      f"Command:\n{cmd_text}\n\n"
      + "=" * 72
      + f"\nFull Terminal Output:\n{full_output}\n"
  )
  log_path.write_text(log_body, encoding="utf-8")
  return log_path


def run_interactive_pty_command(cmd_text: str, cwd: Path) -> Tuple[int, str]:
  """Runs a command inside a pseudo-terminal (PTY) so interactive prompts (admin_session/blaze) stream live without buffering."""
  master_fd, slave_fd = pty.openpty()
  proc = subprocess.Popen(
      cmd_text,
      shell=True,
      executable="/bin/bash",
      cwd=cwd,
      stdin=slave_fd,
      stdout=slave_fd,
      stderr=slave_fd,
      close_fds=True,
  )
  os.close(slave_fd)

  captured: List[bytes] = []
  stdin_fd: Optional[int] = None
  try:
    if sys.stdin.isatty():
      stdin_fd = sys.stdin.fileno()
  except Exception:
    stdin_fd = None

  try:
    while True:
      watch_fds = [master_fd]
      if stdin_fd is not None:
        watch_fds.append(stdin_fd)
      rlist, _, _ = select.select(watch_fds, [], [], 0.1)
      if master_fd in rlist:
        try:
          data = os.read(master_fd, 4096)
        except OSError:
          break
        if not data:
          break
        os.write(sys.stdout.fileno(), data)
        captured.append(data)
      if stdin_fd is not None and stdin_fd in rlist:
        try:
          user_bytes = os.read(stdin_fd, 1024)
          if user_bytes:
            os.write(master_fd, user_bytes)
        except OSError:
          stdin_fd = None
      if proc.poll() is not None and not rlist:
        break
  finally:
    os.close(master_fd)

  proc.wait()
  output_text = b"".join(captured).decode("utf-8", errors="replace").strip()
  return proc.returncode, output_text


def execute_backfill_commands(
    unique_proxy_ids: List[str],
    device_num_ids: Dict[str, str],
    cli_files: Dict[str, Path],
    step_label: str = "Step 6/6",
) -> List[Dict[str, str]]:
  """Runs `cd "$(p4 g4d backfill)" && g4 sync && pwd` once, then executes each proxyId's CLI command."""
  results: List[Dict[str, str]] = []

  print(
      f'\n[{step_label}] Running initialization: cd "$(p4 g4d backfill)" && g4'
      " sync && pwd"
  )
  init_proc = subprocess.run(
      'cd "$(p4 g4d backfill)" && g4 sync >&2 && pwd',
      shell=True,
      executable="/bin/bash",
      text=True,
      stdout=subprocess.PIPE,
  )

  if init_proc.returncode != 0:
    err_msg = (
        'Initialization \'cd "$(p4 g4d backfill)" && g4 sync && pwd\' failed'
        f" with exit code {init_proc.returncode}."
    )
    print(f"  Error: {err_msg}", file=sys.stderr)
    for proxy_id in unique_proxy_ids:
      cli_path = cli_files[proxy_id]
      log_path = write_error_log(
          cli_path, 'cd "$(p4 g4d backfill)" && g4 sync && pwd', err_msg
      )
      results.append({
          "proxy_id": proxy_id,
          "device_num_id": device_num_ids[proxy_id],
          "status": "FAILED (INIT)",
          "return_code": str(init_proc.returncode),
          "cli_file": str(cli_path),
          "error_log_file": str(log_path),
          "bt_delete_commands": "",
          "output": err_msg,
      })
    return results

  lines = [
      line.strip()
      for line in (init_proc.stdout or "").splitlines()
      if line.strip()
  ]
  if not lines:
    err_msg = "Initialization succeeded but did not return a working directory."
    print(f"  Error: {err_msg}", file=sys.stderr)
    for proxy_id in unique_proxy_ids:
      cli_path = cli_files[proxy_id]
      log_path = write_error_log(
          cli_path, 'cd "$(p4 g4d backfill)" && g4 sync && pwd', err_msg
      )
      results.append({
          "proxy_id": proxy_id,
          "device_num_id": device_num_ids[proxy_id],
          "status": "FAILED (INIT)",
          "return_code": "1",
          "cli_file": str(cli_path),
          "error_log_file": str(log_path),
          "bt_delete_commands": "",
          "output": err_msg,
      })
    return results

  g4d_dir = Path(lines[-1])
  print(f"  Synced workspace directory: {g4d_dir}\n")

  total = len(unique_proxy_ids)
  for idx, proxy_id in enumerate(unique_proxy_ids, start=1):
    cli_path = cli_files[proxy_id]
    cmd_text = cli_path.read_text(encoding="utf-8").strip()

    data_csv_match = re.search(r'--data_file="([^"]+)"', cmd_text)
    if data_csv_match:
      data_csv_path = Path(data_csv_match.group(1))
      if data_csv_path.is_file():
        with data_csv_path.open("r", encoding="utf-8", errors="ignore") as df:
          has_header = bool(df.readline())
          has_data = bool(df.readline())
        if has_header and not has_data:
          msg = (
              f"Skipped '{proxy_id}': data CSV '{data_csv_path.name}' has 0"
              " matched data rows."
          )
          print(f"--- [{idx}/{total}] {msg} ---\n")
          results.append({
              "proxy_id": proxy_id,
              "device_num_id": device_num_ids[proxy_id],
              "status": "SKIPPED (EMPTY CSV)",
              "return_code": "N/A",
              "cli_file": str(cli_path),
              "error_log_file": "",
              "bt_delete_commands": "",
              "output": msg,
          })
          continue

    print(f"--- [{idx}/{total}] Running command for proxyId '{proxy_id}' ---")
    print(f"Command:\n{cmd_text}\n", flush=True)

    returncode, output_text = run_interactive_pty_command(cmd_text, cwd=g4d_dir)
    print()

    bt_cmds = "\n".join(extract_bt_delete_commands(output_text))

    if returncode == 0:
      status = "SUCCEEDED"
      error_log_str = ""
    else:
      status = "FAILED"
      log_path = write_error_log(cli_path, cmd_text, output_text)
      error_log_str = str(log_path)
      print(f"  Full error log saved to: {log_path}")

    print(
        f"--- [{idx}/{total}] {proxy_id}: {status} (exit code"
        f" {returncode}) ---\n"
    )

    trimmed_output = (
        output_text[-2000:] if len(output_text) > 2000 else output_text
    )
    results.append({
        "proxy_id": proxy_id,
        "device_num_id": device_num_ids[proxy_id],
        "status": status,
        "return_code": str(returncode),
        "cli_file": str(cli_path),
        "error_log_file": error_log_str,
        "bt_delete_commands": bt_cmds,
        "output": trimmed_output,
    })

  return results


def prompt_non_empty(prompt_text: str) -> str:
  """Prompts the user until a non-empty string is entered."""
  while True:
    val = input(prompt_text).strip()
    if val:
      return val
    print("Input cannot be blank. Please try again.")


def prompt_operation_mode() -> str:
  """Prompts the user to choose Populate mode or Delete mode."""
  print(
      "Select Backfill Tool operation mode:\n"
      "  1. Populate Mode (--mode=\"populate\": extract BMS CSV timeseries &"
      " backfill)\n"
      "  2. Delete Mode   (--mode=\"delete\": generate & run delete commands"
      " for a start/end window)"
  )
  while True:
    raw = input("\nEnter mode (1 or 2) [default: 1]: ").strip()
    if not raw or raw == "1":
      return "populate"
    if raw == "2":
      return "delete"
    print("Please enter 1 or 2.")


def prompt_start_step(description: str, max_step: int) -> int:
  """Prompts the user for which workflow step (1..max_step) to start from."""
  print("\n" + description.strip())
  while True:
    raw = input(
        f"\nEnter step number to start from (1-{max_step}) [default: 1]: "
    ).strip()
    if not raw:
      return 1
    if raw.isdigit() and 1 <= int(raw) <= max_step:
      return int(raw)
    print(f"Please enter a valid number between 1 and {max_step}.")


def run_delete_workflow(root_dir: Path, report_path: Path) -> None:
  """Executes the Delete Mode workflow."""
  start_step = prompt_start_step(DELETE_STEPS_DESCRIPTION, max_step=3)

  unique_proxy_ids: List[str] = []
  device_num_ids: Dict[str, str] = {}
  cli_files: Dict[str, Path] = {}

  if start_step <= 2:
    mapping_input = prompt_non_empty(
        "Enter path to proxy_id <-> external_id mappings CSV (.csv): "
    )
    mapping_csv_path = Path(mapping_input).expanduser().resolve()
    if not mapping_csv_path.is_file():
      sys.exit(f"Error: Mapping CSV file not found at '{mapping_csv_path}'.")

    print("\n[Step 1/3] Reading proxy_id <-> external_id mappings CSV...")
    device_num_ids = read_device_num_id_mappings(mapping_csv_path)
    if not device_num_ids:
      sys.exit("Error: No valid proxy_id -> device_num_id rows found in CSV.")
    unique_proxy_ids = list(device_num_ids.keys())
    print(f"  Loaded {len(unique_proxy_ids):,} proxyId(s).")

    building_name = prompt_non_empty(
        "Enter building name (e.g., US-MTV-L441): "
    )
    start_ts = prompt_delete_timestamp(
        "Enter delete Start Timestamp (e.g., 2025-03-01 00:00:00-08:00): "
    )
    end_ts = prompt_delete_timestamp(
        "Enter delete End Timestamp   (e.g., 2025-05-06 16:00:00-07:00): "
    )
    bug_number = prompt_non_empty("Enter bug number (e.g., b/580954604): ")

    print("\n[Step 2/3] Generating per-proxyId delete CSVs and CLI files...")
    total_proxies = len(unique_proxy_ids)
    for idx, proxy_id in enumerate(unique_proxy_ids, start=1):
      _, cli_txt_path = write_delete_proxy_outputs(
          root_dir=root_dir,
          building=building_name,
          proxy_id=proxy_id,
          start_ts=start_ts,
          end_ts=end_ts,
          bug_number=bug_number,
          device_num_id=device_num_ids[proxy_id],
      )
      cli_files[proxy_id] = cli_txt_path
      render_progress_bar(
          idx,
          total_proxies,
          f"{idx}/{total_proxies} delete files written",
          proxy_id,
      )
    print()

    proceed = input(
        "\nMake sure 'gcert' is active. Press Enter to run the delete CLI"
        " commands (or type 'skip' to exit): "
    ).strip()
    if proceed.lower() == "skip":
      print(f"\nSkipped CLI execution. Files written under: {root_dir / 'proxy_ids'}")
      return
  else:
    unique_proxy_ids, device_num_ids, cli_files = discover_existing_cli_files(
        root_dir
    )

  cli_results = execute_backfill_commands(
      unique_proxy_ids=unique_proxy_ids,
      device_num_ids=device_num_ids,
      cli_files=cli_files,
      step_label="Step 3/3",
  )

  summary_path = write_delete_summary_file(root_dir, cli_results)
  write_matching_and_cli_report(
      report_path,
      mappings=None,
      broken_timestamps=None,
      cli_results=cli_results,
  )

  succeeded = sum(1 for r in cli_results if r["status"] == "SUCCEEDED")
  failed = sum(1 for r in cli_results if r["status"].startswith("FAILED"))
  print(
      f"\nDone! Delete CLI commands succeeded: {succeeded}, failed: {failed}.\n"
      f"Extracted 'bt delete' commands summary (ready for Carson bug):\n"
      f"  {summary_path}\n"
      f"Excel report updated at:\n  {report_path}"
  )


def run_populate_workflow(root_dir: Path, report_path: Path) -> None:
  """Executes the Populate Mode workflow (Steps 1-6)."""
  start_step = prompt_start_step(POPULATE_STEPS_DESCRIPTION, max_step=6)

  mappings: Optional[List[MappingRow]] = None
  device_num_ids: Dict[str, str] = {}
  unique_proxy_ids: List[str] = []
  building_by_proxy: Dict[str, str] = {}
  timeseries_by_proxy: Dict[str, Dict[str, Dict[str, str]]] = {}
  matched_fields_by_proxy: Dict[str, List[str]] = {}
  broken_timestamps: Optional[Dict[Tuple[str, str, str, str], int]] = None
  cli_files: Dict[str, Path] = {}

  if start_step <= 5:
    key_input = prompt_non_empty(
        "Enter path to BMS CSV standardFieldName mapping key (.xlsx): "
    )
    key_path = Path(key_input).expanduser().resolve()
    if not key_path.is_file():
      sys.exit(f"Error: Mapping key file not found at '{key_path}'.")
    building_name = prompt_non_empty(
        "Enter building name (e.g., US-MTV-L441): "
    )

    print("\n[Step 1/6] Reading BMS CSV standardFieldName mapping key...")
    mappings = read_mapping_key_xlsx(key_path, building_name)
    if not mappings:
      sys.exit(
          "No valid rows with a non-blank proxyId and standardFieldName found."
      )
    print(f"  Loaded {len(mappings):,} mapping row(s).")

    unique_proxy_ids = list(dict.fromkeys(m.proxy_id for m in mappings))
    for m in mappings:
      if m.proxy_id not in building_by_proxy and m.building:
        building_by_proxy[m.proxy_id] = m.building

  if start_step <= 4:
    mapping_input = prompt_non_empty(
        "Enter path to proxy_id <-> external_id mappings CSV (.csv): "
    )
    mapping_csv_path = Path(mapping_input).expanduser().resolve()
    if not mapping_csv_path.is_file():
      sys.exit(f"Error: Mapping CSV file not found at '{mapping_csv_path}'.")

    print("\n[Step 2/6] Reading proxy_id <-> external_id mappings CSV...")
    device_num_ids = read_device_num_id_mappings(mapping_csv_path)
    print(f"  Found {len(unique_proxy_ids):,} unique proxyId(s).")
    for proxy_id in unique_proxy_ids:
      if proxy_id not in device_num_ids:
        device_num_ids[proxy_id] = prompt_non_empty(
            f"  proxyId '{proxy_id}' not found in mapping CSV. Enter"
            " device_num_id: "
        )

  if start_step <= 3:
    bms_input = prompt_non_empty("Enter path to BMS CSV export (.csv): ")
    bms_csv_path = Path(bms_input).expanduser().resolve()
    if not bms_csv_path.is_file():
      sys.exit(f"Error: BMS CSV file not found at '{bms_csv_path}'.")

    print("\n[Step 3/6] Matching BMS timeseries data...")
    assert mappings is not None
    timeseries_by_proxy, matched_fields_by_proxy, broken_timestamps = (
        process_bms_csv(bms_csv_path, mappings)
    )
    matched_count = sum(1 for m in mappings if m.matched_bms_rows > 0)
    unmatched_count = len(mappings) - matched_count
    broken_total = sum(broken_timestamps.values())
    print(
        f"  Mapping rows matched: {matched_count:,} /"
        f" {len(mappings):,} ({unmatched_count:,} unmatched)."
    )
    if broken_total:
      print(
          f"  Irregular/broken timestamps omitted from CSVs: {broken_total:,}"
          f" row(s) ({len(broken_timestamps):,} unique) -> logged to 'broken"
          " timestamps' tab."
      )

  if start_step <= 4:
    bug_number = prompt_non_empty("Enter bug number (e.g., b/580954604): ")
    print("\n[Step 4/6] Generating per-proxyId CSVs and CLI command files...")
    assert mappings is not None
    total_proxies = len(unique_proxy_ids)
    write_ts = start_step <= 3
    for idx, proxy_id in enumerate(unique_proxy_ids, start=1):
      matched_fields = matched_fields_by_proxy.get(proxy_id, [])
      ts_data = timeseries_by_proxy.get(proxy_id, {}) if write_ts else None
      if not matched_fields:
        matched_fields = list(
            dict.fromkeys(
                m.standard_field_name
                for m in mappings
                if m.proxy_id == proxy_id
            )
        )

      _, cli_txt_path = write_populate_proxy_outputs(
          root_dir=root_dir,
          building=building_by_proxy[proxy_id],
          proxy_id=proxy_id,
          matched_fields=matched_fields,
          ts_data=ts_data,
          bug_number=bug_number,
          device_num_id=device_num_ids[proxy_id],
          write_timeseries_csv=write_ts,
      )
      cli_files[proxy_id] = cli_txt_path
      render_progress_bar(
          idx,
          total_proxies,
          f"{idx}/{total_proxies} proxyIds written",
          proxy_id,
      )
    print()

  if start_step <= 5:
    if start_step <= 3 or not report_path.is_file():
      print(f"\n[Step 5/6] Writing matching report to '{report_path}'...")
      write_matching_and_cli_report(
          report_path,
          mappings,
          broken_timestamps=broken_timestamps,
          cli_results=None,
      )
      print("  Matching report saved.")
    else:
      print(
          f"\n[Step 5/6] Keeping existing Mapping Key Matching report at"
          f" '{report_path}'."
      )

  if start_step >= 5:
    unique_proxy_ids, device_num_ids, cli_files = discover_existing_cli_files(
        root_dir
    )

  if start_step < 6:
    proceed = input(
        "\nMake sure 'gcert' is active. Press Enter to run the backfill CLI"
        " commands (or type 'skip' to exit): "
    ).strip()
    if proceed.lower() == "skip":
      print(f"\nSkipped CLI execution. Report available at: {report_path}")
      return

  cli_results = execute_backfill_commands(
      unique_proxy_ids=unique_proxy_ids,
      device_num_ids=device_num_ids,
      cli_files=cli_files,
      step_label="Step 6/6",
  )

  write_matching_and_cli_report(
      report_path,
      mappings if start_step <= 3 else None,
      broken_timestamps=broken_timestamps if start_step <= 3 else None,
      cli_results=cli_results,
  )
  succeeded = sum(1 for r in cli_results if r["status"] == "SUCCEEDED")
  failed = sum(1 for r in cli_results if r["status"].startswith("FAILED"))
  skipped = len(cli_results) - succeeded - failed
  print(
      f"\nDone! CLI commands succeeded: {succeeded}, failed: {failed}, skipped:"
      f" {skipped}.\nFull report saved to:\n  {report_path}"
  )


def main() -> None:
  mode = prompt_operation_mode()

  root_input = input(
      "Enter output root directory [default: ./root]: "
  ).strip()
  root_dir = (
      Path(root_input).expanduser().resolve()
      if root_input
      else (Path.cwd() / "root").resolve()
  )
  report_path = (root_dir / "backfill_report.xlsx").resolve()

  if mode == "delete":
    run_delete_workflow(root_dir, report_path)
  else:
    run_populate_workflow(root_dir, report_path)


if __name__ == "__main__":
  main()