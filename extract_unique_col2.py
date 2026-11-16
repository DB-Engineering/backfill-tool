#!/usr/bin/env python3
"""Streams a large CSV file, extracts unique values from column 2, and exports them to an .xlsx file."""

import csv
import io
from pathlib import Path
import sys
import time

try:
  import openpyxl
except ImportError:
  sys.exit(
      "Error: 'openpyxl' is required to write .xlsx files. Install it via"
      " 'pip install openpyxl'."
  )


def render_progress(
    bytes_read: int, total_bytes: int, row_count: int, width: int = 35
) -> None:
  """Renders an inline terminal progress bar based on bytes read."""
  if total_bytes <= 0:
    return
  ratio = min(1.0, bytes_read / total_bytes)
  filled = int(width * ratio)
  bar = "#" * filled + "-" * (width - filled)
  mb_read = bytes_read / (1024 * 1024)
  mb_total = total_bytes / (1024 * 1024)
  sys.stdout.write(
      f"\r[{bar}] {ratio * 100:5.1f}% ({mb_read:.1f}/{mb_total:.1f} MB |"
      f" {row_count:,} rows)"
  )
  sys.stdout.flush()


def main() -> None:
  csv_input = input("Enter path to the large CSV file (.csv): ").strip()
  if not csv_input:
    sys.exit("Error: CSV path cannot be blank.")

  csv_path = Path(csv_input).expanduser().resolve()
  if not csv_path.is_file():
    sys.exit(f"Error: File not found at '{csv_path}'.")

  default_out = csv_path.with_name(f"{csv_path.stem}_unique_col2.xlsx")
  out_input = input(
      f"Enter output .xlsx path [default: {default_out}]: "
  ).strip()
  out_path = (
      Path(out_input).expanduser().resolve() if out_input else default_out
  )

  total_bytes = csv_path.stat().st_size
  unique_values = set()
  row_count = 0
  last_update = 0.0

  print(f"Scanning '{csv_path}'...")
  with csv_path.open("rb") as raw_f:
    text_f = io.TextIOWrapper(raw_f, encoding="utf-8-sig", newline="")
    reader = csv.reader(text_f)
    for row in reader:
      if not row:
        continue
      if len(row) == 1 and "\t" in row[0]:
        row = row[0].split("\t")
      if len(row) < 2:
        continue
      row_count += 1
      val = row[1].strip()
      if val:
        unique_values.add(val)

      now = time.monotonic()
      if now - last_update >= 0.1:
        render_progress(raw_f.tell(), total_bytes, row_count)
        last_update = now

  render_progress(total_bytes, total_bytes, row_count)
  print()

  sorted_values = sorted(unique_values)

  wb = openpyxl.Workbook(write_only=True)
  ws = wb.create_sheet(title="Unique Column 2")
  ws.append(["unique_col2_value"])
  for val in sorted_values:
    ws.append([val])

  out_path.parent.mkdir(parents=True, exist_ok=True)
  wb.save(out_path)

  print(
      f"Done! Scanned {row_count:,} rows and exported {len(sorted_values):,}"
      f" unique values to:\n{out_path}"
  )


if __name__ == "__main__":
  main()