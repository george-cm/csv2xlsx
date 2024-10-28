"""Convert CSV file to XLSX file."""

import csv
import re
import sys
from pathlib import Path
from typing import Optional

import xlsxwriter
from chardet.universaldetector import UniversalDetector

ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
EXCEL_ROW_LIMIT = 1_048_576


class IllegalCharacterError(Exception):
    """The data submitted which cannot be used directly in Excel files. It
    must be removed or escaped."""


def detect_file_encoding(fpath: Path) -> str | None:
    """Detect the encoding of a file.

    Args: fpath: Path to the file to detect encoding for.
    Returns: The detected encoding if it could be determined, otherwise None.
    """
    detector = UniversalDetector()
    with fpath.open("rb") as fin:
        for line in fin:
            detector.feed(line)
            if detector.done:
                break
        detector.close()
    print(detector.result)
    return detector.result["encoding"]


# TODO: Add a function to parse command line arguments using argparse
# One argument will determine whether to autodetect encoding or not.
# Right now the default is to use utf-8. Autodetection was turned off
# on 2024/01/26 because it was causing problems with some files
# detecting Japanese encoding instead of utf-8.


def csv2xlsx(fpath: Path, silent: bool = False, detect_encoding: bool = False) -> None:
    """Convert CSV file to XLSX file.

    Args:
        fpath: Path to CSV file
        silent: If true, suppress print statements
    """
    encoding: Optional[str] = "utf-8"
    if detect_encoding:
        encoding = detect_file_encoding(fpath)
        if encoding is None:
            raise IOError("Could not detect CSV file encoding.")
    output_file_row_margin = 10
    output_file_row_limit = EXCEL_ROW_LIMIT - output_file_row_margin
    with fpath.open("r", encoding=encoding) as inf:
        if not silent:
            print(f"file: {fpath.as_posix()}")
        out_xlsx = fpath.parent / (fpath.stem + ".xlsx")
        reader = csv.reader(inf)
        wb = xlsxwriter.Workbook(out_xlsx)
        ws_count: int = 1
        ws = wb.add_worksheet(name=f"Sheet{ws_count}")  # type: ignore
        try:
            first_line = next(reader)
        except UnicodeDecodeError as e:
            raise e

        ws.write_row(0, 0, first_line)  # type: ignore
        column_count = len(first_line)
        line_count = 1
        for line in reader:
            if line_count >= output_file_row_limit:
                if not silent:
                    print(f"\trows: {line_count-1}, columns: {column_count}")
                ws.add_table(
                    0,
                    0,
                    line_count - 1,
                    column_count - 1,
                    {"columns": [{"header": x} for x in first_line]},
                )  # type: ignore
                ws_count += 1
                ws = wb.add_worksheet(name=f"Sheet{ws_count}")  # type: ignore
                ws.write_row(0, 0, first_line)
                line_count = 1
            try:
                ws.write_row(line_count, 0, line)  # type: ignore
            except IllegalCharacterError:
                new_line = [ILLEGAL_CHARACTERS_RE.sub(r"", x) for x in line]
                ws.write_row(line_count, 0, new_line)

            line_count += 1
        if not silent:
            print(f"\trows: {line_count-1}, columns: {column_count}")
        ws.add_table(
            0,
            0,
            line_count - 1,
            column_count - 1,
            {"columns": [{"header": x} for x in first_line]},
        )  # type: ignore
        ws.autofit()
        wb.close()  # type: ignore
        # input()


def main():
    """Main function"""
    for f in sys.argv[1:]:
        fpath = Path(f)

        if fpath.is_dir():
            csv_files = fpath.glob("*.csv")
            for csv_file in csv_files:
                csv2xlsx(csv_file, silent=False, detect_encoding=False)
        elif fpath.is_file() and fpath.suffix == ".csv":
            csv2xlsx(fpath, silent=False)
        else:
            curr_dir = Path()
            csv_files = curr_dir.glob(f)
            found_csvs = False
            for csv_file in csv_files:
                if not found_csvs:
                    found_csvs = True
                csv2xlsx(csv_file, silent=False, detect_encoding=False)
            if not found_csvs:
                print(f"Could not find any file using the pattern {f}")


if __name__ == "__main__":
    main()
