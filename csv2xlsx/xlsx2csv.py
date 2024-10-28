import csv
import sys
from pathlib import Path

from openpyxl import load_workbook


def fix_header_duplicate_fields(header):
    fields = list()
    new_header = list()
    for field in header:
        count = fields.count(field)
        if count == 0:
            new_header.append(field)
        else:
            new_header.append(f"{field}{count}")
        fields.append(field)
    return new_header


def xlsx2csv(xlsx_file: Path):
    wb = load_workbook(xlsx_file.as_posix(), read_only=True)
    for sh_name in wb.sheetnames:
        sh = wb[sh_name]
        csv_file = xlsx_file.parent / f"{xlsx_file.stem}_{sh.title}.csv"
        with csv_file.open("w", newline="", encoding="utf-8") as outf:
            writer = csv.writer(outf, dialect="excel")
            for i, row in enumerate(sh.rows):
                if i == 0:
                    header = [x.value for x in next(sh.rows)]
                    writer.writerow(fix_header_duplicate_fields(header))
                else:
                    line = [x.value for x in row]
                    writer.writerow(line)


def main():
    for xlsx_fpath in sys.argv[1:]:
        xlsx_file = Path(xlsx_fpath)
        xlsx2csv(xlsx_file)


if __name__ == "__main__":
    main()
