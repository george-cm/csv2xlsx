# Backlog

## Bugs

- **A CSV field containing only `=` is written to XLSX as the formula result `0`.** XlsxWriter interprets strings beginning with `=` as formulas. Preserve a bare equals sign as the literal string `=` during CSV-to-XLSX conversion. Reproduction file: `C:\Users\George.Murga\projects\leadbot\segmentation_rules\csv\540239_Bookboon.csv`.

- **`xlsx2csv` header read is redundant and confusing.** In the loop at line 43-49, when `i == 0` you call `next(sh.rows)` to read the header, but `sh.rows` is a property that creates a new generator each time - so this reads the first row from a throwaway generator while `row` from the outer loop's generator is the same first row. It works, but `row` should be used directly instead of `next(sh.rows)`.

- **`xlsx2csv` return type mismatch.** The signature says `List[Path | None]` but `written_csvs` is typed as `List[Path]` and only `Path` values are ever appended. The `| None` is dead.

- **`csv2xlsx` catches `IllegalCharacterError` but XlsxWriter never raises it.** `IllegalCharacterError` is defined locally but XlsxWriter raises its own exception type. This `except` block is likely dead code.

- **`main()` in `xlsx2csv.py` does not exit on error.** Lines 58-61 print an error if the file does not exist or is not a file, but then fall through to call `xlsx2csv()` anyway, which will crash with a less helpful traceback.

## Code Quality

- **Debug `print` left in `detect_file_encoding`** at line 34. `print(detector.result)` should be removed or gated behind `silent`.

- **`# input()` on line 108.** Dead commented-out code should be removed.

- **`except UnicodeDecodeError as e: raise e`** at lines 69-70 is a no-op; the exception would propagate anyway.

- **No `argparse` for `csv2xlsx`.** The `xlsx2csv` side already uses argparse, so the `csv2xlsx` entry point is inconsistent: raw `sys.argv` parsing and no `--help`.

- **`typing.Optional` mixed with `X | Y` union syntax.** Line 52 uses `Optional[str]` while other places use `str | None`. Use `X | None` consistently because Python 3.11 or later is required.

## Missing Pieces

- **Empty README.** No usage instructions, installation steps, or examples.

- **No tests.** pytest is configured as a development dependency but no test files exist.

- **No CI configuration.** No GitHub Actions, pre-commit hooks, or similar.

- **`.ruff_cache` not in `.gitignore`.** It should be added.

## Minor

- `output_file_row_margin = 10` is arbitrary and undocumented. Explain why the margin is required.

- `csv2xlsx` auto-formats data as an Excel table, but the CLI cannot disable it.
