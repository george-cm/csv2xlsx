# Issues & Suggestions

## Bugs

1. **`xlsx2csv` header read is redundant and confusing.** In the loop at line 43-49, when `i == 0` you call `next(sh.rows)` to read the header, but `sh.rows` is a property that creates a new generator each time -- so this reads the first row from a throwaway generator while `row` (from the outer loop's generator) is the same first row. It works, but `row` should be used directly instead of `next(sh.rows)`.

2. **`xlsx2csv` return type mismatch.** The signature says `List[Path | None]` but `written_csvs` is typed as `List[Path]` and only `Path` values are ever appended. The `| None` is dead.

3. **`csv2xlsx` catches `IllegalCharacterError` but xlsxwriter never raises it.** `IllegalCharacterError` is defined locally but xlsxwriter raises its own exception type. This `except` block is likely dead code.

4. **`main()` in `xlsx2csv.py` doesn't exit on error.** Lines 58-61 print an error if the file doesn't exist or isn't a file, but then fall through to call `xlsx2csv()` anyway, which will crash with a less helpful traceback.

## Code Quality

5. **Debug `print` left in `detect_file_encoding`** (line 34) -- `print(detector.result)` should be removed or gated behind `silent`.

6. **`# input()` on line 108** -- dead commented-out code, should be removed.

7. **`except UnicodeDecodeError as e: raise e`** (lines 69-70) -- this is a no-op; the exception would propagate anyway.

8. **No `argparse` for `csv2xlsx`** -- there's a TODO comment about it. The `xlsx2csv` side already uses argparse, so the `csv2xlsx` entry point is inconsistent (raw `sys.argv` parsing, no `--help`).

9. **`typing.Optional` mixed with `X | Y` union syntax** -- line 52 uses `Optional[str]` while other places use `str | None`. Pick one style for consistency (since you target 3.11+, `X | None` is fine everywhere).

## Missing Pieces

10. **Empty README** -- no usage instructions, installation steps, or examples.

11. **No tests** -- pytest is configured as a dev dependency but no test files exist.

12. **No CI configuration** -- no GitHub Actions, pre-commit hooks, or similar.

13. **`.ruff_cache` not in `.gitignore`** -- should be added.

## Minor

14. The `output_file_row_margin = 10` is arbitrary and undocumented -- a comment explaining why would help.

15. `csv2xlsx` auto-formats data as an Excel table, which is nice, but there's no way to disable it from the CLI.
