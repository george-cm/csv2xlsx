# Project Guidelines

## Workflow

- **TDD always.** Write failing tests before changing any code. Fix the code to make the tests pass.

## Git

- Do NOT include `Co-Authored-By`, `Generated with Devin`, or any Devin-related metadata in commit messages.

## Commands

- `uv sync --all-groups` -- install all dependencies
- `uv run pytest` -- run tests
- `uv run mypy csv2xlsx` -- type check
- `uv run ruff check csv2xlsx` -- lint
