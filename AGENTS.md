# Project Guidelines

> **Global rules** apply from `C:\Users\George.Murga\.agents\AGENTS.md`.
> Rules below are project-specific and extend or override the global rules.

## Workflow

- **TDD always.** Write failing tests before changing any code. Fix the code to make the tests pass.

## Commands

- `uv sync --all-groups` -- install all dependencies
- `uv run pytest` -- run tests
- `uv run mypy csv2xlsx` -- type check
- `uv run ruff check csv2xlsx` -- lint
