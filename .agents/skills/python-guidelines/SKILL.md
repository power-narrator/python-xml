---
name: python-guidelines
description: Python workflow and coding conventions for this repository. Use when Codex is writing, editing, reviewing, testing, linting, formatting, or type-checking Python code in this project, especially for package-manager usage, project commands, docstring conventions, and pptx XML assumptions.
---

# Python Guidelines

Follow these project-specific rules when working on Python code.

## Use The Project Toolchain

Use `uv` for project commands.

```bash
# Run (dev)
uv run scripts/run.py

# Build app
uv run scripts/build.py

# Build pptx cli
uv run scripts/build.py pptx

# Tests
uv run pytest

# Lint / format
uv run ruff check --select I --fix .
uv run ruff format .

# Typecheck
uv run ty check .
```

## Write Docstrings

Use Google-style docstrings.

Document exceptions only when the function explicitly raises them or the called function's docstring says they are raised.

Do not add a `Returns` entry for `None`.

## Structure Code

Place helper functions before the functions that rely on them.

Keep closely related functions near each other, especially when one is only used by one other function.

## PPTX XML Assumptions

Ancestor tags of the target may be missing.

Assume required sibling tags are present.

## XPath Constants

Keep an XPath as a shared constant only when it is reused across modules, parameterized, represents a non-obvious PPTX structure with domain meaning, or belongs to a small coherent selector family.

Inline an XPath when it is used once, very short and obvious, or tightly coupled to one function's local control flow.
