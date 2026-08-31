#!/usr/bin/env python3
"""Lightweight static checks for exported VBA source files.

This is not a VBA compiler. It catches repository-level errors that are easy to
introduce during modular development: duplicate public UDF names, duplicate
module names, missing Option Explicit, and missing core workflow functions.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

MODULE_RE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.MULTILINE | re.IGNORECASE)
PUBLIC_FUNCTION_RE = re.compile(r'^Public\s+Function\s+([A-Za-z][A-Za-z0-9_]*)\b', re.MULTILINE | re.IGNORECASE)

REQUIRED_UDFS = {
    "annualizedreturn",
    "cumulativereturn",
    "aggregatereturns",
    "modifieddietzreturn",
    "timeweightedreturn",
    "moneyweightedreturn",
    "trackingerror",
    "informationratio",
    "regressionstatistics",
    "maxdrawdown",
    "brinsonfachlerattribution",
    "carinolinkattribution",
    "compositereturn",
}


def main() -> int:
    bas_files = sorted(SRC.glob("*.bas"))
    if not bas_files:
        raise SystemExit("No src/*.bas files found")

    modules: dict[str, Path] = {}
    functions: dict[str, list[Path]] = defaultdict(list)

    for path in bas_files:
        text = path.read_text(encoding="utf-8")
        if "Option Explicit" not in text:
            raise SystemExit(f"{path}: missing Option Explicit")

        module_match = MODULE_RE.search(text)
        if not module_match:
            raise SystemExit(f"{path}: missing Attribute VB_Name")
        module_name = module_match.group(1).lower()
        if module_name in modules:
            raise SystemExit(f"Duplicate VBA module name {module_name}: {modules[module_name]} and {path}")
        modules[module_name] = path

        # ToolkitCore intentionally exposes Public helpers inside Option Private Module.
        if "Option Private Module" in text:
            continue

        for name in PUBLIC_FUNCTION_RE.findall(text):
            functions[name.lower()].append(path)

    duplicates = {name: paths for name, paths in functions.items() if len(paths) > 1}
    if duplicates:
        details = "; ".join(f"{name}: {', '.join(map(str, paths))}" for name, paths in duplicates.items())
        raise SystemExit(f"Duplicate public worksheet functions: {details}")

    missing = sorted(REQUIRED_UDFS - set(functions))
    if missing:
        raise SystemExit("Missing required v2 UDFs: " + ", ".join(missing))

    print(f"Checked {len(bas_files)} modules and {len(functions)} public v2 UDFs: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
