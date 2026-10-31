#!/usr/bin/env python
"""List and validate XXXKEYWORD_NAMEXXX placeholders in a PAR template."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import yaml

KEYWORD_PATTERN = re.compile(r"XXX[A-Z][A-Z0-9_]*?XXX", re.IGNORECASE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path, help="Path to the .par template")
    parser.add_argument("--mapping", type=Path, default=Path("config/keyword_mapping.yaml"), help="Keyword mapping YAML")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    template = args.template.read_text(encoding="utf-8")
    mapping = yaml.safe_load(args.mapping.read_text(encoding="utf-8"))["template_keywords"]
    found = sorted({keyword.upper() for keyword in KEYWORD_PATTERN.findall(template)})
    unknown = sorted(set(found) - set(mapping))

    print("Detected keywords:")
    for keyword in found:
        print(f"  {keyword}")
    if unknown:
        print("Unmapped keywords:")
        for keyword in unknown:
            print(f"  {keyword}")
        return 2
    print("All detected keywords have mappings.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
