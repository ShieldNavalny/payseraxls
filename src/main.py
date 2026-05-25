"""
Entry point for the Paysera → Excel converter.

Usage:
    python src/main.py <input.csv> [output.xlsx]

Examples:
    python src/main.py data/statement.csv
    python src/main.py data/statement.csv reports/january.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

# Make sure we can import sibling modules regardless of CWD
sys.path.insert(0, str(Path(__file__).parent))

from paysera_parser import convert  # noqa: E402


def main() -> None:
    args = sys.argv[1:]

    if not args:
        print(
            "Usage: python src/main.py <input.csv> [output.xlsx]\n"
            "\n"
            "  input.csv   – Paysera account statement in CSV format\n"
            "  output.xlsx – optional output path (default: same dir as input, .xlsx)"
        )
        sys.exit(1)

    csv_path = Path(args[0])
    if not csv_path.exists():
        print(f"Error: file not found – {csv_path}")
        sys.exit(2)

    output_path = Path(args[1]) if len(args) > 1 else None
    convert(csv_path, output_path)


if __name__ == "__main__":
    main()
