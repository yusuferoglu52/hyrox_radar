"""Generate one radar chart per participant from an Excel score sheet.

Usage: python radar.py scores.xlsx --out output/
"""
from __future__ import annotations

import argparse
import sys

from hyrox_radar.loader import InputError
from hyrox_radar.normalize import SCALES
from hyrox_radar.runner import Options, parse_axis_min, run

LANGS = ("tr", "en")


def axis_min_arg(text: str):
    """'auto' or a number from 0 to 95."""
    try:
        return parse_axis_min(text)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from None


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Radar chart per participant, with the group average as an overlay.")
    p.add_argument("xlsx", help="Excel file: first column = name, other columns = stations")
    p.add_argument("--out", default="output", help="output folder (default: output)")
    p.add_argument("--sheet", default=None,
                   help="sheet name, or 1-based sheet number (default: first sheet)")
    p.add_argument("--scale", choices=SCALES, default="pct-of-best",
                   help="pct-of-best: best=100, others as a percentage of the best (default). "
                        "minmax: best=100, worst=0")
    p.add_argument("--axis-min", type=axis_min_arg, default="auto", metavar="N|auto",
                   help="where the radial axis starts. 'auto' (default) picks a floor "
                        "just under the lowest score to spread out close scores; a "
                        "number sets it yourself; 0 turns the zoom off. The chart "
                        "footnote states the floor")
    p.add_argument("--group-col", default=None, metavar="COLUMN",
                   help="header of the category column (gender, age group). A column "
                        "named Cinsiyet, Kategori, Gender or Category is found automatically")
    p.add_argument("--compare", choices=("group", "all"), default="group",
                   help="group: score each participant inside their own category when a "
                        "category column exists (default). all: one group for everyone")
    p.add_argument("--lang", choices=LANGS, default="tr",
                   help="language of the text on the charts (default: tr)")
    p.add_argument("--lower-better", action="append", default=[], metavar="STATION",
                   help="force a station to lower-is-better (repeatable)")
    p.add_argument("--higher-better", action="append", default=[], metavar="STATION",
                   help="force a station to higher-is-better (repeatable)")
    p.add_argument("--dpi", type=int, default=200, help="PNG resolution (default: 200)")
    return p


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:  # never crash on a console that cannot print a Turkish character
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)
    opts = Options(xlsx=args.xlsx, out=args.out, sheet=args.sheet, scale=args.scale,
                   axis_min=args.axis_min, lang=args.lang, group_col=args.group_col,
                   compare=args.compare, lower_better=args.lower_better,
                   higher_better=args.higher_better, dpi=args.dpi)
    try:
        run(opts)
    except InputError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"ERROR: could not write the charts ({exc}).", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
