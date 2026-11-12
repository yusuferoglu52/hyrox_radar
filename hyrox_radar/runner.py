"""The whole pipeline in one call, shared by the command line and the window app."""
from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .loader import load_scores
from .normalize import SCALE_NOTES, ordered_categories, score_by_category


def parse_axis_min(text: str):
    """'auto' or a number from 0 to 95. Raises ValueError with a short reason."""
    if str(text).strip().lower() == "auto":
        return "auto"
    try:
        value = float(str(text).strip().replace(",", "."))
    except ValueError:
        raise ValueError("use a number from 0 to 95, or 'auto'") from None
    if not 0 <= value <= 95:
        raise ValueError("must be between 0 and 95")
    return value


@dataclass
class Options:
    xlsx: str
    out: str = "output"
    sheet: str | None = None
    scale: str = "pct-of-best"
    axis_min: object = "auto"  # "auto" or a number
    lang: str = "tr"
    group_col: str | None = None
    compare: str = "group"
    lower_better: list[str] = field(default_factory=list)
    higher_better: list[str] = field(default_factory=list)
    dpi: int = 200


@dataclass
class RunResult:
    written: int
    out_dir: Path
    floor: float


def _print_warning(message: str) -> None:
    print(f"WARNING: {message}", file=sys.stderr)


def run(opts: Options, info: Callable[[str], None] = print,
        warn: Callable[[str], None] = _print_warning,
        progress: Callable[[int, int], None] | None = None) -> RunResult:
    """Read the sheet and write the charts.

    info/warn receive the report lines; progress(done, total) is called per chart.
    Raises InputError for problems in the file or options, OSError if a chart
    cannot be written.
    """
    data = load_scores(opts.xlsx, sheet=opts.sheet, lower_better=opts.lower_better,
                       higher_better=opts.higher_better, group_col=opts.group_col)

    # Imported late so input errors are reported without loading matplotlib.
    from .chart import (auto_floor, category_label, chart_width, render_chart,
                        safe_filename, unique_filenames)

    split = data.categories is not None and opts.compare == "group"
    scores, averages = score_by_category(data.values(), data.lower_better(), opts.scale,
                                         data.categories if split else None)

    def label(key: str) -> str:
        return category_label(key, data.category_names, opts.lang)

    info(f"Sheet: {data.sheet}  |  {len(data.participants)} participants, "
         f"{len(data.stations)} stations")
    if data.categories is None:
        info("Category column: none (everyone is scored as one group)")
    else:
        keys = ordered_categories(data.categories)
        counts = ", ".join(f"{label(k)} ({data.categories.count(k)})" for k in keys)
        info(f"Category column: {data.group_col}")
        info(f"Categories: {counts}")
        info("Compare: within each category" if split
             else "Compare: all participants as one group (--compare all)")
        if split:
            for k in keys:
                if data.categories.count(k) == 1:
                    data.warnings.append(f"Category '{label(k)}' has one participant, "
                                         "so every score in it is 100.")
    info(f"Scale: {opts.scale} ({SCALE_NOTES[opts.scale]})")
    lowest = float(scores.min().min())
    if opts.axis_min == "auto":
        floor = auto_floor(lowest)
        info(f"Axis floor: {floor:g} (auto; lowest score is {lowest:.0f}; "
             "use --axis-min 0 to start at 0)")
    else:
        floor = float(opts.axis_min)
        info(f"Axis floor: {floor:g}" + (" (zoom off)" if floor == 0 else ""))
    info(f"Chart language: {opts.lang}")
    info("Stations:")
    for s in data.stations:
        info(f"  - {s.name}: {s.describe()}")
    for w in data.warnings:
        warn(w)
    below = int((scores < floor).sum().sum())
    if below:
        warn(f"{below} score(s) are below the axis floor of {floor:g} and are "
             "drawn at the centre. Lower --axis-min or use 'auto'.")

    out_dir = Path(opts.out)
    stems = unique_filenames(data.participants)
    width = chart_width(data.station_names, opts.lang)

    total = len(data.participants)
    written = 0
    for i, (name, stem) in enumerate(zip(data.participants, stems)):
        row = scores.iloc[i].to_numpy()
        if not (row == row).any():  # all NaN
            warn(f"{name} has no results at any station; no chart written.")
        else:
            raw = [s.display(i) for s in data.stations]
            if split:
                key = data.categories[i]
                category = label(key)
                path = out_dir / safe_filename(category) / f"{stem}.png"
            else:
                key, category = None, None
                path = out_dir / f"{stem}.png"
            render_chart(name, data.station_names, row, averages[key].to_numpy(), raw,
                         opts.scale, path, dpi=opts.dpi, lang=opts.lang, axis_min=floor,
                         width=width, category=category)
            written += 1
            info(f"  wrote {path}")
        if progress:
            progress(i + 1, total)
    info(f"Done: {written} chart(s) in {out_dir.resolve()}")
    return RunResult(written, out_dir.resolve(), floor)
