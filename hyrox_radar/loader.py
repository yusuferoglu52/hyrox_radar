"""Read the score sheet, validate it and detect each station's value type."""
from __future__ import annotations

import datetime as dt
import math
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException


class InputError(Exception):
    """Raised for problems the user can fix in the input file or CLI flags."""


# Text that means "did not attempt". An empty cell is handled separately as blank.
DASHES = {"-", "--", "–", "—"}
DNF_TOKENS = DASHES | {"x", "dnf", "dns", "dq", "n/a", "na"}

# m:ss, mm:ss, mmm:ss or h:mm:ss, optional fractional seconds (dot or comma).
TIME_RE = re.compile(r"^(?:(\d{1,2}):)?(\d{1,3}):([0-5]\d)(?:[.,](\d{1,3}))?$")
NUMBER_RE = re.compile(r"^[+-]?\d+(?:[.,]\d+)?$")

EXCEL_EPOCH = dt.datetime(1899, 12, 30)

BLANK, DNF, TIME, NUMBER, BAD = "blank", "dnf", "time", "number", "bad"

# Headers that mark the category column (compared after tr_fold).
GROUP_HEADERS = {"cinsiyet", "kategori", "gender", "category"}
# Category keys: FEMALE/MALE for the known gender spellings, "" for a blank cell,
# otherwise the folded text itself (so age groups, Pro/Open etc. also work).
FEMALE, MALE, UNSPECIFIED = "F", "M", ""
_FEMALE_WORDS = {"k", "kadın", "kadin", "kadınlar", "kadinlar", "kız", "kiz", "bayan",
                 "f", "female", "w", "woman", "women"}
_MALE_WORDS = {"e", "erkek", "erkekler", "bay", "m", "male", "man", "men"}


def tr_fold(text) -> str:
    """Case-fold with Turkish rules: I -> ı and İ -> i, trimmed, single spaces.

    Plain lower() turns 'KADIN' into 'kadin', which would not match 'Kadın'.
    """
    clean = unicodedata.normalize("NFC", " ".join(str(text).split()))
    return clean.replace("I", "ı").replace("İ", "i").lower()


def category_key(value) -> str:
    """Canonical key for a category cell, so 'K', 'KADIN' and 'Kadın ' are one group."""
    folded = tr_fold(value)
    if folded in _FEMALE_WORDS:
        return FEMALE
    if folded in _MALE_WORDS:
        return MALE
    return folded


@dataclass
class Parsed:
    kind: str
    value: float | None = None  # seconds for TIME, the number for NUMBER
    native_time: bool = False  # came from an Excel time/duration cell, not text
    text: str = ""  # original text for DNF cells
    hmm_format: bool = False  # Excel cell formatted as hours:minutes, no seconds


_FORMAT_NOISE = re.compile(r'"[^"]*"|\\.|\[\$[^\]]*\]')


def format_lacks_seconds(number_format: str | None) -> bool:
    """True for hour-based Excel formats without seconds, such as 'h:mm'.

    Excel gives exactly this format to a hand-typed '4:32', which it stores as
    4 hours 32 minutes.
    """
    fmt = _FORMAT_NOISE.sub("", number_format or "").lower()
    return "h" in fmt and "s" not in fmt


def _whole_minutes(seconds: float) -> bool:
    return round(seconds, 3) % 60 == 0


def parse_time_string(text: str) -> float | None:
    """Return seconds for 'm:ss' / 'h:mm:ss' strings, or None if not a time."""
    m = TIME_RE.match(text.strip())
    if not m:
        return None
    hours, minutes, seconds, frac = m.groups()
    if hours is not None and int(minutes) > 59:
        return None
    total = int(hours or 0) * 3600 + int(minutes) * 60 + int(seconds)
    if frac:
        total += float("0." + frac)
    return float(total)


def parse_cell(value, number_format: str | None = None) -> Parsed:
    """Classify one cell value. number_format is the Excel format of the cell."""
    parsed = _parse_value(value)
    if parsed.native_time:
        parsed.hmm_format = format_lacks_seconds(number_format)
    return parsed


def _parse_value(value) -> Parsed:
    if value is None:
        return Parsed(BLANK)
    if isinstance(value, bool):
        return Parsed(BAD)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return Parsed(BLANK)
        return Parsed(NUMBER, float(value))
    if isinstance(value, dt.timedelta):
        return Parsed(TIME, value.total_seconds(), native_time=True)
    if isinstance(value, dt.datetime):
        # Durations of 24 h or more can come back as dates near the Excel epoch.
        if value.year <= 1900:
            return Parsed(TIME, (value - EXCEL_EPOCH).total_seconds(), native_time=True)
        return Parsed(BAD)
    if isinstance(value, dt.time):
        seconds = (value.hour * 3600 + value.minute * 60 + value.second
                   + value.microsecond / 1e6)
        return Parsed(TIME, float(seconds), native_time=True)
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return Parsed(BLANK)
        if text.casefold() in DNF_TOKENS:
            return Parsed(DNF, text=text.upper())
        seconds = parse_time_string(text)
        if seconds is not None:
            return Parsed(TIME, seconds)
        if NUMBER_RE.match(text):
            return Parsed(NUMBER, float(text.replace(",", ".")))
    return Parsed(BAD)


def format_time(seconds: float) -> str:
    """Seconds -> 'm:ss' (or 'h:mm:ss'), keeping one decimal if present."""
    tenths = int(round(float(seconds) * 10))
    whole, frac = divmod(tenths, 10)
    hours, rest = divmod(whole, 3600)
    minutes, secs = divmod(rest, 60)
    tail = f"{secs:02d}" + (f".{frac}" if frac else "")
    if hours:
        return f"{hours}:{minutes:02d}:{tail}"
    return f"{minutes}:{tail}"


def format_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


@dataclass
class Station:
    name: str
    kind: str  # TIME or NUMBER
    lower_better: bool
    overridden: bool
    cells: list[Parsed]

    def describe(self) -> str:
        kind = "time" if self.kind == TIME else "number"
        direction = "lower is better" if self.lower_better else "higher is better"
        suffix = " [override]" if self.overridden else ""
        return f"{kind}, {direction}{suffix}"

    def display(self, row: int) -> str:
        """Raw value as the participant should read it."""
        cell = self.cells[row]
        if cell.kind == BLANK:
            return "–"
        if cell.kind == DNF:
            return "DNF" if cell.text in DASHES | {"X"} else cell.text
        if self.kind == TIME:
            return format_time(cell.value)
        return format_number(cell.value)


@dataclass
class Scores:
    sheet: str
    participants: list[str]
    stations: list[Station]
    warnings: list[str] = field(default_factory=list)
    group_col: str | None = None  # header of the category column, if any
    categories: list[str] | None = None  # category key per participant
    category_names: dict[str, str] = field(default_factory=dict)  # key -> first spelling

    @property
    def station_names(self) -> list[str]:
        return [s.name for s in self.stations]

    def values(self) -> pd.DataFrame:
        """Raw numeric values (seconds for times); NaN where there is no result."""
        data = {
            s.name: [c.value if c.kind in (TIME, NUMBER) else float("nan") for c in s.cells]
            for s in self.stations
        }
        return pd.DataFrame(data, columns=self.station_names, dtype=float)

    def lower_better(self) -> dict[str, bool]:
        return {s.name: s.lower_better for s in self.stations}


def _is_empty(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _open_sheet(path: Path, sheet: str | None):
    if not path.exists():
        raise InputError(f"File not found: {path}")
    if path.suffix.lower() == ".xls":
        raise InputError("Old .xls files are not supported. Re-save the file as .xlsx in Excel.")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise InputError(f"Expected an .xlsx file, got '{path.name}'.")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, OSError) as exc:
        raise InputError(f"Could not read '{path.name}' as an Excel file ({exc}).") from exc

    names = wb.sheetnames
    if sheet is None:
        return wb[names[0]]
    if sheet in names:
        return wb[sheet]
    if sheet.isdigit():
        index = int(sheet)
        if 1 <= index <= len(names):
            return wb[names[index - 1]]
        raise InputError(f"--sheet {index} is out of range: the file has {len(names)} sheet(s).")
    raise InputError(f"Sheet '{sheet}' not found. Available sheets: {', '.join(names)}")


def load_scores(path, sheet: str | None = None,
                lower_better: tuple[str, ...] | list[str] = (),
                higher_better: tuple[str, ...] | list[str] = (),
                group_col: str | None = None) -> Scores:
    """Load and validate the score sheet. Raises InputError with a clear message."""
    path = Path(path)
    ws = _open_sheet(path, sheet)
    if not hasattr(ws, "iter_rows"):
        raise InputError(f"Sheet '{ws.title}' is a chart sheet, not a data sheet.")
    rows = [list(r) for r in ws.iter_rows()]
    if not rows or all(_is_empty(c.value) for r in rows for c in r):
        raise InputError(f"Sheet '{ws.title}' is empty.")

    ncols = max((i + 1 for r in rows for i, c in enumerate(r) if not _is_empty(c.value)),
                default=0)
    header = rows[0][:ncols]
    if all(_is_empty(c.value) for c in header):
        raise InputError("Row 1 is empty. It must hold the headers: name column, then one "
                         "column per station.")
    if ncols < 2:
        raise InputError("No station columns found. Column A is the participant name and "
                         "every column after it is a station.")

    names: list[str] = []
    for cell in header[1:]:
        if _is_empty(cell.value):
            raise InputError(f"Column {cell.column_letter} has data but no station name in "
                             f"row 1 (cell {cell.coordinate}).")
        names.append(str(cell.value).strip())
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        raise InputError(f"Duplicate station name(s) in row 1: {', '.join(dupes)}. "
                         "Each station needs a unique header.")

    # Category column: named with --group-col, or recognised by its header.
    gidx: int | None = None
    if group_col is not None:
        wanted = tr_fold(group_col)
        hits = [j for j, n in enumerate(names) if n == group_col] or \
               [j for j, n in enumerate(names) if tr_fold(n) == wanted]
        if not hits:
            raise InputError(f"--group-col: no column named '{group_col}'. "
                             f"Columns after the name column: {names}")
        gidx = hits[0]
    else:
        hits = [j for j, n in enumerate(names) if tr_fold(n) in GROUP_HEADERS]
        if len(hits) > 1:
            raise InputError(f"More than one category column found "
                             f"({', '.join(names[j] for j in hits)}). Pick one with "
                             '--group-col "NAME"; the others are then read as stations.')
        if hits:
            gidx = hits[0]
    station_names = [n for j, n in enumerate(names) if j != gidx]

    participants: list[str] = []
    data_rows: list[list] = []
    for r in rows[1:]:
        cells = r[:ncols]
        if all(_is_empty(c.value) for c in cells):
            continue
        if _is_empty(cells[0].value):
            raise InputError(f"Row {cells[0].row} has scores but no participant name in "
                             "column A.")
        participants.append(str(cells[0].value).strip())
        data_rows.append(cells[1:])
    if not participants:
        raise InputError("No participants found below the header row.")

    for flag, cols in (("--lower-better", lower_better), ("--higher-better", higher_better)):
        unknown = [c for c in cols if c not in station_names]
        if unknown:
            raise InputError(f"{flag}: unknown station(s) {unknown}. "
                             f"Valid station names: {station_names}")
    both = sorted(set(lower_better) & set(higher_better))
    if both:
        raise InputError(f"Station(s) {both} given to both --lower-better and --higher-better.")

    warnings: list[str] = []
    errors: list[str] = []

    categories: list[str] | None = None
    category_names: dict[str, str] = {}
    if gidx is not None:
        categories = []
        for i, row in enumerate(data_rows):
            cell = row[gidx]
            if _is_empty(cell.value):
                key = UNSPECIFIED
                warnings.append(f"{participants[i]}: blank '{names[gidx]}' cell "
                                f"({cell.coordinate}); placed in the 'unspecified' category.")
            else:
                key = category_key(cell.value)
                category_names.setdefault(key, " ".join(str(cell.value).split()))
            categories.append(key)

    stations: list[Station] = []
    for j, name in enumerate(names):
        if j == gidx:
            continue
        cells = [row[j] for row in data_rows]
        parsed = [parse_cell(c.value, c.number_format) for c in cells]
        coords = [c.coordinate for c in cells]

        bad = [f"{co} ({c.value!r})" for p, co, c in zip(parsed, coords, cells)
               if p.kind == BAD]
        if bad:
            errors.append(f"Station '{name}': cannot read {', '.join(bad[:8])}"
                          f"{' ...' if len(bad) > 8 else ''}. Use a number, a time like 4:32, "
                          "or DNF/DNS.")
            # A column of words only is probably a category column with another header.
            if all(isinstance(c.value, str) for p, c in zip(parsed, cells) if p.kind == BAD) \
                    and not any(p.kind in (TIME, NUMBER) for p in parsed):
                errors.append(f"Hint: if '{name}' is a category column (gender, age group), "
                              f'run again with --group-col "{name}".')
            continue
        time_at = [co for p, co in zip(parsed, coords) if p.kind == TIME]
        num_at = [co for p, co in zip(parsed, coords) if p.kind == NUMBER]
        if time_at and num_at:
            errors.append(f"Station '{name}' mixes time values ({', '.join(time_at[:5])}) and "
                          f"plain numbers ({', '.join(num_at[:5])}). Make the whole column "
                          "one kind.")
            continue
        if not time_at and not num_at:
            warnings.append(f"Station '{name}' has no results at all and was left out of "
                            "the charts.")
            continue

        kind = TIME if time_at else NUMBER
        if kind == TIME:
            native = [(p, co) for p, co in zip(parsed, coords)
                      if p.kind == TIME and p.native_time]
            # Rule 1 (per cell): Excel stores a hand-typed "4:32" as 4 h 32 min with an
            # h:mm format. A cell with that format and no seconds is read as mm:ss,
            # whether it is under or over one hour (0:48 and 1:03 both work).
            typed = [(p, co) for p, co in native if p.hmm_format and _whole_minutes(p.value)]
            if typed:
                first, first_at = typed[0]
                shown = format_time(first.value / 60)
                for p, _ in typed:
                    p.value /= 60.0
                warnings.append(
                    f"Station '{name}': {len(typed)} Excel time cell(s) have an "
                    f"hours:minutes format without seconds (e.g. {first_at} shows {shown}), "
                    "which is how Excel stores a hand-typed time. They were read as "
                    f"minutes:seconds ({shown} = {int(first.value // 60)} min "
                    f"{int(first.value % 60)} s). If they really are hours and minutes, "
                    "type them with seconds, e.g. 1:05:00. See the README tip on time cells.")
            # Rule 2 (whole column): other formats, e.g. mm:ss applied before typing.
            # If every remaining Excel time cell is at least one hour with zero seconds,
            # it was almost surely typed as mm:ss too.
            done = {id(p) for p, _ in typed}
            rest = [(p, co) for p, co in native if id(p) not in done]
            if rest and all(p.value >= 3600 and _whole_minutes(p.value) for p, _ in rest):
                first, first_at = rest[0]
                before = format_time(first.value)
                for p, _ in rest:
                    p.value /= 60.0
                warnings.append(
                    f"Station '{name}': {len(rest)} Excel time cell(s) are all one hour or "
                    f"more with zero seconds (e.g. {first_at} = {before}), so they were "
                    f"read as minutes:seconds ({format_time(first.value)}). If these "
                    "really are hours, see the README tip on time cells.")

        if name in lower_better:
            lb, overridden = True, True
        elif name in higher_better:
            lb, overridden = False, True
        else:
            lb, overridden = kind == TIME, False

        for i, p in enumerate(parsed):
            if p.kind == BLANK:
                warnings.append(f"{participants[i]}: blank cell for '{name}' ({coords[i]}); "
                                "axis point skipped.")
            elif p.kind == DNF:
                warnings.append(f"{participants[i]}: '{p.text}' for '{name}' ({coords[i]}); "
                                "treated as not attempted, axis point skipped.")
        stations.append(Station(name, kind, lb, overridden, parsed))

    if errors:
        raise InputError("\n  ".join(["Problems in the score sheet:"] + errors))
    if len(stations) < 3:
        raise InputError(f"A radar chart needs at least 3 stations with results; "
                         f"found {len(stations)}.")
    if len(participants) == 1:
        warnings.append("Only one participant: every score will be 100.")
    return Scores(ws.title, participants, stations, warnings,
                  group_col=names[gidx] if gidx is not None else None,
                  categories=categories, category_names=category_names)
