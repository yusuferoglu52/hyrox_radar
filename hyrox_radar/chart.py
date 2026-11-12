"""Draw one radar chart per participant with a raw-value table underneath."""
from __future__ import annotations

import math
import re
import textwrap
import unicodedata
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Patch, Rectangle  # noqa: E402

# Dark "race poster" theme.
BG = "#0D0E11"        # page
PANEL = "#16181D"     # header band, radar disc, table rows
PANEL_ALT = "#1B1E24"
LINE = "#2A2E37"      # hairlines
PLAYER = "#FFD60A"    # accent
AVG_FILL = "#8B93A1"
AVG_EDGE = "#B4BCC8"
GRID = "#2F343E"
TEXT = "#F2F4F7"
MUTED = "#7C8594"
TICK = "#A3ACB9"
GOOD = "#3DDC84"      # score above the average
BAD = "#FF6B7A"       # score below the average
DASH = "–"

GLOW = ((9.0, 0.07), (5.5, 0.12))  # (line width, alpha) halos under the outline

BASE_WIDTH = 8.5  # inches; widened automatically when axis labels need more room
RADAR = 5.0

STRINGS = {
    "tr": {
        "no_result": "(sonuç yok)",
        "legend_avg": "{who} (normalize puanların\nortalaması; boş ve DNF/DNS hariç)",
        "headers": ["İstasyon", "Sonuç", f"Puan (0{DASH}100)"],
        "who_group": "Grup ortalaması",
        "who_cat": "{label} ortalaması",
        "avg_header_group": "Grup ort. puanı",
        "avg_header_cat": "{label} ort. puanı",
        "avg_header_long": "Kategori ort. puanı",
        "best_group": "Grubun en iyisi",
        "best_cat": "Kategorinin en iyisi",
        "subtitle": "Kategori: {label}",
        "cat_F": "Kadın",
        "cat_M": "Erkek",
        "cat_none": "Belirtilmemiş",
        "scale": {
            "minmax": "Ölçek: istasyon bazında min-maks. {best} = 100, "
                      "en kötüsü = 0. Dışa doğru = daha iyi.",
            "pct-of-best": "Ölçek: istasyon bazında en iyinin yüzdesi. {best} = 100. "
                           "Dışa doğru = daha iyi.",
        },
        "avg_note": "{who} = istasyon bazında normalize puanların ortalaması; "
                    "boş hücreler ve DNF/DNS hariç.",
        "floor_note": "Dikkat: eksen 0'dan değil {floor} puandan başlar (merkez = {floor}); "
                      "farklar olduğundan büyük görünür.",
    },
    "en": {
        "no_result": "(no result)",
        "legend_avg": "{who} (mean of normalised scores,\nexcl. blanks and DNF/DNS)",
        "headers": ["Station", "Result", f"Score (0{DASH}100)"],
        "who_group": "Group average",
        "who_cat": "{label} average",
        "avg_header_group": "Group avg score",
        "avg_header_cat": "{label} avg score",
        "avg_header_long": "Category avg score",
        "best_group": "Best in group",
        "best_cat": "Best in category",
        "subtitle": "Category: {label}",
        "cat_F": "Women",
        "cat_M": "Men",
        "cat_none": "Unspecified",
        "scale": {
            "minmax": "Scale: min-max per station. {best} = 100, worst = 0. "
                      "Outward = better.",
            "pct-of-best": "Scale: % of best per station. {best} = 100. "
                           "Outward = better.",
        },
        "avg_note": "{who} = mean of normalised scores per station, "
                    "excluding blanks and DNF/DNS.",
        "floor_note": "Note: the axis starts at {floor}, not 0 (centre = {floor}), "
                      "so gaps look larger than they are.",
    },
}
LANGS = tuple(STRINGS)

_FORBIDDEN = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
             *(f"LPT{i}" for i in range(1, 10))}


def safe_filename(name: str) -> str:
    """Strip only what Windows forbids; Turkish and other Unicode letters stay."""
    clean = unicodedata.normalize("NFC", name)
    clean = _FORBIDDEN.sub("", clean)
    clean = re.sub(r"\s+", " ", clean).strip(" .")
    if not clean:
        clean = "participant"
    if clean.upper() in _RESERVED:
        clean = "_" + clean
    return clean[:120]


def unique_filenames(names: list[str]) -> list[str]:
    """Safe file stems, with _2, _3... for names that collide (case-insensitive)."""
    seen: dict[str, int] = {}
    result = []
    for name in names:
        stem = safe_filename(name)
        key = stem.casefold()
        seen[key] = seen.get(key, 0) + 1
        result.append(stem if seen[key] == 1 else f"{stem}_{seen[key]}")
    return result


def valid_runs(ok) -> list[list[int]]:
    """Circular runs of neighbouring stations that have a result.

    All valid -> one run with every index (a closed polygon). A missing station
    splits the outline, so no line is drawn across an axis without a result.
    """
    ok = [bool(x) for x in ok]
    n = len(ok)
    if all(ok):
        return [list(range(n))]
    if not any(ok):
        return []
    start = ok.index(False)
    runs, current = [], []
    for k in range(1, n + 1):
        i = (start + k) % n
        if ok[i]:
            current.append(i)
        elif current:
            runs.append(current)
            current = []
    return runs


def auto_floor(min_score: float) -> float:
    """Axis floor for --axis-min auto: a multiple of 10 at least 5 below the lowest score."""
    floor = math.floor((min_score - 5) / 10) * 10
    return float(max(0, min(floor, 90)))


def ring_ticks(floor: float) -> list[int]:
    """Ring values above the floor, using the coarsest step that gives 3+ rings."""
    for step in (20, 10, 5, 2, 1):
        ticks = [t for t in range(step, 101, step) if t > floor + 1e-9]
        if len(ticks) >= 3:
            return ticks
    return [100]


def _score_text(value: float) -> str:
    return DASH if np.isnan(value) else f"{value:.0f}"


def _num(value: float) -> str:
    return f"{value:g}"


def category_label(key: str, names: dict[str, str], lang: str) -> str:
    """Display name of a category key from the loader."""
    s = STRINGS[lang]
    if key == "F":
        return s["cat_F"]
    if key == "M":
        return s["cat_M"]
    if key == "":
        return s["cat_none"]
    return names.get(key, key)


def _draw(name, stations, scores, average, raw, scale, lang, axis_min, width,
          category=None):
    """Build the figure. Returns (fig, ax, title_text).

    category: display name of the participant's category when scoring is per
    category, or None when everyone is compared in one group.
    """
    s = STRINGS[lang]
    n = len(stations)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False)
    floor = float(axis_min)
    if category:
        who = s["who_cat"].format(label=category)
        best = s["best_cat"]
        avg_header = (s["avg_header_cat"].format(label=category) if len(category) <= 8
                      else s["avg_header_long"])
    else:
        who, best, avg_header = s["who_group"], s["best_group"], s["avg_header_group"]

    # Layout in inches, top to bottom.
    title_h = 1.15 if category else 0.85
    label_pad, legend_h = 0.85, 0.75
    foot_h = 0.95 if floor > 0 else 0.75
    table_h = 0.36 * (n + 1)
    height = title_h + label_pad + RADAR + label_pad + legend_h + table_h + foot_h

    fig = plt.figure(figsize=(width, height), facecolor=BG)
    # Header band with an accent stripe under it.
    fig.add_artist(Rectangle((0, 1 - title_h / height), 1, title_h / height,
                             transform=fig.transFigure, facecolor=PANEL, linewidth=0,
                             zorder=0))
    fig.add_artist(Rectangle((0, 1 - (title_h + 0.045) / height), 1, 0.045 / height,
                             transform=fig.transFigure, facecolor=PLAYER, linewidth=0,
                             zorder=0))
    title = fig.text(0.5, 1 - (0.46 if category else 0.43) / height, name, ha="center",
                     va="center", fontsize=24, fontweight="bold", color=TEXT)
    if category:
        fig.text(0.5, 1 - 0.89 / height, s["subtitle"].format(label=category),
                 ha="center", va="center", fontsize=10.5, fontweight="bold", color=BG,
                 bbox=dict(boxstyle="round,pad=0.35,rounding_size=0.9", facecolor=PLAYER,
                           edgecolor="none"))

    radar_y = legend_h + table_h + foot_h + label_pad
    ax = fig.add_axes([(width - RADAR) / 2 / width, radar_y / height,
                       RADAR / width, RADAR / height], projection="polar")
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)
    ax.set_ylim(floor, 100)
    ticks = ring_ticks(floor)
    ax.set_yticks(ticks)
    ax.set_yticklabels([])
    ax.grid(color=GRID, linewidth=0.9)
    ax.spines["polar"].set_color("#4A505C")
    ax.spines["polar"].set_linewidth(1.4)
    ax.set_facecolor(PANEL)

    # Scores under the floor are drawn at the centre.
    shown = np.where(np.isnan(scores), np.nan, np.maximum(scores, floor))
    shown_avg = np.where(np.isnan(average), np.nan, np.maximum(average, floor))

    # Group average behind the participant.
    ok = ~np.isnan(shown_avg)
    if ok.sum() >= 3:
        a, v = angles[ok], shown_avg[ok]
        ax.fill(a, v, color=AVG_FILL, alpha=0.28, zorder=2, linewidth=0)
        ax.plot(np.append(a, a[0]), np.append(v, v[0]), color=AVG_EDGE, linewidth=1.3,
                linestyle=(0, (4, 2)), zorder=2.5, clip_on=False)

    # Participant: the outline breaks at stations without a result.
    ok = ~np.isnan(shown)
    for run in valid_runs(ok):
        a, v = angles[run], shown[run]
        if ok.all():
            ax.fill(a, v, color=PLAYER, alpha=0.20, zorder=3, linewidth=0)
            for lw, alpha in GLOW:
                ax.plot(np.append(a, a[0]), np.append(v, v[0]), color=PLAYER, linewidth=lw,
                        alpha=alpha, zorder=3.5, clip_on=False, solid_joinstyle="round")
            ax.plot(np.append(a, a[0]), np.append(v, v[0]), color=PLAYER, linewidth=2.6,
                    zorder=4, clip_on=False, solid_joinstyle="round")
        elif len(run) >= 2:
            ax.fill(np.append(a, a[-1]), np.append(v, floor), color=PLAYER, alpha=0.20,
                    zorder=3, linewidth=0)
            for lw, alpha in GLOW:
                ax.plot(a, v, color=PLAYER, linewidth=lw, alpha=alpha, zorder=3.5,
                        clip_on=False, solid_joinstyle="round", solid_capstyle="round")
            ax.plot(a, v, color=PLAYER, linewidth=2.6, zorder=4, clip_on=False,
                    solid_joinstyle="round", solid_capstyle="round")
    ax.plot(angles[ok], shown[ok], linestyle="none", marker="o", markersize=7,
            color=PLAYER, markeredgecolor=BG, markeredgewidth=1.6, zorder=5,
            clip_on=False)

    # Ring labels drawn on top of everything so they are always readable.
    box = dict(boxstyle="round,pad=0.18", facecolor=BG, edgecolor="none", alpha=0.75)
    for t in ticks:
        ax.text(np.pi / n, t, str(t), ha="center", va="center", fontsize=8.5, color=TICK,
                zorder=10, bbox=box)
    if floor > 0:
        ax.text(0, floor, _num(floor), ha="center", va="center", fontsize=8.5, color=TICK,
                fontweight="bold", zorder=10, bbox=box)

    labels = []
    for station, score in zip(stations, scores):
        text = "\n".join(textwrap.wrap(station, 16)) or station
        labels.append(text + (f"\n{s['no_result']}" if np.isnan(score) else ""))
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=11.5, fontweight="bold")
    ax.tick_params(axis="x", pad=10)
    for label, angle, score in zip(ax.get_xticklabels(), angles, scores):
        x = np.sin(angle)
        label.set_horizontalalignment("center" if abs(x) < 0.15
                                      else ("left" if x > 0 else "right"))
        label.set_color(MUTED if np.isnan(score) else TEXT)

    handles = [
        Line2D([0], [0], color=PLAYER, linewidth=2.6, marker="o", markersize=7,
               markeredgecolor=BG, label=name),
        Patch(facecolor=AVG_FILL, alpha=0.45, edgecolor=AVG_EDGE, linestyle="--",
              label=s["legend_avg"].format(who=who)),
    ]
    fig.legend(handles=handles, loc="center", ncol=2, frameon=False, fontsize=10.5,
               bbox_to_anchor=(0.5, (foot_h + table_h + legend_h / 2) / height),
               columnspacing=2.5, handlelength=2.2, labelcolor=TEXT)

    # Raw values table.
    table_w = 7.3 / width
    tax = fig.add_axes([(1 - table_w) / 2, foot_h / height, table_w, table_h / height])
    tax.axis("off")
    rows = [[st if len(st) <= 34 else st[:33] + "…", r, _score_text(sc), _score_text(av)]
            for st, r, sc, av in zip(stations, raw, scores, average)]
    table = tax.table(cellText=rows, colLabels=s["headers"] + [avg_header],
                      colWidths=[0.39, 0.18, 0.20, 0.23], cellLoc="center",
                      bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(10.5)
    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor(LINE)
        cell.set_linewidth(0.8)
        if row == 0:
            cell.set_facecolor(BG)
            cell.set_text_props(color=PLAYER, fontweight="bold")
        else:
            score, avg = scores[row - 1], average[row - 1]
            cell.set_facecolor(PANEL_ALT if row % 2 == 0 else PANEL)
            cell.set_text_props(color=TICK if col == 3 else TEXT)
            if col == 1:
                cell.set_text_props(fontweight="bold")
            if col == 2 and not (np.isnan(score) or np.isnan(avg)):
                # Compare what is printed, so equal-looking numbers stay neutral.
                diff = float(_score_text(score)) - float(_score_text(avg))
                if diff:
                    cell.set_text_props(color=GOOD if diff > 0 else BAD, fontweight="bold")
            if np.isnan(score) and col in (1, 2):
                cell.set_text_props(color=MUTED)
        if col == 0:
            cell.set_text_props(ha="left")
            cell.PAD = 0.03

    notes = [s["scale"][scale].format(best=best), s["avg_note"].format(who=who)]
    if floor > 0:
        notes.append(s["floor_note"].format(floor=_num(floor)))
    y = foot_h - 0.33
    for i, note in enumerate(notes):
        fig.text(0.5, (y - 0.2 * i) / height, note, ha="center", va="center", fontsize=8.5,
                 color=PLAYER if i == 2 else MUTED,
                 fontweight="bold" if i == 2 else "normal")
    return fig, ax, title


def chart_width(stations: list[str], lang: str = "tr") -> float:
    """Figure width (inches) that fits every axis label, shared by all charts of a run."""
    n = len(stations)
    nan = np.full(n, np.nan)
    with plt.rc_context({"font.family": "DejaVu Sans"}):
        # All-NaN scores give every label its widest form (with the no-result line).
        fig, ax, _ = _draw("", stations, nan, nan, [""] * n, "minmax", lang, 0.0, BASE_WIDTH)
        fig.canvas.draw()
        renderer = fig.canvas.get_renderer()
        overflow = 0.0
        for label in ax.get_xticklabels():
            box = label.get_window_extent(renderer)
            overflow = max(overflow, -box.x0, box.x1 - fig.bbox.width)
        plt.close(fig)
    need = overflow / fig.dpi + 0.25  # keep a quarter inch of margin
    return BASE_WIDTH + (math.ceil(2 * need * 10) / 10 if need > 0 else 0.0)


def render_chart(name: str, stations: list[str], scores, average, raw: list[str],
                 scale: str, out_path: Path, dpi: int = 200, lang: str = "tr",
                 axis_min: float = 0.0, width: float | None = None,
                 category: str | None = None) -> None:
    """scores/average: 0-100 per station, NaN = no result. raw: display strings."""
    scores = np.asarray(scores, dtype=float)
    average = np.asarray(average, dtype=float)
    if width is None:
        width = chart_width(stations, lang)
    with plt.rc_context({"font.family": "DejaVu Sans", "text.color": TEXT}):
        fig, _, title = _draw(name, stations, scores, average, raw, scale, lang,
                              axis_min, width, category)
        # Shrink a very long name until the title fits.
        fig.canvas.draw()
        renderer = fig.canvas.get_renderer()
        while (title.get_window_extent(renderer).width > 0.92 * fig.bbox.width
               and title.get_fontsize() > 11):
            title.set_fontsize(title.get_fontsize() - 1)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=dpi, facecolor=BG)
        plt.close(fig)
