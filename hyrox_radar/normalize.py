"""Scale each station to 0-100 so that a higher score is always better."""
from __future__ import annotations

import numpy as np
import pandas as pd

SCALES = ("minmax", "pct-of-best")

SCALE_NOTES = {
    "minmax": "min-max per station: best in group = 100, worst = 0",
    "pct-of-best": "% of best per station: best in group = 100, "
                   "others = their percentage of the best result",
}


def normalize(values: pd.DataFrame, lower_better: dict[str, bool],
              scale: str = "minmax") -> pd.DataFrame:
    """Return 0-100 scores per station; NaN stays NaN (no result)."""
    if scale not in SCALES:
        raise ValueError(f"Unknown scale '{scale}'. Choose from {SCALES}.")
    out = pd.DataFrame(np.nan, index=values.index, columns=values.columns, dtype=float)
    for col in values.columns:
        v = values[col].astype(float)
        present = v.notna()
        if not present.any():
            continue
        lo, hi = v.min(), v.max()
        lower = lower_better[col]
        best = lo if lower else hi
        if hi == lo:
            score = pd.Series(100.0, index=v.index)
        elif scale == "minmax":
            score = (hi - v) / (hi - lo) * 100 if lower else (v - lo) / (hi - lo) * 100
        elif best <= 0:
            # A ratio to a zero/negative best is undefined: best gets 100, others 0.
            score = pd.Series(np.where(v == best, 100.0, 0.0), index=v.index)
        elif lower:
            score = best / v * 100
        else:
            score = v / best * 100
        out[col] = score.where(present).clip(0, 100)
    return out


def ordered_categories(categories: list[str]) -> list[str]:
    """Women, men, then other categories in first-seen order, unspecified last."""
    seen = list(dict.fromkeys(categories))
    rank = {"F": 0, "M": 1, "": 3}
    return sorted(seen, key=lambda key: (rank.get(key, 2), seen.index(key)))


def score_by_category(values: pd.DataFrame, lower_better: dict[str, bool], scale: str,
                      categories: list[str] | None):
    """Normalize within each category.

    Returns (scores, averages): scores has the shape of values; averages maps each
    category key to its per-station mean. With categories=None everyone is one
    group and the only key is None.
    """
    if categories is None:
        scores = normalize(values, lower_better, scale)
        return scores, {None: group_average(scores)}
    scores = pd.DataFrame(np.nan, index=values.index, columns=values.columns, dtype=float)
    averages = {}
    for key in ordered_categories(categories):
        rows = [i for i, c in enumerate(categories) if c == key]
        part = normalize(values.iloc[rows], lower_better, scale)
        scores.iloc[rows] = part.to_numpy()
        averages[key] = group_average(part)
    return scores, averages


def group_average(scores: pd.DataFrame) -> pd.Series:
    """Mean of the normalized scores per station, ignoring missing results."""
    return scores.mean(axis=0, skipna=True)
