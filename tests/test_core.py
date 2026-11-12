"""Tests for the core logic: time parsing, column detection, normalization."""
import datetime as dt
import math
import shutil
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from hyrox_radar.chart import (auto_floor, ring_ticks, safe_filename, unique_filenames,
                               valid_runs)
from hyrox_radar.loader import (BLANK, DNF, NUMBER, TIME, InputError, category_key,
                                format_lacks_seconds, format_time, load_scores, parse_cell,
                                parse_time_string, tr_fold)
from hyrox_radar.normalize import (group_average, normalize, ordered_categories,
                                   score_by_category)


def write_xlsx(rows, column_formats=None) -> Path:
    """column_formats: {"B": "h:mm"} applies a number format to data cells of a column."""
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    for column, fmt in (column_formats or {}).items():
        for cell in ws[column][1:]:
            cell.number_format = fmt
    folder = tempfile.mkdtemp(prefix="hyrox_test_")  # system temp, never the project
    _TEMP_DIRS.append(folder)
    path = Path(folder) / "scores.xlsx"
    wb.save(path)
    return path


_TEMP_DIRS: list[str] = []


def tearDownModule():
    for folder in _TEMP_DIRS:
        shutil.rmtree(folder, ignore_errors=True)


class TimeParsing(unittest.TestCase):
    def test_strings(self):
        self.assertEqual(parse_time_string("4:32"), 272)
        self.assertEqual(parse_time_string("04:32"), 272)
        self.assertEqual(parse_time_string("1:04:32"), 3872)
        self.assertEqual(parse_time_string("75:30"), 4530)
        self.assertAlmostEqual(parse_time_string("0:48.5"), 48.5)
        self.assertIsNone(parse_time_string("4:75"))
        self.assertIsNone(parse_time_string("1:75:00"))
        self.assertIsNone(parse_time_string("432"))

    def test_cells(self):
        self.assertEqual(parse_cell(None).kind, BLANK)
        self.assertEqual(parse_cell("  ").kind, BLANK)
        self.assertEqual(parse_cell("dnf").kind, DNF)
        self.assertEqual(parse_cell("DNS").kind, DNF)
        self.assertEqual(parse_cell(42).kind, NUMBER)
        self.assertEqual(parse_cell("42,5").value, 42.5)
        self.assertEqual(parse_cell(dt.timedelta(minutes=4, seconds=32)).value, 272)
        self.assertEqual(parse_cell(dt.time(0, 4, 32)).value, 272)
        self.assertEqual(parse_cell("4:32").kind, TIME)
        self.assertEqual(parse_cell("abc").kind, "bad")

    def test_format(self):
        self.assertEqual(format_time(272), "4:32")
        self.assertEqual(format_time(48.5), "0:48.5")
        self.assertEqual(format_time(3872), "1:04:32")
        self.assertEqual(format_time(59.96), "1:00")


class Loading(unittest.TestCase):
    def test_detection_and_missing(self):
        path = write_xlsx([
            ["Name", "Run", "Sled", "Reps"],
            ["A", dt.timedelta(minutes=4, seconds=32), "0:48", 40],
            ["B", dt.timedelta(minutes=3, seconds=58), "DNF", None],
            ["C", dt.timedelta(minutes=5, seconds=10), "1:03", 55],
        ])
        data = load_scores(path)
        kinds = {s.name: (s.kind, s.lower_better) for s in data.stations}
        self.assertEqual(kinds, {"Run": (TIME, True), "Sled": (TIME, True),
                                 "Reps": (NUMBER, False)})
        values = data.values()
        # Real Excel durations with seconds: the h:mm heuristic must not fire.
        self.assertEqual([round(v) for v in values["Run"]], [272, 238, 310])
        self.assertFalse(any("minutes:seconds" in w for w in data.warnings))
        self.assertTrue(math.isnan(values["Sled"][1]))
        self.assertTrue(math.isnan(values["Reps"][1]))
        self.assertEqual(data.stations[1].display(1), "DNF")
        self.assertEqual(data.stations[2].display(1), "–")
        self.assertEqual(sum("axis point skipped" in w for w in data.warnings), 2)

    def test_format_detection(self):
        for fmt in ("h:mm", "hh:mm", "[h]:mm", "h:mm AM/PM", 'h:mm" s"'):
            self.assertTrue(format_lacks_seconds(fmt), fmt)
        for fmt in ("h:mm:ss", "mm:ss", "[h]:mm:ss", "mm:ss.0", "General", "0.00", None):
            self.assertFalse(format_lacks_seconds(fmt), fmt)

    def test_hmm_format_mixing_under_and_over_one_hour(self):
        # Hand-typed 0:48, 1:03, 4:32 -> Excel stores 48 min, 1 h 3 min, 4 h 32 min as h:mm.
        path = write_xlsx([
            ["Name", "Sled", "Reps", "Metres"],
            ["A", dt.time(0, 48), 1, 2],
            ["B", dt.time(1, 3), 2, 3],
            ["C", dt.time(4, 32), 3, 4],
            ["D", dt.time(0, 5), 4, 5],
        ], {"B": "h:mm"})
        data = load_scores(path)
        self.assertEqual(data.values()["Sled"].tolist(), [48, 63, 272, 5])
        self.assertEqual([data.stations[0].display(i) for i in range(4)],
                         ["0:48", "1:03", "4:32", "0:05"])
        hits = [w for w in data.warnings if "hours:minutes format" in w]
        self.assertEqual(len(hits), 1)
        self.assertIn("4 Excel time cell(s)", hits[0])
        self.assertIn("B2 shows 0:48", hits[0])

    def test_hmm_format_all_under_one_hour(self):
        path = write_xlsx([
            ["Name", "Sled", "Reps", "Metres"],
            ["A", dt.time(0, 48), 1, 2],
            ["B", dt.time(0, 52), 2, 3],
        ], {"B": "h:mm"})
        data = load_scores(path)
        self.assertEqual(data.values()["Sled"].tolist(), [48, 52])
        self.assertTrue(any("hours:minutes format" in w for w in data.warnings))

    def test_hmm_format_mixed_with_text_times(self):
        path = write_xlsx([
            ["Name", "Sled", "Reps", "Metres"],
            ["A", dt.time(0, 48), 1, 2],
            ["B", "1:03", 2, 3],
            ["C", dt.time(1, 10), 3, 4],
        ], {"B": "h:mm"})
        data = load_scores(path)
        self.assertEqual(data.values()["Sled"].tolist(), [48, 63, 70])
        self.assertIn("2 Excel time cell(s)", " ".join(data.warnings))

    def test_proper_time_cells_are_not_converted(self):
        # 0:04:12 with seconds in the format is already correct, over or under an hour.
        path = write_xlsx([
            ["Name", "Run", "Hold", "Metres"],
            ["A", dt.time(0, 4, 12), dt.time(0, 2, 0), 2],
            ["B", dt.time(0, 3, 58), dt.time(0, 3, 0), 3],
        ], {"B": "mm:ss", "C": "h:mm:ss"})
        data = load_scores(path)
        self.assertEqual(data.values()["Run"].tolist(), [252, 238])
        self.assertEqual(data.values()["Hold"].tolist(), [120, 180])  # whole minutes kept
        self.assertEqual(data.warnings, [])

    def test_hmm_format_with_real_seconds_is_not_converted(self):
        # The value carries seconds, so it was not typed as h:mm; only the display hides them.
        path = write_xlsx([
            ["Name", "Run", "Reps", "Metres"],
            ["A", dt.time(0, 4, 12), 1, 2],
            ["B", dt.time(0, 3, 58), 2, 3],
        ], {"B": "h:mm"})
        data = load_scores(path)
        self.assertEqual(data.values()["Run"].tolist(), [252, 238])
        self.assertEqual(data.warnings, [])

    def test_column_fallback_for_other_formats(self):
        # mm:ss format applied first, then 4:32 typed: stored as 4 h 32 min, format has seconds.
        path = write_xlsx([
            ["Name", "Run", "Reps", "Metres"],
            ["A", dt.time(4, 32), 1, 2],
            ["B", dt.time(3, 58), 2, 3],
        ], {"B": "mm:ss"})
        data = load_scores(path)
        self.assertEqual(data.values()["Run"].tolist(), [272, 238])
        self.assertTrue(any("one hour or more with zero seconds" in w for w in data.warnings))

    def test_overrides_and_sheet(self):
        path = write_xlsx([["Name", "A", "B", "C"], ["x", 1, 2, 3], ["y", 2, 3, 4]])
        data = load_scores(path, sheet="1", lower_better=["A"])
        self.assertTrue(data.stations[0].lower_better)
        self.assertFalse(data.stations[1].lower_better)
        self.assertEqual(load_scores(path, sheet="Sheet").sheet, "Sheet")
        with self.assertRaises(InputError):
            load_scores(path, lower_better=["Nope"])
        with self.assertRaises(InputError):
            load_scores(path, sheet="Missing")
        with self.assertRaises(InputError):
            load_scores(path, sheet="5")

    def test_malformed(self):
        mixed = write_xlsx([["Name", "A", "B", "C"], ["x", "4:32", 2, 3], ["y", 50, 3, 4]])
        with self.assertRaisesRegex(InputError, "mixes time values"):
            load_scores(mixed)
        junk = write_xlsx([["Name", "A", "B", "C"], ["x", "fast", 2, 3]])
        with self.assertRaisesRegex(InputError, "cannot read B2"):
            load_scores(junk)
        few = write_xlsx([["Name", "A", "B"], ["x", 1, 2]])
        with self.assertRaisesRegex(InputError, "at least 3 stations"):
            load_scores(few)
        dupe = write_xlsx([["Name", "A", "A", "C"], ["x", 1, 2, 3]])
        with self.assertRaisesRegex(InputError, "Duplicate station"):
            load_scores(dupe)
        noname = write_xlsx([["Name", "A", "B", "C"], [None, 1, 2, 3]])
        with self.assertRaisesRegex(InputError, "no participant name"):
            load_scores(noname)
        with self.assertRaisesRegex(InputError, "File not found"):
            load_scores("does_not_exist.xlsx")


class Normalizing(unittest.TestCase):
    def setUp(self):
        self.values = pd.DataFrame({
            "time": [240.0, 300.0, 480.0, float("nan")],
            "reps": [40.0, 50.0, float("nan"), 30.0],
            "same": [80.0, 80.0, 80.0, float("nan")],
        })
        self.lower = {"time": True, "reps": False, "same": False}

    def test_minmax(self):
        s = normalize(self.values, self.lower, "minmax")
        self.assertEqual(s["time"].tolist()[:3], [100.0, 75.0, 0.0])
        self.assertTrue(math.isnan(s["time"][3]))
        self.assertEqual(s["reps"][0], 50.0)
        self.assertEqual(s["reps"][1], 100.0)
        self.assertEqual(s["reps"][3], 0.0)
        self.assertEqual(s["same"].tolist()[:3], [100.0, 100.0, 100.0])
        self.assertTrue(math.isnan(s["same"][3]))

    def test_pct_of_best(self):
        s = normalize(self.values, self.lower, "pct-of-best")
        self.assertEqual(s["time"].tolist()[:3], [100.0, 80.0, 50.0])
        self.assertEqual(s["reps"][0], 80.0)
        self.assertEqual(s["reps"][3], 60.0)
        self.assertEqual(s["same"][0], 100.0)

    def test_pct_of_best_zero_best(self):
        v = pd.DataFrame({"reps": [0.0, 0.0], "t": [0.0, 10.0]})
        s = normalize(v, {"reps": False, "t": True}, "pct-of-best")
        self.assertEqual(s["reps"].tolist(), [100.0, 100.0])
        self.assertEqual(s["t"].tolist(), [100.0, 0.0])

    def test_average_excludes_missing(self):
        s = normalize(self.values, self.lower, "minmax")
        avg = group_average(s)
        self.assertAlmostEqual(avg["time"], (100 + 75 + 0) / 3)
        self.assertAlmostEqual(avg["reps"], (50 + 100 + 0) / 3)


class Categories(unittest.TestCase):
    ROWS = [
        ["Ad", "Cinsiyet", "A", "B", "C"],
        ["a", "E", 10, 1, 5],
        ["b", "Kadın ", 20, 2, 5],
        ["c", "ERKEK", 30, 3, 5],
        ["d", "KADIN", 40, 4, 5],
        ["e", "k", 50, 5, 5],
    ]

    def test_turkish_fold(self):
        self.assertEqual(tr_fold("KADIN"), "kadın")
        self.assertEqual(tr_fold("  Kadın  "), "kadın")
        self.assertEqual(tr_fold("İSTANBUL  ILIK"), "istanbul ılık")
        self.assertEqual(tr_fold("CİNSİYET"), "cinsiyet")
        for text in ("K", "k", "Kadın", "KADIN", "Kadın ", "kadin", "Kadin", "F", "Female"):
            self.assertEqual(category_key(text), "F", text)
        for text in ("E", "e", "Erkek", "ERKEK", " erkek", "M", "Male"):
            self.assertEqual(category_key(text), "M", text)
        self.assertEqual(category_key("Masters 40+"), "masters 40+")
        self.assertEqual(category_key("MASTERS  40+"), "masters 40+")

    def test_bayan_and_bay_aliases(self):
        for text in ("Bayan", "bayan", "BAYAN", " Bayan "):
            self.assertEqual(category_key(text), "F", text)
        for text in ("Bay", "bay", "BAY", " Bay "):
            self.assertEqual(category_key(text), "M", text)

    def test_whitespace_is_trimmed_for_any_category_value(self):
        # Spaces, tabs, line breaks and non-breaking spaces, before, after and inside.
        for text in ("  K", "K  ", "\tKadın\t", " KADIN\n", " Kadın "):
            self.assertEqual(category_key(text), "F", repr(text))
        for text in ("  E", "E  ", "\tErkek ", " ERKEK"):
            self.assertEqual(category_key(text), "M", repr(text))
        for text in ("Pro", "  Pro", "Pro  ", "\tPRO\n", " pro "):
            self.assertEqual(category_key(text), "pro", repr(text))
        self.assertEqual(category_key("  Masters \t 40+  "), "masters 40+")

    def test_aliases_and_whitespace_through_the_loader(self):
        rows = [["Ad", " Cinsiyet ", "A", "B", "C"],
                ["a", " Bay ", 1, 2, 3], ["b", "BAYAN", 2, 3, 4], ["c", "bay", 3, 4, 5],
                ["d", "\tBayan", 4, 5, 6], ["e", "  Pro ", 5, 6, 7], ["f", "PRO", 6, 7, 8]]
        data = load_scores(write_xlsx(rows))
        self.assertEqual(data.group_col, "Cinsiyet")
        self.assertEqual(data.categories, ["M", "F", "M", "F", "pro", "pro"])
        self.assertEqual(data.category_names["pro"], "Pro")  # stored without the padding

    def test_category_column_is_not_a_station(self):
        data = load_scores(write_xlsx(self.ROWS))
        self.assertEqual(data.group_col, "Cinsiyet")
        self.assertEqual(data.station_names, ["A", "B", "C"])
        self.assertEqual(data.categories, ["M", "F", "M", "F", "F"])

    def test_header_variants_and_position(self):
        rows = [["Ad", "A", "B", "KATEGORİ", "C"], ["a", 1, 2, "Pro", 3], ["b", 2, 3, "open", 4],
                ["c", 3, 4, "PRO", 5]]
        data = load_scores(write_xlsx(rows))
        self.assertEqual(data.group_col, "KATEGORİ")
        self.assertEqual(data.station_names, ["A", "B", "C"])
        self.assertEqual(data.categories, ["pro", "open", "pro"])
        self.assertEqual(data.category_names, {"pro": "Pro", "open": "open"})

    def test_group_col_option_and_hint(self):
        rows = [["Ad", "Lig", "A", "B", "C"], ["a", "Pro", 1, 2, 3], ["b", "Open", 2, 3, 4]]
        path = write_xlsx(rows)
        with self.assertRaisesRegex(InputError, r'--group-col "Lig"'):
            load_scores(path)  # 'Lig' is not a known header, so it is read as a station
        data = load_scores(path, group_col="Lig")
        self.assertEqual(data.station_names, ["A", "B", "C"])
        self.assertEqual(data.categories, ["pro", "open"])
        with self.assertRaisesRegex(InputError, "no column named"):
            load_scores(path, group_col="Nope")

    def test_blank_category_and_no_column(self):
        rows = [r[:] for r in self.ROWS]
        rows[2][1] = None
        data = load_scores(write_xlsx(rows))
        self.assertEqual(data.categories[1], "")
        self.assertTrue(any("unspecified" in w for w in data.warnings))
        plain = load_scores(write_xlsx([["Ad", "A", "B", "C"], ["a", 1, 2, 3]]))
        self.assertIsNone(plain.categories)
        self.assertIsNone(plain.group_col)

    def test_scoring_per_category(self):
        data = load_scores(write_xlsx(self.ROWS))
        scores, averages = score_by_category(data.values(), data.lower_better(),
                                             "pct-of-best", data.categories)
        # Men: 10, 30 -> best 30. Women: 20, 40, 50 -> best 50.
        self.assertEqual([round(v, 1) for v in scores["A"]], [33.3, 40.0, 100.0, 80.0, 100.0])
        self.assertEqual(list(averages), ["F", "M"])
        self.assertAlmostEqual(averages["F"]["A"], (40 + 80 + 100) / 3)
        self.assertAlmostEqual(averages["M"]["A"], (100 / 3 + 100) / 2)
        # --compare all: one group, best is 50.
        scores, averages = score_by_category(data.values(), data.lower_better(),
                                             "pct-of-best", None)
        self.assertEqual(scores["A"].tolist(), [20.0, 40.0, 60.0, 80.0, 100.0])
        self.assertEqual(list(averages), [None])

    def test_category_order(self):
        self.assertEqual(ordered_categories(["M", "pro", "", "F", "open", "M"]),
                         ["F", "M", "pro", "open", ""])


class CommandLine(unittest.TestCase):
    def test_defaults(self):
        from radar import build_parser
        args = build_parser().parse_args(["scores.xlsx"])
        self.assertEqual(args.axis_min, "auto")
        self.assertEqual(args.scale, "pct-of-best")
        self.assertEqual(args.lang, "tr")
        self.assertEqual(args.compare, "group")

    def test_axis_min_values(self):
        from radar import build_parser
        parse = build_parser().parse_args
        self.assertEqual(parse(["f.xlsx", "--axis-min", "0"]).axis_min, 0.0)  # zoom off
        self.assertEqual(parse(["f.xlsx", "--axis-min", "50"]).axis_min, 50.0)
        self.assertEqual(parse(["f.xlsx", "--axis-min", "AUTO"]).axis_min, "auto")


class WindowApp(unittest.TestCase):
    ROWS = [["Ad", "Cinsiyet", "A", "B", "C"], ["Ayşe", "K", "4:10", 30, 500],
            ["Çağrı", "E", "3:50", 40, 600], ["Işıl", "K", "4:30", 35, 550],
            ["Ömer", "E", "4:05", None, 580]]

    def test_form_values_become_options(self):
        import app
        path = str(write_xlsx(self.ROWS))
        opts = app.build_options(path, "out", app.AUTO, app.AUTO, "Herkes tek grupta",
                                 "Min-maks (en kötü = 0)", "Elle gir", "40", "English", "150")
        self.assertEqual((opts.sheet, opts.group_col, opts.compare, opts.scale),
                         (None, None, "all", "minmax"))
        self.assertEqual((opts.axis_min, opts.lang, opts.dpi), (40.0, "en", 150))
        defaults = app.build_options(path, "out", app.AUTO, "Cinsiyet",
                                     "Kendi kategorisi içinde",
                                     "En iyinin yüzdesi (önerilen)", "Otomatik (önerilen)",
                                     "", "Türkçe", "200")
        self.assertEqual((defaults.axis_min, defaults.group_col), ("auto", "Cinsiyet"))
        zero = app.build_options(path, "out", "", "", "", "",
                                 "0'dan başlat (yakınlaştırma yok)", "", "", "200")
        self.assertEqual(zero.axis_min, 0.0)
        for bad in [("", "out", "200"), ("missing.xlsx", "out", "200"), (path, "", "200"),
                    (path, "out", "abc"), (path, "out", "5")]:
            with self.assertRaises(ValueError):
                app.build_options(bad[0], bad[1], "", "", "", "", "", "", "", bad[2])
        with self.assertRaises(ValueError):
            app.build_options(path, "out", "", "", "", "", "Elle gir", "120", "", "200")

    def test_sheet_info_for_drop_downs(self):
        import app
        sheets, headers = app.read_sheet_info(str(write_xlsx(self.ROWS)))
        self.assertEqual(sheets, ["Sheet"])
        self.assertEqual(headers, ["Cinsiyet", "A", "B", "C"])

    def test_window_runs_end_to_end(self):
        """Drive the real window: choose a file, press the button, wait for the charts."""
        import os
        import time
        import app
        try:
            import tkinter
            os.environ["HYROX_GUI_SMOKE"] = "1"
            root = app.run_gui()
        except Exception as exc:  # no display available
            self.skipTest(f"no GUI available: {exc}")
        finally:
            os.environ.pop("HYROX_GUI_SMOKE", None)
        try:
            path = write_xlsx(self.ROWS)
            out = path.parent / "grafikler"
            ui = root.app
            ui["file"].set(str(path))
            ui["out"].set(str(out))
            ui["refresh"]()
            self.assertEqual(list(ui["group_box"]["values"]),
                             [app.AUTO, "Cinsiyet", "A", "B", "C"])
            ui["toggle_adv"]()
            ui["poll"]()
            ui["start"]()
            deadline = time.time() + 120
            while ui["state"]["running"] and time.time() < deadline:
                root.update()
                time.sleep(0.02)
            for _ in range(10):
                root.update()
            log = ui["text"].get("1.0", "end")
            self.assertFalse(ui["state"]["running"], log)
            self.assertIn("Tamamlandı: 4 grafik", log)
            self.assertIn("UYARI:", log)  # Ömer's blank cell
            self.assertEqual(sorted(p.name for p in out.rglob("*.png")),
                             ["Ayşe.png", "Işıl.png", "Çağrı.png", "Ömer.png"])
            self.assertEqual(sorted(p.name for p in out.iterdir()), ["Erkek", "Kadın"])
        finally:
            root.destroy()


class ChartHelpers(unittest.TestCase):
    def test_valid_runs_break_at_missing_station(self):
        self.assertEqual(valid_runs([True] * 4), [[0, 1, 2, 3]])
        self.assertEqual(valid_runs([False] * 4), [])
        # Station 5 missing: one open run that wraps around the top.
        self.assertEqual(valid_runs([True] * 5 + [False, True, True]),
                         [[6, 7, 0, 1, 2, 3, 4]])
        self.assertEqual(valid_runs([True, False, True, True, False, True]),
                         [[2, 3], [5, 0]])
        self.assertEqual(valid_runs([False, True, False, True, True]), [[1], [3, 4]])

    def test_axis_floor_and_ticks(self):
        self.assertEqual(auto_floor(68.0), 60)
        self.assertEqual(auto_floor(62.0), 50)
        self.assertEqual(auto_floor(3.0), 0)
        self.assertEqual(auto_floor(100.0), 90)
        self.assertEqual(ring_ticks(0), [20, 40, 60, 80, 100])
        self.assertEqual(ring_ticks(60), [70, 80, 90, 100])
        self.assertEqual(ring_ticks(90), [92, 94, 96, 98, 100])
        for floor in (0, 12.5, 50, 75, 90, 95):
            ticks = ring_ticks(floor)
            self.assertGreaterEqual(len(ticks), 3)
            self.assertTrue(all(floor < t <= 100 for t in ticks))


class FileNames(unittest.TestCase):
    def test_sanitize(self):
        self.assertEqual(safe_filename("Çağrı Öztürk"), "Çağrı Öztürk")
        self.assertEqual(safe_filename('A/B:C*?"<>|'), "ABC")
        self.assertEqual(safe_filename("  ..  "), "participant")
        self.assertEqual(safe_filename("con"), "_con")
        self.assertEqual(unique_filenames(["Ali", "ali", "Veli"]), ["Ali", "ali_2", "Veli"])


if __name__ == "__main__":
    unittest.main()
