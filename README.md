# Hyrox Radar Charts

Turns an Excel sheet of station results into one radar chart per participant.

Each chart shows:

- the participant's result at every station as a yellow outline, where **further
  out always means better**,
- the average of their group as a grey shape behind it,
- a table with the real times, reps or metres, so nobody has to guess what a
  score means.

The charts are PNG files sized for sharing in a group chat.

## Using the app (no Python needed)

`HyroxRadar.exe` is a single file for Windows. Copy it to any computer and
double-click it. Nothing has to be installed.

1. Click **Seç...** next to "Excel dosyası" and pick the score sheet.
2. The output folder is filled in for you: a `grafikler` folder next to the
   Excel file. Change it if you like.
3. Click **Grafikleri oluştur**. The box below shows what the app understood
   (sheet, categories, station types) and every warning, in orange.
4. Click **Klasörü aç** to see the charts.

**Gelişmiş ayarlar** opens the same choices as the command line: sheet,
category column, comparison mode, score scale, axis start, chart language and
resolution. The defaults are right for a normal competition day.

Things to tell the people you send it to:

- **First start.** Windows may show a "Windows protected your PC" screen,
  because the file is not signed by a publisher. Click **More info**, then
  **Run anyway**. Some antivirus tools also flag unsigned single-file apps.
- **It takes a few seconds to open.** The file unpacks itself on every start.
- **Windows only.** It does not run on a Mac or a phone.
- Close any chart that is open in an image viewer before running again, or the
  app cannot overwrite it.

The report box is in English for now; the buttons, labels and error pop-ups are
in Turkish.

### Building the app yourself

```powershell
.\build.ps1
```

This installs PyInstaller into `.venv` and writes `dist\HyroxRadar.exe`
(about 50 MB). Rebuild after any code change, then send the new file.

To check a build without clicking through the window, start it with arguments.
It then behaves like `radar.py`:

```powershell
dist\HyroxRadar.exe sample_scores.xlsx --out charts\exe_check
```

You can also start the window from source with `python app.py`.

## Contents

1. [Quick start](#quick-start)
2. [Preparing the Excel file](#preparing-the-excel-file)
3. [Entering times correctly](#entering-times-correctly)
4. [Categories: women and men](#categories-women-and-men)
5. [How scores are calculated](#how-scores-are-calculated)
6. [Reading a chart](#reading-a-chart)
7. [Zooming in with the axis floor](#zooming-in-with-the-axis-floor)
8. [All options](#all-options)
9. [Warnings and errors](#warnings-and-errors)
10. [Project layout and tests](#project-layout-and-tests)

## Quick start

All commands are for PowerShell, run from the project folder.

**1. Install once**

```powershell
uv venv .venv
uv pip install --python .venv/Scripts/python.exe -r requirements.txt
```

No uv? Use `python -m venv .venv`, activate it, then `pip install -r requirements.txt`.

**2. Activate the environment** (every new terminal)

```powershell
.venv\Scripts\activate
```

**3. Make the charts**

```powershell
python radar.py sample_scores.xlsx --out output/
```

`sample_scores.xlsx` in the project root is the reference file. Replace it with
your own file name for a real competition day.

The tool prints what it understood before it draws anything:

```
Sheet: Skorlar  |  8 participants, 8 stations
Category column: Cinsiyet
Categories: Kadın (4), Erkek (4)
Compare: within each category
Scale: pct-of-best (...)
Axis floor: 60 (auto; lowest score is 74; use --axis-min 0 to start at 0)
Chart language: tr
Stations:
  - SkiErg: time, lower is better
  - Wall Balls: time, lower is better
  ...
```

Read this block every time. If a station shows the wrong type or direction, the
charts will be wrong too, and this is where you catch it.

## Preparing the Excel file

| Katılımcı | Cinsiyet | SkiErg | Sled Push | Wall Balls (tekrar) |
| --- | --- | --- | --- | --- |
| Ahmet Yılmaz | E | 4:12 | 2:30 | 55 |
| Elif Çelik | Kadın | 4:40 | 3:50 | 48 |
| Zeynep Kaya | K | 4:25 | DNF | 51 |

The rules:

- **Row 1 is the header row.** The header texts become the axis labels, so
  write them the way participants should read them. Put units in the header if
  you want them shown, for example `Row (m)`.
- **Column A is the participant name.**
- **Every other column is one station**, except the optional category column
  (see [Categories](#categories-women-and-men)).
- **The number of stations is free.** Add or remove columns from day to day.
  The chart gets exactly that many axes. The minimum is 3.
- **The first sheet is used** unless you pass `--sheet`.

### How the tool decides what a column means

| What the column contains | Read as | Better is |
| --- | --- | --- |
| Times such as `4:32`, `04:32`, `1:04:32`, `0:48.5`, or Excel time cells | time | lower |
| Plain numbers (reps, metres, kilos) | number | higher |

A column must be one kind. A column that mixes times and plain numbers is
rejected with a message naming the cells.

If the tool guesses a direction wrong, for example a "penalty points" column
where lower is better, override it:

```powershell
python radar.py scores.xlsx --lower-better "Penalty Points"
python radar.py scores.xlsx --higher-better "Some Station"
```

Both flags can be repeated for several stations. The name must match the header
exactly.

### Missing results

A cell counts as "no result" when it is blank or contains one of
`DNF`, `DNS`, `DQ`, `N/A`, `-`, `x`.

For that participant and station:

- the axis label turns grey and says "(sonuç yok)",
- the outline **stops** at the neighbouring stations instead of cutting across,
  so the chart never suggests a score that does not exist,
- the table shows `DNF` or a dash,
- the result is left out of the best value and the average,
- a warning names the participant, the station and the cell.

## Entering times correctly

This is the one place where Excel can silently change your data.

**The problem.** Type `4:32` into a normal Excel cell and Excel stores
**4 hours 32 minutes**, not 4 minutes 32 seconds. It also gives the cell an
hours:minutes format.

**What the tool does about it.**

1. Any Excel time cell with an hours:minutes format and no seconds is read as
   minutes:seconds. This is decided cell by cell, so a column containing `0:48`
   and `1:03` is read as 48 s and 63 s.
2. For other formats, if every Excel time cell in a column is at least one hour
   with zero seconds, the column is read as minutes:seconds too.

Both cases **always print a warning**, so a conversion never happens unnoticed:

```
WARNING: Station 'Wall Balls': 8 Excel time cell(s) have an hours:minutes format
without seconds (e.g. J2 shows 5:50) ... They were read as minutes:seconds
(5:50 = 5 min 50 s).
```

**How to avoid the guesswork entirely.** Pick one:

- Format the column as **Text** before typing, then enter `4:32`.
- Type the full form `0:04:32`. Excel stores that correctly.

**If a station really is measured in hours.** Type the seconds as well, for
example `1:05:00`, so it is not mistaken for 1 minute 5 seconds.

## Categories: women and men

Add a column that says which category each participant is in. The tool then
compares women with women and men with men.

**Naming the column.** A column called `Cinsiyet`, `Kategori`, `Gender` or
`Category` is found automatically, in any letter case and anywhere after the
name column. For any other header, name it yourself:

```powershell
python radar.py scores.xlsx --group-col "Lig"
```

**Writing the values.** These all mean the same thing:

| Category | Accepted spellings |
| --- | --- |
| Kadın | `K`, `Kadın`, `KADIN`, `Kadin`, `Bayan`, `Kız`, `F`, `Female`, `Woman`, `Women` |
| Erkek | `E`, `Erkek`, `ERKEK`, `Bay`, `M`, `Male`, `Man`, `Men` |

Letter case does not matter and Turkish letters are handled properly, so
`KADIN` and `Kadın` match. Spaces before or after a value are ignored.

Any other value becomes its own category. That means the same column works for
age groups or `Pro` / `Open`. A blank cell goes into an "unspecified" category
and prints a warning.

**What changes when a category column exists.**

- The best result and the average come from the participant's own category.
  The fastest woman gets 100 and the fastest man gets 100.
- The chart shows "Kategori: Kadın" under the name, and the legend and table say
  "Kadın ortalaması" instead of "Grup ortalaması".
- Charts are written to one subfolder per category:

```
output\
  Kadın\   Elif Çelik.png, Zeynep Kaya.png, ...
  Erkek\   Ahmet Yılmaz.png, Can Öztürk.png, ...
```

**Comparing everyone together anyway.**

```powershell
python radar.py scores.xlsx --compare all
```

This ignores the categories for scoring: one best result, one average, one
output folder.

## How scores are calculated

Raw results cannot share one chart, because a time, a rep count and a distance
have different units and directions. So every station is converted to a score
from 0 to 100 where higher is always better. Each station is scored on its own.

### Default: percentage of the best (`--scale pct-of-best`)

The best result in the group gets 100. Everyone else gets their percentage of it.

| Station type | Formula |
| --- | --- |
| Time (lower is better) | best time ÷ your time × 100 |
| Reps, distance (higher is better) | your value ÷ best value × 100 |

Example from the reference file, Wall Balls, men: the fastest time is 4:50
(290 s). Ahmet did 5:50 (350 s). His score is 290 ÷ 350 × 100 = **83**.

This scale is proportional to real performance: being 10% slower costs about 10
points, however the rest of the group did.

### Alternative: min-max (`--scale minmax`)

The best result gets 100 and the **worst gets 0**, with everyone else spread in
between.

It makes differences very visible, but it has two drawbacks. A gap of a few
seconds can become a gap of 100 points. And whoever is last at every station
scores 0 everywhere, so their chart collapses to a dot. That is why it is not
the default.

### Special cases

- **Everyone has the same result** at a station: everyone gets 100.
- **No result**: no score, and the result is ignored for the best and the average.
- **A category with one participant**: every score is 100, and a warning says so.

### The grey shape

The average is the mean of the **scores** at each station, not of the raw times.
Blank cells and DNF/DNS are left out. With categories it is the average of the
participant's own category.

## Reading a chart

From top to bottom:

1. **Name**, and the category when categories are in use.
2. **The radar.** One axis per station, in the same order and position on every
   chart, so two charts can be compared side by side. Rings are labelled with
   their score. The yellow outline is the participant; the grey shape is the
   average. Yellow outside grey means better than average at that station.
3. **Legend.**
4. **Table**, one row per station: the real result, the participant's score, and
   the average score.
5. **Footnotes** stating the scale used and how the average is calculated.

## Zooming in with the axis floor

With percentage-of-best, scores often sit between 70 and 100. Drawn on a 0 to
100 axis, everyone's shape looks similar and the middle of the chart is empty.

So by default the axis starts at a higher value, which zooms in on the range
where the scores actually are. **Scores and the table do not change, only the
drawing.**

```powershell
python radar.py scores.xlsx                 # auto floor (default)
python radar.py scores.xlsx --axis-min 50   # your own floor
python radar.py scores.xlsx --axis-min 0    # zoom off, axis starts at 0
```

- `auto` (the default) picks the floor: a multiple of 10 that is at least 5
  points below the lowest score. With a lowest score of 74 the floor is 60.
  When scores are widely spread the floor comes out at or near 0, so the zoom
  only applies when it is needed.
- A number sets the floor yourself, from 0 to 95. Scores below it are drawn at
  the centre and a warning is printed.
- `0` turns the zoom off.
- One floor is used for all categories, so every chart stays on the same axis.

**The trade-off.** Zooming makes gaps look bigger than they are. To keep the
chart honest, the floor value is printed at the centre and a bold footnote says
the axis does not start at 0. If the charts are used for an official ranking
rather than for feedback, use `--axis-min 0`.

## All options

```powershell
python radar.py FILE.xlsx [options]
```

| Option | Default | What it does |
| --- | --- | --- |
| `--out DIR` | `output` | Folder for the PNG files. Created if missing. |
| `--sheet NAME` | first sheet | Sheet to read: its name, or its number starting at 1. |
| `--scale pct-of-best` or `minmax` | `pct-of-best` | How scores are calculated. |
| `--axis-min N` or `auto` | `auto` | Where the radial axis starts. `0` turns the zoom off. |
| `--group-col "Header"` | auto | Header of the category column. |
| `--compare group` or `all` | `group` | Score inside each category, or everyone together. |
| `--lang tr` or `en` | `tr` | Language of the text on the charts. |
| `--lower-better "Station"` | auto | Force lower-is-better for a station. Repeatable. |
| `--higher-better "Station"` | auto | Force higher-is-better for a station. Repeatable. |
| `--dpi N` | `200` | Image resolution. |

A typical competition-day command:

```powershell
python radar.py scores.xlsx --out charts\2026-11-12
```

Notes:

- `--lang` changes only the text on the charts. Console messages stay in English.
- Station names are shown exactly as typed in the header and are never translated.
- File names are the participant names. Only characters Windows forbids are
  removed, so Turkish letters are kept. Two participants with the same name get
  `_2`, `_3` added.
- The image gets wider automatically when a station name is long, and a very
  long participant name is shown in a smaller title font.

## Warnings and errors

**Warnings** do not stop the run. They tell you about something the tool decided
for you, so check that the decision was right.

| Warning | Meaning |
| --- | --- |
| `blank cell for 'X'; axis point skipped` | That participant has no result at X. |
| `'DNF' for 'X'; treated as not attempted` | Same, with a did-not-finish marker. |
| `hours:minutes format without seconds ... read as minutes:seconds` | Hand-typed Excel times were converted. See [Entering times correctly](#entering-times-correctly). |
| `blank 'Cinsiyet' cell; placed in the 'unspecified' category` | A participant has no category. |
| `Category 'X' has one participant` | Every score in that category is 100. |
| `score(s) are below the axis floor` | Your `--axis-min` is higher than some scores. |
| `Station 'X' has no results at all` | The column is empty and was left out. |

**Errors** stop the run before any chart is written. Each one names the cell or
column to fix.

| Error | Fix |
| --- | --- |
| `File not found` / `Expected an .xlsx file` | Check the path. Old `.xls` files must be re-saved as `.xlsx`. |
| `Sheet 'X' not found` | The message lists the sheets that exist. |
| `cannot read B5 ('fast')` | That cell is neither a number, a time, nor a no-result marker. |
| `Hint: if 'X' is a category column ...` | The column holds words. Run again with `--group-col "X"`. |
| `mixes time values and plain numbers` | Make the whole column one kind. |
| `Duplicate station name(s)` | Give each station a unique header. |
| `has data but no station name` | A column has values but an empty header. |
| `has scores but no participant name` | A row has values but an empty name cell. |
| `needs at least 3 stations` | A radar chart cannot be drawn with fewer. |
| `unknown station(s)` for an override | The name must match the header exactly. |

## Project layout and tests

```
radar.py              command-line entry point
app.py                window app (the source of HyroxRadar.exe)
build.ps1             builds dist\HyroxRadar.exe
hyrox_radar\
  loader.py           reads the sheet, detects column types, times and categories
  normalize.py        turns results into 0-100 scores and averages
  chart.py            draws the radar, the table and the footnotes
  runner.py           the full pipeline, shared by radar.py and app.py
tests\test_core.py    unit tests for the logic above
sample_scores.xlsx    reference file for manual runs
requirements.txt      pandas, openpyxl, matplotlib
```

Run the tests:

```powershell
python -m unittest discover -s tests -v
```

The tests build their own small workbooks in the system temp folder and delete
them afterwards. They never read or change `sample_scores.xlsx`.
