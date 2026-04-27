# WedstrijdCalculator

A local Python tool for scoring sailing competitions using the **Portsmouth Yardstick (PY)** handicap system. Supports both a command-line interface and a modern graphical interface.

> Dutch documentation: see [HANDLEIDING.md](HANDLEIDING.md)

---

## Features

- Portsmouth Yardstick corrected-time calculation
- Drop-worst result (schrapresultaat) support
- Built-in PY table for common dinghy and catamaran classes
- Extend with your own boat types via `extra_boottypes.csv`
- Input via CSV or Excel file
- Export results to a formatted Excel workbook (3 tabs: standings, race detail, position matrix)
- Interactive GUI built with `customtkinter`
- Demo mode with built-in sample data

---

## Requirements

Python 3.10 or higher.

```
pip install pandas openpyxl customtkinter
```

---

## Files

| File | Description |
|---|---|
| `wedstrijd_calculator.py` | Core scoring engine + CLI |
| `wedstrijd_gui.py` | Graphical user interface |
| `voorbeeld_invoer.csv` | Sample input file (6 participants, 5 races) |
| `extra_boottypes.csv` | Optional: add custom boat types / PY values |
| `wedstrijd_resultaten.xlsx` | Generated after running (not committed) |

---

## Usage

### GUI (recommended)

```bash
python wedstrijd_gui.py
```

### CLI - interactive menu

```bash
python wedstrijd_calculator.py
```

### CLI - direct processing

```bash
python wedstrijd_calculator.py --invoer voorbeeld_invoer.csv
python wedstrijd_calculator.py --invoer mijn_wedstrijd.xlsx
python wedstrijd_calculator.py --demo
```

### CLI options

| Option | Description |
|---|---|
| `--invoer` / `-i` | Path to CSV or Excel input file |
| `--uitvoer` / `-o` | Output Excel file name (default: `wedstrijd_resultaten.xlsx`) |
| `--schrap` | Apply drop-worst rule (default: on) |
| `--geen-schrap` | Disable drop-worst rule |
| `--extra-py` | CSV file with additional boat type / PY values |
| `--demo` | Run with built-in sample data |

---

## Input Format (CSV or Excel)

Required columns (case-insensitive):

| Column | Description |
|---|---|
| `naam` | Participant name |
| `boottype` | Boat class (must match a PY entry) |
| `reeks` | Race number (1, 2, 3, ...) |
| `minuten` | Elapsed time - minutes |
| `seconden` | Elapsed time - seconds |

Optional column:

| Column | Description |
|---|---|
| `py` | Manual PY override (overrides the built-in table) |

See `voorbeeld_invoer.csv` for a ready-to-use example.

---

## PY Formula

```
corrected_time = elapsed_time_seconds * 1000 / PY
```

Lower corrected time = better performance. Race ranking is based on corrected time.

---

## Scoring

- 1st place = 1 point
- 2nd place = 2 points
- etc.

Lowest total points wins. With drop-worst active, each participant's single worst race is excluded from their total.

---

## Customizing Boat Types

Edit the `BOAT_PY` dictionary at the top of `wedstrijd_calculator.py`, or add entries to `extra_boottypes.csv` without touching the source code:

```
boottype,py
MijnBoot,1050
```

---

## License

MIT
