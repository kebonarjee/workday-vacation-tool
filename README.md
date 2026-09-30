# Workday Vacation Tool

[![CI](https://github.com/kebonarjee/workday-vacation-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/kebonarjee/workday-vacation-tool/actions/workflows/ci.yml)

Convert a Workday absence export into an aligned Markdown vacation summary
and, optionally, all-day **Out of Office** events for Outlook. Files are
processed locally.

## Installation

Requires **Python 3.10 or newer**. Microsoft Excel is not required.

Clone the repository:

```console
git clone https://github.com/kebonarjee/workday-vacation-tool.git
cd workday-vacation-tool
```

Create a virtual environment and install the dependencies.

**Windows (PowerShell):**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

If PowerShell blocks activation, use `.\.venv\Scripts\python` for the `pip`
and script commands instead of `python`.

**macOS or Linux:**

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
```

## Export your absences from Workday

1. Sign in to your organization's Workday tenant.
2. Open **Manage Absence**.
3. In the **Actions** menu, select **My Absence**.
4. In the displayed absence table, click **Export to Excel** to save an
   `.xlsx` workbook.

The export must contain `Date`, `Type`, and `Status` columns, in any order.
Dates can be Excel date cells or `YYYY-MM-DD` text.

## Usage

Run the script with your export's filename. Replace `AbsenceRequests.xlsx`
in the examples with the name of your file.

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx
```

By default, the script prints approved vacation dated today or later to the
terminal, using your computer's local date.

Example summary:

```markdown
# Forthcoming vacation

| Date range               | Vacation days |
| ------------------------ | ------------: |
| Mon 6 - Wed 8 May 2030   |             3 |
| Fri 10 - Mon 13 May 2030 |             2 |
| **Total**                |         **5** |
```

### Save the summary

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --output vacation.md
```

Use `--output` to write a UTF-8 Markdown file. Shell redirection is also supported:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx > vacation.md
```

Existing output files are overwritten.

### Date ranges

Consecutive Monday-Friday vacation dates are grouped together. For example,
a Friday and the following Monday form one range with two vacation days.

Use `--grouping calendar` to group only consecutive calendar dates:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --grouping calendar
```

Use `--from-date YYYY-MM-DD` to change the starting date, or `--include-past`
to include all matching dates. Duplicate dates are counted once.

### Export an Outlook calendar

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --ics
```

This creates `AbsenceRequests.ics` beside the workbook and still prints the
Markdown summary. Each vacation range becomes an all-day **Out of Office**
event titled **Planned vacation**, covering the entire range, including any
weekends between vacation days.

Use `--ics PATH` to choose a filename or `--ics-title TEXT` to change the event
title.

### Import into Outlook

- **Outlook on the web or new Outlook:** Open
  **Calendar > Add calendar > Upload from file** and select the `.ics` file
  and destination calendar.
- **Classic Outlook for Windows:** Open
  **File > Open & Export > Import/Export** and import the iCalendar file.

Import into the calendar used for your work availability. Confirm that each
imported event has **Show as** set to **Out of Office**.

The import is a snapshot, not a live link to Workday. Importing again may
create duplicate events. Automatic email replies must be configured separately.

## Command-line options

Run `python workday_vacation_tool.py --help` for full help.

| Option | Purpose |
| --- | --- |
| `input` | Required path to an `.xlsx` workbook |
| `-o PATH`, `--output PATH` | Write Markdown to a file instead of stdout |
| `--sheet NAME` | Select a worksheet; default is the first with the required headers |
| `--absence-type VALUE` | Absence type to include; default `Vacation` |
| `--status VALUE` | Absence status to include; default `Approved` |
| `--from-date YYYY-MM-DD` | Earliest date to include; default today |
| `--include-past` | Include all matching dates; cannot combine with `--from-date` |
| `--grouping MODE` | Use `workdays` (default) or `calendar` |
| `--title TEXT` | Markdown heading; default `Forthcoming vacation` |
| `--ics [PATH]` | Export a calendar; default is the workbook path with an `.ics` suffix |
| `--ics-title TEXT` | Calendar event title; default `Planned vacation` |

## Limitations

- Each unique matching date is treated as a full vacation day. Partial-day
  and hourly absences are not supported.
- Workday grouping assumes a Monday-Friday schedule and does not account for
  public holidays or custom workweeks.
- The required English column headers must appear within the first 50 rows
  of a worksheet.

## Development

From the repository root with your virtual environment active:

```console
python -m pip install -r requirements-dev.txt
python -m ruff format .
python -m ruff check .
python -m ruff format --check .
python -m pytest
```

GitHub Actions runs lint, formatting checks, and tests on Linux and Windows.

## License

[MIT](LICENSE). This project is independent of and not affiliated with Workday
or Microsoft.
