# Workday Vacation Tool

[![CI](https://github.com/kebonarjee/workday-vacation-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/kebonarjee/workday-vacation-tool/actions/workflows/ci.yml)

Turn a Workday absence export into an at-a-glance vacation summary and,
optionally, an Outlook-compatible calendar.

The tool selects approved vacation from today onward, combines consecutive
workdays into ranges, and prints an aligned Markdown table. It can also generate
one `.ics` file containing multiple all-day **Planned vacation** events marked
**Out of Office** for Outlook.

This is a standalone Python script, not a Workday integration. It processes
files locally and does not upload your absence data or access your accounts.

## Requirements and installation

Python **3.10 or newer** is required. The only runtime dependency is
[`openpyxl`](https://openpyxl.readthedocs.io/); Microsoft Excel is not required.

Clone the repository:

```console
git clone https://github.com/kebonarjee/workday-vacation-tool.git
cd workday-vacation-tool
```

Create and activate a virtual environment in Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

If PowerShell prevents activation, use the environment's Python directly:

```powershell
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python workday_vacation_tool.py --help
```

On macOS or Linux:

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
```

You can also share just `workday_vacation_tool.py`; its module documentation
includes installation instructions for its dependency.

## Prepare your export

Export your absence list from Workday:

1. Sign in to your organization's Workday tenant.
2. Open **Manage Absence**.
3. In the **Actions** menu, select **My Absence**.
4. In the displayed absence table, click **Export to Excel** and save the
   **`.xlsx`** workbook.

Keep the export on your computer; do not commit it to this public repository.

The workbook must have these headers, in any column order:

| Header | Expected values |
| --- | --- |
| `Date` | An Excel date cell or `YYYY-MM-DD` text |
| `Type` | `Vacation` by default |
| `Status` | `Approved` by default |

The script searches the first 50 rows for the header, so a title such as
`Absence Requests` can appear above it. Extra columns and blank trailing rows
are allowed. Header names and filter values are case-insensitive; extra
whitespace is ignored.

By default, the first worksheet with these headers is selected. Use `--sheet`
to choose another worksheet.

## Quick start

Print the Markdown summary to standard output:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx
```

Example output using fictional dates:

```markdown
# Forthcoming vacation

| Date range               | Vacation days |
| ------------------------ | ------------: |
| Mon 6 - Wed 8 May 2030   |             3 |
| Fri 10 - Mon 13 May 2030 |             2 |
| **Total**                |         **5** |
```

The columns align in a terminal or plain-text editor as well as in rendered
Markdown. The real output also includes a note explaining the grouping.

### Write Markdown to a file

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --output output\vacation.md
```

Shell redirection is also supported:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx > vacation.md
```

Explicit `--output` writes UTF-8 regardless of the shell's redirection encoding.
Output directories are created if needed. Existing output files are overwritten.
Errors and file-written notifications go to stderr, leaving stdout safe to pipe.

### Generate an Outlook calendar

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --ics
```

This still prints Markdown and also writes `AbsenceRequests.ics` alongside the
workbook. A single iCalendar file can contain multiple events; no ZIP is needed.

To choose both output paths:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx `
  --output output\vacation.md `
  --ics output\vacation.ics
```

Each event:

- Is an all-day event spanning the corresponding displayed range.
- Is titled **Planned vacation**, unless `--ics-title` is provided.
- Uses Outlook's `X-MICROSOFT-CDO-BUSYSTATUS:OOF` and
  `X-MICROSOFT-CDO-INTENDEDSTATUS:OOF` properties, with `TRANSP:OPAQUE` for busy
  status in other compatible calendar applications.

For Outlook on the web or new Outlook, open **Calendar**, then **Add calendar**
and **Upload from file**. Select the generated `.ics` file and import it into
the calendar your colleagues use to check your availability. In classic Outlook
for Windows, use **File > Open & Export > Import/Export** and select the
iCalendar file.

Verify the imported dates and **Show as: Out of Office** in your Outlook client.
Other clients may interpret the Outlook-specific properties as simply busy.
Importing these events does not enable automatic email replies or synchronize
future Workday changes. Re-importing may create duplicates; event IDs are
deterministic for unchanged filenames, worksheet names, titles, and ranges,
but client import behavior varies.

### Date filtering and grouping

Past dates are ignored automatically, so the export does not need precleaning.
Today is included, using the computer's local date.

Override the cutoff or include past dates:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --from-date 2030-01-01
python workday_vacation_tool.py AbsenceRequests.xlsx --include-past
```

The default `--grouping workdays` treats Friday followed by Monday as continuous.
Weekend dates bridged by a range are not added to the vacation-day count, but
**the corresponding calendar event covers those weekends**.

Use strict calendar adjacency to keep Friday and Monday as separate events:

```powershell
python workday_vacation_tool.py AbsenceRequests.xlsx --grouping calendar --ics
```

## Command-line options

Run `python workday_vacation_tool.py --help` for the complete CLI help.

| Option | Purpose |
| --- | --- |
| `input` | Required path to an `.xlsx` workbook |
| `-o PATH`, `--output PATH` | Write Markdown to a file instead of stdout |
| `--sheet NAME` | Select a worksheet |
| `--absence-type VALUE` | Filter `Type`; default `Vacation` |
| `--status VALUE` | Filter `Status`; default `Approved` |
| `--from-date YYYY-MM-DD` | Override the inclusive cutoff; default today |
| `--include-past` | Disable the date cutoff; cannot combine with `--from-date` |
| `--grouping workdays\|calendar` | Choose adjacency; default `workdays` |
| `--title TEXT` | Markdown heading; default `Forthcoming vacation` |
| `--ics [PATH]` | Also write iCalendar; default filename has an `.ics` suffix |
| `--ics-title TEXT` | Calendar event title; default `Planned vacation` |

If using `--ics` without a path, put the workbook argument before `--ics`.

## Limitations and privacy

The tool assumes a Monday-Friday workweek and full-day vacation. Each unique
matching date counts as one day; the export's `Requested` and `Unit of Time`
columns are not used. Half days, hourly absence, public holidays, custom
workweeks, and recurring events are not calculated.

The workbook must use the English header names shown above. Numeric cells
without Excel date formatting and ambiguous localized date strings are rejected.
Formula cells require cached values from Excel because `openpyxl` does not
calculate formulas.

Malformed matching rows and empty results are reported as errors with a nonzero
exit code, rather than producing a misleading empty vacation summary.

The `.gitignore` excludes Excel exports, calendar files, and the `output`
directory. Markdown summaries also contain personal information; keep them in
`output` or outside the repository, and review your staged files before committing.

## Development

From the repository root with your virtual environment active:

```console
python -m pip install -r requirements-dev.txt
python -m ruff format .
python -m ruff check .
python -m ruff format --check .
python -m pytest
```

There is no compilation or packaging step for this standalone script.

Tests generate synthetic Excel workbooks in temporary directories. They cover
filtering, grouping, raw Markdown alignment, stdout/file modes, error handling,
and calendar parsing with the `icalendar` library, including UTF-8 folding and
Outlook OOF properties. No real absence exports are included.

GitHub Actions runs lint and formatting checks plus the tests on Linux and
Windows with Python 3.10 and 3.14. Actions are pinned to immutable commit hashes,
workflow permissions are read-only, and Dependabot checks dependencies weekly.

## License

[MIT](LICENSE). This project is independent of and not affiliated with Workday
or Microsoft.
