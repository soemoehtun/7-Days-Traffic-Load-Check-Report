# 7-Day Network Traffic Report Automation

An Excel VBA tool for importing Huawei and Ericsson network traffic reports, calculating UMTS/LTE load indicators, and generating a consolidated summary.

The report is generated inside the macro-enabled workbook. Source files are opened read-only and closed without saving changes.

## Important: Back Up Your Workbook

Before running `ReportProcess`, be aware that:

- All worksheets except `Control` and `EP Information` are deleted **before the file-selection dialog opens**.
- Canceling file selection does not restore the deleted worksheets.
- During summary generation, entire rows **4–1000** on `Control` are deleted, including any unrelated content in those rows.
- `EP Information` is preserved, but its formatting may be changed.

Use a dedicated reporting workbook and keep a backup.

## Requirements

- Microsoft Excel desktop for Windows with VBA support.
- A macro-enabled workbook, such as `.xlsm`.
- A worksheet named exactly `Control`.
- Source reports with the expected column layouts.

The `EP Information` worksheet is optional. Its contents are not used for calculations or lookups.

The code uses a late-bound `Scripting.Dictionary`, so no manual VBA reference needs to be added.

## Installation

1. Open or create an Excel workbook.
2. Save it as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Create a worksheet named `Control`.
4. Optionally create an `EP Information` worksheet.
5. Press **Alt + F11** to open the VBA editor.
6. Select **Insert → Module**.
7. Paste the VBA code into the module.
8. Prepare the summary headings in rows 1–3 of `Control`, using the column mapping below.
9. Save the workbook.

The macro writes summary data starting at row 4; it does not create the summary headings.

## Input Files

Files are identified by case-insensitive text matching in their filenames.

| Dataset | Filename must contain | Source columns copied | Destination worksheet | Destination columns |
|---|---|---|---|---|
| Huawei UMTS / 3G | `hw_data_update_3g_nbh_cell` | `A, D, E, I, J, K` | `HW UMTS` | `A:F` |
| Ericsson UMTS / 3G | `eric_data_update_3g_nbh_cell` | `A, C, F, G` | `Eric UMTS` | `A:D` |
| Huawei LTE / 4G | `hw_data_update_4g_dbh` | `A, E, I:M` | `HW LTE` | `A:G` |
| Ericsson LTE / 4G | `eric_data_update_4g_dbh` | `A:H, J:K` | `Eric LTE` | `A:J` |

### Source Data Assumptions

- The first sheet in each source file contains the data and must be a worksheet.
- Row 1 contains headers; records begin in row 2.
- Column A contains the grouping key used in the summary.
- Grouping keys must be consistent across all four datasets.
- The last imported row is determined from column A.
- Columns are selected by position, not by header name.
- Files should already contain the intended reporting period. The macro does not filter dates or verify that exactly seven days are present.

### File Selection Behavior

- Select all required files in one operation.
- For a complete report, select one file for each dataset.
- If multiple selected files match the same dataset, the last matching file is used; files are not merged.
- Unmatched files are ignored.
- Missing datasets are not rejected. Their destination worksheets are still created but remain empty.

## Usage

1. Make the macro-containing workbook the active workbook.
2. Press **Alt + F8**.
3. Select `ReportProcess`.
4. Click **Run**.
5. In the **Import 7Days Traffic Data** dialog, select the source files.
6. Review the generated worksheets and the `Control` summary.
7. Save the workbook manually.

No separate report file is exported automatically.

> The current completion dialog displays a success message even after cancellation or a handled runtime error. Always verify that the expected outputs were generated.

## Processing Workflow

`ReportProcess` performs the following steps:

1. Disables screen updating, alerts, events, and automatic calculation.
2. Deletes previous report worksheets.
3. Imports the selected source files.
4. Calculates throughput, utilization, and load flags.
5. Aggregates records by grouping key.
6. Writes the summary to `Control`.
7. Formats non-Control worksheets.
8. Restores the captured calculation mode and re-enables events, alerts, and screen updating.
9. Displays elapsed processing time.

## Calculation Rules

Column references below refer to the **destination worksheets after import**, not the original source files.

A load flag is `1` when its condition is true and `0` otherwise, subject to the missing-value and error handling described below.

### Huawei UMTS — `HW UMTS`

| Column | Metric | Formula or condition |
|---|---|---|
| G | HSDPA user throughput | `D / (E * 2)` |
| H | Mean TCP in watts | `10^(F / 10) / 1000` |
| I | Maximum cell transmit power in watts | `10^(C / 100) / 1000` |
| J | TCP utilization (%) | `H / I * 100` |
| K | Throughput-only load flag | `G < 1000` |
| L | Combined load flag | `G < 1000 AND J > 0.8` |

**Threshold note:** The combined Huawei UMTS check uses `J > 0.8` exactly as coded, even though J is multiplied by 100. It does **not** use `J > 80`. Confirm this threshold against the intended reporting rule.

### Ericsson UMTS — `Eric UMTS`

Imported column C contains throughput, and column D contains utilization.

| Column | Metric | Condition |
|---|---|---|
| E | Throughput-only load flag | `C < 1000` |
| F | Combined load flag | `C < 1000 AND D > 80` |

If throughput in column C is blank, both flags are left blank after the temporary `NA` values are removed.

### Huawei LTE — `HW LTE`

| Column | Metric | Formula or condition |
|---|---|---|
| H | Throughput (kbps) | `(C - D) / E` |
| I | PRB utilization (%) | `F / G * 100` |
| J | Load flag | `H < 2000 AND I > 70` |

### Ericsson LTE — `Eric LTE`

| Column | Metric | Formula or condition |
|---|---|---|
| K | Throughput (kbps) | `((C - D) / E) * 1000` |
| L | PRB utilization (%) | `((F + G + H + I) / J) * 100` |
| M | Load flag | `K < 2000 AND L > 70` |

### Calculated Values and Errors

- Calculated formulas are converted to static values.
- `#DIV/0!` values are replaced with blanks in the calculated ranges for Huawei UMTS, Huawei LTE, and Ericsson LTE.
- Temporary `NA` text is replaced with blanks in Ericsson UMTS load flags.
- Other error types are not explicitly handled.
- Editing imported data afterward does not automatically update the calculated results.

## Summary Output

Summary data is written to `Control`, beginning at row 4.

### UMTS Summary — Columns I:P

| Column | Suggested heading | Description |
|---|---|---|
| I | Grouping Key | Key from Huawei UMTS column A |
| J | HW UMTS Total | Huawei UMTS record count |
| K | HW UMTS Low Throughput | Count of Huawei UMTS throughput-only flags |
| L | HW UMTS Combined Load | Count of Huawei UMTS combined flags |
| M | Eric UMTS Total | Ericsson UMTS record count |
| N | Eric UMTS Low Throughput | Count of Ericsson UMTS throughput-only flags |
| O | Eric UMTS Combined Load | Count of Ericsson UMTS combined flags |
| P | UMTS Load Ratio | `(L + O) / (J + M)` |

### LTE Summary — Columns R:W

| Column | Suggested heading | Description |
|---|---|---|
| R | Grouping Key | Same key list used by the UMTS summary |
| S | HW LTE Total | Huawei LTE record count |
| T | HW LTE Load | Count of Huawei LTE load flags |
| U | Eric LTE Total | Ericsson LTE record count |
| V | Eric LTE Load | Count of Ericsson LTE load flags |
| W | LTE Load Ratio | `(T + V) / (S + U)` |

### Aggregation Behavior

- Summary keys come **only from non-empty Huawei UMTS column A values**.
- Keys found only in Ericsson UMTS or either LTE dataset are excluded.
- Without Huawei UMTS records, no new summary rows are written.
- Matching is case-insensitive; leading and trailing spaces are not removed.
- Totals count records, not unique cells or unique days.
- Records with blank load flags still contribute to their dataset's total.
- Duplicate records are not removed.
- A zero denominator produces `#DIV/0!` in the summary ratio.

Columns P and W contain fractional ratios. Apply percentage formatting on `Control` if percentage display is required.

## Formatting

All worksheets except `Control` receive:

- Calibri, size 9
- Centered horizontal and vertical alignment
- Thin borders
- Bold, shaded headers
- Header row height of 20
- Two-decimal number formatting from column C onward
- Automatically fitted column widths

This includes `EP Information` when present.

## Performance Approach

The code reduces worksheet interaction by using:

- Bulk formula assignment
- In-memory arrays for reading and writing summary data
- Dictionary-based aggregation
- Disabled screen updating and events during processing
- Manual calculation mode during processing

Actual runtime depends on workbook size, source data volume, and the Excel environment.

## Procedure Reference

| Procedure | Purpose |
|---|---|
| `OpenFile` | Opens the multi-file selection dialog |
| `ReportProcess` | Main entry point; manages workflow, timing, and cleanup |
| `DeleteAllSheet` | Deletes worksheets other than the preserved sheets |
| `ImportFile` | Matches filenames and imports selected columns |
| `Process` | Calculates metrics and load flags |
| `Summary` | Aggregates records and writes the Control summary |
| `FormatAllSheets` | Formats non-Control worksheets |

## Known Limitations

1. **Completion status is not reliable:** Cancellation and runtime errors reach the same cleanup block as successful processing, and no error details are displayed.
2. **Input validation is limited:** Missing files, incorrect layouts, unexpected data types, and reporting periods are not validated.
3. **Calculation should be verified:** The code uses manual calculation and does not explicitly call `Calculate` before converting formulas to values. Validate derived results before relying on the report.
4. **Control cleanup is limited:** Only rows 4–1000 are deleted. Older content below row 1000 can shift upward and leave stale data.
5. **The active workbook matters:** Some worksheet references in `ImportFile` are unqualified. Run the macro with the report workbook active.
6. **Thresholds are hard-coded:** Change the formulas in `Process` if reporting rules change.
