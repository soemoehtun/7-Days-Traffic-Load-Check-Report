# Daily 3G/4G Throughput Combination Report

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-Macro-blue?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Stable-success?style=for-the-badge)

Automated Excel VBA solution to consolidate, calculate, and summarize daily 3G (UMTS) & 4G (LTE) throughput data from **Huawei** and **Ericsson** network equipment.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [How to Use](#how-to-use)
- [Required Input Files](#required-input-files)
- [Output Structure](#output-structure)
- [Formula Reference](#formula-reference)
  - [3G (UMTS) Formulas](#3g-umts-formulas)
  - [4G (LTE) Formulas](#4g-lte-formulas)
  - [Summary Calculations](#summary-calculations)
  - [Date / Region Calculations](#date--region-calculations)
- [Module Architecture](#module-architecture)
- [Troubleshooting](#troubleshooting)
- [License](#license)

---

## Overview

This macro processes raw daily traffic exports from Huawei and Ericsson OSS/PM systems, extracts only the necessary columns, applies vendor-specific KPI formulas (throughput, PRB/TCP utilization, load checks), enriches data with Region & Township mappings, and produces a consolidated **Control** dashboard.

---

## Features

- **One-click import** of multiple source files
- **Vendor-aware column extraction** (only imports needed columns for speed)
- **Automatic calculation** of throughput, power, utilization, and congestion flags
- **Region / Township enrichment** via Control sheet lookup table
- **By-date summary** aggregated to the Control sheet
- **By-date & by-region** congestion breakdown for UMTS and LTE
- **Auto-formatting** (fonts, borders, number formats, column widths)
- **Speed optimized** with `ScreenUpdating=False`, `Calculation=Manual`, and bulk array writes

---

## Prerequisites

| Requirement | Detail |
|-------------|--------|
| **Excel** | 2010 or newer (`.xlsm`) |
| **Trust Center** | Macros must be enabled |
| **Input Files** | Placed in one folder, following naming conventions below |
| **Control Sheet** | Must contain Region/Township mapping in columns `C:F` starting at row 4 |

---

## How to Use

1. Place all required source files in a single folder.
2. Open the workbook `Report.xlsm`.
3. Press `Alt + F8`, select **`ReportProcess`**, and click **Run**.
4. In the file picker, select **all 4 files** and click **Open**.
5. Wait for the completion message showing total execution time.

> The macro automatically deletes previous report sheets, rebuilds them, and restores Excel settings even if an error occurs.

---

## Required Input Files

The macro identifies files by partial filename (case-insensitive):

| Technology | Vendor | Filename must contain | Destination Sheet |
|------------|--------|----------------------|-------------------|
| 3G UMTS | Huawei | `hw_data_update_3g_nbh_cell` | `HW UMTS` |
| 3G UMTS | Ericsson | `eric_data_update_3g_nbh_cell` | `Eric UMTS` |
| 4G LTE | Huawei | `hw_data_update_4g_dbh` | `HW LTE` |
| 4G LTE | Ericsson | `eric_data_update_4g_dbh` | `Eric LTE` |

If a file is missing from the selection, that sheet is created but left empty.

---

## Output Structure

After a successful run, the workbook contains:

| Sheet Name | Purpose |
|------------|---------|
| **Control** | Summary tables (by-date, by-date-region), Region mapping |
| **EP Information** | Preserved configuration / reference sheet (never deleted) |
| **HW UMTS** | Huawei 3G data + calculated KPIs |
| **Eric UMTS** | Ericsson 3G data + calculated KPIs |
| **HW LTE** | Huawei 4G data + calculated KPIs |
| **Eric LTE** | Ericsson 4G data + calculated KPIs |

---

## Formula Reference

All formulas below are written exactly as they appear in the VBA `Process` sub using R1C1 notation. Column letters are shown in parentheses for clarity.

### 3G (UMTS) Formulas

#### Huawei 3G (`HW UMTS`)

Imported source columns: `A, D, E, I, J, K`

| Col | Header | R1C1 Formula | Plain Excel (Row 2) | Logic |
|-----|--------|--------------|---------------------|-------|
| **G** | `HSDPA_UserThroughput` | `=IFERROR(RC4/(RC5*2),"")` | `=IFERROR(D2/(E2*2),"")` | HSDPA Traffic Volume ÷ (HSDPA User Number × 2) |
| **H** | `VS.MeanTCP(W)` | `=IFERROR(10^(RC6/10)/1000,"")` | `=IFERROR(10^(F2/10)/1000,"")` | Convert dBm → Watts |
| **I** | `Max Transmit Power of Cell_watt` | `=IFERROR(10^(RC3/100)/1000,"")` | `=IFERROR(10^(C2/100)/1000,"")` | Convert dBm → Watts |
| **J** | `TCP Utilization(%)` | `=IFERROR(RC8/RC9*100,"")` | `=IFERROR(H2/I2*100,"")` | VS.MeanTCP ÷ Max Transmit Power × 100 |
| **K** | `Load Check_thrp` | `=IFERROR(IF(RC7<1000,1,0),"")` | `=IFERROR(IF(G2<1000,1,0),"")` | `1` = congested by throughput only (< 1000 kbps) |
| **L** | `Load Check_both` | `=IFERROR(IF(AND(RC7<1000,RC10>80),1,0),"")` | `=IFERROR(IF(AND(G2<1000,J2>80),1,0),"")` | `1` = throughput < 1000 kbps **AND** TCP > 80 % |
| **M** | `Region` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,3,0),"")` | Extract 6-char code from cell name → lookup Region (Control col E) |
| **N** | `Township` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,4,0),"")` | Lookup Township (Control col F) |

#### Ericsson 3G (`Eric UMTS`)

Imported source columns: `A, C, F, G`

| Col | Header | R1C1 Formula | Plain Excel (Row 2) | Logic |
|-----|--------|--------------|---------------------|-------|
| **E** | `Load Check_thrp` | `=IFERROR(IF(RC3="","",IF(RC3<1000,1,0)),"")` | `=IFERROR(IF(C2="","",IF(C2<1000,1,0)),"")` | `1` = throughput < 1000 kbps; blank if no data |
| **F** | `Load Check_both` | `=IFERROR(IF(RC3="","",IF(AND(RC3<1000,RC4>80),1,0)),"")` | `=IFERROR(IF(C2="","",IF(AND(C2<1000,D2>80),1,0)),"")` | `1` = throughput < 1000 **AND** PRB/Util > 80 % |
| **G** | `Region` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,3,0),"")` | Same mapping as Huawei |
| **H** | `Township` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,4,0),"")` | Same mapping as Huawei |

---

### 4G (LTE) Formulas

#### Huawei 4G (`HW LTE`)

Imported source columns: `A, E, I:M`

| Col | Header | R1C1 Formula | Plain Excel (Row 2) | Logic |
|-----|--------|--------------|---------------------|-------|
| **H** | `Throughput(kbps)` | `=IFERROR((RC3-RC4)/RC5,"")` | `=IFERROR((C2-D2)/E2,"")` | (DL Traffic − Retrans) ÷ PRB Used |
| **I** | `PRB Utilization (%)` | `=IFERROR(RC6/RC7*100,"")` | `=IFERROR(F2/G2*100,"")` | PRB Used ÷ PRB Available × 100 |
| **J** | `Load Check` | `=IFERROR(IF(AND(RC8<2000,RC9>70),1,0),"")` | `=IFERROR(IF(AND(H2<2000,I2>70),1,0),"")` | `1` = throughput < 2000 kbps **AND** PRB > 70 % |
| **K** | `Region` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,3,0),"")` | Region lookup |
| **L** | `Township` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,4,0),"")` | Township lookup |

#### Ericsson 4G (`Eric LTE`)

Imported source columns: `A:H, J:K`

| Col | Header | R1C1 Formula | Plain Excel (Row 2) | Logic |
|-----|--------|--------------|---------------------|-------|
| **K** | `Throughput(kbps)` | `=IFERROR(((RC3-RC4)/RC5)*1000,"")` | `=IFERROR(((C2-D2)/E2)*1000,"")` | (DL Traffic − Retrans) ÷ PRB Used × 1000 |
| **L** | `PRB Utilization(%)` | `=IFERROR(((RC6+RC7+RC8+RC9)/RC10)*100,"")` | `=IFERROR(((F2+G2+H2+I2)/J2)*100,"")` | Sum of 4 PRB counters ÷ Total PRB × 100 |
| **M** | `Load Check` | `=IFERROR(IF(AND(RC11<2000,RC12>70),1,0),"")` | `=IFERROR(IF(AND(K2<2000,L2>70),1,0),"")` | `1` = throughput < 2000 kbps **AND** PRB > 70 % |
| **N** | `Region` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,3,0),"")` | Region lookup |
| **O** | `Township` | `=IFERROR(VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0),"")` | `=IFERROR(VLOOKUP(MID(B2,5,6),Control!$C$4:$F$7,4,0),"")` | Township lookup |

---

### Summary Calculations

Data is aggregated in VBA using a `Scripting.Dictionary` keyed by **Date** (Column A of source sheets) and written to the **Control** sheet in one bulk operation.

#### UMTS Summary (`Control` columns I:P)

| Col | Metric | Dictionary Index | Excel Logic |
|-----|--------|------------------|-------------|
| **I** | Date | Key | Unique dates from `HW UMTS` Column A |
| **J** | HW UMTS Total | `stats(0)` | Count of all Huawei 3G rows for that date |
| **K** | HW UMTS `Load Check_thrp` = 1 | `stats(1)` | Count where `HW UMTS!K = 1` |
| **L** | HW UMTS `Load Check_both` = 1 | `stats(2)` | Count where `HW UMTS!L = 1` |
| **M** | Eric UMTS Total | `stats(3)` | Count of all Ericsson 3G rows for that date |
| **N** | Eric UMTS `Load Check_thrp` = 1 | `stats(4)` | Count where `Eric UMTS!E = 1` |
| **O** | Eric UMTS `Load Check_both` = 1 | `stats(5)` | Count where `Eric UMTS!F = 1` |
| **P** | Congestion Rate | — | `(L + O) / (J + M)` |

#### LTE Summary (`Control` columns R:W)

| Col | Metric | Dictionary Index | Excel Logic |
|-----|--------|------------------|-------------|
| **R** | Date | Key | Unique dates (duplicate of column I for layout) |
| **S** | HW LTE Total | `stats(6)` | Count of all Huawei 4G rows for that date |
| **T** | HW LTE `Load Check` = 1 | `stats(7)` | Count where `HW LTE!J = 1` |
| **U** | Eric LTE Total | `stats(8)` | Count of all Ericsson 4G rows for that date |
| **V** | Eric LTE `Load Check` = 1 | `stats(9)` | Count where `Eric LTE!M = 1` |
| **W** | Congestion Rate | — | `(T + V) / (S + U)` |

---

### Date / Region Calculations

A second block writes **Date × 5 Regions** (`North West`, `North East`, `South West`, `South East`, `Yangon`) to the Control sheet using `COUNTIFS` across all four technology sheets, then converts results to static values.

#### UMTS Block (`Control` columns Y:AC)

| Col | Header | Formula Written by VBA |
|-----|--------|------------------------|
| **Y** | Date | Static date value (`d-mmm` format) |
| **Z** | Region | Static text from region array |
| **AA** | Congested (`thrp` only) | `=COUNTIFS('Eric UMTS'!G:G,Z4,'Eric UMTS'!A:A,Y4,'Eric UMTS'!E:E,1)+COUNTIFS('HW UMTS'!M:M,Z4,'HW UMTS'!A:A,Y4,'HW UMTS'!K:K,1)` |
| **AB** | Congested (`both`) | `=COUNTIFS('Eric UMTS'!G:G,Z4,'Eric UMTS'!A:A,Y4,'Eric UMTS'!F:F,1)+COUNTIFS('HW UMTS'!M:M,Z4,'HW UMTS'!A:A,Y4,'HW UMTS'!L:L,1)` |
| **AC** | Total Cells | `=COUNTIFS('Eric UMTS'!G:G,Z4,'Eric UMTS'!A:A,Y4)+COUNTIFS('HW UMTS'!M:M,Z4,'HW UMTS'!A:A,Y4)` |

#### LTE Block (`Control` columns AE:AH)

| Col | Header | Formula Written by VBA |
|-----|--------|------------------------|
| **AE** | Date | Static date value (`d-mmm` format) |
| **AF** | Region | Static text from region array |
| **AG** | Congested (`Load Check`) | `=COUNTIFS('Eric LTE'!N:N,AF4,'Eric LTE'!A:A,AE4,'Eric LTE'!M:M,1)+COUNTIFS('HW LTE'!K:K,AF4,'HW LTE'!A:A,AE4,'HW LTE'!J:J,1)` |
| **AH** | Total Cells | `=COUNTIFS('Eric LTE'!N:N,AF4,'Eric LTE'!A:A,AE4)+COUNTIFS('HW LTE'!K:K,AF4,'HW LTE'!A:A,AE4)` |

> After calculation, formulas in `AA:AC` and `AG:AH` are replaced by their computed values to eliminate external links and speed up the workbook.

---

## Module Architecture

| Sub / Function | Responsibility |
|----------------|----------------|
| `ReportProcess` | Entry point; sets speed flags, calls worker subs, restores state in `CleanUp` |
| `DeleteAllSheet` | Removes old report sheets; clears Control output ranges |
| `ImportFile` | Opens file picker, routes files by name, delegates column extraction |
| `ImportColumns` | Copies only required columns from source to destination sheet |
| `Process` | Applies vendor-specific formulas and converts them to values |
| `Summary` | Aggregates 4 sheets into a Dictionary and writes by-date summary |
| `FillDateRegion` | Builds Date × Region matrix, applies COUNTIFS, converts to values |
| `ClearFromRow4` | Helper to clear old data below row 4 for given column ranges |
| `FormatAllSheets` | Applies fonts, borders, number formats, and autofit to all report sheets |

---

## Troubleshooting

| Symptom | Cause | Solution |
|---------|-------|----------|
| `Type mismatch` / `Subscript out of range` | One of the 4 files was not selected | Re-run and select **all 4 files** together |
| Region/Township columns show `#N/A` | Cell name code missing from Control mapping | Add mapping to `Control!C4:F7` (or extend range) |
| Dates are missing in `HW UMTS` col A | Source file has dates in a different format | Ensure dates are real Excel dates, not text |
| Macro runs slowly | `Calculation` is not set to Manual | Do not interrupt; macro sets it internally |
| `Run-time error 1004` on delete | A report sheet is active when deletion starts | Close any other workbooks and re-run |

---

## License

Released under the MIT License.  
Feel free to fork, modify, and distribute.

---

**Maintained by:** [soemoehtun](https://github.com/soemoehtun)
