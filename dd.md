# Daily 3G/4G Throughput Combination Report

Automated Excel VBA macro for combining and analyzing daily 3G (UMTS) and 4G (LTE) network throughput data from Huawei and Ericsson vendors.

---

## Overview

This tool imports raw traffic data from 4 source files, calculates KPIs per cell, flags congested cells, and writes a consolidated summary to the Control sheet — grouped by date and region.

---

## Input Files

| File Name Pattern | Technology | Vendor |
|---|---|---|
| `hw_data_update_3g_nbh_cell_*.xlsx` | 3G UMTS | Huawei |
| `eric_data_update_3g_nbh_cell_*.xlsx` | 3G UMTS | Ericsson |
| `hw_data_update_4g_dbh_*.xlsx` | 4G LTE | Huawei |
| `eric_data_update_4g_dbh_*.xlsx` | 4G LTE | Ericsson |

---

## How to Run

1. Open `Report.xlsm`
2. Enable macros
3. Click Run Report button or press `Alt+F8` then run `ReportProcess`
4. Select all 4 input files when prompted
5. Wait for completion message

---

## Output Sheets

| Sheet | Description |
|---|---|
| Control | Summary tables, date/region breakdown |
| HW UMTS | Huawei 3G raw data + calculated columns |
| Eric UMTS | Ericsson 3G raw data + calculated columns |
| HW LTE | Huawei 4G raw data + calculated columns |
| Eric LTE | Ericsson 4G raw data + calculated columns |

---

## Macro Execution Flow
