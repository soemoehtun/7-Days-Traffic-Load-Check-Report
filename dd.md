# Daily 3G/4G Throughput Combination Report

Excel VBA macro that imports daily busy-hour cell traffic exports from **Huawei (HW)** and **Ericsson (Eric)** for **3G (UMTS)** and **4G (LTE)**, computes throughput / utilization KPIs, flags *loaded* cells, and writes a summary per **Date** and per **Date × Region** into the `Control` sheet.

---

## 1. Workbook Structure

| Sheet | Purpose | Created by macro? |
|---|---|---|
| `Control` | Lookup table (Township → Region) + all summary outputs | No (must exist) |
| `EP Information` | Reference info, untouched | No (must exist) |
| `HW UMTS` | Huawei 3G raw + calculated columns | Yes |
| `Eric UMTS` | Ericsson 3G raw + calculated columns | Yes |
| `HW LTE` | Huawei 4G raw + calculated columns | Yes |
| `Eric LTE` | Ericsson 4G raw + calculated columns | Yes |

> All other sheets are **deleted** at the start of every run.

### Control sheet layout

| Range | Content |
|---|---|
| `D4:G7` | Lookup table: `D` = 6-char township code, `F` = Region, `G` = Township |
| `I:P` | UMTS summary per Date |
| `R:W` | LTE summary per Date |
| `Y:AC` | UMTS summary per Date × Region |
| `AE:AH` | LTE summary per Date × Region |

---

## 2. Input Files

Select **all 4 files at once** in the file dialog. Files are matched by file name (case-insensitive):

| Pattern in file name | Loaded into |
|---|---|
| `hw_data_update_3g_nbh_cell` | `HW UMTS` |
| `eric_data_update_3g_nbh_cell` | `Eric UMTS` |
| `hw_data_update_4g_dbh` | `HW LTE` |
| `eric_data_update_4g_dbh` | `Eric LTE` |

Only the first sheet of each file is read. Missing files are skipped.

### Columns imported

| Sheet | Source columns | Destination |
|---|---|---|
| HW UMTS | `A, D, E, I, J, K` | `A:F` |
| Eric UMTS | `A, C, F, G` | `A:D` |
| HW LTE | `A, E, I:M` | `A:G` |
| Eric LTE | `A:H, J:K` | `A:J` |

Column `A` must be **Date**, column `B` must be **Cell Name** (township code is `MID(CellName, 5, 6)`).

---

## 3. KPI Formulas

### 3.1 HW UMTS (`HW UMTS` sheet)

Imported: `A` Date · `B` Cell · `C` MaxTxPower (×0.1 dBm) · `D` HSDPA Throughput · `E` HSDPA Users · `F` VS.MeanTCP (dBm)

| Col | Header | Formula | Meaning |
|---|---|---|---|
| G | HSDPA_UserThroughput | `= D / (E × 2)` | kbps per user |
| H | VS.MeanTCP(W) | `= 10^(F/10) / 1000` | dBm → Watt |
| I | Max Transmit Power of Cell_watt | `= 10^(C/100) / 1000` | 0.1 dBm → Watt |
| J | TCP Utilization(%) | `= H / I × 100` | Power utilization |
| K | Load Check_thrp | `= IF(G < 1000, 1, 0)` | Low throughput |
| L | Load Check_both | `= IF(AND(G < 1000, J > 80), 1, 0)` | Low thrp **and** high power |
| M | Region | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 3, 0)` | |
| N | Township | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 4, 0)` | |

### 3.2 Eric UMTS (`Eric UMTS` sheet)

Imported: `A` Date · `B` Cell · `C` User Throughput (kbps) · `D` TCP Utilization (%)

| Col | Header | Formula |
|---|---|---|
| E | Load Check_thrp | `= IF(C < 1000, 1, 0)` |
| F | Load Check_both | `= IF(AND(C < 1000, D > 80), 1, 0)` |
| G | Region | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 3, 0)` |
| H | Township | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 4, 0)` |

### 3.3 HW LTE (`HW LTE` sheet)

Imported: `A` Date · `B` Cell · `C` DL Traffic Volume · `D` DL Volume in Last TTI · `E` DL Transmit Time · `F` PRB Used · `G` PRB Available

| Col | Header | Formula |
|---|---|---|
| H | Throughput(kbps) | `= (C − D) / E` |
| I | PRB Utilization (%) | `= F / G × 100` |
| J | Load Check | `= IF(AND(H < 2000, I > 70), 1, 0)` |
| K | Region | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 3, 0)` |
| L | Township | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 4, 0)` |

### 3.4 Eric LTE (`Eric LTE` sheet)

Imported: `A` Date · `B` Cell · `C` DL Volume · `D` DL Volume Last TTI · `E` DL Time · `F,G,H,I` PRB Used (per type) · `J` PRB Available

| Col | Header | Formula |
|---|---|---|
| K | Throughput(kbps) | `= ((C − D) / E) × 1000` |
| L | PRB Utilization(%) | `= ((F + G + H + I) / J) × 100` |
| M | Load Check | `= IF(AND(K < 2000, L > 70), 1, 0)` |
| N | Region | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 3, 0)` |
| O | Township | `= VLOOKUP(MID(B,5,6), Control!D4:G7, 4, 0)` |

### Load thresholds

| Tech | Throughput | Utilization |
|---|---|---|
| 3G (UMTS) | `< 1000 kbps` | TCP `> 80 %` |
| 4G (LTE) | `< 2000 kbps` | PRB `> 70 %` |

All formulas are wrapped in `IFERROR(..., "")` and converted to values after calculation.

---

## 4. Summary per Date (Control `I:W`)

Keys = unique dates found in `HW UMTS!A`.

**UMTS**

| Col | Formula |
|---|---|
| I | Date |
| J | `COUNT(HW UMTS cells)` |
| K | `COUNT(HW UMTS, K = 1)` |
| L | `COUNT(HW UMTS, L = 1)` |
| M | `COUNT(Eric UMTS cells)` |
| N | `COUNT(Eric UMTS, E = 1)` |
| O | `COUNT(Eric UMTS, F = 1)` |
| P | **Loaded Ratio** `= (L + O) / (J + M)` |

**LTE**

| Col | Formula |
|---|---|
| R | Date |
| S | `COUNT(HW LTE cells)` |
| T | `COUNT(HW LTE, J = 1)` |
| U | `COUNT(Eric LTE cells)` |
| V | `COUNT(Eric LTE, M = 1)` |
| W | **Loaded Ratio** `= (T + V) / (S + U)` |

---

## 5. Summary per Date × Region (Control `Y:AH`)

Regions: `North West`, `North East`, `South West`, `South East`, `Yangon` — 5 rows per date, dates sorted ascending.

**UMTS**

| Col | Formula |
|---|---|
| Y | Date |
| Z | Region |
| AA | `COUNTIFS(Eric UMTS, Region, Date, E=1) + COUNTIFS(HW UMTS, Region, Date, K=1)` |
| AB | `COUNTIFS(Eric UMTS, Region, Date, F=1) + COUNTIFS(HW UMTS, Region, Date, L=1)` |
| AC | `COUNTIFS(Eric UMTS, Region, Date) + COUNTIFS(HW UMTS, Region, Date)` |

**LTE**

| Col | Formula |
|---|---|
| AE | Date |
| AF | Region |
| AG | `COUNTIFS(Eric LTE, Region, Date, M=1) + COUNTIFS(HW LTE, Region, Date, J=1)` |
| AH | `COUNTIFS(Eric LTE, Region, Date) + COUNTIFS(HW LTE, Region, Date)` |

---

## 6. Process Flow
