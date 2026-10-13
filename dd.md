# 7-Days Traffic Load Check Report — Excel VBA

This tool answers one core Radio-Optimization question:

> ## *"How many cells are giving LOW THROUGHPUT but operating at HIGH LOAD?"*

It imports 7-day busy-hour traffic data for **Huawei/Ericsson UMTS (3G) and LTE (4G)**,
derives throughput and resource-utilization KPIs from raw vendor counters, flags
**possible load-limited cells**, and aggregates the results **date-wise** into a
summary sheet.

| Technology | High-Load Indicator Used |
|---|---|
| Huawei UMTS | **TCP Power Utilization > 80%** |
| Ericsson UMTS | **Power Utilization > 80%** |
| Huawei LTE | **PRB Utilization > 70%** |
| Ericsson LTE | **PRB Utilization > 70%** |

---

## Table of Contents

1. [Requirements & Setup](#1-requirements--setup)
2. [How to Run](#2-how-to-run)
3. [Source File Naming](#3-source-file-naming)
4. [Processing Pipeline](#4-processing-pipeline)
5. [Part A — Huawei UMTS Formulas](#5-part-a--huawei-umts-3g)
6. [Part B — Ericsson UMTS Formulas](#6-part-b--ericsson-umts-3g)
7. [Part C — Huawei LTE Formulas](#7-part-c--huawei-lte-4g)
8. [Part D — Ericsson LTE Formulas](#8-part-d--ericsson-lte-4g)
9. [Summary Sheet (Control)](#9-summary-sheet-control)
10. [Thresholds Summary](#10-thresholds-summary)
11. [Notes & Known Issues](#11-notes--known-issues)

---

## 1. Requirements & Setup

- Microsoft Excel (Windows), macros enabled, workbook saved as `.xlsm`
- Two permanent sheets must exist (never deleted):
  - **`Control`** — receives the summary output
  - **`EP Information`** — reference sheet
- Source files must have: data on the **first sheet**, headers in **row 1**,
  and the **Date in column A**

**Setup:** `Alt + F11` → Insert → Module → paste the code → save as `.xlsm`.

## 2. How to Run

1. Run the **`ReportProcess`** macro (`Alt + F8`, or assign it to a button).
2. Select one or more source files in the dialog (Ctrl/Shift = multi-select).
3. A message box shows the elapsed processing time when done.

All data sheets are deleted and rebuilt on every run.

## 3. Source File Naming

Files are matched by **case-insensitive substring** in the file name:

| Destination Sheet | File name must contain | Vendor | Tech |
|---|---|---|---|
| `HW UMTS` | `hw_data_update_3g_nbh_cell` | Huawei | UMTS |
| `Eric UMTS` | `eric_data_update_3g_nbh_cell` | Ericsson | UMTS |
| `HW LTE` | `hw_data_update_4g_dbh` | Huawei | LTE |
| `Eric LTE` | `eric_data_update_4g_dbh` | Ericsson | LTE |

## 4. Processing Pipeline

```
ReportProcess
 ├── DeleteAllSheet   – remove all sheets except "Control" / "EP Information"
 ├── ImportFile       – open source files read-only, copy required columns only
 ├── Process          – add calculated columns (formulas below)
 ├── Summary          – aggregate date-wise into "Control"
 └── FormatAllSheets  – fonts, borders, header styling, number formats
```

---

## 5. Part A — Huawei UMTS (3G)

### 5.1 Imported Columns

Source columns `A, D, E, I, J, K` → sheet columns `A–F`:

| Sheet Col | Source Col | KPI Counter |
|---|---|---|
| A | A | **Date** (grouping key) |
| B | D | — (carried, not used in formulas) |
| C | E | **Max Transmit Power of Cell** |
| D | I | **VS.HSDPA.DataOutput.Traffic(bit)** |
| E | J | **VS.HSDPA.DataTtiNum.User** |
| F | K | **VS.MeanTCP(dBm)** |

### 5.2 Calculated Columns

#### G — `HSDPA_UserThroughput`

```
Excel:   =D2/(E2*2)
KPI:     VS.HSDPA.DataOutput.Traffic(bit) / (VS.HSDPA.DataTtiNum.User × 2)
```

$$
\text{HSDPA User Throughput} = \frac{\text{Total HSDPA bits transmitted}}{\text{User TTI count} \times 2\,\text{ms}}
$$

**Why × 2?** In HSDPA, **TTI (Transmission Time Interval) = 2 ms**.
Total transmission time = (number of user TTIs) × 2 ms, so throughput = bits ÷ time.
We are deriving cell-level HSDPA user throughput from raw counters.

#### H — `VS.MeanTCP(W)`

```
Excel:   =10^(F2/10)/1000
KPI:     10^(VS.MeanTCP(dBm)/10) / 1000
```

$$
P_{(Watt)} = \frac{10^{(dBm/10)}}{1000}
$$

**Why convert?** The counter is in **dBm**, but the cell max power is converted to
**Watts** — both must be in the same unit before dividing.
(dBm → mW: `10^(dBm/10)`, then mW → W: ÷ 1000.)

**Examples:** 43 dBm ≈ 20 W · 40 dBm = 10 W

#### I — `Max Transmit Power of Cell_watt`

```
Excel:   =10^(C2/100)/1000
KPI:     10^(Max Transmit Power of Cell/100) / 1000
```

$$
P_{max(Watt)} = \frac{10^{(value/100)}}{1000}
$$

**Why `/100` instead of `/10`?** The Huawei export stores this column in a **scaled
format** (value ≈ dBm × 10), so dividing by 100 instead of 10 compensates.
Example: `430` → 10^4.3 / 1000 ≈ **20 W**.

#### J — `TCP Utilization(%)`

```
Excel:   =H2/I2*100
KPI:     VS.MeanTCP(W) / Max Transmit Power of Cell_watt × 100
```

$$
\text{TCP Utilization (\%)} = \frac{\text{Average TCP Power used}}{\text{Cell Maximum Transmit Power}} \times 100
$$

**Interpretation:** How much of the cell's downlink power is used on average.
Close to **100%** → power-heavy / power-limited cell.

#### K — `Load Check_thrp`

```
Excel:   =IF(G2<1000,1,0)
```

> **IF HSDPA User Throughput < 1000 → 1, ELSE 0** (low-throughput cell)

#### L — `Load Check_both`

```
Excel:   =IF(AND(G2<1000,J2>0.8),1,0)
```

> **IF (Throughput < 1000) AND (TCP Utilization > 80%) → 1, ELSE 0**

⚠️ **See [Known Issues](#11-notes--known-issues)** — column J is in %, so the
threshold `0.8` should likely be `80`.

After calculation, formulas are converted to values and `#DIV/0!` cells are cleared
(Excel equivalent of the Python `nan_count` check — caused by zero TTIs, zero max
power, or missing counters).

---

## 6. Part B — Ericsson UMTS (3G)

### 6.1 Imported Columns

Source columns `A, C, F, G` → sheet columns `A–D`:

| Sheet Col | Source Col | KPI |
|---|---|---|
| A | A | **Date** (grouping key) |
| B | C | — (carried, not used) |
| C | F | **HS User Throughput(Kbps)** |
| D | G | **Power Utilization** |

### 6.2 Calculated Columns

Ericsson **already provides** user throughput and power utilization directly —
**no counter derivation needed** (unlike Huawei).

#### E — `Load Check_thrp`

```
Excel:   =IF(C2="","NA",IF(C2<1000,1,0))
```

> **IF HS User Throughput < 1000 Kbps → 1, ELSE 0** (blank throughput → NA, cleared later)

#### F — `Load Check_both`

```
Excel:   =IF(C2="","NA",IF(AND(C2<1000,D2>80),1,0))
```

> **IF (HS User Throughput < 1000 Kbps) AND (Power Utilization > 80%) → 1, ELSE 0**

**Unit difference vs Huawei:** Huawei uses a **ratio** (0.8 = 80%);
Ericsson reports **percentage** directly (> 80). Same logic, different units.

---

## 7. Part C — Huawei LTE (4G)

### 7.1 Imported Columns

Source columns `A, E, I:M` → sheet columns `A–G`:

| Sheet Col | Source Col | KPI Counter |
|---|---|---|
| A | A | **Date** (grouping key) |
| B | E | — (carried, not used) |
| C | I | **L.Thrp.bits.DL(bit)** |
| D | J | **L.Thrp.bits.DL.LastTTI(bit)** |
| E | K | **L.Thrp.Time.DL.RmvLastTTI(ms)** |
| F | L | **L.ChMeas.PRB.DL.Used.Avg** |
| G | M | **L.ChMeas.PRB.DL.Avail** |

### 7.2 Calculated Columns

#### H — `Throughput(kbps)`

```
Excel:   =(C2-D2)/E2
KPI:     (L.Thrp.bits.DL(bit) − L.Thrp.bits.DL.LastTTI(bit)) / L.Thrp.Time.DL.RmvLastTTI(ms)
```

$$
\text{Throughput} = \frac{\text{DL bits (excluding last TTI)}}{\text{DL throughput time (excluding last TTI, ms)}}
$$

**Why subtract LastTTI?** Huawei's total DL bits **include** the last TTI's bits,
which fall outside the measured time window — removing them keeps numerator and
denominator consistent.

**Units:** `1 bit/ms = 1000 bit/s = 1 kbps`, so bits ÷ ms directly gives **kbps**.

#### I — `PRB Utilization (%)`

```
Excel:   =F2/G2*100
KPI:     L.ChMeas.PRB.DL.Used.Avg / L.ChMeas.PRB.DL.Avail × 100
```

$$
\text{PRB Utilization (\%)} = \frac{\text{Average DL PRBs Used}}{\text{Average DL PRBs Available}} \times 100
$$

**PRB = Physical Resource Block** — all LTE DL user data is scheduled on PRBs.
Examples: 50 used/100 avail = 50% · 85/100 = 85%.

#### J — `Load Check`

```
Excel:   =IF(AND(H2<2000,I2>70),1,0)
```

> **IF (DL Throughput < 2000 kbps) AND (PRB Utilization > 70%) → 1, ELSE 0**

**Meaning:** the cell is heavily loaded in PRBs but users still get low throughput →
potential load-limited cell.

---

## 8. Part D — Ericsson LTE (4G)

### 8.1 Imported Columns

Source columns `A:H, J:K` → sheet columns `A–J` (source column I is skipped):

| Sheet Col | Source Col | KPI Counter |
|---|---|---|
| A | A | **Date** (grouping key) |
| B | B | — (carried) |
| C | C | **pmPdcpVolDlDrb** |
| D | D | **pmPdcpVolDlDrbLastTTI** |
| E | E | **pmUeThpTimeDl** |
| F | F | **pmPrbUsedDlDtch** |
| G | G | **pmPrbUsedDlSrbFirstTrans** |
| H | H | **pmPrbUsedDlAilgTot** |
| I | J | **pmPrbUsedDlPcch** |
| J | K | **pmPrbAvailDl** |

### 8.2 Calculated Columns

#### K — `Throughput(kbps)`

```
Excel:   =((C2-D2)/E2)*1000
KPI:     (pmPdcpVolDlDrb − pmPdcpVolDlDrbLastTTI) / pmUeThpTimeDl × 1000
```

$$
\text{Throughput} = \frac{\text{PDCP DL volume} - \text{Last TTI volume}}{\text{UE throughput time}} \times 1000
$$

- Same "remove LastTTI" logic as Huawei.
- **Why × 1000?** Ericsson's `pmUeThpTimeDl` uses a different time base than
  Huawei's millisecond counter; the × 1000 factor normalizes the result into
  **kbps** so the same 2000 kbps threshold applies. (Verify against your report's
  counter unit definitions.)

#### L — `PRB Utilization(%)`

```
Excel:   =((F2+G2+H2+I2)/J2)*100
KPI:     (pmPrbUsedDlDtch + pmPrbUsedDlSrbFirstTrans + pmPrbUsedDlAilgTot + pmPrbUsedDlPcch)
         / pmPrbAvailDl × 100
```

$$
\text{PRB Utilization (\%)} = \frac{pmPrbUsedDlDtch + pmPrbUsedDlSrbFirstTrans + pmPrbUsedDlAilgTot + pmPrbUsedDlPcch}{pmPrbAvailDl} \times 100
$$

**Why sum 4 counters?** Ericsson splits DL PRB usage by channel type —
DTCH (user data), SRB first transmissions (signaling), AILG, PCCH (paging) —
so they must be **summed** to get total DL PRB usage. Huawei provides this as a
single averaged counter.

#### M — `Load Check`

```
Excel:   =IF(AND(K2<2000,L2>70),1,0)
```

> **IF (Throughput < 2000 kbps) AND (PRB Utilization > 70%) → 1, ELSE 0**

Identical threshold to Huawei LTE.

---

## 9. Summary Sheet (Control)

Rows 4:1000 are cleared, then results are written from **row 4**, grouped by the
**Date** (column A of the imported data). This is the Excel equivalent of the
Python `groupby('Date').agg(['sum','count'])`:

- **`count`** → total cell records reported that date
- **`sum`** → how many of those cells were flagged (= 1)

### UMTS Block (columns I–P)

| Col | Field | Formula |
|---|---|---|
| I | Date | grouping key |
| J | HW UMTS total records | count of rows |
| K | HW UMTS `Load Check_thrp` hits | count of G = 1 |
| L | HW UMTS `Load Check_both` hits | count of L = 1 |
| M | Eric UMTS total records | count of rows |
| N | Eric UMTS `Load Check_thrp` hits | count of E = 1 |
| O | Eric UMTS `Load Check_both` hits | count of F = 1 |
| P | **UMTS load ratio** | `=(L+O)/(J+M)` |

### LTE Block (columns R–W)

| Col | Field | Formula |
|---|---|---|
| R | Date | grouping key |
| S | HW LTE total records | count of rows |
| T | HW LTE `Load Check` hits | count of J = 1 |
| U | Eric LTE total records | count of rows |
| V | Eric LTE `Load Check` hits | count of M = 1 |
| W | **LTE load ratio** | `=(T+V)/(S+U)` |

**Reading the ratios:** `P` and `W` are the fraction (0–1) of cell-records that day
which were **low throughput + high load**, combined across both vendors.

---

## 10. Thresholds Summary

| Check | Condition | Flag |
|---|---|---|
| UMTS low throughput (both vendors) | User Throughput **< 1000 kbps** | `Load Check_thrp` |
| Huawei UMTS power | TCP Utilization **> 80%** | `Load Check_both` |
| Ericsson UMTS power | Power Utilization **> 80%** | `Load Check_both` |
| LTE low throughput (both vendors) | DL Throughput **< 2000 kbps** | `Load Check` |
| LTE PRB (both vendors) | PRB Utilization **> 70%** | `Load Check` |

All thresholds live in the `Process` sub and can be edited there.

## 11. Notes & Known Issues

1. **⚠️ HW UMTS `Load Check_both` threshold:** column J is multiplied by × 100
   (percentage), but the condition tests `J > 0.8` (i.e., 0.8%). The original
   Python script compares the **ratio** against 0.8 (= 80%). The Excel version
   likely should use `J2>80`. Please verify.
2. **Summary keys come only from HW UMTS.** Dates/cells present only in Eric or
   LTE files are not added to the summary — those rows are matched against the
   HW UMTS key set only. If the HW UMTS file is missing, the summary is empty.
3. **Success message on error:** `CleanUp` always shows the MsgBox, even if a
   runtime error occurred mid-run (application settings are still restored).
4. Source workbooks are opened **read-only** and closed without saving.
5. Data sheets are **deleted and recreated** every run — do not store manual work
   in them.
6. The last row is detected via column A — avoid blank rows inside the data.
7. `Max Transmit Power of Cell` uses the `/100` conversion (scaled vendor format) —
   confirm with a known cell (e.g., 43 dBm ≈ 20 W).
