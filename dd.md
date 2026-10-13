# Daily 3G/4G Throughput Combination Report

## Overview

An Excel VBA macro that imports daily busy-hour cell traffic exports
from **Huawei (HW)** and **Ericsson (Eric)** for **3G (UMTS)** and **4G
(LTE)**. It calculates throughput and utilization KPIs, flags loaded
cells, and writes summary reports by **Date** and **Date × Region** to
the `Control` worksheet.

------------------------------------------------------------------------

## 1. Workbook Structure

  -----------------------------------------------------------------------
  Sheet                   Purpose                 Created by macro?
  ----------------------- ----------------------- -----------------------
  `Control`               Township-to-region      No --- must exist
                          lookup table and all    
                          summary outputs         

  `EP Information`        Reference information;  No --- must exist
                          left untouched          

  `HW UMTS`               Huawei 3G raw data and  Yes
                          calculated columns      

  `Eric UMTS`             Ericsson 3G raw data    Yes
                          and calculated columns  

  `HW LTE`                Huawei 4G raw data and  Yes
                          calculated columns      

  `Eric LTE`              Ericsson 4G raw data    Yes
                          and calculated columns  
  -----------------------------------------------------------------------

> **Important:** All other worksheets are deleted at the start of each
> run. Save a backup before running the macro if the workbook contains
> other sheets you need.

### Control worksheet layout

  -----------------------------------------------------------------------
  Range                               Content
  ----------------------------------- -----------------------------------
  `D4:G7`                             Township lookup table. Column D =
                                      6-character township code, F =
                                      Region, G = Township

  `I:P`                               UMTS summary by Date

  `R:W`                               LTE summary by Date

  `Y:AC`                              UMTS summary by Date × Region

  `AE:AH`                             LTE summary by Date × Region
  -----------------------------------------------------------------------

The township code is extracted from the cell name using
`MID(CellName,5,6)`.

------------------------------------------------------------------------

## 2. Input Files

Select all four input files in the file-selection dialog. The macro
identifies each file by a case-insensitive substring in its filename.

  Filename contains                Loaded into worksheet
  -------------------------------- -----------------------
  `hw_data_update_3g_nbh_cell`     `HW UMTS`
  `eric_data_update_3g_nbh_cell`   `Eric UMTS`
  `hw_data_update_4g_dbh`          `HW LTE`
  `eric_data_update_4g_dbh`        `Eric LTE`

-   Only the first worksheet of each selected file is read.
-   If a matching file is not selected, that input is skipped.
-   The source columns listed below are copied into the corresponding
    destination columns.

### Imported source columns

  Destination worksheet   Source columns       Destination columns
  ----------------------- -------------------- ---------------------
  `HW UMTS`               `A, D, E, I, J, K`   `A:F`
  `Eric UMTS`             `A, C, F, G`         `A:D`
  `HW LTE`                `A, E, I:M`          `A:G`
  `Eric LTE`              `A:H, J:K`           `A:J`

**Expected data fields:** Column A must contain the Date, and column B
must contain the cell identifier/name. For regional mapping, the macro
extracts the six-character township code from characters 5--10 of column
B.

------------------------------------------------------------------------

## 3. Data Layout After Import and KPI Calculation

The following tables describe the expected final column layout on each
data worksheet after import and calculation.

### 3.1 HW UMTS

  ----------------------------------------------------------------------------------
  Column                  Header                             Description
  ----------------------- ---------------------------------- -----------------------
  A                       Date                               Reporting date

  B                       Cell Name                          Cell identifier/name

  C                       Max Transmit Power of Cell         Imported transmit-power
                                                             value (×0.1 dBm,
                                                             according to the source
                                                             definition)

  D                       VS.HSDPA.DataOutput.Traffic(bit)   HSDPA traffic/output
                                                             counter

  E                       VS.HSDPA.DataTtiNum.User           HSDPA user TTI counter

  F                       VS.MeanTCP(dBm)                    Mean TCP power in dBm

  G                       HSDPA_UserThroughput               Calculated user
                                                             throughput (kbps)

  H                       VS.MeanTCP(W)                      Mean TCP converted to
                                                             watts

  I                       Max Transmit Power of Cell_watt    Maximum transmit power
                                                             converted to watts

  J                       TCP Utilization(%)                 TCP power utilization

  K                       Load Check_thrp                    Flag for low throughput

  L                       Load Check_both                    Flag for low throughput
                                                             and high TCP
                                                             utilization

  M                       Region                             Region looked up from
                                                             `Control`

  N                       Township                           Township looked up from
                                                             `Control`
  ----------------------------------------------------------------------------------

### 3.2 Eric UMTS

  -----------------------------------------------------------------------
  Column                  Header                  Description
  ----------------------- ----------------------- -----------------------
  A                       Date                    Reporting date

  B                       Ucell Id                Cell identifier

  C                       HS User                 Imported user
                          Throughput(Kbps)        throughput

  D                       Power Utilization       Imported power
                                                  utilization

  E                       Load Check_thrp         Flag for low throughput

  F                       Load Check_both         Flag for low throughput
                                                  and high power
                                                  utilization

  G                       Region                  Region looked up from
                                                  `Control`

  H                       Township                Township looked up from
                                                  `Control`
  -----------------------------------------------------------------------

### 3.3 HW LTE

  -------------------------------------------------------------------------------
  Column                  Header                          Description
  ----------------------- ------------------------------- -----------------------
  A                       Date                            Reporting date

  B                       Cell Name                       Cell identifier/name

  C                       L.Thrp.bits.DL(bit)             Downlink traffic volume

  D                       L.Thrp.bits.DL.LastTTI(bit)     Downlink volume in the
                                                          last TTI

  E                       L.Thrp.Time.DL.RmvLastTTI(ms)   Downlink transmission
                                                          time excluding last TTI

  F                       L.ChMeas.PRB.DL.Used.Avg        Average downlink PRBs
                                                          used

  G                       L.ChMeas.PRB.DL.Avail           Available downlink PRBs

  H                       Throughput(kbps)                Calculated downlink
                                                          throughput

  I                       PRB Utilization (%)             Calculated PRB
                                                          utilization

  J                       Load Check                      Flag for low throughput
                                                          and high PRB
                                                          utilization

  K                       Region                          Region looked up from
                                                          `Control`

  L                       Township                        Township looked up from
                                                          `Control`
  -------------------------------------------------------------------------------

### 3.4 Eric LTE

  --------------------------------------------------------------------------
  Column                  Header                     Description
  ----------------------- -------------------------- -----------------------
  A                       Date                       Reporting date

  B                       EUtranCell Id              Cell identifier

  C                       pmPdcpVolDlDrb             Downlink PDCP volume

  D                       pmPdcpVolDlDrbLastTTI      Downlink PDCP volume in
                                                     the last TTI

  E                       pmUeThpTimeDl              Downlink throughput
                                                     measurement time

  F                       pmPrbUsedDlDtch            PRB used for DL DTCH

  G                       pmPrbUsedDlSrbFirstTrans   PRB used for
                                                     first-transmission DL
                                                     SRB

  H                       pmPrbUsedDlAilgTot         PRB-used counter for
                                                     the named LTE
                                                     measurement

  I                       pmPrbUsedDlPcch            PRB used for DL PCCH

  J                       pmPrbAvailDl               Available downlink PRBs

  K                       Throughput(kbps)           Calculated downlink
                                                     throughput

  L                       PRB Utilization(%)         Calculated PRB
                                                     utilization

  M                       Load Check                 Flag for low throughput
                                                     and high PRB
                                                     utilization

  N                       Region                     Region looked up from
                                                     `Control`

  O                       Township                   Township looked up from
                                                     `Control`
  --------------------------------------------------------------------------

------------------------------------------------------------------------

## 4. KPI Formulas and Load Checks

The formulas below show the intended calculation logic. In the workbook,
formulas are wrapped in `IFERROR(...,"")` and converted to values after
calculation.

### 4.1 Huawei UMTS --- `HW UMTS`

  -----------------------------------------------------------------------------------------------------
  Column            Header                 Formula / logic                            Meaning
  ----------------- ---------------------- ------------------------------------------ -----------------
  G                 HSDPA_UserThroughput   `=D/(E*2)`                                 User throughput
                                                                                      (kbps, subject to
                                                                                      source units)

  H                 VS.MeanTCP(W)          `=10^(F/10)/1000`                          Convert the
                                                                                      imported mean TCP
                                                                                      value to watts,
                                                                                      as specified by
                                                                                      the macro

  I                 Max Transmit Power of  `=10^(C/100)/1000`                         Convert the
                    Cell_watt                                                         imported power
                                                                                      value to watts,
                                                                                      as specified by
                                                                                      the macro

  J                 TCP Utilization(%)     `=H/I*100`                                 TCP power
                                                                                      utilization

  K                 Load Check_thrp        `=IF(G<1000,1,0)`                          1 when throughput
                                                                                      is below 1,000
                                                                                      kbps

  L                 Load Check_both        `=IF(AND(G<1000,J>80),1,0)`                1 when throughput
                                                                                      is below 1,000
                                                                                      kbps and TCP
                                                                                      utilization is
                                                                                      above 80%

  M                 Region                 `=VLOOKUP(MID(B,5,6),Control!D4:G7,3,0)`   Region lookup

  N                 Township               `=VLOOKUP(MID(B,5,6),Control!D4:G7,4,0)`   Township lookup
  -----------------------------------------------------------------------------------------------------

### 4.2 Ericsson UMTS --- `Eric UMTS`

  Column   Header            Formula / logic
  -------- ----------------- ------------------------------------------
  E        Load Check_thrp   `=IF(C<1000,1,0)`
  F        Load Check_both   `=IF(AND(C<1000,D>80),1,0)`
  G        Region            `=VLOOKUP(MID(B,5,6),Control!D4:G7,3,0)`
  H        Township          `=VLOOKUP(MID(B,5,6),Control!D4:G7,4,0)`

### 4.3 Huawei LTE --- `HW LTE`

  Column   Header                Formula / logic
  -------- --------------------- ------------------------------------------
  H        Throughput(kbps)      `=(C-D)/E`
  I        PRB Utilization (%)   `=F/G*100`
  J        Load Check            `=IF(AND(H<2000,I>70),1,0)`
  K        Region                `=VLOOKUP(MID(B,5,6),Control!D4:G7,3,0)`
  L        Township              `=VLOOKUP(MID(B,5,6),Control!D4:G7,4,0)`

### 4.4 Ericsson LTE --- `Eric LTE`

  Column   Header                Formula / logic
  -------- --------------------- ------------------------------------------
  K        Throughput(kbps)      `=((C-D)/E)*1000`
  L        PRB Utilization (%)   `=((F+G+H+I)/J)*100`
  M        Load Check            `=IF(AND(K<2000,L>70),1,0)`
  N        Region                `=VLOOKUP(MID(B,5,6),Control!D4:G7,3,0)`
  O        Township              `=VLOOKUP(MID(B,5,6),Control!D4:G7,4,0)`

### Load thresholds

  ---------------------------------------------------------------------------------
  Technology      Throughput threshold          Utilization Loaded-cell rule
                                                  threshold 
  --------------- -------------------- -------------------- -----------------------
  3G (UMTS)           Below 1,000 kbps      TCP utilization `Load Check_both = 1`
                                                  above 80% 

  4G (LTE)            Below 2,000 kbps      PRB utilization `Load Check = 1`
                                                  above 70% 
  ---------------------------------------------------------------------------------

**Flag interpretation:** A value of `1` means the row meets the
corresponding condition; `0` means it does not. The throughput-only flag
and the combined throughput-plus-utilization flag are separate checks
for UMTS.

> **Formula validation note:** Confirm source units and counter
> definitions before using these formulas for engineering decisions. In
> particular, verify the UMTS power conversion and the throughput unit
> conversions for both vendors against the vendor counter definitions.
> The formulas above document the stated macro logic rather than
> independently validating the counter units.

------------------------------------------------------------------------

## 5. Summary by Date --- `Control!I:W`

The date keys are unique dates found in `HW UMTS!A`, sorted in ascending
order.

### 5.1 UMTS summary --- `Control!I:P`

  -----------------------------------------------------------------------
  Column                  Output                  Calculation
  ----------------------- ----------------------- -----------------------
  I                       Date                    Unique reporting date

  J                       HW UMTS cell count      Count of Huawei UMTS
                                                  cells for the date

  K                       HW UMTS low-throughput  Count of `HW UMTS` rows
                          count                   where column K = 1

  L                       HW UMTS loaded count    Count of `HW UMTS` rows
                                                  where column L = 1

  M                       Eric UMTS cell count    Count of Ericsson UMTS
                                                  cells for the date

  N                       Eric UMTS               Count of `Eric UMTS`
                          low-throughput count    rows where column E = 1

  O                       Eric UMTS loaded count  Count of `Eric UMTS`
                                                  rows where column F = 1

  P                       Loaded Ratio            `(L+O)/(J+M)`
  -----------------------------------------------------------------------

### 5.2 LTE summary --- `Control!R:W`

  -----------------------------------------------------------------------
  Column                  Output                  Calculation
  ----------------------- ----------------------- -----------------------
  R                       Date                    Reporting date

  S                       HW LTE cell count       Count of Huawei LTE
                                                  cells for the date

  T                       HW LTE loaded count     Count of `HW LTE` rows
                                                  where column J = 1

  U                       Eric LTE cell count     Count of Ericsson LTE
                                                  cells for the date

  V                       Eric LTE loaded count   Count of `Eric LTE`
                                                  rows where column M = 1

  W                       Loaded Ratio            `(T+V)/(S+U)`
  -----------------------------------------------------------------------

The loaded ratio is the total loaded-cell count divided by the total
cell count for the corresponding technology and date. If the total cell
count is zero, the macro should avoid a divide-by-zero result.

------------------------------------------------------------------------

## 6. Summary by Date × Region --- `Control!Y:AH`

The report uses these five regions, with five rows per date and dates
sorted ascending:

1.  North West
2.  North East
3.  South West
4.  South East
5.  Yangon

### 6.1 UMTS summary --- `Control!Y:AC`

  -----------------------------------------------------------------------
  Column                  Output                  Calculation
  ----------------------- ----------------------- -----------------------
  Y                       Date                    Reporting date

  Z                       Region                  One of the five
                                                  configured regions

  AA                      Low-throughput count    `COUNTIFS` for Ericsson
                                                  UMTS column E = 1 plus
                                                  Huawei UMTS column K =
                                                  1, matching Region and
                                                  Date

  AB                      Loaded count            `COUNTIFS` for Ericsson
                                                  UMTS column F = 1 plus
                                                  Huawei UMTS column L =
                                                  1, matching Region and
                                                  Date

  AC                      Total cell count        Ericsson UMTS cell
                                                  count plus Huawei UMTS
                                                  cell count, matching
                                                  Region and Date
  -----------------------------------------------------------------------

### 6.2 LTE summary --- `Control!AE:AH`

  -----------------------------------------------------------------------
  Column                  Output                  Calculation
  ----------------------- ----------------------- -----------------------
  AE                      Date                    Reporting date

  AF                      Region                  One of the five
                                                  configured regions

  AG                      Loaded count            `COUNTIFS` for Ericsson
                                                  LTE column M = 1 plus
                                                  Huawei LTE column J =
                                                  1, matching Region and
                                                  Date

  AH                      Total cell count        Ericsson LTE cell count
                                                  plus Huawei LTE cell
                                                  count, matching Region
                                                  and Date
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 7. Process Flow

``` mermaid
flowchart TD
    A([Start macro]) --> B[Keep Control and EP Information]
    B --> C[Delete all other worksheets]
    C --> D[Select the four daily source files]
    D --> E[Match files by filename pattern]
    E --> F[Import selected source columns]
    F --> G[Calculate UMTS throughput and TCP utilization]
    G --> H[Calculate LTE throughput and PRB utilization]
    H --> I[Apply load-check flags]
    I --> J[Map cell IDs to Region and Township]
    J --> K[Convert formulas to values]
    K --> L[Build daily UMTS and LTE summaries]
    L --> M[Build Date × Region summaries]
    M --> N([Finish])
```

### Process steps

1.  **Prepare workbook:** Retain `Control` and `EP Information`; delete
    other worksheets.
2.  **Select inputs:** Choose the daily Huawei/Ericsson UMTS/LTE files.
3.  **Identify and import:** Match files using their filename patterns
    and import the specified columns from each file's first worksheet.
4.  **Calculate KPIs:** Calculate throughput and utilization for each
    vendor/technology dataset.
5.  **Flag loaded cells:** Apply the configured throughput and
    utilization thresholds.
6.  **Map geography:** Extract the township code from the cell
    identifier and look up Region and Township on `Control`.
7.  **Finalize data:** Calculate formulas, handle formula errors as
    blanks, and convert formula results to values.
8.  **Create summaries:** Populate the daily summaries and Date × Region
    summaries on `Control`.

------------------------------------------------------------------------

## 8. Pre-Run and Post-Run Checks

-   Confirm that `Control` and `EP Information` exist before running the
    macro.
-   Check that the township lookup table in `Control!D4:G7` covers the
    required township codes and has the expected column positions.
-   Select the correct four files and confirm their filenames contain
    the expected patterns.
-   Check that each source file's first worksheet contains the expected
    source columns.
-   Confirm that Date and cell identifier fields are in the expected
    positions and formats.
-   Check for missing region/township lookups and blank KPI values after
    the run.
-   Validate throughput and utilization calculations against a few
    manually calculated sample rows.
-   Confirm that the summary dates and region names are complete and
    correctly sorted.
-   Save the completed workbook under a new name if the original
    workbook needs to be preserved.

## 9. Assumptions and Items to Verify

1.  **Lookup range:** `Control!D4:G7` contains only four rows, while the
    report defines five regions. Confirm whether the lookup range is
    intentionally limited or should be expanded to include all township
    mappings.
2.  **VLOOKUP indexing:** With `D:G` as the lookup range, the third
    column index returns column F (Region), and the fourth returns
    column G (Township). This matches the described layout.
3.  **Date source:** Daily summary keys come from `HW UMTS!A`. If the
    Huawei UMTS file is missing or lacks dates, the macro may not create
    a complete date list for other datasets.
4.  **Missing files:** Since missing files are skipped, summaries may be
    incomplete when one or more files are not selected.
5.  **UMTS power conversion:** Verify that the input in column C is
    represented in the expected unit before applying `=10^(C/100)/1000`.
6.  **Throughput units:** Confirm whether each vendor's source volume
    and time counters require a factor of 1,000 or another conversion to
    produce kbps.
7.  **Loaded ratio:** The daily ratio uses the combined Huawei and
    Ericsson loaded counts divided by the combined cell counts. Ensure
    this is the intended KPI definition.
8.  **Worksheet deletion:** The macro deletes every worksheet other than
    `Control` and `EP Information` at the start of each run. Keep a
    backup of any workbook with additional sheets.

------------------------------------------------------------------------

## 10. Report Purpose

This workbook provides a repeatable daily view of 3G/4G cell throughput
and resource utilization across Huawei and Ericsson networks. It helps
identify cells with low throughput, cells simultaneously experiencing
low throughput and high utilization, and regional patterns that may
require capacity or optimization investigation.
