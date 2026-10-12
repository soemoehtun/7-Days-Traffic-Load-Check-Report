# Region-Wise Daily Cell Load Check — NE / NW / SE / SW / YN

Daily cell-level **load (congestion) check** for **Huawei** and **Ericsson** networks
(UMTS 3G + LTE 4G), aggregated by **Date × Region** and exported to CSV.

A cell is flagged as **loaded/congested** when user throughput is low **and**
radio-resource utilization (power / PRB) is high. One summary CSV is produced per
vendor & technology.

---

## 1. Overview

| # | Block | Input files (pattern) | Output file |
|---|-------|----------------------|-------------|
| 1 | Huawei UMTS (3G, NBH) | `HW_Data_Update_3G_NBH_Cell*.csv` | `Region_HW_UMTS_Daily_NBH_Cell_Load_Check_result.csv` |
| 2 | Huawei LTE (4G, DBH) | `HW_Data_Update_4G_DBH*.csv` | `Region_HW_LTE_Daily_DBH_Cell_Load_Check_result.csv` |
| 3 | Ericsson UMTS (3G, NBH) | `ERIC_Data_Update_3G_NBH_Cell*.csv` | `Region_ERIC_UMTS_Daily_NBH_Cell_Load_Check_result.csv` |
| 4 | Ericsson LTE (4G, DBH) | `ERIC_Data_Update_4G_DBH*.csv` | `Region_ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv` |

Input files are searched **recursively** (`rglob`), so they may sit in sub-folders.
All outputs are written into the **same date folder** as the inputs.

---

## 2. Requirements

- Python 3.x
- `pandas` **≤ 1.5.x** — the script uses `df.append()`, removed in pandas 2.0 (see caveats)
- `numpy`
- Windows machine with access to the `D:\ATOM\...` paths

---

## 3. Inputs

### 3.1 Site ↔ Region master list

```python
site_region_list = pd.read_csv(r'D:\ATOM\Region Site List\Site_Region_list_20250707.csv')
```

- Required columns: `Site` (6-character site ID) and `Region` (`NE`, `NW`, `SE`, `SW`, `YN`)
- Every cell is tagged with its region via a **left join** on site ID.

### 3.2 Daily KPI export folder

```python
file_dir_1 = Path(r'D:\ATOM\KPI Monitor\Data Update for load\5-Oct-2026')
```

- **Update this date-stamped path for every run.**
- Files expected, and the columns each block uses:

| File pattern | Columns used |
|---|---|
| `HW_Data_Update_3G_NBH_Cell*.csv` | `Date`, `Cell Name`, `VS.HSDPA.DataOutput.Traffic(bit)`, `VS.HSDPA.DataTtiNum.User`, `VS.MeanTCP(dBm)`, `Max Transmit Power of Cell` |
| `HW_Data_Update_4G_DBH*.csv` | `Date`, `Cell Name`, `L.Thrp.bits.DL(bit)`, `L.Thrp.bits.DL.LastTTI(bit)`, `L.Thrp.Time.DL.RmvLastTTI(ms)`, `L.ChMeas.PRB.DL.Used.Avg`, `L.ChMeas.PRB.DL.Avail` |
| `ERIC_Data_Update_3G_NBH_Cell*.csv` | `Date`, `Ucell Id`, `HS User Throughput(Kbps)`, `Power Utilization` |
| `ERIC_Data_Update_4G_DBH*.csv` | `Date`, `EUtranCell Id`, `pmPdcpVolDlDrb`, `pmPdcpVolDlDrbLastTTI`, `pmUeThpTimeDl`, `pmPrbUsedDlDtch`, `pmPrbUsedDlSrbFirstTrans`, `pmPrbUsedDlAilgTot`, `pmPrbUsedDlPcch`, `pmPrbAvailDl` |

---

## 4. Load-check rules (thresholds)

| Vendor / Tech | Throughput metric | Flag if **<** | Resource metric | Flag if **>** | Flags produced |
|---|---|---|---|---|---|
| Huawei UMTS | HSDPA user throughput (kbps) | 1000 kbps | TCP* utilization | 80 % | `Load Check_thrp`, `Load Check_both` |
| Huawei LTE | DL throughput (kbps) | 2000 kbps | DL PRB utilization | 70 % | `Load Check` |
| Ericsson UMTS | HS user throughput (kbps) | 1000 kbps | Power utilization (%) | 80 | `Load Check_thrp`, `Load Check_both` |
| Ericsson LTE | DL throughput (kbps) | 2000 kbps | DL PRB utilization | 70 % | `Load Check` |

\* TCP = **Transmitted Carrier Power** (not the TCP protocol).

- `Load Check_thrp` (3G only) → throughput condition **alone**.
- `Load Check_both` / `Load Check` → **both** conditions must be true
  (low throughput **+** high utilization = congested cell).

---

## 5. Step-by-step walkthrough

### Step 0 — Setup
1. Load the site ↔ region master list (`site_region_list`).
2. Point `file_dir_1` to the date folder holding the day's raw exports.

### Step 1 — Huawei UMTS (3G, NBH)
1. **Collect:** read every `HW_Data_Update_3G_NBH_Cell*.csv` under `file_dir_1` (recursive) and append into one DataFrame.
2. **Site ID:** `site_id = Cell Name[4:10]` → the 6-character site code (same convention for all vendors).
3. **HSDPA user throughput (kbps):**
   `VS.HSDPA.DataOutput.Traffic(bit) / (VS.HSDPA.DataTtiNum.User × 2)`
   (2 ms = HSDPA TTI, so bits ÷ ms = kbps)
4. **Mean TCP → watts:** `10^(VS.MeanTCP(dBm)/10) / 1000`
5. **Max cell power → watts:** `10^(Max Transmit Power of Cell / 100) / 1000`
   (the ÷100 treats the value as a 0.1 dBm scale, e.g. 430 → 43 dBm → 20 W)
6. **TCP utilization:** mean TCP (W) ÷ max power (W)
7. **Flags:**
   - `Load Check_thrp = 1` if throughput < 1000 kbps, else 0
   - `Load Check_both = 1` if throughput < 1000 kbps **and** TCP utilization > 0.8, else 0
8. **`nan_count`:** count of NaN values in `Load Check_both` (see caveats).
9. **Region tag:** left-merge with `site_region_list` on `site_id ↔ Site`.
10. **Aggregate:** group by `Date, Region` → `sum` + `count` of each flag.
11. **Export:** `Region_HW_UMTS_Daily_NBH_Cell_Load_Check_result.csv`.

### Step 2 — Huawei LTE (4G, DBH)
1. Collect every `HW_Data_Update_4G_DBH*.csv` into one DataFrame.
2. `site_id = Cell Name[4:10]`.
3. **DL throughput (kbps):**
   `(L.Thrp.bits.DL(bit) − L.Thrp.bits.DL.LastTTI(bit)) / L.Thrp.Time.DL.RmvLastTTI(ms)`
   (bits ÷ ms = kbps)
4. **PRB utilization:** `L.ChMeas.PRB.DL.Used.Avg / L.ChMeas.PRB.DL.Avail`
5. **Flag:** `Load Check = 1` if throughput < 2000 kbps **and** PRB utilization > 0.7.
6. nan_count → region merge → group by `Date, Region` (`sum`, `count`) →
   export `Region_HW_LTE_Daily_DBH_Cell_Load_Check_result.csv`.

### Step 3 — Ericsson UMTS (3G, NBH)
1. Collect every `ERIC_Data_Update_3G_NBH_Cell*.csv`.
2. `site_id = Ucell Id[4:10]`.
3. **Flags (throughput comes straight from the file):**
   - `Load Check_thrp = 1` if `HS User Throughput(Kbps)` < 1000
   - `Load Check_both = 1` if throughput < 1000 **and** `Power Utilization` > 80
4. nan_count → region merge → group by `Date, Region` →
   export `Region_ERIC_UMTS_Daily_NBH_Cell_Load_Check_result.csv`.

### Step 4 — Ericsson LTE (4G, DBH)
1. Collect every `ERIC_Data_Update_4G_DBH*.csv`.
2. `site_id = EUtranCell Id[4:10]`.
3. **DL throughput (kbps):**
   `(pmPdcpVolDlDrb − pmPdcpVolDlDrbLastTTI) / pmUeThpTimeDl × 1000`
4. **PRB utilization:**
   `(pmPrbUsedDlDtch + pmPrbUsedDlSrbFirstTrans + pmPrbUsedDlAilgTot + pmPrbUsedDlPcch) / pmPrbAvailDl`
   (user-plane + signalling + other channels ÷ available PRBs)
5. **Flag:** `Load Check = 1` if throughput < 2000 kbps **and** PRB utilization > 0.7.
6. Region merge → group by `Date, Region` →
   export `Region_ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv`.

### Step 5 — Finish
Prints `Regional load check successfully done.`

---

## 6. Output format

Each output CSV is indexed by **Date × Region** and has a **two-row header**
(result of `agg(['sum', 'count'])` creating a MultiIndex):

```
Date,Region,Load Check_thrp,,Load Check_both,,nan_count,,
,,sum,count,sum,count,sum,count
2026-10-05,NE,142,5210,37,5210,0,5210
2026-10-05,NW,98,4380,21,4380,0,4380
```
*(values illustrative)*

| Column | Meaning |
|---|---|
| `sum` | number of cells that met the flag condition (flag = 1) = **loaded cells** |
| `count` | total cell records for that Date × Region |
| Load ratio | `sum / count` (e.g. 37/5210 ≈ 0.7 % of cells congested) |
| `nan_count` | rows where the check could not be evaluated (see caveats) |

Cells whose `site_id` is **not found** in the master list get `Region = NaN`
and appear grouped under a blank region in the output.

---

## 7. How to run

1. Place the day's four export file types into
   `D:\ATOM\KPI Monitor\Data Update for load\<date>` (sub-folders allowed).
2. Update `file_dir_1` in the script to that folder.
3. If the master list changed, update the `Site_Region_list_*.csv` path.
4. Run the script:
   ```bash
   python region_load_check.py
   ```
5. Collect the four `Region_*_Load_Check_result.csv` files from the same folder.

---

## 8. Notes & known caveats

1. **`df.append()`** was deprecated in pandas 1.4 and **removed in pandas 2.0** —
   run with pandas ≤ 1.5.x, or replace the loops with `pd.concat([...])`.
2. **`nan_count`** is computed once as a single scalar and broadcast to every row, so in the
   aggregation its `sum` = scalar × row count and its `count` = row count. Also, since
   `np.where` only returns 0/1, the flag columns never contain NaN and `nan_count` will
   read 0 in practice — compute `.isna().sum()` on the *input* columns instead if you
   want the true missing-data count.
3. **Ericsson LTE throughput** is multiplied by 1000 — verify the unit of `pmUeThpTimeDl`
   in your export (if it is in ms, bits ÷ ms is already kbps, and the ×1000 would yield Mbps).
4. **Hard-coded Windows paths** — the date folder must be edited for every run.
5. **Two-row headers** in the output CSVs — read them back with
   `pd.read_csv(..., header=[0, 1])` if post-processing.
