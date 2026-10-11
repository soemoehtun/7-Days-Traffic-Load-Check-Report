# Step-by-Step Explanation of This Python Load Check Script

> **Goal of this script:**
This script reads multiple vendor cell KPI CSV files (Huawei 3G, Huawei 4G, Ericsson 3G, Ericsson 4G), calculates **"Low Throughput but High Load" cells**, groups them by Date and tells us **how many cells are under that condition per day**.
> 

## 1. Understanding the Purpose of This Code

The main idea behind this script is

> **"Find cells which are giving low user throughput but their radio resource/power is highly utilized."**
> 

This type of check is very common in **Radio Optimization / KPI Monitoring**.

We are trying to identify **possible load-limited cells**.

Example logic:

- **UMTS (3G)** – Check using **HSDPA User Throughput + TCP Power Utilization**
- **LTE (4G)** – Check using **DL User Throughput + PRB Utilization**

## 2. Importing Required Libraries

```python
import pandas as pd
import numpy as np
from pathlib import Path
import os
```

Let's understand each library one by one.

| Library | Purpose |
| --- | --- |
| `pandas` | To read CSV files, merge them, calculate columns, group data |
| `numpy` | To use `np.where()` for conditions (IF-ELSE logic) |
| `pathlib.Path` | To handle folder paths in Python |
| `os` | To join folder path and output file name |

## 3. Defining the Input Folder

```python
# Huawei UMTS
file_dir_1 = Path(r'D:\ATOM\KPI Monitor\Data Update for load\5-Oct-2026')
```

This means

> Python will go to this folder and **search for all required CSV files inside this folder and all its subfolders**.
> 

The `r` before the string means **raw string**. This prevents Python from misreading backslashes `\`.

## 4. PART-A: HUAWEI UMTS (3G) – NBH Cell Load Check

> NBH = Usually **Non-Busy Hour** / Network Busy Hour related report (vendor naming)
> 

### Step 4.1 – Create an empty DataFrame & read all Huawei 3G files

```python
df = pd.DataFrame()
for f in file_dir_1.rglob('HW_Data_Update_3G_NBH_Cell*.csv'):
    df_temp = pd.read_csv(f)
    df = df.append(df_temp)
```

### What does `rglob()` mean?

```
rglob = recursive glob
```

It means:

> **Search this folder + all subfolders for files starting with HW_Data_Update_3G_NBH_Cell and ending with .csv**
> 

### What does `df.append()` do?

It **reads one CSV at a time and appends (stacks vertically)** all files into a single DataFrame.

> After this loop, `df` contains **combined data of all Huawei UMTS NBH cell CSV files**.
> 

> **Note:** In newer pandas versions, `pd.concat()` is recommended instead of `df.append()`. But this code will still work in many environments.
> 

### Step 4.2 – Calculate HSDPA User Throughput

```python
df['HSDPA_UserThroughput'] = df['VS.HSDPA.DataOutput.Traffic(bit)']/(df['VS.HSDPA.DataTtiNum.User']*2)
```

We need to understand the logic behind this formula.

#### What is TTI in UMTS HSPA?

In HSDPA, **TTI (Transmission Time Interval) = 2 ms**

So,

> **Total Transmission Time considered = (Number of User TTIs) × 2 ms**
> 

And

> **Total Data Transferred = VS.HSDPA.DataOutput.Traffic(bit)**
> 

Therefore,

$$
HSDPA\ User\ Throughput\ (bits/ms) = \frac{Total\ Bits}{User\ TTI\ Count \times 2}
$$

But in KPI reports, this calculated value is usually interpreted in **kbps-like units depending on counter definition**.

> **The important thing here is: We are deriving cell-level HSDPA user throughput from counters.**
> 

### Step 4.3 – Convert Mean TCP from dBm to Watts

```python
df['VS.MeanTCP(W)'] = df['VS.MeanTCP(dBm)'].apply(lambda x: 10**(x/10)/1000)
```

#### Why convert dBm to Watts?

KPI `VS.MeanTCP(dBm)` is in **decibel-milliwatts (dBm)**.

But `Max Transmit Power of Cell` in this formula is being converted to **Watts** later.

> **To compare both in the same unit (Watts), we convert TCP power to Watts.**
> 

#### dBm to Watt formula

$$
Power(Watt) = \frac{10^{(dBm/10)}}{1000}
$$

**Example understanding:**

- 43 dBm ≈ 20 Watts
- 40 dBm ≈ 10 Watts

### Step 4.4 – Convert Cell Max Transmit Power to Watts

```python
df['Max Transmit Power of Cell_watt'] = df['Max Transmit Power of Cell'].apply(lambda x: 10**(x/100)/1000)
```

Now look carefully here.

Normally cell power is given as **dBm**.

But here they are doing

```
10^(x/100) / 1000
```

> **This means the column `Max Transmit Power of Cell` is expected in a different numeric format in Huawei CSV.**
> 

The logic idea is still the same:

> **Get Maximum Cell Transmit Power in Watts**
> 

### Step 4.5 – Calculate TCP Utilization

```python
df['TCP Utilization'] = df['VS.MeanTCP(W)']/df['Max Transmit Power of Cell_watt']
```

**TCP = Transmit Channel Power / Total Cell Power used idea**

In simple words:

> **TCP Utilization = (Average TCP Power used by cell) ÷ (Cell Maximum Transmit Power)**
> 

This tells us

> **"How much of the cell downlink power is being used on average?"**
> 

If this value is **close to 1.0**, that means **the cell is power-heavy / power utilized high**.

If this value is **low**, power usage is low.

### Step 4.6 – Load Check Conditions

#### Condition 1 – Low Throughput Check

```python
df['Load Check_thrp'] = np.where((df['HSDPA_UserThroughput'] < 1000), 1, 0)
```

> **If HSDPA User Throughput < 1000 → Mark as 1 else 0**
> 

Meaning:

> **Cell is giving low HSDPA user throughput**
> 

#### Condition 2 – Low Throughput AND High TCP Utilization

```python
df['Load Check_both'] = np.where(((df['HSDPA_UserThroughput'] < 1000) & (df['TCP Utilization'] > 0.8)), 1, 0)
```

Now read this condition in plain English.

> **IF (Low Throughput < 1000) AND (TCP Utilization > 0.8) THEN 1 ELSE 0**
> 

**Interpretation:**

> **Cell has low throughput BUT cell power utilization is more than 80%**
> 

> **This generally indicates the cell is trying to serve users but users are getting low throughput. This is suspicious for a load/power limited cell.**
> 

### Step 4.7 – Count NaN values

```python
df['nan_count'] = df['Load Check_both'].isna().sum()
```

`isna()` finds **blank / missing / NaN** values.

`.sum()` counts how many rows have NaN.

**Why is this added?**

> To quickly check **"Did any calculation give missing values?"**
> 

Common reasons for NaN here:

- Division by zero (`User TTI = 0`)
- `Max Transmit Power of Cell_watt = 0`
- Missing KPI counters in CSV
- Blank cells in source file

### Step 4.8 – Group by Date

```python
df_1 = df.groupby('Date')[['Load Check_thrp', 'Load Check_both','nan_count']].agg(['sum','count'])
df_1 = df_1.reset_index()
```

Now this is a very important step.

We are doing **grouping by Date column**.

`.agg(['sum','count'])` means for each Date:

| Function | Meaning |
| --- | --- |
| `sum` | Total number of cells where value = **1** |
| `count` | Total number of cells **considered** for that column |

So for `Load Check_thrp`

> **sum = How many cells had low throughput on that date**
> 
> 
> **count = Total cells reported on that date**
> 

Similarly for `Load Check_both`

> **sum = How many cells are low throughput + high power utilization on that date**
> 

After `reset_index()`, `Date` comes back as a normal column instead of an index.

### Step 4.9 – Export Result

```python
df_1.to_csv(os.path.join(file_dir_1,'HW_UMTS_Daily_NBH_Cell_Load_Check_result.csv'), index=False, header=True)
```

**Output file name:**

```
HW_UMTS_Daily_NBH_Cell_Load_Check_result.csv
```

This CSV will be saved **inside the same input folder**.

## 5. PART-B: HUAWEI LTE (4G) – DBH Cell Load Check

```python
# Huawei LTE

df = pd.DataFrame()
for f in file_dir_1.rglob('HW_Data_Update_4G_DBH*.csv'):
    df_temp = pd.read_csv(f, sep=',')
    df = df.append(df_temp)
```

Notice here

```python
pd.read_csv(f, sep=',')
```

Comma is the default separator anyway. So this is explicit.

### Step 5.1 – Calculate LTE DL Throughput (kbps)

```python
df['Throughput(kbps)'] = (df['L.Thrp.bits.DL(bit)'] - df['L.Thrp.bits.DL.LastTTI(bit)'])/df['L.Thrp.Time.DL.RmvLastTTI(ms)']
```

#### Why subtract LastTTI?

In Huawei LTE counters, sometimes

> **`L.Thrp.bits.DL(bit)` includes the last TTI bits**
> 
> 
> **`L.Thrp.bits.DL.LastTTI(bit)` gives those last TTI bits**
> 

To get **accurate throughput calculation**, they remove that last TTI.

#### Why divide by time in ms?

Formula logic:

$$
Throughput(bit/ms) = \frac{DL\ Bits\ (excluding\ Last\ TTI)}{DL\ Throughput\ Time\ (excluding\ Last\ TTI)}
$$

Since time is in **milliseconds (ms)**

> **bit/ms = 1000 kbps**
> 

Wait, let’s be clear.

> **1 bit/ms = 1000 bits/sec = 1 kbps**
> 

Yes.

So this calculation gives **Throughput in kbps**.

### Step 5.2 – Calculate DL PRB Utilization

```python
df['PRB Utilization'] = df['L.ChMeas.PRB.DL.Used.Avg']/df['L.ChMeas.PRB.DL.Avail']
```

**PRB = Physical Resource Block**

In LTE, **all user data on DL is scheduled on PRBs**.

> **PRB Utilization = (Average DL PRBs Used) ÷ (Average DL PRBs Available)**
> 

This is a **ratio between 0 and 1**

**Example:**

- 0.50 = 50% PRB used
- 0.75 = 75% PRB used
- 0.85 = 85% PRB used

### Step 5.3 – Huawei LTE Load Check Condition

```python
df['Load Check'] = np.where(((df['Throughput(kbps)'] < 2000) & (df['PRB Utilization'] > 0.7)), 1, 0)
```

**Plain English:**

> **If DL Throughput < 2000 kbps AND DL PRB Utilization > 70% → Mark 1**
> 

**Meaning:**

> **Cell is heavily loaded in PRBs but users are getting low DL throughput**
> 

This matches the exact same optimization idea.

> **Low Throughput + High Resource Utilization = Potential Load Limited Cell**
> 

### Step 5.4 – Group & Save

```python
df['nan_count'] = df['Load Check'].isna().sum()
df_2 = df.groupby('Date')[['Load Check','nan_count']].agg(['sum','count'])
df_2 = df_2.reset_index()
df_2.to_csv(os.path.join(file_dir_1,'HW_LTE_Daily_DBH_Cell_Load_Check_result.csv'), index=False, header=True)
```

Output:

```
HW_LTE_Daily_DBH_Cell_Load_Check_result.csv
```

## 6. PART-C: ERICSSON UMTS (3G) – NBH Cell Load Check

Now vendor changes from **Huawei to Ericsson**.

So **column names are completely different**.

This is extremely important to notice.

```python
df = pd.DataFrame()
for f in file_dir_1.rglob('ERIC_Data_Update_3G_NBH_Cell*.csv'):
    df_temp = pd.read_csv(f, sep=',')
    df = df.append(df_temp)
```

Reading all Ericsson UMTS NBH files.

### Step 6.1 – Low Throughput Check

```python
df['Load Check_thrp'] = np.where((df['HS User Throughput(Kbps)'] < 1000), 1, 0)
```

Unlike Huawei, **Ericsson already gives HS User Throughput directly in Kbps**.

> **We do NOT need to calculate it from counters here.**
> 

### Step 6.2 – Low Throughput + High Power Utilization

```python
df['Load Check_both'] = np.where(((df['HS User Throughput(Kbps)'] < 1000) & (df['Power Utilization'] > 80)), 1, 0)
```

Look at this difference carefully.

**Huawei UMTS used:** `TCP Utilization > 0.8`  **(ratio 0–1)**

**Ericsson UMTS used:** `Power Utilization > 80` **(percentage %)**

> **Same logic. Different unit from vendor report.**
> 

**Meaning:**

> **HS User Throughput < 1000 Kbps AND Power Utilization > 80%**
> 

### Step 6.3 – Daily Summary

```python
df['nan_count'] = df['Load Check_both'].isna().sum()
df_1 = df.groupby('Date')[['Load Check_thrp', 'Load Check_both','nan_count']].agg(['sum','count'])
df_1 = df_1.reset_index()
df_1.to_csv(os.path.join(file_dir_1,'ERIC_UMTS_Daily_NBH_Cell_Load_Check_result.csv'), index=False, header=True)
```

> **Note:** Here variable name `df_1` is reused. It overwrites Huawei UMTS result DataFrame. But we already saved Huawei UMTS CSV, so **that is not a problem**.
> 

## 7. PART-D: ERICSSON LTE (4G) – DBH Cell Load Check

```python
df = pd.DataFrame()
for f in file_dir_1.rglob('ERIC_Data_Update_4G_DBH*.csv'):
    df_temp = pd.read_csv(f, sep=',')
    df = df.append(df_temp)
```

### Step 7.1 – Calculate Ericsson LTE DL Throughput (kbps)

```python
df['Throughput(kbps)'] = (df['pmPdcpVolDlDrb'] - df['pmPdcpVolDlDrbLastTTI'])/df['pmUeThpTimeDl'] * 1000
```

Now compare this with Huawei LTE formula.

**Huawei LTE:**

```
(L.Thrp.bits.DL(bit) - L.Thrp.bits.DL.LastTTI(bit)) / L.Thrp.Time.DL.RmvLastTTI(ms)
```

**Ericsson LTE:**

```
(pmPdcpVolDlDrb - pmPdcpVolDlDrbLastTTI) / pmUeThpTimeDl * 1000
```

#### Why multiply by 1000?

Because in Ericsson, `pmUeThpTimeDl` is usually in **seconds** or in a different time unit.

By multiplying with **1000**, they convert the result into **kbps**.

> **Vendor counter naming differs, but the mathematical idea is identical: (Bits transferred) / (Throughput time) = Throughput**
> 

### Step 7.2 – Calculate Ericsson LTE DL PRB Utilization

```python
df['PRB Utilization'] = (df['pmPrbUsedDlDtch'] + df['pmPrbUsedDlSrbFirstTrans'] + df['pmPrbUsedDlAilgTot'] + df['pmPrbUsedDlPcch'])/df['pmPrbAvailDl']
```

Look here very carefully.

**Huawei LTE takes:** `Single counter` → `L.ChMeas.PRB.DL.Used.Avg`

**Ericsson LTE takes:** `Sum of multiple PRB used counters`

$$
Total\ PRB\ Used = pmPrbUsedDlDtch + pmPrbUsedDlSrbFirstTrans + pmPrbUsedDlAilgTot + pmPrbUsedDlPcch
$$

$$
PRB\ Utilization = \frac{Total\ DL\ PRB\ Used}{pmPrbAvailDl}
$$

> **Reason:** Ericsson splits DL PRB usage into different channels. So we must **sum them** to get total DL PRB usage.
> 

### Step 7.3 – Same LTE Load Condition

```python
df['Load Check'] = np.where(((df['Throughput(kbps)'] < 2000) & (df['PRB Utilization'] > 0.7)), 1, 0)
```

Same threshold as **Huawei LTE**

- Throughput < **2000 kbps**
- PRB Utilization > **0.7 (70%)**

### Step 7.4 – Save Result

```python
df['nan_count'] = df['Load Check'].isna().sum()
df_3 = df.groupby('Date')[['Load Check','nan_count']].agg(['sum','count'])
df_3 = df_3.reset_index()
df_3.to_csv(os.path.join(file_dir_1,'ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv'), index=False, index=False, header=True)  # (note small duplicate index=False)
df_3.to_csv(os.path.join(file_dir_1,'ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv'), index=False, header=True)
```

> Actually in code written: `index=False, index=False` appears once but logically it means saving without row numbers.
> 

Final output:

```
ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv
```

## 8. Final Print

```python
print("Load check successfully done.")
```

After all 4 vendor reports are processed and 4 CSV files are created, this message appears in terminal/Spyder/VS Code.

## 9. Complete Output Files Generated

After running this entire script, you will get **4 CSV files** in

`D:\ATOM\KPI Monitor\Data Update for load\5-Oct-2026\`

| # | File Name | Vendor | Technology |
| --- | --- | --- | --- |
| 1 | `HW_UMTS_Daily_NBH_Cell_Load_Check_result.csv` | Huawei | UMTS 3G |
| 2 | `HW_LTE_Daily_DBH_Cell_Load_Check_result.csv` | Huawei | LTE 4G |
| 3 | `ERIC_UMTS_Daily_NBH_Cell_Load_Check_result.csv` | Ericsson | UMTS 3G |
| 4 | `ERIC_LTE_Daily_DBH_Cell_Load_Check_result.csv` | Ericsson | LTE 4G |

## 10. How to Read These Output CSV Files

Take example of Huawei UMTS output.

Your columns will look like this after MultiIndex from `groupby.agg`

Because of

```python
.agg(['sum','count'])
```

You get columns like:

```
Date
Load Check_thrp	sum
Load Check_thrp	count
Load Check_both	sum
Load Check_both	count
nan_count	sum
nan_count	count
```

### Meaning in simple words

> **"On Date X, out of total NBH 3G cells, how many cells had low throughput? And how many cells had low throughput with high TCP utilization?"**
> 

## 11. Simple Logic Summary (Most Important)

> **Core Question this script answers**
> 
> 
> ## *"How many cells are giving LOW THROUGHPUT but operating at HIGH LOAD?"*
> 

| Technology | High Load Indicator |
| --- | --- |
| **Huawei UMTS** | High **TCP Power Utilization (>80%)** |
| **Ericsson UMTS** | High **Power Utilization (>80%)** |
| **Huawei LTE** | High **PRB Utilization (>70%)** |
| **Ericsson LTE** | High **PRB Utilization (>70%)** |

## 12. Step-by-Step Execution Flow (In Order)

```
Step 1  → Set folder path
Step 2  → Read all Huawei 3G CSVs → Merge
Step 3  → Derive HSDPA Throughput
Step 4  → Convert dBm to Watts
Step 5  → Calculate TCP Utilization
Step 6  → Apply UMTS load conditions
Step 7  → Group by Date & export HW UMTS result
Step 8  → Read all Huawei 4G CSVs → Merge
Step 9  → Calculate LTE Throughput & PRB Utilization
Step 10 → Apply LTE load condition & export HW LTE result
Step 11 → Read all Ericsson 3G CSVs → Merge
Step 12 → Apply Ericsson UMTS condition & export ERIC UMTS result
Step 13 → Read all Ericsson 4G CSVs → Merge
Step 14 → Calculate LTE Throughput & PRB Utilization
Step 15 → Apply Ericsson LTE condition & export ERIC LTE result
Step 16 → Print success message
```

## 13. One Advice to Understand This Better

**Do NOT try to memorize this code. Try to memorize this question.**

> **"For each vendor + each technology, I am filtering cells where throughput is low but resource usage is high and counting them date-wise."**
> 

## 14. My Recommendation For Learning

The smartest way to learn this script is to do this **practically**.

Open **one single CSV file** from your folder in Excel or Notepad++.

Now check these column names exist or not:

### For Huawei 3G

```
Date
VS.HSDPA.DataOutput.Traffic(bit)
VS.HSDPA.DataTtiNum.User
VS.MeanTCP(dBm)
Max Transmit Power of Cell
```

### For Huawei 4G

```
Date
L.Thrp.bits.DL(bit)
L.Thrp.bits.DL.LastTTI(bit)
L.Thrp.Time.DL.RmvLastTTI(ms)
L.ChMeas.PRB.DL.Used.Avg
L.ChMeas.PRB.DL.Avail
```

### For Ericsson 3G

```
Date
HS User Throughput(Kbps)
Power Utilization
```

### For Ericsson 4G

```
Date
pmPdcpVolDlDrb
pmPdcpVolDlDrbLastTTI
pmUeThpTimeDl
pmPrbUsedDlDtch
pmPrbUsedDlSrbFirstTrans
pmPrbUsedDlAilgTot
pmPrbUsedDlPcch
pmPrbAvailDl
```

> **Golden Rule**
> 
> 
> ## **If you understand "Which columns → Which formula → Which condition → Group by Date → Count 1s", you completely understand this entire script.**
>