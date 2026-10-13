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
