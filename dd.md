Here's a comprehensive `README.md` file for your **Daily 3G/4G Throughput Combination Report** project, including detailed formula explanations:

```markdown
# Daily 3G/4G Throughput Combination Report

![Excel VBA](https://img.shields.io/badge/Excel-VBA-blue)
![License](https://img.shields.io/badge/License-MIT-green)

Automated Excel VBA solution for combining and analyzing 3G/4G network throughput data from Huawei and Ericsson equipment.

## 📌 Overview

This tool processes daily traffic data from multiple network vendors (Huawei & Ericsson) for both 3G (UMTS) and 4G (LTE) technologies, calculates key performance indicators, and generates a consolidated report with regional summaries.

## 📁 File Structure

```
Daily-3G-4G-Throughput-Combination/
├── Report.xlsm                # Main Excel file with VBA macro
├── README.md                  # This documentation
└── Input Files/               # Folder for source data files
    ├── hw_data_update_3g_nbh_cell_*.xlsx
    ├── eric_data_update_3g_nbh_cell_*.xlsx
    ├── hw_data_update_4g_dbh_*.xlsx
    └── eric_data_update_4g_dbh_*.xlsx
```

## 🚀 How to Use

1. **Prepare Input Files**
   - Place all 4 required files in the same folder
   - Files must follow naming convention:
     - `hw_data_update_3g_nbh_cell_YYYYMMDD.xlsx`
     - `eric_data_update_3g_nbh_cell_YYYYMMDD.xlsx`
     - `hw_data_update_4g_dbh_YYYYMMDD.xlsx`
     - `eric_data_update_4g_dbh_YYYYMMDD.xlsx`

2. **Run the Report**
   - Open `Report.xlsm`
   - Enable macros when prompted
   - Click the "Run Report" button or press `Alt+F8` and run `ReportProcess` macro
   - Select all 4 input files when prompted

3. **View Results**
   - **Control Sheet**: Contains consolidated summary
   - **HW UMTS/Eric UMTS**: 3G data with calculations
   - **HW LTE/Eric LTE**: 4G data with calculations

## 🧮 Key Formulas Explained

### 3G (UMTS) Calculations

#### Huawei 3G Formulas
| Column | Formula | Description |
|--------|---------|-------------|
| G (HSDPA_UserThroughput) | `=IFERROR(RC4/(RC5*2),"")` | User throughput = HSDPA Traffic Volume / (HSDPA User Number * 2) |
| H (VS.MeanTCP) | `=IFERROR(10^(RC6/10)/1000,"")` | Convert TCP power from dBm to watts |
| I (Max Transmit Power) | `=IFERROR(10^(RC3/100)/1000,"")` | Convert cell power from dBm to watts |
| J (TCP Utilization) | `=IFERROR(RC8/RC9*100,"")` | TCP utilization percentage |
| K (Load Check_thrp) | `=IFERROR(IF(RC7<1000,1,0),"")` | 1 if throughput < 1000kbps (congested) |
| L (Load Check_both) | `=IFERROR(IF(AND(RC7<1000,RC10>80),1,0),"")` | 1 if throughput < 1000kbps AND PRB > 80% |
| M (Region) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0)` | Extract region from cell ID |
| N (Township) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0)` | Extract township from cell ID |

#### Ericsson 3G Formulas
| Column | Formula | Description |
|--------|---------|-------------|
| E (Load Check_thrp) | `=IFERROR(IF(RC3="", "", IF(RC3<1000,1,0)),"")` | 1 if throughput < 1000kbps |
| F (Load Check_both) | `=IFERROR(IF(RC3="", "", IF(AND(RC3<1000,RC4>80),1,0)),"")` | 1 if throughput < 1000kbps AND PRB > 80% |
| G (Region) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0)` | Extract region from cell ID |
| H (Township) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0)` | Extract township from cell ID |

### 4G (LTE) Calculations

#### Huawei 4G Formulas
| Column | Formula | Description |
|--------|---------|-------------|
| H (Throughput) | `=IFERROR((RC3-RC4)/RC5,"")` | Throughput = (DL Traffic - Retransmission)/PRB Utilization |
| I (PRB Utilization) | `=IFERROR(RC6/RC7*100,"")` | PRB utilization percentage |
| J (Load Check) | `=IFERROR(IF(AND(RC8<2000,RC9>70),1,0),"")` | 1 if throughput < 2000kbps AND PRB > 70% |
| K (Region) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0)` | Extract region from cell ID |
| L (Township) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0)` | Extract township from cell ID |

#### Ericsson 4G Formulas
| Column | Formula | Description |
|--------|---------|-------------|
| K (Throughput) | `=IFERROR(((RC3-RC4)/RC5)*1000,"")` | Throughput = (DL Traffic - Retransmission)/PRB Utilization (kbps) |
| L (PRB Utilization) | `=IFERROR(((RC6+RC7+RC8+RC9)/RC10)*100,"")` | PRB utilization percentage |
| M (Load Check) | `=IFERROR(IF(AND(RC11<2000,RC12>70),1,0),"")` | 1 if throughput < 2000kbps AND PRB > 70% |
| N (Region) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,3,0)` | Extract region from cell ID |
| O (Township) | `=VLOOKUP(MID(RC2,5,6),Control!C4:C7,4,0)` | Extract township from cell ID |

## 📊 Summary Calculations

The Control sheet contains two main summary tables:

### UMTS Summary (Columns J-P)
| Column | Calculation | Description |
|--------|-------------|-------------|
| J | HW UMTS Total | Count of all Huawei 3G cells |
| K | HW UMTS K=1 | Count of Huawei 3G cells with throughput congestion |
| L | HW UMTS L=1 | Count of Huawei 3G cells with both throughput and PRB congestion |
| M | Eric UMTS Total | Count of all Ericsson 3G cells |
| N | Eric UMTS E=1 | Count of Ericsson 3G cells with throughput congestion |
| O | Eric UMTS F=1 | Count of Ericsson 3G cells with both throughput and PRB congestion |
| P | Congestion Rate | (HW L=1 + Eric F=1) / (HW Total + Eric Total) |

### LTE Summary (Columns S-W)
| Column | Calculation | Description |
|--------|-------------|-------------|
| S | HW LTE Total | Count of all Huawei 4G cells |
| T | HW LTE J=1 | Count of Huawei 4G cells with congestion |
| U | Eric LTE Total | Count of all Ericsson 4G cells |
| V | Eric LTE M=1 | Count of Ericsson 4G cells with congestion |
| W | Congestion Rate | (HW J=1 + Eric M=1) / (HW Total + Eric Total) |

## ⚠️ Requirements

- Microsoft Excel 2010 or later
- Macros must be enabled
- Input files must follow exact naming convention

## 🛠️ Troubleshooting

**Issue: "File not found" error**
- Verify all 4 input files exist in the same folder
- Check file names match the required pattern

**Issue: Slow performance**
- Close other Excel files before running
- Ensure input files are not too large (>100,000 rows)

**Issue: Incorrect calculations**
- Verify Control sheet has correct region mappings
- Check input files have expected column structure

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
```

## Formula Reference Sheet (marco.md)

For your GitHub repository, here's the `marco.md` file with all formulas in a more technical format:

```markdown
# Daily 3G/4G Throughput Combination - Formula Reference

## 3G (UMTS) Calculations

### Huawei 3G
```excel
G2 (HSDPA_UserThroughput) = IFERROR(D2/(E2*2), "")
H2 (VS.MeanTCP) = IFERROR(10^(F2/10)/1000, "")
I2 (Max Transmit Power) = IFERROR(10^(C2/100)/1000, "")
J2 (TCP Utilization) = IFERROR(H2/I2*100, "")
K2 (Load Check_thrp) = IFERROR(IF(G2<1000,1,0), "")
L2 (Load Check_both) = IFERROR(IF(AND(G2<1000,J2>80),1,0), "")
M2 (Region) = VLOOKUP(MID(B2,5,6),Control!$C$4:$D$7,2,0)
N2 (Township) = VLOOKUP(MID(B2,5,6),Control!$C$4:$E$7,3,0)
```

### Ericsson 3G
```excel
E2 (Load Check_thrp) = IFERROR(IF(C2="", "", IF(C2<1000,1,0)), "")
F2 (Load Check_both) = IFERROR(IF(C2="", "", IF(AND(C2<1000,D2>80),1,0)), "")
G2 (Region) = VLOOKUP(MID(B2,5,6),Control!$C$4:$D$7,2,0)
H2 (Township) = VLOOKUP(MID(B2,5,6),Control!$C$4:$E$7,3,0)
```

## 4G (LTE) Calculations

### Huawei 4G
```excel
H2 (Throughput) = IFERROR((C2-D2)/E2, "")
I2 (PRB Utilization) = IFERROR(F2/G2*100, "")
J2 (Load Check) = IFERROR(IF(AND(H2<2000,I2>70),1,0), "")
K2 (Region) = VLOOKUP(MID(B2,5,6),Control!$C$4:$D$7,2,0)
L2 (Township) = VLOOKUP(MID(B2,5,6),Control!$C$4:$E$7,3,0)
```

### Ericsson 4G
```excel
K2 (Throughput) = IFERROR(((C2-D2)/E2)*1000, "")
L2 (PRB Utilization) = IFERROR(((F2+G2+H2+I2)/J2)*100, "")
M2 (Load Check) = IFERROR(IF(AND(K2<2000,L2>70),1,0), "")
N2 (Region) = VLOOKUP(MID(B2,5,6),Control!$C$4:$D$7,2,0)
O2 (Township) = VLOOKUP(MID(B2,5,6),Control!$C$4:$E$7,3,0)
```

## Summary Calculations

### UMTS Summary (Control Sheet)
```excel
J4 (HW UMTS Total) = COUNTIF('HW UMTS'!A:A, I4)
K4 (HW UMTS K=1) = COUNTIFS('HW UMTS'!A:A, I4, 'HW UMTS'!K:K, 1)
L4 (HW UMTS L=1) = COUNTIFS('HW UMTS'!A:A, I4, 'HW UMTS'!L:L, 1)
M4 (Eric UMTS Total) = COUNTIF('Eric UMTS'!A:A, I4)
N4 (Eric UMTS E=1) = COUNTIFS('Eric UMTS'!A:A, I4, 'Eric UMTS'!E:E, 1)
O4 (Eric UMTS F=1) = COUNTIFS('Eric UMTS'!A:A, I4, 'Eric UMTS'!F:F, 1)
P4 (Congestion Rate) = (L4+O4)/(J4+M4)
```

### LTE Summary (Control Sheet)
```excel
S4 (HW LTE Total) = COUNTIF('HW LTE'!A:A, R4)
T4 (HW LTE J=1) = COUNTIFS('HW LTE'!A:A, R4, 'HW LTE'!J:J, 1)
U4 (Eric LTE Total) = COUNTIF('Eric LTE'!A:A, R4)
V4 (Eric LTE M=1) = COUNTIFS('Eric LTE'!A:A, R4, 'Eric LTE'!M:M, 1)
W4 (Congestion Rate) = (T4+V4)/(S4+U4)
```

## Date/Region Calculations (Control Sheet)
```excel
AA4 (UMTS Congested) = COUNTIFS('Eric UMTS'!G:G, Z4, 'Eric UMTS'!A:A, Y4, 'Eric UMTS'!E:E, 1) +
                       COUNTIFS('HW UMTS'!M:M, Z4, 'HW UMTS'!A:A, Y4, 'HW UMTS'!K:K, 1)
AB4 (UMTS Both Congested) = COUNTIFS('Eric UMTS'!G:G, Z4, 'Eric UMTS'!A:A, Y4, 'Eric UMTS'!F:F, 1) +
                           COUNTIFS('HW UMTS'!M:M, Z4, 'HW UMTS'!A:A, Y4, 'HW UMTS'!L:L, 1)
AC4 (UMTS Total) = COUNTIFS('Eric UMTS'!G:G, Z4, 'Eric UMTS'!A:A, Y4) +
                  COUNTIFS('HW UMTS'!M:M, Z4, 'HW UMTS'!A:A, Y4)
AG4 (LTE Congested) = COUNTIFS('Eric LTE'!N:N, AF4, 'Eric LTE'!A:A, AE4, 'Eric LTE'!M:M, 1) +
                     COUNTIFS('HW LTE'!K:K, AF4, 'HW LTE'!A:A, AE4, 'HW LTE'!J:J, 1)
AH4 (LTE Total) = COUNTIFS('Eric LTE'!N:N, AF4, 'Eric LTE'!A:A, AE4) +
                 COUNTIFS('HW LTE'!K:K, AF4, 'HW LTE'!A:A, AE4)
```
```

These files provide complete documentation for your project, making it easy for others to understand and use your solution. The README.md is user-friendly while marco.md contains the technical formula reference.
