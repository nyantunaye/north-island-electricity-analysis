# North Island Electricity Demand Analysis (MATH803 Project Part B)

A SAS-based time series analysis of hourly electricity demand data for New Zealand's North Island, covering March–May 2020.

**Author:** Nyan Tun Aye (23198368)
**Course:** MATH803 — AUT

---

## Overview

This project analyses 2,208 hourly electricity demand records (92 days) from a North Island node, alongside temperature and wind speed data. Four forecasting approaches are built, evaluated, and compared to identify the best model for predicting the next 7 days of electricity demand.

---

## Dataset

| File | Description |
|------|-------------|
| `data/NorthIslandHourlyElec2-1.xlsx` | Hourly electricity demand (MW), temperature (°C), and wind speed (knots) — March 1 to May 31, 2020 |

**Variables:**

| Variable | Description |
|----------|-------------|
| `MW` | Electricity consumption in megawatts — response variable |
| `Temperature` | Auckland air temperature (°C) |
| `WindSpeed` | Wind speed (knots) |
| `Date` | Hourly datetime timestamp |

---

## Repository Structure

```
north-island-electricity-analysis/
├── data/
│   └── NorthIslandHourlyElec2-1.xlsx
├── ProjectSAS_NyanTunAye_23198368.sas
└── README.md
```

---

## Questions & Methods

### Question 1 — Data Import & Exploration
- Imported the Excel file using `PROC IMPORT` and converted the date column to SAS datetime format
- Plotted hourly electricity demand (March 1 – May 31, 2020) using `PROC GPLOT`

**Key observations:** Demand ranges from ~800 MW to ~2,100 MW with clear daily cycles (peaks in late afternoon/evening), weekly cycles (higher on weekdays), and a seasonal upward trend from autumn into winter.

---

### Question 2 — Time Series Regression & Exponential Smoothing

#### 2(a) Linear Regression + 7-Day Forecast
- Fitted `MW ~ date_value` using `PROC REG` and extrapolated to 168 hours (June 1–7, 2020)

| Metric | Value |
|--------|-------|
| R² | 0.0429 |
| Durbin-Watson | 0.116 (strong autocorrelation) |
| Slope | +2.6 MW/day |

Forecast: demand rises gradually from **1,416.25 MW** (June 1) to **1,429.75 MW** (June 7).

#### 2(b) Multiplicative Holt-Winters
- Applied `PROC ESM` with `METHOD=MULTWINTERS` and a 168-hour forecast lead
- Multiplicative method chosen over Additive because seasonal fluctuations grow proportionally with overall demand level

| Day | Avg MW | Min MW | Max MW |
|-----|--------|--------|--------|
| Jun 1 | 1,259.8 | 945.5 | 1,470.7 |
| Jun 4 | 1,467.7 | 1,013.4 | 1,765.2 |
| Jun 7 | 1,261.4 | 958.8 | 1,478.4 |

---

### Question 3 — ARIMA & ARIMAX

#### 3(a) SARIMA(1,1,1)(1,1,1)[24]
- Used `PROC ARIMA` with `VAR=MW(1,24)` — first-order non-seasonal and seasonal differencing at lag 24
- Captures both the upward trend and the strong 24-hour daily cycle
- Residuals approximate white noise, confirming good model fit

| Day | Avg MW | Min MW | Max MW |
|-----|--------|--------|--------|
| Jun 1 | 1,429.01 | 950.31 | 1,795.35 |
| Jun 7 | 1,227.48 | 913.51 | 1,464.69 |

#### 3(b) ARIMAX(1,1,1)(1,1,1)[24] with Exogenous Variables
- Extended SARIMA by adding `Temperature` and `WindSpeed` as exogenous inputs via `CROSSCORR` and `INPUT=`
- Temperature is statistically significant; WindSpeed has a weaker but positive contribution (p=0.4160)
- Forecast shows demand declining from **1,430.14 MW** (June 1) to **1,227.92 MW** (June 7)

---

### Question 4 — Out-of-Sample Forecast Evaluation

#### 4(a) Train/Test Split
- **Training set:** March 1 – May 23, 2020
- **Test set:** May 24 – May 31, 2020 (last 7 days)

#### 4(b) Accuracy Comparison

| Model | MAE | MSE | RMSE |
|-------|-----|-----|------|
| Linear Regression | 298.91 | 112,635.8 | 335.61 |
| SARIMA | 112.94 | 20,923.34 | 144.65 |
| Holt-Winters (Multiplicative) | 106.77 | 19,577.53 | 139.92 |
| **ARIMAX** | **21.53** | **1,146.58** | **33.86** |

**Winner: ARIMAX** — lowest MAE, MSE, and RMSE across all four models. Including temperature and wind speed as exogenous variables significantly improves forecast accuracy over time-series-only approaches.

---

## How to Run

1. Clone the repository and upload `NorthIslandHourlyElec2-1.xlsx` to your SAS environment
2. Update the `DATAFILE` paths in the `.sas` file to match your SAS server directory
3. Open `ProjectSAS_NyanTunAye_23198368.sas` in SAS Studio or SAS OnDemand for Academics
4. Run all sections sequentially — plots and output tables will be generated per question

---

## Requirements

- SAS (University Edition / SAS Studio / SAS OnDemand for Academics)
- Procedures used: `PROC IMPORT`, `PROC GPLOT`, `PROC REG`, `PROC ESM`, `PROC ARIMA`, `PROC SGPLOT`, `PROC MEANS`, `PROC SQL`
- No additional SAS packages required
