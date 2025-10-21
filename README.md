# Excel Investment Toolkit

This repository contains a VBA module (`FinancialMetrics.bas`) that provides a suite of Excel user-defined functions (UDFs) for portfolio and risk analytics using return series data. Import the module into any macro-enabled workbook to access the functions from Excel formulas.

## Available functions

All functions expect periodic returns expressed as decimals (e.g., 0.01 for 1%). Unless otherwise noted, supply the annual risk-free rate in decimal form and specify `PeriodsPerYear` as the number of observations per calendar year (12 for monthly data, 252 for trading days, etc.).

| Function | Description |
| --- | --- |
| `AnnualizedReturn(Returns, PeriodsPerYear)` | Geometric annualized return for the provided series. |
| `AnnualizedVolatility(Returns, PeriodsPerYear)` | Annualized standard deviation using the sample statistic. |
| `CumulativeReturn(Returns)` | Total compounded return across the full sample. |
| `MonthlyReturns(Returns, ObservationsPerMonth)` | Converts higher-frequency returns into monthly returns via compounding. |
| `RollingAnnualizedReturn(Returns, WindowLength, PeriodsPerYear)` | Rolling geometric annualized return over the specified window. |
| `RollingVolatility(Returns, WindowLength, PeriodsPerYear)` | Rolling annualized volatility. |
| `SharpeRatio(Returns, RiskFreeRate, PeriodsPerYear)` | Annualized Sharpe ratio using excess returns versus the risk-free rate. |
| `RollingSharpeRatio(Returns, RiskFreeRate, PeriodsPerYear, WindowLength)` | Rolling Sharpe ratio series. |
| `SortinoRatio(Returns, RiskFreeRate, PeriodsPerYear [, MinimumAcceptableReturn])` | Annualized Sortino ratio using the risk-free rate or custom MAR. |
| `TrackingError(PortfolioReturns, BenchmarkReturns, PeriodsPerYear)` | Annualized tracking error between two return series. |
| `InformationRatio(PortfolioReturns, BenchmarkReturns, PeriodsPerYear)` | Information ratio using geometric annualized active return and tracking error. |
| `BetaCoefficient(PortfolioReturns, BenchmarkReturns)` | Sample beta of the portfolio versus its benchmark. |
| `RollingBeta(PortfolioReturns, BenchmarkReturns, WindowLength)` | Rolling beta series. |
| `AlphaCoefficient(PortfolioReturns, BenchmarkReturns, RiskFreeRate, PeriodsPerYear)` | Annualized Jensen alpha based on beta and benchmark returns. |
| `JensenAlpha(...)` | Alias for `AlphaCoefficient`. |
| `DownsideDeviationSeries(Returns, MinimumAcceptableReturn, PeriodsPerYear, WindowLength)` | Rolling annualized downside deviation relative to a MAR. |
| `MaxDrawdown(Returns)` | Peak-to-trough drawdown over the sample (reported as a negative value). |

## Usage

1. Open the VBA editor in Excel (`ALT` + `F11`).
2. Use **File → Import File...** and select `FinancialMetrics.bas` to add the module to your workbook or personal macro workbook.
3. Save the workbook as a macro-enabled file (`.xlsm`) if needed.
4. The functions become available like any native Excel function. Example formulas:
   - `=SharpeRatio(B2:B253, 0.02, 252)`
   - `=RollingAnnualizedReturn(B2:B253, 63, 252)` (returns a spill range with one value per completed window)
   - `=MaxDrawdown(B2:B253)`

Ensure the input ranges contain numeric, non-empty values and that paired ranges (e.g., portfolio and benchmark returns) are aligned and of equal length.
