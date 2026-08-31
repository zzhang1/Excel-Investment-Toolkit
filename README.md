# Excel Investment Toolkit

Excel Investment Toolkit is a VBA library for institutional-style investment performance, attribution, and risk analysis directly in worksheet formulas.

The modular v2 toolkit expands the original return-statistics module into a workflow-oriented library covering return measurement, cash flows, benchmark-relative appraisal, Brinson attribution, multi-period linking, composites, currency decomposition, and fixed-income contribution building blocks.

## Installation

### Recommended: modular v2

Import every `.bas` file in `src/` into the same VBA project:

1. Open Excel's VBA editor (`ALT` + `F11`).
2. Choose **File > Import File...** for each file in `src/`.
3. Save the workbook as `.xlsm`, or import the modules into your Personal Macro Workbook/add-in.
4. Use the functions like native Excel formulas.

`ToolkitCore.bas` contains shared helpers and uses `Option Private Module`; import it even though its helper functions are not intended as worksheet UDFs.

### Legacy v1

The root-level `FinancialMetrics.bas` remains available for users who want the original single-file library. Do **not** import it together with the v2 modules because several public function names overlap. The v2 methodology intentionally corrects the information-ratio and Jensen-alpha conventions; see `docs/METHODOLOGY.md`.

## Typical analyst workflow

```excel
=TimeWeightedReturn(B2:B260,A2:A260,C2:C260,"END")
=ModifiedDietzReturn(BeginMV,EndMV,Flows,FlowDates,StartDate,EndDate)
=BrinsonFachlerAttribution(PortWeights,PortReturns,BenchWeights,BenchReturns)
=CarinoLinkAttribution(PeriodEffects,PortReturns,BenchReturns)
=TrackingError(PortReturns,BenchReturns,12)
=InformationRatio(PortReturns,BenchReturns,12)
=RegressionStatistics(PortReturns,BenchReturns,0.03,12)
=MaxDrawdown(PortReturns)
```

Functions returning arrays spill automatically in current versions of Excel.

## Function reference

### Return measurement and calendar utilities

| Function | Purpose |
| --- | --- |
| `AnnualizedReturn(Returns, PeriodsPerYear)` | Geometric annualized return. |
| `CumulativeReturn(Returns)` | Compounded total return. |
| `RollingAnnualizedReturn(Returns, WindowLength, PeriodsPerYear)` | Rolling annualized return. |
| `AggregateReturns(Returns, Dates, Frequency)` | Calendar-aware W/M/Q/Y aggregation; spills period-end date and return. |
| `PeriodReturn(Returns, Dates, StartDate, EndDate)` | Compounded return over an inclusive date interval. |
| `MTDReturn`, `QTDReturn`, `YTDReturn` | Calendar-to-date return through an optional as-of date. |
| `MonthlyReturns(Returns, ObservationsPerMonth)` | Legacy fixed-width block aggregation; prefer `AggregateReturns`. |

### Cash-flow-aware performance

| Function | Purpose |
| --- | --- |
| `ModifiedDietzReturn(...)` | Daily-weighted external-cash-flow return using actual/begin/end timing. |
| `TimeWeightedReturn(...)` | Geometrically linked valuation-to-valuation TWR with external flows. |
| `MoneyWeightedReturn(...)` | XIRR-based money-weighted return with terminal market value appended. |

### Benchmark-relative appraisal

| Function | Purpose |
| --- | --- |
| `TrackingError(...)` | Annualized standard deviation of active returns. |
| `InformationRatio(...)` | Annualized arithmetic active return divided by tracking error. |
| `ActiveReturn(...)` | Annualized arithmetic active return. |
| `ActiveReturnSeries(...)` | Period-by-period active returns. |
| `BetaCoefficient(...)` / `RollingBeta(...)` | Market beta and rolling beta. |
| `CorrelationCoefficient(...)` | Portfolio/benchmark correlation. |
| `RSquared(...)` | R-squared against benchmark. |
| `AlphaCoefficient(...)` / `JensenAlpha(...)` | Annualized periodic excess-return regression intercept. |
| `RegressionStatistics(...)` | Spills Alpha, Beta, R², Correlation, Residual Volatility, Appraisal Ratio. |
| `AppraisalRatio(...)` | Regression alpha divided by residual volatility. |
| `UpCaptureRatio(...)` / `DownCaptureRatio(...)` | Up/down-market capture. |
| `SharpeRatio(...)` / `RollingSharpeRatio(...)` | Sharpe ratio and rolling series. |
| `SortinoRatio(...)` | Sortino ratio versus risk-free rate or custom MAR. |
| `TreynorRatio(...)` | Excess return per unit of beta. |
| `CalmarRatio(...)` | Annualized return divided by absolute maximum drawdown. |
| `OmegaRatio(...)` | Gains above MAR divided by shortfalls below MAR. |
| `BattingAverage(...)` | Fraction of periods outperforming benchmark. |
| `PercentPositivePeriods(...)` | Fraction of positive-return periods. |

### Contribution and attribution

| Function | Purpose |
| --- | --- |
| `ReturnContribution(Weights, Returns)` | Security/category contribution `w*r`. |
| `ActiveContribution(...)` | Portfolio contribution minus benchmark contribution by category. |
| `BrinsonFachlerAttribution(...)` | Allocation, Selection, Interaction, Total effects. |
| `BrinsonHoodBeebowerAttribution(...)` | BHB allocation variant with selection and interaction. |
| `CarinoLinkAttribution(...)` | Multi-period logarithmic smoothing/linking of attribution effects. |
| `FactorContribution(Exposures, FactorReturns)` | Contribution for externally estimated linear factors. |

### Composite performance

| Function | Purpose |
| --- | --- |
| `CompositeReturn(Returns, BeginningMarketValues)` | Beginning-asset-weighted composite return. |
| `CompositeDispersion(Returns)` | Equal-weighted sample dispersion across portfolios. |
| `AssetWeightedCompositeDispersion(...)` | Beginning-asset-weighted dispersion. |

### Risk and drawdowns

| Function | Purpose |
| --- | --- |
| `AnnualizedVolatility(...)` / `RollingVolatility(...)` | Annualized sample volatility. |
| `DownsideDeviationSeries(...)` | Rolling annualized downside deviation. |
| `MaxDrawdown(...)` | Maximum peak-to-trough drawdown. |
| `DrawdownSeries(...)` | Drawdown at every observation. |
| `DrawdownDuration(...)` | Current/last drawdown duration in observations. |
| `LongestDrawdown(...)` | Longest time below a prior wealth peak. |
| `AverageDrawdown(...)` | Mean negative drawdown observation. |
| `HistoricalVaR(...)` | Historical VaR as a positive loss magnitude. |
| `HistoricalCVaR(...)` | Historical expected shortfall as a positive loss magnitude. |
| `Skewness(...)` | Bias-adjusted sample skewness. |
| `ExcessKurtosis(...)` | Bias-adjusted sample excess kurtosis. |
| `TailRatio(...)` | Upper-tail percentile divided by absolute lower-tail percentile. |

### Currency and fixed-income building blocks

| Function | Purpose |
| --- | --- |
| `CurrencyReturnDecomposition(...)` | Local, currency, interaction, and total unhedged return. |
| `CurrencyContribution(...)` | Weighted local/FX/interaction contributions. |
| `DurationContribution(...)` | First-order effective-duration contribution. |
| `SpreadContribution(...)` | First-order spread-duration contribution. |
| `ConvexityContribution(...)` | Convexity contribution approximation. |
| `CarryContribution(...)` | Weighted carry contribution. |
| `FixedIncomeReturnDecomposition(...)` | Carry, rate, spread, and explained-total contributions. |

The currency and fixed-income functions are transparent building blocks, not full Karnosky-Singer or key-rate curve attribution engines. See `docs/PERFORMANCE_WORKFLOWS.md` for assumptions and limitations.

## Input conventions

- Returns, weights, rates, yield changes, and spread changes are decimals (`0.01` = 1%; 10 bp = `0.001`).
- Vector functions accept one row or one column. Rectangular 2-D ranges are rejected rather than flattened.
- Blank observations are generally ignored; paired functions use pairwise-complete observations.
- Paired ranges must be aligned and have the same worksheet length before blank filtering.
- `PeriodsPerYear` is typically 12 for monthly, 52 for weekly, and 252 for daily observations.
- Historical VaR/CVaR return positive loss magnitudes; drawdowns return negative percentages.

Detailed calculation conventions are in `docs/METHODOLOGY.md`.

## Tests

Import `tests/TestKnownIdentities.bas` with the v2 source modules and run:

```text
RunAllToolkitTests
```

The tests exercise known return calculations and reconciliation identities including Brinson effects, TWR, Modified Dietz, composites, currency decomposition, fixed-income duration contribution, and drawdowns.

## Examples

Review `examples/README.md` and the CSV files in `examples/`. They provide auditable datasets for return analysis, attribution, and cash-flow measurement without committing a binary workbook.

## Repository layout

```text
src/                         modular v2 VBA source
  ToolkitCore.bas
  ReturnMeasurement.bas
  CashFlowReturns.bas
  PerformanceAppraisal.bas
  PerformanceAttribution.bas
  RiskAnalytics.bas
  CompositePerformance.bas
  CurrencyAttribution.bas
  FixedIncomeAttribution.bas

tests/                       VBA regression/reconciliation tests
examples/                    reviewable CSV examples
docs/                        methodology and workflow notes
FinancialMetrics.bas         legacy v1 single-file distribution
```

## Scope

The toolkit is intended for transparent Excel analysis and reconciliation. It is not a performance-book-of-record system, pricing engine, or substitute for portfolio accounting controls. Complex fixed-income curve attribution, derivatives, transaction-level intraday performance, and full GIPS composite administration require additional data and methodology choices beyond generic worksheet functions.
