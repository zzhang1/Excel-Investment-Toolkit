# Changelog

## 2.0.0 - Unreleased

### Added
- Modular `src/` architecture with shared private core helpers.
- Calendar-aware weekly/monthly/quarterly/yearly return aggregation plus MTD/QTD/YTD and arbitrary-period returns.
- Modified Dietz, time-weighted, and money-weighted return functions.
- Return and active contribution functions.
- Brinson-Fachler and Brinson-Hood-Beebower attribution.
- Carino multi-period attribution linking.
- Factor contribution building block.
- Composite return and dispersion functions.
- Regression statistics, appraisal ratio, Calmar, Omega, batting average, positive-period percentage, correlation, and R-squared.
- Drawdown series/duration, historical VaR/CVaR, skewness, excess kurtosis, and tail ratio.
- Currency local/FX/interaction decomposition.
- Duration, spread, convexity, carry, and first-order fixed-income decomposition helpers.
- Reproducible CSV examples and VBA reconciliation tests.
- GitHub Actions static checks for module structure and duplicate public UDFs.

### Changed
- v2 `InformationRatio` uses annualized arithmetic mean active return divided by annualized tracking error.
- v2 `AlphaCoefficient` / `JensenAlpha` use a periodic excess-return regression intercept and consistent annualization.
- v2 vector inputs reject unintended rectangular 2-D ranges instead of flattening them.
- Blank observations are ignored by default; paired series use pairwise-complete observations.

### Compatibility
- The root-level `FinancialMetrics.bas` remains unchanged as the legacy v1 single-file module.
- Do not import the legacy v1 module together with the modular v2 source because public function names overlap.
