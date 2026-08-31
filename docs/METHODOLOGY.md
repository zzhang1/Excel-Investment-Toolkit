# Methodology and conventions

The modular `src/` implementation is the v2 API. The root-level `FinancialMetrics.bas` is retained as the legacy single-file v1 module so existing users are not forced to migrate immediately. Do not import both the legacy module and the modular v2 files into the same VBA project because several public function names intentionally overlap.

## Input handling

- Functions expecting a vector accept a single row or a single column, not a rectangular 2-D range.
- Blank cells are ignored by default.
- Paired portfolio/benchmark functions use pairwise-complete observations: if either member of an aligned pair is blank, the pair is skipped.
- Non-numeric nonblank cells and Excel errors return an Excel error rather than being silently flattened or coerced.
- Return inputs are decimal returns (`0.01` = 1%). Annual rates are decimals.

## Return conventions

- `AnnualizedReturn` uses geometric compounding.
- `ActiveReturn` and `InformationRatio` use arithmetic mean periodic active return annualized by `PeriodsPerYear`; tracking error uses sample standard deviation annualized by the square-root-of-time convention.
- `AlphaCoefficient` / `JensenAlpha` estimate the intercept of the periodic excess-return regression and arithmetically annualize the intercept. This avoids mixing geometric annualized returns with a beta estimated from periodic arithmetic returns.
- `AggregateReturns` groups observations by actual calendar week, month, quarter, or year and geometrically compounds within each group. `MonthlyReturns(Returns, ObservationsPerMonth)` remains only as a compatibility function for fixed-width blocks.

## Risk conventions

- Volatility and tracking error use sample standard deviation.
- Historical VaR and CVaR are returned as positive loss magnitudes. For example, a 5th-percentile return of `-0.04` produces a 95% historical VaR of `0.04`.
- Drawdowns are returned as negative returns from the prior wealth peak.
- Drawdown duration functions measure duration in observations rather than calendar days.

## Compatibility

The public names from v1 are preserved in the v2 modules where practical. Two functions have intentionally revised methodology in v2:

1. `InformationRatio` now uses annualized arithmetic active return divided by annualized tracking error.
2. `AlphaCoefficient` / `JensenAlpha` now use a periodic excess-return regression intercept.

These changes are methodological corrections and can produce results that differ from the legacy root module.
