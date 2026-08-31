# Performance workflow functions

## Cash-flow-aware returns

### Modified Dietz

`ModifiedDietzReturn(BeginningValue, EndingValue, CashFlows, CashFlowDates, StartDate, EndDate, [FlowTiming])`

Positive cash flows are contributions to the portfolio. `FlowTiming` accepts `ACTUAL` (default), `BEGIN`, or `END`. With `ACTUAL`, each cash flow is weighted by the fraction of the measurement period remaining after the flow date.

### Time-weighted return

`TimeWeightedReturn(MarketValues, Dates, ExternalCashFlows, [FlowTiming])`

The function geometrically links valuation-to-valuation subperiod returns. `ExternalCashFlows` may contain either one flow per interval (`n-1`) or one value per valuation (`n`, with the first value ignored). Blank flow cells are treated as zero.

`END` means the external flow occurs immediately before the ending valuation; `BEGIN` means it occurs immediately after the beginning valuation. If flows occur intraday or between sparse valuations, use a more precise valuation policy rather than assuming this function can infer timing.

### Money-weighted return

`MoneyWeightedReturn(CashFlows, CashFlowDates, EndingValue, EndingDate, [Guess])`

Uses Excel's XIRR solver after appending the terminal market value. Investor-perspective cash-flow convention is used: contributions are negative, withdrawals/distributions are positive, and terminal market value is positive.

## Contribution and attribution

`ReturnContribution(Weights, Returns)` returns one contribution per security/category.

`ActiveContribution(PortfolioWeights, PortfolioReturns, BenchmarkWeights, BenchmarkReturns)` returns `wp*rp - wb*rb` by category.

`BrinsonFachlerAttribution(...)` spills four columns per category:

1. Allocation
2. Selection
3. Interaction
4. Total active contribution

Brinson-Fachler allocation uses `(wp-wb)*(rb_category-rb_total)`. Portfolio and benchmark weights must each sum to 1 within tolerance.

`BrinsonHoodBeebowerAttribution(...)` uses `(wp-wb)*rb_category` for allocation and otherwise uses the same selection and interaction conventions.

### Multi-period linking

`CarinoLinkAttribution(AttributionEffects, PortfolioReturns, BenchmarkReturns)` applies logarithmic Carino smoothing. The attribution matrix must have one row per return period. The function returns one cumulative linked effect per input effect column.

## Composite performance

`CompositeReturn(PortfolioReturns, BeginningMarketValues)` calculates a beginning-asset-weighted one-period composite return.

`CompositeDispersion(PortfolioReturns)` calculates equal-weighted sample standard deviation across constituent portfolio returns.

`AssetWeightedCompositeDispersion(PortfolioReturns, BeginningMarketValues)` calculates beginning-asset-weighted dispersion around the asset-weighted composite return.

## Currency decomposition

`CurrencyReturnDecomposition(LocalReturns, CurrencyReturns)` decomposes unhedged base-currency return into local, currency, interaction, and total return using `(1+local)*(1+currency)-1`.

`CurrencyContribution(Weights, LocalReturns, CurrencyReturns)` applies the same decomposition to weighted contributions.

These are transparent return decomposition tools, not a full Karnosky-Singer active currency attribution model.

## Fixed-income building blocks

`DurationContribution`, `SpreadContribution`, and `ConvexityContribution` provide transparent duration-based approximations. `FixedIncomeReturnDecomposition` spills Carry, Rate, Spread, and Explained Total columns.

These functions intentionally do not claim to be a complete fixed-income attribution engine. A production curve attribution model requires explicit decisions about key-rate buckets, roll-down/carry, spread curves, optionality, pricing sources, and residual treatment.
