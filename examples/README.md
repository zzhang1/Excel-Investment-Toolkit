# Examples

These CSV files are intentionally plain text so the examples remain reviewable in Git. Open them in Excel and use the formulas below after importing all modules in `src/`.

## Return-series example

Open `PerformanceExample.csv`.

Assuming dates are in `A2:A7`, portfolio returns in `B2:B7`, and benchmark returns in `C2:C7`:

```excel
=AnnualizedReturn(B2:B7,12)
=TrackingError(B2:B7,C2:C7,12)
=InformationRatio(B2:B7,C2:C7,12)
=RegressionStatistics(B2:B7,C2:C7,0.03,12)
=MaxDrawdown(B2:B7)
=BattingAverage(B2:B7,C2:C7)
=AggregateReturns(B2:B7,A2:A7,"Q")
=YTDReturn(B2:B7,A2:A7)
```

## Brinson attribution example

Open `AttributionExample.csv`.

Assuming sector, portfolio weight, portfolio return, benchmark weight, and benchmark return are in columns A:E:

```excel
=BrinsonFachlerAttribution(B2:B6,C2:C6,D2:D6,E2:E6)
```

The spilled result contains Allocation, Selection, Interaction, and Total Active Contribution. Sum the fourth column and reconcile it to:

```excel
=SUMPRODUCT(B2:B6,C2:C6)-SUMPRODUCT(D2:D6,E2:E6)
```

## Cash-flow return example

Open `CashFlowExample.csv`. It contains a valuation series and interval external cash flows. With market values in B2:B5, dates in A2:A5, and interval-ending cash flows in C3:C5:

```excel
=TimeWeightedReturn(B2:B5,A2:A5,C3:C5,"END")
```

For Modified Dietz, use the beginning and ending values plus actual dated external cash flows:

```excel
=ModifiedDietzReturn(1000000,1075000,F2:F3,E2:E3,DATE(2026,1,1),DATE(2026,3,31),"ACTUAL")
```
