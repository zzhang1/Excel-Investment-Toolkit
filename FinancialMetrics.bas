Attribute VB_Name = "FinancialMetrics"
Option Explicit

' =============================================================================
' FinancialMetrics
' A library of Excel User Defined Functions (UDFs) for common investment and
' portfolio analytics using return series data.
'
' All functions expect periodic returns expressed as decimals (e.g. 0.01 for 1%)
' unless otherwise noted. Annual rates, such as the risk-free rate, should be
' supplied in decimal form as well. PeriodsPerYear identifies how many return
' observations fall in a calendar year (12 for monthly, 52 for weekly, 252 for
' daily, etc.).
' =============================================================================

' === Public API ==============================================================

Public Function AnnualizedReturn(Returns As Range, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    data = RangeToVector(Returns, hasError)
    If hasError Then
        AnnualizedReturn = CVErr(xlErrValue)
        Exit Function
    End If
    If PeriodsPerYear <= 0 Then
        AnnualizedReturn = CVErr(xlErrDiv0)
        Exit Function
    End If
    AnnualizedReturn = GeometricAnnualizedReturn(data, PeriodsPerYear)
End Function

Public Function AnnualizedVolatility(Returns As Range, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    data = RangeToVector(Returns, hasError)
    If hasError Then
        AnnualizedVolatility = CVErr(xlErrValue)
        Exit Function
    End If
    If PeriodsPerYear <= 0 Then
        AnnualizedVolatility = CVErr(xlErrDiv0)
        Exit Function
    End If
    AnnualizedVolatility = Sqr(SampleVariance(data)) * Sqr(PeriodsPerYear)
End Function

Public Function CumulativeReturn(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim cumulative As Double
    Dim i As Long

    data = RangeToVector(Returns, hasError)
    If hasError Then
        CumulativeReturn = CVErr(xlErrValue)
        Exit Function
    End If

    cumulative = 1
    For i = 1 To UBound(data)
        cumulative = cumulative * (1 + data(i))
    Next i

    CumulativeReturn = cumulative - 1
End Function

Public Function MonthlyReturns(Returns As Range, ObservationsPerMonth As Long) As Variant
    Dim data As Variant, hasError As Boolean
    Dim monthCount As Long, obsCount As Long, i As Long, j As Long
    Dim result() As Double, monthlyProduct As Double

    data = RangeToVector(Returns, hasError)
    If hasError Then
        MonthlyReturns = CVErr(xlErrValue)
        Exit Function
    End If

    If ObservationsPerMonth <= 0 Then
        MonthlyReturns = CVErr(xlErrDiv0)
        Exit Function
    End If

    obsCount = UBound(data)
    monthCount = obsCount \ ObservationsPerMonth
    If monthCount = 0 Then
        MonthlyReturns = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To monthCount, 1 To 1)
    For i = 1 To monthCount
        monthlyProduct = 1
        For j = 1 To ObservationsPerMonth
            monthlyProduct = monthlyProduct * (1 + data((i - 1) * ObservationsPerMonth + j))
        Next j
        result(i, 1) = monthlyProduct - 1
    Next i

    MonthlyReturns = result
End Function

Public Function RollingAnnualizedReturn(Returns As Range, WindowLength As Long, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Double
    Dim i As Long, j As Long
    Dim cumulative As Double
    Dim obsCount As Long

    data = RangeToVector(Returns, hasError)
    If hasError Then
        RollingAnnualizedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    obsCount = UBound(data)
    If WindowLength <= 0 Or WindowLength > obsCount Then
        RollingAnnualizedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        RollingAnnualizedReturn = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim result(1 To obsCount - WindowLength + 1, 1 To 1)

    For i = WindowLength To obsCount
        cumulative = 1
        For j = i - WindowLength + 1 To i
            cumulative = cumulative * (1 + data(j))
        Next j
        result(i - WindowLength + 1, 1) = cumulative ^ (PeriodsPerYear / WindowLength) - 1
    Next i

    RollingAnnualizedReturn = result
End Function

Public Function RollingVolatility(Returns As Range, WindowLength As Long, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Double
    Dim i As Long
    Dim obsCount As Long
    Dim windowData() As Double

    data = RangeToVector(Returns, hasError)
    If hasError Then
        RollingVolatility = CVErr(xlErrValue)
        Exit Function
    End If

    obsCount = UBound(data)
    If WindowLength <= 1 Or WindowLength > obsCount Then
        RollingVolatility = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        RollingVolatility = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim result(1 To obsCount - WindowLength + 1, 1 To 1)

    For i = WindowLength To obsCount
        windowData = SliceVector(data, i - WindowLength + 1, i)
        result(i - WindowLength + 1, 1) = Sqr(SampleVariance(windowData)) * Sqr(PeriodsPerYear)
    Next i

    RollingVolatility = result
End Function

Public Function SharpeRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim rfPerPeriod As Double
    Dim excessMean As Double
    Dim stdDev As Double

    data = RangeToVector(Returns, hasError)
    If hasError Then
        SharpeRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        SharpeRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    rfPerPeriod = PeriodicRate(RiskFreeRate, PeriodsPerYear)
    excessMean = MeanExcess(data, rfPerPeriod)
    stdDev = Sqr(SampleVariance(data))

    If stdDev = 0 Then
        SharpeRatio = CVErr(xlErrDiv0)
    Else
        SharpeRatio = (excessMean / stdDev) * Sqr(PeriodsPerYear)
    End If
End Function

Public Function RollingSharpeRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double, WindowLength As Long) As Variant
    Dim data As Variant, hasError As Boolean
    Dim rfPerPeriod As Double
    Dim result() As Variant
    Dim obsCount As Long
    Dim windowData() As Double
    Dim stdDev As Double
    Dim i As Long

    data = RangeToVector(Returns, hasError)
    If hasError Then
        RollingSharpeRatio = CVErr(xlErrValue)
        Exit Function
    End If

    obsCount = UBound(data)
    If WindowLength <= 1 Or WindowLength > obsCount Then
        RollingSharpeRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        RollingSharpeRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    rfPerPeriod = PeriodicRate(RiskFreeRate, PeriodsPerYear)
    ReDim result(1 To obsCount - WindowLength + 1, 1 To 1)

    For i = WindowLength To obsCount
        windowData = SliceVector(data, i - WindowLength + 1, i)
        stdDev = Sqr(SampleVariance(windowData))
        If stdDev = 0 Then
            result(i - WindowLength + 1, 1) = CVErr(xlErrDiv0)
        Else
            result(i - WindowLength + 1, 1) = (MeanExcess(windowData, rfPerPeriod) / stdDev) * Sqr(PeriodsPerYear)
        End If
    Next i

    RollingSharpeRatio = result
End Function

Public Function SortinoRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double, Optional MinimumAcceptableReturn As Variant) As Variant
    Dim data As Variant, hasError As Boolean
    Dim marPerPeriod As Double
    Dim annualizedReturnValue As Double
    Dim downsideDev As Double
    Dim marAnnual As Double

    data = RangeToVector(Returns, hasError)
    If hasError Then
        SortinoRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        SortinoRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    If IsMissing(MinimumAcceptableReturn) Then
        marPerPeriod = PeriodicRate(RiskFreeRate, PeriodsPerYear)
    Else
        marPerPeriod = PeriodicRate(CDbl(MinimumAcceptableReturn), PeriodsPerYear)
    End If

    annualizedReturnValue = GeometricAnnualizedReturn(data, PeriodsPerYear)
    marAnnual = (1 + marPerPeriod) ^ PeriodsPerYear - 1
    downsideDev = DownsideDeviation(data, marPerPeriod) * Sqr(PeriodsPerYear)

    If downsideDev = 0 Then
        SortinoRatio = CVErr(xlErrDiv0)
    Else
        SortinoRatio = (annualizedReturnValue - marAnnual) / downsideDev
    End If
End Function

Public Function TrackingError(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim activeReturns() As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        TrackingError = CVErr(xlErrDiv0)
        Exit Function
    End If

    activeReturns = ArrayDifference(portfolio, benchmark)
    TrackingError = Sqr(SampleVariance(activeReturns)) * Sqr(PeriodsPerYear)
End Function

Public Function InformationRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim te As Double
    Dim activeReturnAnnual As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        InformationRatio = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        InformationRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        InformationRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        InformationRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    te = Sqr(SampleVariance(ArrayDifference(portfolio, benchmark))) * Sqr(PeriodsPerYear)

    If te = 0 Then
        InformationRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    activeReturnAnnual = GeometricAnnualizedReturn(portfolio, PeriodsPerYear) - GeometricAnnualizedReturn(benchmark, PeriodsPerYear)
    InformationRatio = activeReturnAnnual / te
End Function

Public Function UpCaptureRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim portFiltered() As Double, benchFiltered() As Double
    Dim count As Long, i As Long
    Dim portfolioAnnual As Double, benchmarkAnnual As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        UpCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        UpCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        UpCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        UpCaptureRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim portFiltered(1 To UBound(portfolio))
    ReDim benchFiltered(1 To UBound(benchmark))
    count = 0

    For i = 1 To UBound(benchmark)
        If benchmark(i) > 0 Then
            count = count + 1
            portFiltered(count) = portfolio(i)
            benchFiltered(count) = benchmark(i)
        End If
    Next i

    If count = 0 Then
        UpCaptureRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim Preserve portFiltered(1 To count)
    ReDim Preserve benchFiltered(1 To count)

    portfolioAnnual = GeometricAnnualizedReturn(portFiltered, PeriodsPerYear)
    benchmarkAnnual = GeometricAnnualizedReturn(benchFiltered, PeriodsPerYear)

    If benchmarkAnnual = 0 Then
        UpCaptureRatio = CVErr(xlErrDiv0)
    Else
        UpCaptureRatio = portfolioAnnual / benchmarkAnnual
    End If
End Function

Public Function DownCaptureRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim portFiltered() As Double, benchFiltered() As Double
    Dim count As Long, i As Long
    Dim portfolioAnnual As Double, benchmarkAnnual As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        DownCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        DownCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        DownCaptureRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        DownCaptureRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim portFiltered(1 To UBound(portfolio))
    ReDim benchFiltered(1 To UBound(benchmark))
    count = 0

    For i = 1 To UBound(benchmark)
        If benchmark(i) < 0 Then
            count = count + 1
            portFiltered(count) = portfolio(i)
            benchFiltered(count) = benchmark(i)
        End If
    Next i

    If count = 0 Then
        DownCaptureRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim Preserve portFiltered(1 To count)
    ReDim Preserve benchFiltered(1 To count)

    portfolioAnnual = GeometricAnnualizedReturn(portFiltered, PeriodsPerYear)
    benchmarkAnnual = GeometricAnnualizedReturn(benchFiltered, PeriodsPerYear)

    If benchmarkAnnual = 0 Then
        DownCaptureRatio = CVErr(xlErrDiv0)
    Else
        DownCaptureRatio = portfolioAnnual / benchmarkAnnual
    End If
End Function

Public Function BetaCoefficient(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim covar As Double
    Dim benchVar As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        BetaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        BetaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        BetaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    covar = SampleCovariance(portfolio, benchmark)
    benchVar = SampleVariance(benchmark)

    If benchVar = 0 Then
        BetaCoefficient = CVErr(xlErrDiv0)
    Else
        BetaCoefficient = covar / benchVar
    End If
End Function

Public Function RollingBeta(PortfolioReturns As Range, BenchmarkReturns As Range, WindowLength As Long) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim obsCount As Long
    Dim result() As Variant
    Dim i As Long
    Dim windowPortfolio() As Double
    Dim windowBenchmark() As Double
    Dim benchVar As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        RollingBeta = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        RollingBeta = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        RollingBeta = CVErr(xlErrValue)
        Exit Function
    End If

    obsCount = UBound(portfolio)
    If WindowLength <= 1 Or WindowLength > obsCount Then
        RollingBeta = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To obsCount - WindowLength + 1, 1 To 1)

    For i = WindowLength To obsCount
        windowPortfolio = SliceVector(portfolio, i - WindowLength + 1, i)
        windowBenchmark = SliceVector(benchmark, i - WindowLength + 1, i)
        benchVar = SampleVariance(windowBenchmark)
        If benchVar = 0 Then
            result(i - WindowLength + 1, 1) = CVErr(xlErrDiv0)
        Else
            result(i - WindowLength + 1, 1) = SampleCovariance(windowPortfolio, windowBenchmark) / benchVar
        End If
    Next i

    RollingBeta = result
End Function

Public Function TreynorRatio(PortfolioReturns As Range, BenchmarkReturns As Range, RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim betaVal As Double
    Dim benchVar As Double
    Dim excessReturn As Double
    Dim portfolioAnnual As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        TreynorRatio = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        TreynorRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        TreynorRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        TreynorRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    benchVar = SampleVariance(benchmark)
    If benchVar = 0 Then
        TreynorRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    betaVal = SampleCovariance(portfolio, benchmark) / benchVar

    If betaVal = 0 Then
        TreynorRatio = CVErr(xlErrDiv0)
        Exit Function
    End If

    portfolioAnnual = GeometricAnnualizedReturn(portfolio, PeriodsPerYear)
    excessReturn = portfolioAnnual - RiskFreeRate

    TreynorRatio = excessReturn / betaVal
End Function

Public Function AlphaCoefficient(PortfolioReturns As Range, BenchmarkReturns As Range, RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim portfolio As Variant, benchmark As Variant
    Dim hasError As Boolean
    Dim betaVal As Variant
    Dim portfolioAnnual As Double, benchmarkAnnual As Double
    Dim rfAnnual As Double

    portfolio = RangeToVector(PortfolioReturns, hasError)
    If hasError Then
        AlphaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    benchmark = RangeToVector(BenchmarkReturns, hasError)
    If hasError Then
        AlphaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    If Not ValidateMatchingLengths(portfolio, benchmark) Then
        AlphaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        AlphaCoefficient = CVErr(xlErrDiv0)
        Exit Function
    End If

    betaVal = BetaCoefficient(PortfolioReturns, BenchmarkReturns)
    If IsError(betaVal) Then
        AlphaCoefficient = betaVal
        Exit Function
    End If

    portfolioAnnual = GeometricAnnualizedReturn(portfolio, PeriodsPerYear)
    benchmarkAnnual = GeometricAnnualizedReturn(benchmark, PeriodsPerYear)
    rfAnnual = RiskFreeRate

    AlphaCoefficient = portfolioAnnual - (rfAnnual + betaVal * (benchmarkAnnual - rfAnnual))
End Function

Public Function JensenAlpha(PortfolioReturns As Range, BenchmarkReturns As Range, RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    JensenAlpha = AlphaCoefficient(PortfolioReturns, BenchmarkReturns, RiskFreeRate, PeriodsPerYear)
End Function

Public Function DownsideDeviationSeries(Returns As Range, MinimumAcceptableReturn As Double, PeriodsPerYear As Double, WindowLength As Long) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Double
    Dim i As Long
    Dim obsCount As Long
    Dim windowData() As Double
    Dim marPerPeriod As Double

    data = RangeToVector(Returns, hasError)
    If hasError Then
        DownsideDeviationSeries = CVErr(xlErrValue)
        Exit Function
    End If

    If PeriodsPerYear <= 0 Then
        DownsideDeviationSeries = CVErr(xlErrDiv0)
        Exit Function
    End If

    obsCount = UBound(data)
    If WindowLength <= 1 Or WindowLength > obsCount Then
        DownsideDeviationSeries = CVErr(xlErrValue)
        Exit Function
    End If

    marPerPeriod = PeriodicRate(MinimumAcceptableReturn, PeriodsPerYear)

    ReDim result(1 To obsCount - WindowLength + 1, 1 To 1)

    For i = WindowLength To obsCount
        windowData = SliceVector(data, i - WindowLength + 1, i)
        result(i - WindowLength + 1, 1) = DownsideDeviation(windowData, marPerPeriod) * Sqr(PeriodsPerYear)
    Next i

    DownsideDeviationSeries = result
End Function

Public Function MaxDrawdown(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim cumulative As Double
    Dim peak As Double
    Dim drawdown As Double
    Dim maxDraw As Double
    Dim i As Long

    data = RangeToVector(Returns, hasError)
    If hasError Then
        MaxDrawdown = CVErr(xlErrValue)
        Exit Function
    End If

    cumulative = 1
    peak = 1
    maxDraw = 0

    For i = 1 To UBound(data)
        cumulative = cumulative * (1 + data(i))
        If cumulative > peak Then
            peak = cumulative
        End If
        If peak <> 0 Then
            drawdown = cumulative / peak - 1
            If drawdown < maxDraw Then
                maxDraw = drawdown
            End If
        End If
    Next i

    MaxDrawdown = maxDraw
End Function

' === Helper routines =========================================================

Private Function RangeToVector(inputRange As Range, ByRef hasError As Boolean) As Variant
    Dim values As Variant
    Dim data() As Double
    Dim r As Long, c As Long
    Dim count As Long
    Dim cellValue As Variant

    If inputRange.Count = 1 Then
        cellValue = inputRange.Value
        If IsEmpty(cellValue) Or Not IsNumeric(cellValue) Then
            hasError = True
            Exit Function
        End If
        ReDim data(1 To 1)
        data(1) = CDbl(cellValue)
        RangeToVector = data
        Exit Function
    End If

    values = inputRange.Value
    ReDim data(1 To inputRange.Count)

    count = 0
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            cellValue = values(r, c)
            If IsEmpty(cellValue) Or Not IsNumeric(cellValue) Then
                hasError = True
                Exit Function
            End If
            count = count + 1
            data(count) = CDbl(cellValue)
        Next c
    Next r

    ReDim Preserve data(1 To count)
    RangeToVector = data
End Function

Private Function SliceVector(data As Variant, startIndex As Long, endIndex As Long) As Double()
    Dim result() As Double
    Dim i As Long
    Dim pos As Long

    ReDim result(1 To endIndex - startIndex + 1)
    pos = 1
    For i = startIndex To endIndex
        result(pos) = data(i)
        pos = pos + 1
    Next i

    SliceVector = result
End Function

Private Function ValidateMatchingLengths(arr1 As Variant, arr2 As Variant) As Boolean
    On Error GoTo ErrHandler
    ValidateMatchingLengths = (UBound(arr1) = UBound(arr2))
    Exit Function
ErrHandler:
    ValidateMatchingLengths = False
End Function

Private Function GeometricAnnualizedReturn(data As Variant, PeriodsPerYear As Double) As Double
    Dim cumulative As Double
    Dim i As Long
    Dim n As Long

    cumulative = 1
    n = UBound(data)

    For i = 1 To n
        cumulative = cumulative * (1 + data(i))
    Next i

    GeometricAnnualizedReturn = cumulative ^ (PeriodsPerYear / n) - 1
End Function

Private Function PeriodicRate(annualRate As Double, PeriodsPerYear As Double) As Double
    PeriodicRate = (1 + annualRate) ^ (1 / PeriodsPerYear) - 1
End Function

Private Function Mean(data As Variant) As Double
    Dim total As Double
    Dim i As Long

    total = 0
    For i = 1 To UBound(data)
        total = total + data(i)
    Next i

    Mean = total / UBound(data)
End Function

Private Function MeanExcess(data As Variant, referenceRate As Double) As Double
    Dim total As Double
    Dim i As Long

    total = 0
    For i = 1 To UBound(data)
        total = total + (data(i) - referenceRate)
    Next i

    MeanExcess = total / UBound(data)
End Function

Private Function SampleVariance(data As Variant) As Double
    Dim meanValue As Double
    Dim sumSq As Double
    Dim i As Long
    Dim n As Long

    n = UBound(data)
    If n <= 1 Then
        SampleVariance = 0
        Exit Function
    End If

    meanValue = Mean(data)
    sumSq = 0
    For i = 1 To n
        sumSq = sumSq + (data(i) - meanValue) ^ 2
    Next i

    SampleVariance = sumSq / (n - 1)
End Function

Private Function SampleCovariance(arr1 As Variant, arr2 As Variant) As Double
    Dim mean1 As Double, mean2 As Double
    Dim sumCov As Double
    Dim i As Long
    Dim n As Long

    n = UBound(arr1)
    If n <= 1 Then
        SampleCovariance = 0
        Exit Function
    End If

    mean1 = Mean(arr1)
    mean2 = Mean(arr2)

    sumCov = 0
    For i = 1 To n
        sumCov = sumCov + (arr1(i) - mean1) * (arr2(i) - mean2)
    Next i

    SampleCovariance = sumCov / (n - 1)
End Function

Private Function ArrayDifference(arr1 As Variant, arr2 As Variant) As Double()
    Dim result() As Double
    Dim i As Long
    Dim n As Long

    n = UBound(arr1)
    ReDim result(1 To n)
    For i = 1 To n
        result(i) = arr1(i) - arr2(i)
    Next i

    ArrayDifference = result
End Function

Private Function DownsideDeviation(data As Variant, marPerPeriod As Double) As Double
    Dim sumSq As Double
    Dim i As Long
    Dim diff As Double
    Dim n As Long

    n = UBound(data)
    If n = 0 Then
        DownsideDeviation = 0
        Exit Function
    End If

    sumSq = 0
    For i = 1 To n
        diff = data(i) - marPerPeriod
        If diff < 0 Then
            sumSq = sumSq + diff ^ 2
        End If
    Next i

    DownsideDeviation = Sqr(sumSq / n)
End Function

