Attribute VB_Name = "PerformanceAppraisal"
Option Explicit

Public Function SharpeRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim rfPeriodic As Variant, sd As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        SharpeRatio = CVErr(xlErrValue)
        Exit Function
    End If
    rfPeriodic = EITPeriodicRate(RiskFreeRate, PeriodsPerYear)
    If IsError(rfPeriodic) Then
        SharpeRatio = rfPeriodic
        Exit Function
    End If
    sd = Sqr(EITSampleVariance(data))
    If sd = 0 Then
        SharpeRatio = CVErr(xlErrDiv0)
    Else
        SharpeRatio = ((EITMean(data) - CDbl(rfPeriodic)) / sd) * Sqr(PeriodsPerYear)
    End If
End Function

Public Function RollingSharpeRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double, WindowLength As Long) As Variant
    Dim data As Variant, windowData As Variant
    Dim hasError As Boolean, result() As Variant
    Dim rfPeriodic As Variant, sd As Double
    Dim i As Long, n As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        RollingSharpeRatio = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(data)
    If WindowLength <= 1 Or WindowLength > n Then
        RollingSharpeRatio = CVErr(xlErrNum)
        Exit Function
    End If
    rfPeriodic = EITPeriodicRate(RiskFreeRate, PeriodsPerYear)
    If IsError(rfPeriodic) Then
        RollingSharpeRatio = rfPeriodic
        Exit Function
    End If

    ReDim result(1 To n - WindowLength + 1, 1 To 1)
    For i = WindowLength To n
        windowData = EITSliceVector(data, i - WindowLength + 1, i)
        sd = Sqr(EITSampleVariance(windowData))
        If sd = 0 Then
            result(i - WindowLength + 1, 1) = CVErr(xlErrDiv0)
        Else
            result(i - WindowLength + 1, 1) = ((EITMean(windowData) - CDbl(rfPeriodic)) / sd) * Sqr(PeriodsPerYear)
        End If
    Next i
    RollingSharpeRatio = result
End Function

Public Function SortinoRatio(Returns As Range, RiskFreeRate As Double, PeriodsPerYear As Double, _
                             Optional MinimumAcceptableReturn As Variant) As Variant
    Dim data As Variant, hasError As Boolean
    Dim marAnnual As Double, marPeriodic As Variant
    Dim annualReturn As Variant, downside As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        SortinoRatio = CVErr(xlErrValue)
        Exit Function
    End If

    If IsMissing(MinimumAcceptableReturn) Then
        marAnnual = RiskFreeRate
    Else
        If Not IsNumeric(MinimumAcceptableReturn) Then
            SortinoRatio = CVErr(xlErrValue)
            Exit Function
        End If
        marAnnual = CDbl(MinimumAcceptableReturn)
    End If

    marPeriodic = EITPeriodicRate(marAnnual, PeriodsPerYear)
    annualReturn = EITGeometricAnnualizedReturn(data, PeriodsPerYear)
    If IsError(marPeriodic) Or IsError(annualReturn) Then
        SortinoRatio = CVErr(xlErrNum)
        Exit Function
    End If
    downside = EITDownsideDeviation(data, CDbl(marPeriodic)) * Sqr(PeriodsPerYear)
    If downside = 0 Then
        SortinoRatio = CVErr(xlErrDiv0)
    Else
        SortinoRatio = (CDbl(annualReturn) - marAnnual) / downside
    End If
End Function

Public Function TrackingError(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim p As Variant, b As Variant, active As Variant
    Dim hasError As Boolean

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        TrackingError = CVErr(xlErrValue)
        Exit Function
    End If
    active = EITArrayDifference(p, b)
    TrackingError = Sqr(EITSampleVariance(active)) * Sqr(PeriodsPerYear)
End Function

' Arithmetic information ratio: annualized mean periodic active return / annualized tracking error.
Public Function InformationRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim p As Variant, b As Variant, active As Variant
    Dim hasError As Boolean, te As Double

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        InformationRatio = CVErr(xlErrValue)
        Exit Function
    End If
    active = EITArrayDifference(p, b)
    te = Sqr(EITSampleVariance(active)) * Sqr(PeriodsPerYear)
    If te = 0 Then
        InformationRatio = CVErr(xlErrDiv0)
    Else
        InformationRatio = (EITMean(active) * PeriodsPerYear) / te
    End If
End Function

Public Function ActiveReturnSeries(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim p As Variant, b As Variant, active As Variant
    Dim hasError As Boolean, result() As Double
    Dim i As Long

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Then
        ActiveReturnSeries = CVErr(xlErrValue)
        Exit Function
    End If
    active = EITArrayDifference(p, b)
    ReDim result(1 To UBound(active), 1 To 1)
    For i = 1 To UBound(active)
        result(i, 1) = active(i)
    Next i
    ActiveReturnSeries = result
End Function

Public Function ActiveReturn(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    Dim p As Variant, b As Variant, active As Variant
    Dim hasError As Boolean
    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        ActiveReturn = CVErr(xlErrValue)
        Exit Function
    End If
    active = EITArrayDifference(p, b)
    ActiveReturn = EITMean(active) * PeriodsPerYear
End Function

Public Function BetaCoefficient(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim p As Variant, b As Variant
    Dim hasError As Boolean, varianceB As Double

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Then
        BetaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If
    varianceB = EITSampleVariance(b)
    If varianceB = 0 Then
        BetaCoefficient = CVErr(xlErrDiv0)
    Else
        BetaCoefficient = EITSampleCovariance(p, b) / varianceB
    End If
End Function

Public Function RollingBeta(PortfolioReturns As Range, BenchmarkReturns As Range, WindowLength As Long) As Variant
    Dim p As Variant, b As Variant, wp As Variant, wb As Variant
    Dim hasError As Boolean, result() As Variant
    Dim n As Long, i As Long, varianceB As Double

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Then
        RollingBeta = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(p)
    If WindowLength <= 1 Or WindowLength > n Then
        RollingBeta = CVErr(xlErrNum)
        Exit Function
    End If

    ReDim result(1 To n - WindowLength + 1, 1 To 1)
    For i = WindowLength To n
        wp = EITSliceVector(p, i - WindowLength + 1, i)
        wb = EITSliceVector(b, i - WindowLength + 1, i)
        varianceB = EITSampleVariance(wb)
        If varianceB = 0 Then
            result(i - WindowLength + 1, 1) = CVErr(xlErrDiv0)
        Else
            result(i - WindowLength + 1, 1) = EITSampleCovariance(wp, wb) / varianceB
        End If
    Next i
    RollingBeta = result
End Function

Public Function CorrelationCoefficient(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim p As Variant, b As Variant
    Dim hasError As Boolean, sdP As Double, sdB As Double

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Then
        CorrelationCoefficient = CVErr(xlErrValue)
        Exit Function
    End If
    sdP = Sqr(EITSampleVariance(p)): sdB = Sqr(EITSampleVariance(b))
    If sdP = 0 Or sdB = 0 Then
        CorrelationCoefficient = CVErr(xlErrDiv0)
    Else
        CorrelationCoefficient = EITSampleCovariance(p, b) / (sdP * sdB)
    End If
End Function

Public Function RSquared(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim corr As Variant
    corr = CorrelationCoefficient(PortfolioReturns, BenchmarkReturns)
    If IsError(corr) Then RSquared = corr Else RSquared = CDbl(corr) ^ 2
End Function

' Jensen alpha estimated as the intercept of periodic excess-return regression and
' arithmetically annualized for consistency with the regression specification.
Public Function AlphaCoefficient(PortfolioReturns As Range, BenchmarkReturns As Range, _
                                 RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim p As Variant, b As Variant
    Dim pExcess() As Double, bExcess() As Double
    Dim hasError As Boolean, rfPeriodic As Variant
    Dim beta As Double, varianceB As Double, i As Long

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        AlphaCoefficient = CVErr(xlErrValue)
        Exit Function
    End If
    rfPeriodic = EITPeriodicRate(RiskFreeRate, PeriodsPerYear)
    If IsError(rfPeriodic) Then
        AlphaCoefficient = rfPeriodic
        Exit Function
    End If

    ReDim pExcess(1 To UBound(p)): ReDim bExcess(1 To UBound(b))
    For i = 1 To UBound(p)
        pExcess(i) = p(i) - CDbl(rfPeriodic)
        bExcess(i) = b(i) - CDbl(rfPeriodic)
    Next i
    varianceB = EITSampleVariance(bExcess)
    If varianceB = 0 Then
        AlphaCoefficient = CVErr(xlErrDiv0)
        Exit Function
    End If
    beta = EITSampleCovariance(pExcess, bExcess) / varianceB
    AlphaCoefficient = (EITMean(pExcess) - beta * EITMean(bExcess)) * PeriodsPerYear
End Function

Public Function JensenAlpha(PortfolioReturns As Range, BenchmarkReturns As Range, _
                            RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    JensenAlpha = AlphaCoefficient(PortfolioReturns, BenchmarkReturns, RiskFreeRate, PeriodsPerYear)
End Function

Public Function TreynorRatio(PortfolioReturns As Range, BenchmarkReturns As Range, _
                             RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim p As Variant, b As Variant
    Dim hasError As Boolean, varianceB As Double, beta As Double
    Dim annualReturn As Variant

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        TreynorRatio = CVErr(xlErrValue)
        Exit Function
    End If
    varianceB = EITSampleVariance(b)
    If varianceB = 0 Then
        TreynorRatio = CVErr(xlErrDiv0)
        Exit Function
    End If
    beta = EITSampleCovariance(p, b) / varianceB
    If beta = 0 Then
        TreynorRatio = CVErr(xlErrDiv0)
        Exit Function
    End If
    annualReturn = EITGeometricAnnualizedReturn(p, PeriodsPerYear)
    TreynorRatio = (CDbl(annualReturn) - RiskFreeRate) / beta
End Function

Public Function UpCaptureRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    UpCaptureRatio = CaptureRatioInternal(PortfolioReturns, BenchmarkReturns, PeriodsPerYear, True)
End Function

Public Function DownCaptureRatio(PortfolioReturns As Range, BenchmarkReturns As Range, PeriodsPerYear As Double) As Variant
    DownCaptureRatio = CaptureRatioInternal(PortfolioReturns, BenchmarkReturns, PeriodsPerYear, False)
End Function

Public Function AppraisalRatio(PortfolioReturns As Range, BenchmarkReturns As Range, _
                               RiskFreeRate As Double, PeriodsPerYear As Double) As Variant
    Dim stats As Variant
    stats = RegressionStatistics(PortfolioReturns, BenchmarkReturns, RiskFreeRate, PeriodsPerYear)
    If IsError(stats) Then
        AppraisalRatio = stats
    Else
        AppraisalRatio = stats(1, 6)
    End If
End Function

' Spills: Alpha, Beta, R-squared, Correlation, Residual Volatility, Appraisal Ratio.
Public Function RegressionStatistics(PortfolioReturns As Range, BenchmarkReturns As Range, _
                                     Optional RiskFreeRate As Double = 0, _
                                     Optional PeriodsPerYear As Double = 12) As Variant
    Dim p As Variant, b As Variant
    Dim px() As Double, bx() As Double, residuals() As Double
    Dim hasError As Boolean, rfPeriodic As Variant
    Dim varianceB As Double, beta As Double, alphaPeriodic As Double
    Dim corr As Double, sdP As Double, sdB As Double, residualVol As Double
    Dim alphaAnnual As Double, result(1 To 1, 1 To 6) As Variant
    Dim i As Long

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        RegressionStatistics = CVErr(xlErrValue)
        Exit Function
    End If
    rfPeriodic = EITPeriodicRate(RiskFreeRate, PeriodsPerYear)
    If IsError(rfPeriodic) Then
        RegressionStatistics = rfPeriodic
        Exit Function
    End If

    ReDim px(1 To UBound(p)): ReDim bx(1 To UBound(b)): ReDim residuals(1 To UBound(p))
    For i = 1 To UBound(p)
        px(i) = p(i) - CDbl(rfPeriodic)
        bx(i) = b(i) - CDbl(rfPeriodic)
    Next i
    varianceB = EITSampleVariance(bx)
    If varianceB = 0 Then
        RegressionStatistics = CVErr(xlErrDiv0)
        Exit Function
    End If

    beta = EITSampleCovariance(px, bx) / varianceB
    alphaPeriodic = EITMean(px) - beta * EITMean(bx)
    alphaAnnual = alphaPeriodic * PeriodsPerYear

    sdP = Sqr(EITSampleVariance(px)): sdB = Sqr(EITSampleVariance(bx))
    If sdP = 0 Or sdB = 0 Then corr = 0 Else corr = EITSampleCovariance(px, bx) / (sdP * sdB)

    For i = 1 To UBound(px)
        residuals(i) = px(i) - alphaPeriodic - beta * bx(i)
    Next i
    residualVol = Sqr(EITSampleVariance(residuals)) * Sqr(PeriodsPerYear)

    result(1, 1) = alphaAnnual
    result(1, 2) = beta
    result(1, 3) = corr ^ 2
    result(1, 4) = corr
    result(1, 5) = residualVol
    If residualVol = 0 Then result(1, 6) = CVErr(xlErrDiv0) Else result(1, 6) = alphaAnnual / residualVol
    RegressionStatistics = result
End Function

Public Function CalmarRatio(Returns As Range, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim annualReturn As Variant, maxDD As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        CalmarRatio = CVErr(xlErrValue)
        Exit Function
    End If
    annualReturn = EITGeometricAnnualizedReturn(data, PeriodsPerYear)
    maxDD = EITMaxDrawdownFromVector(data)
    If maxDD = 0 Then CalmarRatio = CVErr(xlErrDiv0) Else CalmarRatio = CDbl(annualReturn) / Abs(maxDD)
End Function

Public Function OmegaRatio(Returns As Range, MinimumAcceptableReturn As Double, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim marPeriodic As Variant, diff As Double
    Dim gains As Double, losses As Double, i As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        OmegaRatio = CVErr(xlErrValue)
        Exit Function
    End If
    marPeriodic = EITPeriodicRate(MinimumAcceptableReturn, PeriodsPerYear)
    If IsError(marPeriodic) Then
        OmegaRatio = marPeriodic
        Exit Function
    End If

    For i = 1 To UBound(data)
        diff = data(i) - CDbl(marPeriodic)
        If diff > 0 Then gains = gains + diff Else losses = losses - diff
    Next i
    If losses = 0 Then OmegaRatio = CVErr(xlErrDiv0) Else OmegaRatio = gains / losses
End Function

Public Function BattingAverage(PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim p As Variant, b As Variant
    Dim hasError As Boolean, wins As Long, i As Long
    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Then
        BattingAverage = CVErr(xlErrValue)
        Exit Function
    End If
    For i = 1 To UBound(p)
        If p(i) > b(i) Then wins = wins + 1
    Next i
    BattingAverage = wins / UBound(p)
End Function

Public Function PercentPositivePeriods(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim positives As Long, i As Long
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        PercentPositivePeriods = CVErr(xlErrValue)
        Exit Function
    End If
    For i = 1 To UBound(data)
        If data(i) > 0 Then positives = positives + 1
    Next i
    PercentPositivePeriods = positives / UBound(data)
End Function

Private Function CaptureRatioInternal(PortfolioReturns As Range, BenchmarkReturns As Range, _
                                      PeriodsPerYear As Double, UpMarket As Boolean) As Variant
    Dim p As Variant, b As Variant
    Dim fp() As Double, fb() As Double
    Dim hasError As Boolean, count As Long, i As Long
    Dim pAnnual As Variant, bAnnual As Variant

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError
    If hasError Or PeriodsPerYear <= 0 Then
        CaptureRatioInternal = CVErr(xlErrValue)
        Exit Function
    End If
    ReDim fp(1 To UBound(p)): ReDim fb(1 To UBound(b))

    For i = 1 To UBound(b)
        If (UpMarket And b(i) > 0) Or ((Not UpMarket) And b(i) < 0) Then
            count = count + 1
            fp(count) = p(i): fb(count) = b(i)
        End If
    Next i
    If count = 0 Then
        CaptureRatioInternal = CVErr(xlErrNA)
        Exit Function
    End If
    ReDim Preserve fp(1 To count): ReDim Preserve fb(1 To count)
    pAnnual = EITGeometricAnnualizedReturn(fp, PeriodsPerYear)
    bAnnual = EITGeometricAnnualizedReturn(fb, PeriodsPerYear)
    If CDbl(bAnnual) = 0 Then CaptureRatioInternal = CVErr(xlErrDiv0) Else CaptureRatioInternal = CDbl(pAnnual) / CDbl(bAnnual)
End Function
