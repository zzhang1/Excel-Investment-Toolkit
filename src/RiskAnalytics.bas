Attribute VB_Name = "RiskAnalytics"
Option Explicit

Public Function AnnualizedVolatility(Returns As Range, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        AnnualizedVolatility = CVErr(xlErrValue)
        Exit Function
    End If
    AnnualizedVolatility = Sqr(EITSampleVariance(data)) * Sqr(PeriodsPerYear)
End Function

Public Function RollingVolatility(Returns As Range, WindowLength As Long, PeriodsPerYear As Double) As Variant
    Dim data As Variant, windowData As Variant
    Dim result() As Variant
    Dim hasError As Boolean, i As Long, n As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        RollingVolatility = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(data)
    If WindowLength <= 1 Or WindowLength > n Or PeriodsPerYear <= 0 Then
        RollingVolatility = CVErr(xlErrNum)
        Exit Function
    End If

    ReDim result(1 To n - WindowLength + 1, 1 To 1)
    For i = WindowLength To n
        windowData = EITSliceVector(data, i - WindowLength + 1, i)
        result(i - WindowLength + 1, 1) = Sqr(EITSampleVariance(windowData)) * Sqr(PeriodsPerYear)
    Next i
    RollingVolatility = result
End Function

Public Function DownsideDeviationSeries(Returns As Range, MinimumAcceptableReturn As Double, _
                                        PeriodsPerYear As Double, WindowLength As Long) As Variant
    Dim data As Variant, windowData As Variant
    Dim result() As Variant
    Dim hasError As Boolean, i As Long, n As Long
    Dim marPeriodic As Variant

    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        DownsideDeviationSeries = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(data)
    If WindowLength <= 1 Or WindowLength > n Then
        DownsideDeviationSeries = CVErr(xlErrNum)
        Exit Function
    End If

    marPeriodic = EITPeriodicRate(MinimumAcceptableReturn, PeriodsPerYear)
    If IsError(marPeriodic) Then
        DownsideDeviationSeries = marPeriodic
        Exit Function
    End If

    ReDim result(1 To n - WindowLength + 1, 1 To 1)
    For i = WindowLength To n
        windowData = EITSliceVector(data, i - WindowLength + 1, i)
        result(i - WindowLength + 1, 1) = EITDownsideDeviation(windowData, CDbl(marPeriodic)) * Sqr(PeriodsPerYear)
    Next i
    DownsideDeviationSeries = result
End Function

Public Function MaxDrawdown(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        MaxDrawdown = CVErr(xlErrValue)
    Else
        MaxDrawdown = EITMaxDrawdownFromVector(data)
    End If
End Function

Public Function DrawdownSeries(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Double
    Dim wealth As Double, peak As Double
    Dim i As Long, n As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        DrawdownSeries = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(data)
    ReDim result(1 To n, 1 To 1)
    wealth = 1#: peak = 1#
    For i = 1 To n
        wealth = wealth * (1# + CDbl(data(i)))
        If wealth > peak Then peak = wealth
        result(i, 1) = wealth / peak - 1#
    Next i
    DrawdownSeries = result
End Function

' Number of observations spent below the prior peak across the current/last drawdown.
Public Function DrawdownDuration(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim wealth As Double, peak As Double
    Dim duration As Long, i As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        DrawdownDuration = CVErr(xlErrValue)
        Exit Function
    End If

    wealth = 1#: peak = 1#
    For i = 1 To UBound(data)
        wealth = wealth * (1# + CDbl(data(i)))
        If wealth >= peak Then
            peak = wealth
            duration = 0
        Else
            duration = duration + 1
        End If
    Next i
    DrawdownDuration = duration
End Function

Public Function LongestDrawdown(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim wealth As Double, peak As Double
    Dim duration As Long, maxDuration As Long, i As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        LongestDrawdown = CVErr(xlErrValue)
        Exit Function
    End If

    wealth = 1#: peak = 1#
    For i = 1 To UBound(data)
        wealth = wealth * (1# + CDbl(data(i)))
        If wealth >= peak Then
            peak = wealth
            duration = 0
        Else
            duration = duration + 1
            If duration > maxDuration Then maxDuration = duration
        End If
    Next i
    LongestDrawdown = maxDuration
End Function

Public Function AverageDrawdown(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim wealth As Double, peak As Double, dd As Double
    Dim total As Double, count As Long, i As Long

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        AverageDrawdown = CVErr(xlErrValue)
        Exit Function
    End If

    wealth = 1#: peak = 1#
    For i = 1 To UBound(data)
        wealth = wealth * (1# + CDbl(data(i)))
        If wealth > peak Then peak = wealth
        dd = wealth / peak - 1#
        If dd < 0 Then
            total = total + dd
            count = count + 1
        End If
    Next i

    If count = 0 Then AverageDrawdown = 0# Else AverageDrawdown = total / count
End Function

Public Function HistoricalVaR(Returns As Range, Optional ConfidenceLevel As Double = 0.95) As Variant
    Dim data As Variant, hasError As Boolean
    Dim q As Variant

    If ConfidenceLevel <= 0# Or ConfidenceLevel >= 1# Then
        HistoricalVaR = CVErr(xlErrNum)
        Exit Function
    End If
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        HistoricalVaR = CVErr(xlErrValue)
        Exit Function
    End If
    q = EITPercentile(data, 1# - ConfidenceLevel)
    If IsError(q) Then HistoricalVaR = q Else HistoricalVaR = -CDbl(q)
End Function

Public Function HistoricalCVaR(Returns As Range, Optional ConfidenceLevel As Double = 0.95) As Variant
    Dim data As Variant, hasError As Boolean
    Dim q As Variant
    Dim total As Double, count As Long, i As Long

    If ConfidenceLevel <= 0# Or ConfidenceLevel >= 1# Then
        HistoricalCVaR = CVErr(xlErrNum)
        Exit Function
    End If
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        HistoricalCVaR = CVErr(xlErrValue)
        Exit Function
    End If
    q = EITPercentile(data, 1# - ConfidenceLevel)
    If IsError(q) Then
        HistoricalCVaR = q
        Exit Function
    End If

    For i = 1 To UBound(data)
        If CDbl(data(i)) <= CDbl(q) Then
            total = total + CDbl(data(i))
            count = count + 1
        End If
    Next i
    If count = 0 Then HistoricalCVaR = CVErr(xlErrNA) Else HistoricalCVaR = -(total / count)
End Function

Public Function Skewness(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim n As Long, i As Long
    Dim mu As Double, s As Double, total As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        Skewness = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(data)
    If n < 3 Then
        Skewness = CVErr(xlErrNum)
        Exit Function
    End If

    mu = EITMean(data)
    s = Sqr(EITSampleVariance(data))
    If s = 0 Then
        Skewness = CVErr(xlErrDiv0)
        Exit Function
    End If
    For i = 1 To n
        total = total + ((CDbl(data(i)) - mu) / s) ^ 3
    Next i
    Skewness = n * total / ((n - 1#) * (n - 2#))
End Function

Public Function ExcessKurtosis(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    Dim n As Long, i As Long
    Dim mu As Double, s As Double, sum4 As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        ExcessKurtosis = CVErr(xlErrValue)
        Exit Function
    End If
    n = UBound(data)
    If n < 4 Then
        ExcessKurtosis = CVErr(xlErrNum)
        Exit Function
    End If

    mu = EITMean(data)
    s = Sqr(EITSampleVariance(data))
    If s = 0 Then
        ExcessKurtosis = CVErr(xlErrDiv0)
        Exit Function
    End If
    For i = 1 To n
        sum4 = sum4 + ((CDbl(data(i)) - mu) / s) ^ 4
    Next i

    ExcessKurtosis = (n * (n + 1#) * sum4 / ((n - 1#) * (n - 2#) * (n - 3#))) - _
                     (3# * (n - 1#) ^ 2 / ((n - 2#) * (n - 3#)))
End Function

Public Function TailRatio(Returns As Range, Optional UpperPercentile As Double = 0.95) As Variant
    Dim data As Variant, hasError As Boolean
    Dim upperQ As Variant, lowerQ As Variant

    If UpperPercentile <= 0.5 Or UpperPercentile >= 1# Then
        TailRatio = CVErr(xlErrNum)
        Exit Function
    End If
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        TailRatio = CVErr(xlErrValue)
        Exit Function
    End If
    upperQ = EITPercentile(data, UpperPercentile)
    lowerQ = EITPercentile(data, 1# - UpperPercentile)
    If IsError(upperQ) Or IsError(lowerQ) Or CDbl(lowerQ) = 0 Then
        TailRatio = CVErr(xlErrDiv0)
    Else
        TailRatio = CDbl(upperQ) / Abs(CDbl(lowerQ))
    End If
End Function
