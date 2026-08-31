Attribute VB_Name = "ReturnMeasurement"
Option Explicit

' Core return-series measurement functions.
' All return inputs are decimal periodic returns (0.01 = 1%).

Public Function AnnualizedReturn(Returns As Range, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    data = EITRangeToVector(Returns, hasError)
    If hasError Or PeriodsPerYear <= 0 Then
        AnnualizedReturn = CVErr(xlErrValue)
        Exit Function
    End If
    AnnualizedReturn = EITGeometricAnnualizedReturn(data, PeriodsPerYear)
End Function

Public Function CumulativeReturn(Returns As Range) As Variant
    Dim data As Variant, hasError As Boolean
    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        CumulativeReturn = CVErr(xlErrValue)
    Else
        CumulativeReturn = EITGeometricReturn(data)
    End If
End Function

Public Function RollingAnnualizedReturn(Returns As Range, WindowLength As Long, PeriodsPerYear As Double) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Variant
    Dim i As Long, n As Long
    Dim windowData As Variant

    data = EITRangeToVector(Returns, hasError)
    If hasError Then
        RollingAnnualizedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(data) - LBound(data) + 1
    If WindowLength <= 0 Or WindowLength > n Or PeriodsPerYear <= 0 Then
        RollingAnnualizedReturn = CVErr(xlErrNum)
        Exit Function
    End If

    ReDim result(1 To n - WindowLength + 1, 1 To 1)
    For i = WindowLength To n
        windowData = EITSliceVector(data, i - WindowLength + 1, i)
        result(i - WindowLength + 1, 1) = EITGeometricAnnualizedReturn(windowData, PeriodsPerYear)
    Next i
    RollingAnnualizedReturn = result
End Function

' Legacy fixed-observation aggregation retained for compatibility.
' Prefer AggregateReturns for calendar-aware daily/weekly data.
Public Function MonthlyReturns(Returns As Range, ObservationsPerMonth As Long) As Variant
    Dim data As Variant, hasError As Boolean
    Dim result() As Double
    Dim n As Long, monthCount As Long
    Dim i As Long, j As Long, wealth As Double

    data = EITRangeToVector(Returns, hasError)
    If hasError Or ObservationsPerMonth <= 0 Then
        MonthlyReturns = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(data) - LBound(data) + 1
    monthCount = n \ ObservationsPerMonth
    If monthCount = 0 Then
        MonthlyReturns = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To monthCount, 1 To 1)
    For i = 1 To monthCount
        wealth = 1#
        For j = 1 To ObservationsPerMonth
            wealth = wealth * (1# + CDbl(data((i - 1) * ObservationsPerMonth + j)))
        Next j
        result(i, 1) = wealth - 1#
    Next i
    MonthlyReturns = result
End Function

' Calendar-aware aggregation. Returns a two-column spill range:
' [period ending observation date, compounded return].
' Frequency accepts W, M, Q, or Y.
Public Function AggregateReturns(Returns As Range, Dates As Range, Frequency As String) As Variant
    Dim r() As Double, d() As Double, hasError As Boolean
    Dim n As Long, i As Long, groupCount As Long
    Dim result() As Variant
    Dim currentKey As Double, nextKey As Double
    Dim wealth As Double, freq As String

    LoadDateReturnPairs Returns, Dates, r, d, hasError
    If hasError Then
        AggregateReturns = CVErr(xlErrValue)
        Exit Function
    End If

    If Not EITIsAscending(d) Then
        AggregateReturns = CVErr(xlErrNum)
        Exit Function
    End If

    freq = UCase$(Trim$(Frequency))
    If freq <> "W" And freq <> "M" And freq <> "Q" And freq <> "Y" Then
        AggregateReturns = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(r)
    groupCount = 1
    currentKey = PeriodKey(CDate(d(1)), freq)
    For i = 2 To n
        nextKey = PeriodKey(CDate(d(i)), freq)
        If nextKey <> currentKey Then
            groupCount = groupCount + 1
            currentKey = nextKey
        End If
    Next i

    ReDim result(1 To groupCount, 1 To 2)
    groupCount = 1
    currentKey = PeriodKey(CDate(d(1)), freq)
    wealth = 1#

    For i = 1 To n
        nextKey = PeriodKey(CDate(d(i)), freq)
        If nextKey <> currentKey Then
            result(groupCount, 1) = CDate(d(i - 1))
            result(groupCount, 2) = wealth - 1#
            groupCount = groupCount + 1
            currentKey = nextKey
            wealth = 1#
        End If
        wealth = wealth * (1# + r(i))
    Next i

    result(groupCount, 1) = CDate(d(n))
    result(groupCount, 2) = wealth - 1#
    AggregateReturns = result
End Function

Public Function PeriodReturn(Returns As Range, Dates As Range, StartDate As Date, EndDate As Date) As Variant
    Dim r() As Double, d() As Double, hasError As Boolean
    Dim i As Long, count As Long
    Dim wealth As Double

    If EndDate < StartDate Then
        PeriodReturn = CVErr(xlErrNum)
        Exit Function
    End If

    LoadDateReturnPairs Returns, Dates, r, d, hasError
    If hasError Then
        PeriodReturn = CVErr(xlErrValue)
        Exit Function
    End If

    wealth = 1#
    For i = 1 To UBound(r)
        If CDate(d(i)) >= StartDate And CDate(d(i)) <= EndDate Then
            wealth = wealth * (1# + r(i))
            count = count + 1
        End If
    Next i

    If count = 0 Then
        PeriodReturn = CVErr(xlErrNA)
    Else
        PeriodReturn = wealth - 1#
    End If
End Function

Public Function MTDReturn(Returns As Range, Dates As Range, Optional AsOfDate As Variant) As Variant
    Dim asOf As Date
    asOf = ResolveAsOfDate(Dates, AsOfDate)
    If asOf = 0 Then
        MTDReturn = CVErr(xlErrValue)
    Else
        MTDReturn = PeriodReturn(Returns, Dates, DateSerial(Year(asOf), Month(asOf), 1), asOf)
    End If
End Function

Public Function QTDReturn(Returns As Range, Dates As Range, Optional AsOfDate As Variant) As Variant
    Dim asOf As Date, quarterMonth As Long
    asOf = ResolveAsOfDate(Dates, AsOfDate)
    If asOf = 0 Then
        QTDReturn = CVErr(xlErrValue)
        Exit Function
    End If
    quarterMonth = ((Month(asOf) - 1) \ 3) * 3 + 1
    QTDReturn = PeriodReturn(Returns, Dates, DateSerial(Year(asOf), quarterMonth, 1), asOf)
End Function

Public Function YTDReturn(Returns As Range, Dates As Range, Optional AsOfDate As Variant) As Variant
    Dim asOf As Date
    asOf = ResolveAsOfDate(Dates, AsOfDate)
    If asOf = 0 Then
        YTDReturn = CVErr(xlErrValue)
    Else
        YTDReturn = PeriodReturn(Returns, Dates, DateSerial(Year(asOf), 1, 1), asOf)
    End If
End Function

Private Sub LoadDateReturnPairs(Returns As Range, Dates As Range, _
                                ByRef returnData() As Double, ByRef dateData() As Double, _
                                ByRef hasError As Boolean)
    Dim i As Long, count As Long
    Dim rv As Variant, dv As Variant

    hasError = False
    If Returns.Count <> Dates.Count Then
        hasError = True
        Exit Sub
    End If
    If (Returns.Rows.Count > 1 And Returns.Columns.Count > 1) Or _
       (Dates.Rows.Count > 1 And Dates.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim returnData(1 To Returns.Count)
    ReDim dateData(1 To Dates.Count)

    For i = 1 To Returns.Count
        rv = Returns.Cells(i).Value
        dv = Dates.Cells(i).Value
        If Not (IsEmpty(rv) Or IsEmpty(dv) Or Len(Trim$(CStr(rv))) = 0 Or Len(Trim$(CStr(dv))) = 0) Then
            If Not IsNumeric(rv) Or Not (IsDate(dv) Or IsNumeric(dv)) Then
                hasError = True
                Exit Sub
            End If
            count = count + 1
            returnData(count) = CDbl(rv)
            dateData(count) = CDbl(CDate(dv))
        End If
    Next i

    If count = 0 Then
        hasError = True
        Exit Sub
    End If
    ReDim Preserve returnData(1 To count)
    ReDim Preserve dateData(1 To count)
End Sub

Private Function PeriodKey(d As Date, freq As String) As Double
    Select Case freq
        Case "W": PeriodKey = CDbl(d - Weekday(d, vbMonday) + 1)
        Case "M": PeriodKey = CDbl(DateSerial(Year(d), Month(d), 1))
        Case "Q": PeriodKey = CDbl(DateSerial(Year(d), ((Month(d) - 1) \ 3) * 3 + 1, 1))
        Case "Y": PeriodKey = CDbl(DateSerial(Year(d), 1, 1))
    End Select
End Function

Private Function ResolveAsOfDate(Dates As Range, AsOfDate As Variant) As Date
    Dim d As Variant, hasError As Boolean
    If Not IsMissing(AsOfDate) Then
        If IsDate(AsOfDate) Or IsNumeric(AsOfDate) Then ResolveAsOfDate = CDate(AsOfDate)
        Exit Function
    End If

    d = EITRangeToDateVector(Dates, hasError)
    If hasError Then Exit Function
    ResolveAsOfDate = CDate(d(UBound(d)))
End Function
