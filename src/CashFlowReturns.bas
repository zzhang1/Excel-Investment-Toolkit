Attribute VB_Name = "CashFlowReturns"
Option Explicit

' Cash-flow-aware return measurement.
' Modified Dietz convention: positive cash flows are contributions to the portfolio.
' Money-weighted return convention: investor cash flows are negative contributions and
' positive withdrawals/distributions; EndingValue is appended as a positive terminal flow.

Public Function ModifiedDietzReturn(BeginningValue As Double, EndingValue As Double, _
                                    CashFlows As Range, CashFlowDates As Range, _
                                    StartDate As Date, EndDate As Date, _
                                    Optional FlowTiming As String = "ACTUAL") As Variant
    Dim flows As Variant, dates As Variant, hasError As Boolean
    Dim weightedFlows As Double, totalFlows As Double
    Dim weight As Double, denominator As Double
    Dim i As Long, timing As String

    If EndDate <= StartDate Then
        ModifiedDietzReturn = CVErr(xlErrNum)
        Exit Function
    End If

    LoadCashFlowPairs CashFlows, CashFlowDates, flows, dates, hasError
    If hasError Then
        ModifiedDietzReturn = CVErr(xlErrValue)
        Exit Function
    End If

    timing = UCase$(Trim$(FlowTiming))
    If timing <> "ACTUAL" And timing <> "BEGIN" And timing <> "END" Then
        ModifiedDietzReturn = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To UBound(flows)
        If CDate(dates(i)) < StartDate Or CDate(dates(i)) > EndDate Then
            ModifiedDietzReturn = CVErr(xlErrNum)
            Exit Function
        End If

        Select Case timing
            Case "BEGIN": weight = 1#
            Case "END": weight = 0#
            Case Else
                weight = (CDbl(EndDate) - CDbl(CDate(dates(i)))) / (CDbl(EndDate) - CDbl(StartDate))
        End Select

        totalFlows = totalFlows + CDbl(flows(i))
        weightedFlows = weightedFlows + weight * CDbl(flows(i))
    Next i

    denominator = BeginningValue + weightedFlows
    If denominator = 0 Then
        ModifiedDietzReturn = CVErr(xlErrDiv0)
    Else
        ModifiedDietzReturn = (EndingValue - BeginningValue - totalFlows) / denominator
    End If
End Function

' True time-weighted return from a sequence of valuations and external cash flows.
' MarketValues and Dates contain n aligned observations. ExternalCashFlows may contain
' n-1 values (one for each interval) or n values (the first value is ignored).
' END convention: flow occurs just before ending valuation => (EndMV - Flow) / BeginMV - 1.
' BEGIN convention: flow occurs just after beginning valuation => EndMV / (BeginMV + Flow) - 1.
Public Function TimeWeightedReturn(MarketValues As Range, Dates As Range, _
                                   ExternalCashFlows As Range, _
                                   Optional FlowTiming As String = "END") As Variant
    Dim values() As Double, dateData() As Double, flows() As Double
    Dim hasError As Boolean, timing As String
    Dim n As Long, i As Long, flowOffset As Long
    Dim cf As Double, subReturn As Double, wealth As Double, denominator As Double

    LoadValuationSeries MarketValues, Dates, values, dateData, hasError
    If hasError Then
        TimeWeightedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(values)
    If n < 2 Then
        TimeWeightedReturn = CVErr(xlErrNum)
        Exit Function
    End If
    If Not EITIsAscending(dateData) Then
        TimeWeightedReturn = CVErr(xlErrNum)
        Exit Function
    End If

    LoadFlowVector ExternalCashFlows, flows, hasError
    If hasError Then
        TimeWeightedReturn = CVErr(xlErrValue)
        Exit Function
    End If
    If UBound(flows) = n Then
        flowOffset = 0
    ElseIf UBound(flows) = n - 1 Then
        flowOffset = -1
    Else
        TimeWeightedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    timing = UCase$(Trim$(FlowTiming))
    If timing <> "END" And timing <> "BEGIN" Then
        TimeWeightedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    wealth = 1#
    For i = 2 To n
        cf = flows(i + flowOffset)
        If timing = "END" Then
            If values(i - 1) = 0 Then
                TimeWeightedReturn = CVErr(xlErrDiv0)
                Exit Function
            End If
            subReturn = (values(i) - cf) / values(i - 1) - 1#
        Else
            denominator = values(i - 1) + cf
            If denominator = 0 Then
                TimeWeightedReturn = CVErr(xlErrDiv0)
                Exit Function
            End If
            subReturn = values(i) / denominator - 1#
        End If
        If subReturn <= -1# Then
            TimeWeightedReturn = -1#
            Exit Function
        End If
        wealth = wealth * (1# + subReturn)
    Next i
    TimeWeightedReturn = wealth - 1#
End Function

' Appends EndingValue at EndingDate and calls Excel's XIRR solver.
Public Function MoneyWeightedReturn(CashFlows As Range, CashFlowDates As Range, _
                                    EndingValue As Double, EndingDate As Date, _
                                    Optional Guess As Double = 0.1) As Variant
    Dim flows As Variant, dates As Variant, hasError As Boolean
    Dim values() As Double, xDates() As Double
    Dim n As Long, i As Long
    Dim hasPositive As Boolean, hasNegative As Boolean

    LoadCashFlowPairs CashFlows, CashFlowDates, flows, dates, hasError
    If hasError Then
        MoneyWeightedReturn = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(flows)
    ReDim values(1 To n + 1)
    ReDim xDates(1 To n + 1)
    For i = 1 To n
        values(i) = CDbl(flows(i))
        xDates(i) = CDbl(CDate(dates(i)))
        If values(i) > 0 Then hasPositive = True
        If values(i) < 0 Then hasNegative = True
    Next i

    values(n + 1) = EndingValue
    xDates(n + 1) = CDbl(EndingDate)
    If EndingValue > 0 Then hasPositive = True
    If EndingValue < 0 Then hasNegative = True

    If Not hasPositive Or Not hasNegative Then
        MoneyWeightedReturn = CVErr(xlErrNum)
        Exit Function
    End If

    On Error GoTo SolverError
    MoneyWeightedReturn = Application.WorksheetFunction.Xirr(values, xDates, Guess)
    Exit Function

SolverError:
    MoneyWeightedReturn = CVErr(xlErrNum)
End Function

Private Sub LoadCashFlowPairs(CashFlows As Range, CashFlowDates As Range, _
                              ByRef flows As Variant, ByRef dates As Variant, _
                              ByRef hasError As Boolean)
    Dim f() As Double, d() As Double
    Dim i As Long, count As Long
    Dim fv As Variant, dv As Variant

    hasError = False
    If CashFlows.Count <> CashFlowDates.Count Then
        hasError = True
        Exit Sub
    End If
    If (CashFlows.Rows.Count > 1 And CashFlows.Columns.Count > 1) Or _
       (CashFlowDates.Rows.Count > 1 And CashFlowDates.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim f(1 To CashFlows.Count)
    ReDim d(1 To CashFlowDates.Count)
    For i = 1 To CashFlows.Count
        fv = CashFlows.Cells(i).Value
        dv = CashFlowDates.Cells(i).Value
        If Not (IsEmpty(fv) Or IsEmpty(dv) Or Len(Trim$(CStr(fv))) = 0 Or Len(Trim$(CStr(dv))) = 0) Then
            If Not IsNumeric(fv) Or Not (IsDate(dv) Or IsNumeric(dv)) Then
                hasError = True
                Exit Sub
            End If
            count = count + 1
            f(count) = CDbl(fv)
            d(count) = CDbl(CDate(dv))
        End If
    Next i
    If count = 0 Then
        hasError = True
        Exit Sub
    End If
    ReDim Preserve f(1 To count): ReDim Preserve d(1 To count)
    flows = f: dates = d
End Sub

Private Sub LoadValuationSeries(MarketValues As Range, Dates As Range, _
                                ByRef values() As Double, ByRef dateData() As Double, _
                                ByRef hasError As Boolean)
    Dim i As Long
    Dim mv As Variant, dv As Variant

    hasError = False
    If MarketValues.Count <> Dates.Count Then
        hasError = True
        Exit Sub
    End If
    If (MarketValues.Rows.Count > 1 And MarketValues.Columns.Count > 1) Or _
       (Dates.Rows.Count > 1 And Dates.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim values(1 To MarketValues.Count): ReDim dateData(1 To Dates.Count)
    For i = 1 To MarketValues.Count
        mv = MarketValues.Cells(i).Value
        dv = Dates.Cells(i).Value
        If IsEmpty(mv) Or IsEmpty(dv) Or Not IsNumeric(mv) Or Not (IsDate(dv) Or IsNumeric(dv)) Then
            hasError = True
            Exit Sub
        End If
        values(i) = CDbl(mv)
        dateData(i) = CDbl(CDate(dv))
    Next i
End Sub

Private Sub LoadFlowVector(ExternalCashFlows As Range, ByRef flows() As Double, ByRef hasError As Boolean)
    Dim i As Long, value As Variant
    hasError = False
    If ExternalCashFlows.Rows.Count > 1 And ExternalCashFlows.Columns.Count > 1 Then
        hasError = True
        Exit Sub
    End If
    ReDim flows(1 To ExternalCashFlows.Count)
    For i = 1 To ExternalCashFlows.Count
        value = ExternalCashFlows.Cells(i).Value
        If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
            flows(i) = 0#
        ElseIf IsNumeric(value) Then
            flows(i) = CDbl(value)
        Else
            hasError = True
            Exit Sub
        End If
    Next i
End Sub
