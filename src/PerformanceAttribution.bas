Attribute VB_Name = "PerformanceAttribution"
Option Explicit

' Single-period contribution and Brinson attribution plus Carino multi-period linking.
' Inputs are decimal weights and returns. Brinson functions return four columns per row:
' Allocation, Selection, Interaction, Total Active Contribution.

Public Function ReturnContribution(Weights As Range, Returns As Range) As Variant
    Dim w As Variant, r As Variant, hasError As Boolean
    Dim result() As Double, i As Long

    EITPairedVectors Weights, Returns, w, r, hasError
    If hasError Then
        ReturnContribution = CVErr(xlErrValue)
        Exit Function
    End If
    ReDim result(1 To UBound(w), 1 To 1)
    For i = 1 To UBound(w)
        result(i, 1) = w(i) * r(i)
    Next i
    ReturnContribution = result
End Function

Public Function ActiveContribution(PortfolioWeights As Range, PortfolioReturns As Range, _
                                   BenchmarkWeights As Range, BenchmarkReturns As Range) As Variant
    Dim pw() As Double, pr() As Double, bw() As Double, br() As Double
    Dim hasError As Boolean, result() As Double, i As Long

    LoadAttributionInputs PortfolioWeights, PortfolioReturns, BenchmarkWeights, BenchmarkReturns, _
                          pw, pr, bw, br, hasError
    If hasError Then
        ActiveContribution = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To UBound(pw), 1 To 1)
    For i = 1 To UBound(pw)
        result(i, 1) = pw(i) * pr(i) - bw(i) * br(i)
    Next i
    ActiveContribution = result
End Function

Public Function BrinsonFachlerAttribution(PortfolioWeights As Range, PortfolioReturns As Range, _
                                          BenchmarkWeights As Range, BenchmarkReturns As Range) As Variant
    Dim pw() As Double, pr() As Double, bw() As Double, br() As Double
    Dim hasError As Boolean, result() As Double
    Dim benchmarkTotal As Double
    Dim allocation As Double, selection As Double, interaction As Double
    Dim i As Long, n As Long

    LoadAttributionInputs PortfolioWeights, PortfolioReturns, BenchmarkWeights, BenchmarkReturns, _
                          pw, pr, bw, br, hasError
    If hasError Then
        BrinsonFachlerAttribution = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(pw)
    If Not WeightsApproximatelyEqualOne(pw) Or Not WeightsApproximatelyEqualOne(bw) Then
        BrinsonFachlerAttribution = CVErr(xlErrNum)
        Exit Function
    End If

    For i = 1 To n
        benchmarkTotal = benchmarkTotal + bw(i) * br(i)
    Next i

    ReDim result(1 To n, 1 To 4)
    For i = 1 To n
        allocation = (pw(i) - bw(i)) * (br(i) - benchmarkTotal)
        selection = bw(i) * (pr(i) - br(i))
        interaction = (pw(i) - bw(i)) * (pr(i) - br(i))
        result(i, 1) = allocation
        result(i, 2) = selection
        result(i, 3) = interaction
        result(i, 4) = allocation + selection + interaction
    Next i
    BrinsonFachlerAttribution = result
End Function

Public Function BrinsonHoodBeebowerAttribution(PortfolioWeights As Range, PortfolioReturns As Range, _
                                               BenchmarkWeights As Range, BenchmarkReturns As Range) As Variant
    Dim pw() As Double, pr() As Double, bw() As Double, br() As Double
    Dim hasError As Boolean, result() As Double
    Dim allocation As Double, selection As Double, interaction As Double
    Dim i As Long, n As Long

    LoadAttributionInputs PortfolioWeights, PortfolioReturns, BenchmarkWeights, BenchmarkReturns, _
                          pw, pr, bw, br, hasError
    If hasError Then
        BrinsonHoodBeebowerAttribution = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(pw)
    If Not WeightsApproximatelyEqualOne(pw) Or Not WeightsApproximatelyEqualOne(bw) Then
        BrinsonHoodBeebowerAttribution = CVErr(xlErrNum)
        Exit Function
    End If

    ReDim result(1 To n, 1 To 4)
    For i = 1 To n
        allocation = (pw(i) - bw(i)) * br(i)
        selection = bw(i) * (pr(i) - br(i))
        interaction = (pw(i) - bw(i)) * (pr(i) - br(i))
        result(i, 1) = allocation
        result(i, 2) = selection
        result(i, 3) = interaction
        result(i, 4) = allocation + selection + interaction
    Next i
    BrinsonHoodBeebowerAttribution = result
End Function

' Links a period-by-effect attribution matrix using Carino logarithmic smoothing.
' PortfolioReturns and BenchmarkReturns must have one observation per matrix row.
' Returns a one-row spill range with the linked cumulative effect for each input column.
Public Function CarinoLinkAttribution(AttributionEffects As Range, _
                                      PortfolioReturns As Range, BenchmarkReturns As Range) As Variant
    Dim p As Variant, b As Variant, hasError As Boolean
    Dim effects As Variant, result() As Double
    Dim rows As Long, cols As Long, i As Long, j As Long
    Dim pTotal As Double, bTotal As Double
    Dim localK As Double, totalK As Double, scale As Double

    EITPairedVectors PortfolioReturns, BenchmarkReturns, p, b, hasError, False, False
    If hasError Then
        CarinoLinkAttribution = CVErr(xlErrValue)
        Exit Function
    End If

    effects = AttributionEffects.Value
    If AttributionEffects.Count = 1 Then
        rows = 1: cols = 1
    Else
        rows = AttributionEffects.Rows.Count
        cols = AttributionEffects.Columns.Count
    End If

    If rows <> UBound(p) Then
        CarinoLinkAttribution = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To rows
        If p(i) <= -1# Or b(i) <= -1# Then
            CarinoLinkAttribution = CVErr(xlErrNum)
            Exit Function
        End If
    Next i

    pTotal = LinkedReturnFromVector(p)
    bTotal = LinkedReturnFromVector(b)
    If pTotal <= -1# Or bTotal <= -1# Then
        CarinoLinkAttribution = CVErr(xlErrNum)
        Exit Function
    End If
    totalK = CarinoK(pTotal, bTotal)
    If totalK = 0 Then
        CarinoLinkAttribution = CVErr(xlErrDiv0)
        Exit Function
    End If

    ReDim result(1 To 1, 1 To cols)
    For i = 1 To rows
        localK = CarinoK(p(i), b(i))
        scale = localK / totalK
        For j = 1 To cols
            If AttributionEffects.Count = 1 Then
                If Not IsNumeric(effects) Then
                    CarinoLinkAttribution = CVErr(xlErrValue)
                    Exit Function
                End If
                result(1, 1) = result(1, 1) + CDbl(effects) * scale
            Else
                If IsEmpty(effects(i, j)) Or Not IsNumeric(effects(i, j)) Then
                    CarinoLinkAttribution = CVErr(xlErrValue)
                    Exit Function
                End If
                result(1, j) = result(1, j) + CDbl(effects(i, j)) * scale
            End If
        Next j
    Next i
    CarinoLinkAttribution = result
End Function

' Simple factor contribution for an already-estimated linear factor model.
' This deliberately does not estimate exposures or factor returns.
Public Function FactorContribution(FactorExposures As Range, FactorReturns As Range) As Variant
    Dim e As Variant, f As Variant, hasError As Boolean
    Dim result() As Double, i As Long
    EITPairedVectors FactorExposures, FactorReturns, e, f, hasError
    If hasError Then
        FactorContribution = CVErr(xlErrValue)
        Exit Function
    End If
    ReDim result(1 To UBound(e), 1 To 1)
    For i = 1 To UBound(e)
        result(i, 1) = e(i) * f(i)
    Next i
    FactorContribution = result
End Function

Private Sub LoadAttributionInputs(PortfolioWeights As Range, PortfolioReturns As Range, _
                                  BenchmarkWeights As Range, BenchmarkReturns As Range, _
                                  ByRef pw() As Double, ByRef pr() As Double, _
                                  ByRef bw() As Double, ByRef br() As Double, _
                                  ByRef hasError As Boolean)
    Dim i As Long, count As Long, n As Long
    Dim a As Variant, b As Variant, c As Variant, d As Variant

    hasError = False
    n = PortfolioWeights.Count
    If PortfolioReturns.Count <> n Or BenchmarkWeights.Count <> n Or BenchmarkReturns.Count <> n Then
        hasError = True
        Exit Sub
    End If
    If (PortfolioWeights.Rows.Count > 1 And PortfolioWeights.Columns.Count > 1) Or _
       (PortfolioReturns.Rows.Count > 1 And PortfolioReturns.Columns.Count > 1) Or _
       (BenchmarkWeights.Rows.Count > 1 And BenchmarkWeights.Columns.Count > 1) Or _
       (BenchmarkReturns.Rows.Count > 1 And BenchmarkReturns.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim pw(1 To n): ReDim pr(1 To n): ReDim bw(1 To n): ReDim br(1 To n)
    For i = 1 To n
        a = PortfolioWeights.Cells(i).Value
        b = PortfolioReturns.Cells(i).Value
        c = BenchmarkWeights.Cells(i).Value
        d = BenchmarkReturns.Cells(i).Value
        If Not (IsEmpty(a) Or IsEmpty(b) Or IsEmpty(c) Or IsEmpty(d)) Then
            If Not IsNumeric(a) Or Not IsNumeric(b) Or Not IsNumeric(c) Or Not IsNumeric(d) Then
                hasError = True
                Exit Sub
            End If
            count = count + 1
            pw(count) = CDbl(a): pr(count) = CDbl(b)
            bw(count) = CDbl(c): br(count) = CDbl(d)
        End If
    Next i
    If count = 0 Then
        hasError = True
        Exit Sub
    End If
    ReDim Preserve pw(1 To count): ReDim Preserve pr(1 To count)
    ReDim Preserve bw(1 To count): ReDim Preserve br(1 To count)
End Sub

Private Function WeightsApproximatelyEqualOne(weights() As Double) As Boolean
    Dim i As Long, total As Double
    For i = 1 To UBound(weights)
        total = total + weights(i)
    Next i
    WeightsApproximatelyEqualOne = (Abs(total - 1#) <= 0.000001)
End Function

Private Function LinkedReturnFromVector(data As Variant) As Double
    Dim wealth As Double, i As Long
    wealth = 1#
    For i = 1 To UBound(data)
        wealth = wealth * (1# + CDbl(data(i)))
    Next i
    LinkedReturnFromVector = wealth - 1#
End Function

Private Function CarinoK(portfolioReturn As Double, benchmarkReturn As Double) As Double
    If Abs(portfolioReturn - benchmarkReturn) < 0.000000000001 Then
        CarinoK = 1# / (1# + portfolioReturn)
    Else
        CarinoK = (Log(1# + portfolioReturn) - Log(1# + benchmarkReturn)) / _
                  (portfolioReturn - benchmarkReturn)
    End If
End Function
