Attribute VB_Name = "FixedIncomeAttribution"
Option Explicit

' Transparent first-order fixed-income attribution primitives.
' Yield and spread changes are decimals (10 bp = 0.001).
' These functions are approximations, not a replacement for a full key-rate/curve model.

Public Function DurationContribution(Weights As Range, EffectiveDurations As Range, YieldChanges As Range) As Variant
    DurationContribution = FirstOrderContribution(Weights, EffectiveDurations, YieldChanges, -1#)
End Function

Public Function SpreadContribution(Weights As Range, SpreadDurations As Range, SpreadChanges As Range) As Variant
    SpreadContribution = FirstOrderContribution(Weights, SpreadDurations, SpreadChanges, -1#)
End Function

Public Function ConvexityContribution(Weights As Range, Convexities As Range, YieldChanges As Range) As Variant
    Dim w() As Double, c() As Double, dy() As Double
    Dim hasError As Boolean, result() As Double, i As Long

    LoadThreeNumericVectors Weights, Convexities, YieldChanges, w, c, dy, hasError
    If hasError Then
        ConvexityContribution = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To UBound(w), 1 To 1)
    For i = 1 To UBound(w)
        result(i, 1) = 0.5 * w(i) * c(i) * dy(i) ^ 2
    Next i
    ConvexityContribution = result
End Function

Public Function CarryContribution(Weights As Range, CarryReturns As Range) As Variant
    Dim w As Variant, carry As Variant, hasError As Boolean
    Dim result() As Double, i As Long
    EITPairedVectors Weights, CarryReturns, w, carry, hasError
    If hasError Then
        CarryContribution = CVErr(xlErrValue)
        Exit Function
    End If
    ReDim result(1 To UBound(w), 1 To 1)
    For i = 1 To UBound(w)
        result(i, 1) = w(i) * carry(i)
    Next i
    CarryContribution = result
End Function

' Returns four columns per security: Carry, Rate, Spread, Explained Total.
' Rate and spread effects use first-order duration approximations.
Public Function FixedIncomeReturnDecomposition(Weights As Range, CarryReturns As Range, _
                                               EffectiveDurations As Range, YieldChanges As Range, _
                                               SpreadDurations As Range, SpreadChanges As Range) As Variant
    Dim w() As Double, carry() As Double, dur() As Double, dy() As Double
    Dim spreadDur() As Double, ds() As Double
    Dim result() As Double, hasError As Boolean
    Dim n As Long, i As Long

    LoadSixNumericVectors Weights, CarryReturns, EffectiveDurations, YieldChanges, _
                          SpreadDurations, SpreadChanges, w, carry, dur, dy, spreadDur, ds, hasError
    If hasError Then
        FixedIncomeReturnDecomposition = CVErr(xlErrValue)
        Exit Function
    End If

    n = UBound(w)
    ReDim result(1 To n, 1 To 4)
    For i = 1 To n
        result(i, 1) = w(i) * carry(i)
        result(i, 2) = -w(i) * dur(i) * dy(i)
        result(i, 3) = -w(i) * spreadDur(i) * ds(i)
        result(i, 4) = result(i, 1) + result(i, 2) + result(i, 3)
    Next i
    FixedIncomeReturnDecomposition = result
End Function

Private Function FirstOrderContribution(Weights As Range, Sensitivities As Range, Changes As Range, _
                                        SignMultiplier As Double) As Variant
    Dim w() As Double, s() As Double, chg() As Double
    Dim result() As Double, hasError As Boolean, i As Long

    LoadThreeNumericVectors Weights, Sensitivities, Changes, w, s, chg, hasError
    If hasError Then
        FirstOrderContribution = CVErr(xlErrValue)
        Exit Function
    End If
    ReDim result(1 To UBound(w), 1 To 1)
    For i = 1 To UBound(w)
        result(i, 1) = SignMultiplier * w(i) * s(i) * chg(i)
    Next i
    FirstOrderContribution = result
End Function

Private Sub LoadThreeNumericVectors(firstRange As Range, secondRange As Range, thirdRange As Range, _
                                    ByRef a() As Double, ByRef b() As Double, ByRef c() As Double, _
                                    ByRef hasError As Boolean)
    Dim n As Long, i As Long, count As Long
    Dim va As Variant, vb As Variant, vc As Variant

    hasError = False
    n = firstRange.Count
    If secondRange.Count <> n Or thirdRange.Count <> n Then hasError = True: Exit Sub
    If (firstRange.Rows.Count > 1 And firstRange.Columns.Count > 1) Or _
       (secondRange.Rows.Count > 1 And secondRange.Columns.Count > 1) Or _
       (thirdRange.Rows.Count > 1 And thirdRange.Columns.Count > 1) Then hasError = True: Exit Sub

    ReDim a(1 To n): ReDim b(1 To n): ReDim c(1 To n)
    For i = 1 To n
        va = firstRange.Cells(i).Value: vb = secondRange.Cells(i).Value: vc = thirdRange.Cells(i).Value
        If Not (IsEmpty(va) Or IsEmpty(vb) Or IsEmpty(vc)) Then
            If Not IsNumeric(va) Or Not IsNumeric(vb) Or Not IsNumeric(vc) Then hasError = True: Exit Sub
            count = count + 1
            a(count) = CDbl(va): b(count) = CDbl(vb): c(count) = CDbl(vc)
        End If
    Next i
    If count = 0 Then hasError = True: Exit Sub
    ReDim Preserve a(1 To count): ReDim Preserve b(1 To count): ReDim Preserve c(1 To count)
End Sub

Private Sub LoadSixNumericVectors(r1 As Range, r2 As Range, r3 As Range, r4 As Range, r5 As Range, r6 As Range, _
                                  ByRef a() As Double, ByRef b() As Double, ByRef c() As Double, _
                                  ByRef d() As Double, ByRef e() As Double, ByRef f() As Double, _
                                  ByRef hasError As Boolean)
    Dim n As Long, i As Long, count As Long
    Dim v1 As Variant, v2 As Variant, v3 As Variant, v4 As Variant, v5 As Variant, v6 As Variant

    hasError = False
    n = r1.Count
    If r2.Count <> n Or r3.Count <> n Or r4.Count <> n Or r5.Count <> n Or r6.Count <> n Then hasError = True: Exit Sub
    If (r1.Rows.Count > 1 And r1.Columns.Count > 1) Or (r2.Rows.Count > 1 And r2.Columns.Count > 1) Or _
       (r3.Rows.Count > 1 And r3.Columns.Count > 1) Or (r4.Rows.Count > 1 And r4.Columns.Count > 1) Or _
       (r5.Rows.Count > 1 And r5.Columns.Count > 1) Or (r6.Rows.Count > 1 And r6.Columns.Count > 1) Then hasError = True: Exit Sub

    ReDim a(1 To n): ReDim b(1 To n): ReDim c(1 To n): ReDim d(1 To n): ReDim e(1 To n): ReDim f(1 To n)
    For i = 1 To n
        v1 = r1.Cells(i).Value: v2 = r2.Cells(i).Value: v3 = r3.Cells(i).Value
        v4 = r4.Cells(i).Value: v5 = r5.Cells(i).Value: v6 = r6.Cells(i).Value
        If Not (IsEmpty(v1) Or IsEmpty(v2) Or IsEmpty(v3) Or IsEmpty(v4) Or IsEmpty(v5) Or IsEmpty(v6)) Then
            If Not IsNumeric(v1) Or Not IsNumeric(v2) Or Not IsNumeric(v3) Or Not IsNumeric(v4) Or Not IsNumeric(v5) Or Not IsNumeric(v6) Then hasError = True: Exit Sub
            count = count + 1
            a(count) = CDbl(v1): b(count) = CDbl(v2): c(count) = CDbl(v3)
            d(count) = CDbl(v4): e(count) = CDbl(v5): f(count) = CDbl(v6)
        End If
    Next i
    If count = 0 Then hasError = True: Exit Sub
    ReDim Preserve a(1 To count): ReDim Preserve b(1 To count): ReDim Preserve c(1 To count)
    ReDim Preserve d(1 To count): ReDim Preserve e(1 To count): ReDim Preserve f(1 To count)
End Sub
