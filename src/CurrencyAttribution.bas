Attribute VB_Name = "CurrencyAttribution"
Option Explicit

' Transparent currency return decomposition.
' For an unhedged foreign asset, base-currency return is:
' (1 + local return) * (1 + currency return) - 1
' = local + currency + local*currency.

' Returns four columns per observation: Local, Currency, Interaction, Total.
Public Function CurrencyReturnDecomposition(LocalReturns As Range, CurrencyReturns As Range) As Variant
    Dim localR As Variant, fxR As Variant, hasError As Boolean
    Dim result() As Double, i As Long

    EITPairedVectors LocalReturns, CurrencyReturns, localR, fxR, hasError
    If hasError Then
        CurrencyReturnDecomposition = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To UBound(localR), 1 To 4)
    For i = 1 To UBound(localR)
        result(i, 1) = localR(i)
        result(i, 2) = fxR(i)
        result(i, 3) = localR(i) * fxR(i)
        result(i, 4) = result(i, 1) + result(i, 2) + result(i, 3)
    Next i
    CurrencyReturnDecomposition = result
End Function

' Weighted contribution version for securities/countries.
' Returns Local Contribution, Currency Contribution, Interaction, Total Contribution.
Public Function CurrencyContribution(Weights As Range, LocalReturns As Range, CurrencyReturns As Range) As Variant
    Dim w() As Double, localR() As Double, fxR() As Double
    Dim hasError As Boolean, result() As Double
    Dim i As Long

    LoadThreeVectors Weights, LocalReturns, CurrencyReturns, w, localR, fxR, hasError
    If hasError Then
        CurrencyContribution = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To UBound(w), 1 To 4)
    For i = 1 To UBound(w)
        result(i, 1) = w(i) * localR(i)
        result(i, 2) = w(i) * fxR(i)
        result(i, 3) = w(i) * localR(i) * fxR(i)
        result(i, 4) = result(i, 1) + result(i, 2) + result(i, 3)
    Next i
    CurrencyContribution = result
End Function

Private Sub LoadThreeVectors(firstRange As Range, secondRange As Range, thirdRange As Range, _
                             ByRef a() As Double, ByRef b() As Double, ByRef c() As Double, _
                             ByRef hasError As Boolean)
    Dim n As Long, i As Long, count As Long
    Dim va As Variant, vb As Variant, vc As Variant

    hasError = False
    n = firstRange.Count
    If secondRange.Count <> n Or thirdRange.Count <> n Then
        hasError = True
        Exit Sub
    End If
    If (firstRange.Rows.Count > 1 And firstRange.Columns.Count > 1) Or _
       (secondRange.Rows.Count > 1 And secondRange.Columns.Count > 1) Or _
       (thirdRange.Rows.Count > 1 And thirdRange.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim a(1 To n): ReDim b(1 To n): ReDim c(1 To n)
    For i = 1 To n
        va = firstRange.Cells(i).Value: vb = secondRange.Cells(i).Value: vc = thirdRange.Cells(i).Value
        If Not (IsEmpty(va) Or IsEmpty(vb) Or IsEmpty(vc)) Then
            If Not IsNumeric(va) Or Not IsNumeric(vb) Or Not IsNumeric(vc) Then
                hasError = True
                Exit Sub
            End If
            count = count + 1
            a(count) = CDbl(va): b(count) = CDbl(vb): c(count) = CDbl(vc)
        End If
    Next i
    If count = 0 Then
        hasError = True
        Exit Sub
    End If
    ReDim Preserve a(1 To count): ReDim Preserve b(1 To count): ReDim Preserve c(1 To count)
End Sub
