Attribute VB_Name = "ToolkitCore"
Option Explicit
Option Private Module

' Shared helpers for the modular v2 toolkit.
' Option Private Module keeps these helpers out of the Excel function wizard while
' allowing the other modules in the same VBA project to call them.

Public Function EITRangeToVector(inputRange As Range, ByRef hasError As Boolean, _
                                 Optional IgnoreBlanks As Boolean = True, _
                                 Optional IgnoreErrors As Boolean = False) As Variant
    Dim data() As Double
    Dim cell As Range
    Dim count As Long
    Dim value As Variant

    hasError = False

    If inputRange Is Nothing Then
        hasError = True
        Exit Function
    End If

    If inputRange.Rows.Count > 1 And inputRange.Columns.Count > 1 Then
        hasError = True
        Exit Function
    End If

    ReDim data(1 To inputRange.Count)
    count = 0

    For Each cell In inputRange.Cells
        value = cell.Value

        If IsError(value) Then
            If Not IgnoreErrors Then
                hasError = True
                Exit Function
            End If
        ElseIf IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
            If Not IgnoreBlanks Then
                hasError = True
                Exit Function
            End If
        ElseIf IsNumeric(value) Then
            count = count + 1
            data(count) = CDbl(value)
        Else
            hasError = True
            Exit Function
        End If
    Next cell

    If count = 0 Then
        hasError = True
        Exit Function
    End If

    ReDim Preserve data(1 To count)
    EITRangeToVector = data
End Function

Public Function EITRangeToDateVector(inputRange As Range, ByRef hasError As Boolean, _
                                     Optional IgnoreBlanks As Boolean = True) As Variant
    Dim data() As Double
    Dim cell As Range
    Dim count As Long
    Dim value As Variant

    hasError = False

    If inputRange Is Nothing Then
        hasError = True
        Exit Function
    End If

    If inputRange.Rows.Count > 1 And inputRange.Columns.Count > 1 Then
        hasError = True
        Exit Function
    End If

    ReDim data(1 To inputRange.Count)
    count = 0

    For Each cell In inputRange.Cells
        value = cell.Value
        If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
            If Not IgnoreBlanks Then
                hasError = True
                Exit Function
            End If
        ElseIf IsDate(value) Or IsNumeric(value) Then
            count = count + 1
            data(count) = CDbl(CDate(value))
        Else
            hasError = True
            Exit Function
        End If
    Next cell

    If count = 0 Then
        hasError = True
        Exit Function
    End If

    ReDim Preserve data(1 To count)
    EITRangeToDateVector = data
End Function

Public Sub EITPairedVectors(firstRange As Range, secondRange As Range, _
                            ByRef firstData As Variant, ByRef secondData As Variant, _
                            ByRef hasError As Boolean, _
                            Optional IgnoreBlanks As Boolean = True, _
                            Optional IgnoreErrors As Boolean = False)
    Dim a() As Double, b() As Double
    Dim i As Long, count As Long
    Dim va As Variant, vb As Variant

    hasError = False

    If firstRange Is Nothing Or secondRange Is Nothing Then
        hasError = True
        Exit Sub
    End If

    If firstRange.Count <> secondRange.Count Then
        hasError = True
        Exit Sub
    End If

    If (firstRange.Rows.Count > 1 And firstRange.Columns.Count > 1) Or _
       (secondRange.Rows.Count > 1 And secondRange.Columns.Count > 1) Then
        hasError = True
        Exit Sub
    End If

    ReDim a(1 To firstRange.Count)
    ReDim b(1 To secondRange.Count)
    count = 0

    For i = 1 To firstRange.Count
        va = firstRange.Cells(i).Value
        vb = secondRange.Cells(i).Value

        If IsError(va) Or IsError(vb) Then
            If Not IgnoreErrors Then
                hasError = True
                Exit Sub
            End If
        ElseIf IsEmpty(va) Or IsEmpty(vb) Or Len(Trim$(CStr(va))) = 0 Or Len(Trim$(CStr(vb))) = 0 Then
            If Not IgnoreBlanks Then
                hasError = True
                Exit Sub
            End If
        ElseIf IsNumeric(va) And IsNumeric(vb) Then
            count = count + 1
            a(count) = CDbl(va)
            b(count) = CDbl(vb)
        Else
            hasError = True
            Exit Sub
        End If
    Next i

    If count = 0 Then
        hasError = True
        Exit Sub
    End If

    ReDim Preserve a(1 To count)
    ReDim Preserve b(1 To count)
    firstData = a
    secondData = b
End Sub

Public Function EITMean(data As Variant) As Double
    Dim i As Long
    Dim total As Double
    For i = LBound(data) To UBound(data)
        total = total + CDbl(data(i))
    Next i
    EITMean = total / (UBound(data) - LBound(data) + 1)
End Function

Public Function EITSampleVariance(data As Variant) As Double
    Dim n As Long, i As Long
    Dim mu As Double, ss As Double

    n = UBound(data) - LBound(data) + 1
    If n <= 1 Then
        EITSampleVariance = 0
        Exit Function
    End If

    mu = EITMean(data)
    For i = LBound(data) To UBound(data)
        ss = ss + (CDbl(data(i)) - mu) ^ 2
    Next i
    EITSampleVariance = ss / (n - 1)
End Function

Public Function EITSampleCovariance(firstData As Variant, secondData As Variant) As Double
    Dim n As Long, i As Long
    Dim meanA As Double, meanB As Double, total As Double

    n = UBound(firstData) - LBound(firstData) + 1
    If n <= 1 Then
        EITSampleCovariance = 0
        Exit Function
    End If

    meanA = EITMean(firstData)
    meanB = EITMean(secondData)

    For i = LBound(firstData) To UBound(firstData)
        total = total + (CDbl(firstData(i)) - meanA) * (CDbl(secondData(i)) - meanB)
    Next i
    EITSampleCovariance = total / (n - 1)
End Function

Public Function EITGeometricReturn(data As Variant) As Double
    Dim i As Long
    Dim wealth As Double
    wealth = 1#

    For i = LBound(data) To UBound(data)
        If CDbl(data(i)) <= -1# Then
            EITGeometricReturn = -1#
            Exit Function
        End If
        wealth = wealth * (1# + CDbl(data(i)))
    Next i
    EITGeometricReturn = wealth - 1#
End Function

Public Function EITGeometricAnnualizedReturn(data As Variant, PeriodsPerYear As Double) As Variant
    Dim n As Long
    Dim totalReturn As Double

    If PeriodsPerYear <= 0 Then
        EITGeometricAnnualizedReturn = CVErr(xlErrNum)
        Exit Function
    End If

    n = UBound(data) - LBound(data) + 1
    totalReturn = EITGeometricReturn(data)
    If totalReturn <= -1# Then
        EITGeometricAnnualizedReturn = -1#
    Else
        EITGeometricAnnualizedReturn = (1# + totalReturn) ^ (PeriodsPerYear / n) - 1#
    End If
End Function

Public Function EITPeriodicRate(AnnualRate As Double, PeriodsPerYear As Double) As Variant
    If PeriodsPerYear <= 0 Or AnnualRate <= -1# Then
        EITPeriodicRate = CVErr(xlErrNum)
    Else
        EITPeriodicRate = (1# + AnnualRate) ^ (1# / PeriodsPerYear) - 1#
    End If
End Function

Public Function EITSliceVector(data As Variant, startIndex As Long, endIndex As Long) As Variant
    Dim result() As Double
    Dim i As Long, j As Long

    ReDim result(1 To endIndex - startIndex + 1)
    j = 1
    For i = startIndex To endIndex
        result(j) = CDbl(data(i))
        j = j + 1
    Next i
    EITSliceVector = result
End Function

Public Function EITArrayDifference(firstData As Variant, secondData As Variant) As Variant
    Dim result() As Double
    Dim i As Long, n As Long

    n = UBound(firstData) - LBound(firstData) + 1
    ReDim result(1 To n)
    For i = 1 To n
        result(i) = CDbl(firstData(LBound(firstData) + i - 1)) - CDbl(secondData(LBound(secondData) + i - 1))
    Next i
    EITArrayDifference = result
End Function

Public Function EITDownsideDeviation(data As Variant, MARPerPeriod As Double) As Double
    Dim i As Long, n As Long
    Dim d As Double, ss As Double

    n = UBound(data) - LBound(data) + 1
    For i = LBound(data) To UBound(data)
        d = CDbl(data(i)) - MARPerPeriod
        If d < 0 Then ss = ss + d ^ 2
    Next i
    EITDownsideDeviation = Sqr(ss / n)
End Function

Public Function EITMaxDrawdownFromVector(data As Variant) As Double
    Dim wealth As Double, peak As Double, dd As Double, maxDD As Double
    Dim i As Long

    wealth = 1#: peak = 1#: maxDD = 0#
    For i = LBound(data) To UBound(data)
        wealth = wealth * (1# + CDbl(data(i)))
        If wealth > peak Then peak = wealth
        If peak <> 0 Then
            dd = wealth / peak - 1#
            If dd < maxDD Then maxDD = dd
        End If
    Next i
    EITMaxDrawdownFromVector = maxDD
End Function

Public Function EITPercentile(data As Variant, percentile As Double) As Variant
    Dim sorted() As Double
    Dim n As Long, i As Long, j As Long
    Dim temp As Double, position As Double, lo As Long, hi As Long

    If percentile < 0# Or percentile > 1# Then
        EITPercentile = CVErr(xlErrNum)
        Exit Function
    End If

    n = UBound(data) - LBound(data) + 1
    ReDim sorted(1 To n)
    For i = 1 To n
        sorted(i) = CDbl(data(LBound(data) + i - 1))
    Next i

    For i = 2 To n
        temp = sorted(i)
        j = i - 1
        Do While j >= 1
            If sorted(j) <= temp Then Exit Do
            sorted(j + 1) = sorted(j)
            j = j - 1
        Loop
        sorted(j + 1) = temp
    Next i

    If n = 1 Then
        EITPercentile = sorted(1)
        Exit Function
    End If

    position = 1# + percentile * (n - 1#)
    lo = Int(position)
    hi = lo + 1
    If hi > n Then
        EITPercentile = sorted(n)
    Else
        EITPercentile = sorted(lo) + (position - lo) * (sorted(hi) - sorted(lo))
    End If
End Function

Public Function EITIsAscending(data As Variant) As Boolean
    Dim i As Long
    EITIsAscending = True
    For i = LBound(data) + 1 To UBound(data)
        If CDbl(data(i)) < CDbl(data(i - 1)) Then
            EITIsAscending = False
            Exit Function
        End If
    Next i
End Function
