Attribute VB_Name = "CompositePerformance"
Option Explicit

' Composite-level performance utilities.

' Beginning-market-value weighted composite return for one period.
Public Function CompositeReturn(PortfolioReturns As Range, BeginningMarketValues As Range) As Variant
    Dim r As Variant, mv As Variant, hasError As Boolean
    Dim numerator As Double, denominator As Double, i As Long

    EITPairedVectors PortfolioReturns, BeginningMarketValues, r, mv, hasError
    If hasError Then
        CompositeReturn = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To UBound(r)
        If mv(i) < 0 Then
            CompositeReturn = CVErr(xlErrNum)
            Exit Function
        End If
        numerator = numerator + r(i) * mv(i)
        denominator = denominator + mv(i)
    Next i
    If denominator = 0 Then CompositeReturn = CVErr(xlErrDiv0) Else CompositeReturn = numerator / denominator
End Function

' Equal-weighted sample standard deviation of constituent portfolio returns.
Public Function CompositeDispersion(PortfolioReturns As Range) As Variant
    Dim r As Variant, hasError As Boolean
    r = EITRangeToVector(PortfolioReturns, hasError)
    If hasError Then
        CompositeDispersion = CVErr(xlErrValue)
    ElseIf UBound(r) < 2 Then
        CompositeDispersion = CVErr(xlErrNum)
    Else
        CompositeDispersion = Sqr(EITSampleVariance(r))
    End If
End Function

' Beginning-market-value weighted dispersion around the asset-weighted composite return.
Public Function AssetWeightedCompositeDispersion(PortfolioReturns As Range, _
                                                 BeginningMarketValues As Range) As Variant
    Dim r As Variant, mv As Variant, hasError As Boolean
    Dim composite As Double, totalMV As Double, weightedSS As Double
    Dim i As Long

    EITPairedVectors PortfolioReturns, BeginningMarketValues, r, mv, hasError
    If hasError Then
        AssetWeightedCompositeDispersion = CVErr(xlErrValue)
        Exit Function
    End If

    For i = 1 To UBound(r)
        If mv(i) < 0 Then
            AssetWeightedCompositeDispersion = CVErr(xlErrNum)
            Exit Function
        End If
        totalMV = totalMV + mv(i)
        composite = composite + r(i) * mv(i)
    Next i
    If totalMV = 0 Then
        AssetWeightedCompositeDispersion = CVErr(xlErrDiv0)
        Exit Function
    End If
    composite = composite / totalMV

    For i = 1 To UBound(r)
        weightedSS = weightedSS + mv(i) * (r(i) - composite) ^ 2
    Next i
    AssetWeightedCompositeDispersion = Sqr(weightedSS / totalMV)
End Function
