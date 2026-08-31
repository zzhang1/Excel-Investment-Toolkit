Attribute VB_Name = "TestKnownIdentities"
Option Explicit

' Manual regression tests for the modular v2 toolkit.
' Import this module together with all src/*.bas files and run RunAllToolkitTests.

Public Sub RunAllToolkitTests()
    Dim ws As Worksheet
    Dim oldAlerts As Boolean

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets("EIT_TEST_TMP").Delete
    On Error GoTo TestFailed

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "EIT_TEST_TMP"

    TestReturnMath ws
    TestModifiedDietz ws
    TestTimeWeightedReturn ws
    TestBrinsonReconciliation ws
    TestCompositeReturn ws
    TestCurrencyDecomposition ws
    TestFixedIncomeContribution ws
    TestDrawdown ws

    ws.Delete
    Application.DisplayAlerts = oldAlerts
    MsgBox "Excel Investment Toolkit tests passed.", vbInformation
    Exit Sub

TestFailed:
    On Error Resume Next
    ws.Delete
    Application.DisplayAlerts = oldAlerts
    MsgBox "Toolkit test failed: " & Err.Description, vbCritical
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub TestReturnMath(ws As Worksheet)
    ws.Range("A1:A2").Value = Application.Transpose(Array(0.1, -0.1))
    AssertNear CumulativeReturn(ws.Range("A1:A2")), -0.01, 0.000000001, "CumulativeReturn"

    ws.Range("B1:B4").Value = Application.Transpose(Array(0.02, 0.01, -0.01, 0.03))
    ws.Range("C1:C4").Value = Application.Transpose(Array(0.01, 0, -0.02, 0.01))
    ' Mean active return = 1% per period; with 12 periods/year => 12% arithmetic active return.
    AssertNear ActiveReturn(ws.Range("B1:B4"), ws.Range("C1:C4"), 12), _
               0.12, 0.000000001, "ActiveReturn"
End Sub

Private Sub TestModifiedDietz(ws As Worksheet)
    ws.Range("D1").Value = 10
    ws.Range("E1").Value = DateSerial(2026, 1, 16)
    AssertNear ModifiedDietzReturn(100, 115, ws.Range("D1"), ws.Range("E1"), _
               DateSerial(2026, 1, 1), DateSerial(2026, 1, 31), "ACTUAL"), _
               5# / 105#, 0.000000001, "ModifiedDietzReturn"
End Sub

Private Sub TestTimeWeightedReturn(ws As Worksheet)
    ws.Range("F1:F3").Value = Application.Transpose(Array(100, 110, 121))
    ws.Range("G1:G3").Value = Application.Transpose(Array(DateSerial(2026, 1, 1), DateSerial(2026, 1, 2), DateSerial(2026, 1, 3)))
    ws.Range("H1:H2").Value = Application.Transpose(Array(0, 0))
    AssertNear TimeWeightedReturn(ws.Range("F1:F3"), ws.Range("G1:G3"), ws.Range("H1:H2")), _
               0.21, 0.000000001, "TimeWeightedReturn"
End Sub

Private Sub TestBrinsonReconciliation(ws As Worksheet)
    Dim result As Variant
    Dim totalEffects As Double

    ws.Range("I1:I2").Value = Application.Transpose(Array(0.6, 0.4))
    ws.Range("J1:J2").Value = Application.Transpose(Array(0.1, 0.02))
    ws.Range("K1:K2").Value = Application.Transpose(Array(0.5, 0.5))
    ws.Range("L1:L2").Value = Application.Transpose(Array(0.08, 0.03))

    result = BrinsonFachlerAttribution(ws.Range("I1:I2"), ws.Range("J1:J2"), ws.Range("K1:K2"), ws.Range("L1:L2"))
    totalEffects = result(1, 4) + result(2, 4)
    AssertNear totalEffects, (0.6 * 0.1 + 0.4 * 0.02) - (0.5 * 0.08 + 0.5 * 0.03), _
               0.000000001, "Brinson reconciliation"
End Sub

Private Sub TestCompositeReturn(ws As Worksheet)
    ws.Range("M1:M2").Value = Application.Transpose(Array(0.1, 0))
    ws.Range("N1:N2").Value = Application.Transpose(Array(100, 300))
    AssertNear CompositeReturn(ws.Range("M1:M2"), ws.Range("N1:N2")), 0.025, 0.000000001, "CompositeReturn"
End Sub

Private Sub TestCurrencyDecomposition(ws As Worksheet)
    Dim result As Variant
    ws.Range("O1").Value = 0.1
    ws.Range("P1").Value = 0.05
    result = CurrencyReturnDecomposition(ws.Range("O1"), ws.Range("P1"))
    AssertNear result(1, 4), 0.155, 0.000000001, "CurrencyReturnDecomposition"
End Sub

Private Sub TestFixedIncomeContribution(ws As Worksheet)
    Dim result As Variant
    ws.Range("Q1").Value = 1
    ws.Range("R1").Value = 5
    ws.Range("S1").Value = 0.001
    result = DurationContribution(ws.Range("Q1"), ws.Range("R1"), ws.Range("S1"))
    AssertNear result(1, 1), -0.005, 0.000000001, "DurationContribution"
End Sub

Private Sub TestDrawdown(ws As Worksheet)
    ws.Range("T1:T3").Value = Application.Transpose(Array(0.1, -0.2, 0.05))
    AssertNear MaxDrawdown(ws.Range("T1:T3")), -0.2, 0.000000001, "MaxDrawdown"
    If LongestDrawdown(ws.Range("T1:T3")) <> 2 Then
        Err.Raise vbObjectError + 1002, "TestKnownIdentities", "LongestDrawdown expected 2 observations."
    End If
End Sub

Private Sub AssertNear(actual As Variant, expected As Double, tolerance As Double, testName As String)
    If IsError(actual) Then
        Err.Raise vbObjectError + 1000, "TestKnownIdentities", testName & " returned an Excel error."
    End If
    If Abs(CDbl(actual) - expected) > tolerance Then
        Err.Raise vbObjectError + 1001, "TestKnownIdentities", _
                  testName & " expected " & CStr(expected) & " but got " & CStr(actual)
    End If
End Sub
