Attribute VB_Name = "modDashboard"
Option Explicit

' Entry point invoked by the "Run Analysis" button on the Dashboard sheet.
' Reads inputs from Config, calculates each scenario, writes results.

Public Sub RunAnalysis()
    Dim baseline As Double, cost As Double
    Dim rA As Double, rB As Double, rC As Double
    Dim wsResults As Worksheet

    On Error GoTo Fail

    baseline = modConfig.GetConfigNum("baseline_visits_per_month")
    cost = modConfig.GetConfigNum("cost_per_visit")
    rA = modConfig.GetConfigNum("reduction_A")
    rB = modConfig.GetConfigNum("reduction_B")
    rC = modConfig.GetConfigNum("reduction_C")

    Set wsResults = ThisWorkbook.Worksheets("Dashboard")
    wsResults.Range("A10:E14").ClearContents
    wsResults.Range("A10:E10").Value = Array("Scenario", "Reduction", "Visits Avoided", "Monthly Savings", "Annual Savings")

    Dim sA As ScenarioResult, sB As ScenarioResult, sC As ScenarioResult, sCmb As ScenarioResult
    sA = modScenarios.CalcSingle("A: Telehealth triage", baseline, cost, rA)
    sB = modScenarios.CalcSingle("B: Care coordination", baseline, cost, rB)
    sC = modScenarios.CalcSingle("C: Patient education", baseline, cost, rC)
    sCmb = modScenarios.CalcCombined(baseline, cost, rA, rB, rC)

    modScenarios.WriteScenarioRow wsResults.Range("A11"), sA
    modScenarios.WriteScenarioRow wsResults.Range("A12"), sB
    modScenarios.WriteScenarioRow wsResults.Range("A13"), sC
    modScenarios.WriteScenarioRow wsResults.Range("A14"), sCmb

    wsResults.Range("B11:B14").NumberFormat = "0.0%"
    wsResults.Range("D11:E14").NumberFormat = "$#,##0"

    MsgBox "Analysis complete.", vbInformation
    Exit Sub
Fail:
    MsgBox "RunAnalysis failed: " & Err.Description, vbCritical
End Sub

Public Sub RunSensitivity()
    Dim wsSens As Worksheet
    Dim baseline As Double, cost As Double
    Dim rA As Double, rB As Double, rC As Double
    Dim deltas As Variant, i As Long, row As Long

    baseline = modConfig.GetConfigNum("baseline_visits_per_month")
    cost = modConfig.GetConfigNum("cost_per_visit")
    rA = modConfig.GetConfigNum("reduction_A")
    rB = modConfig.GetConfigNum("reduction_B")
    rC = modConfig.GetConfigNum("reduction_C")

    Set wsSens = ThisWorkbook.Worksheets("Sensitivity")
    wsSens.Cells.ClearContents
    wsSens.Range("A1:D1").Value = Array("Parameter", "Delta", "Combined Reduction", "Annual Savings")

    deltas = Array(-0.05, -0.025, 0#, 0.025, 0.05)
    row = 2
    For i = LBound(deltas) To UBound(deltas)
        Dim s As ScenarioResult
        s = modScenarios.CalcCombined(baseline, cost, rA + deltas(i), rB, rC)
        wsSens.Cells(row, 1).Value = "reduction_A"
        wsSens.Cells(row, 2).Value = deltas(i)
        wsSens.Cells(row, 3).Value = s.Reduction
        wsSens.Cells(row, 4).Value = s.AnnualSavings
        row = row + 1
    Next i

    wsSens.Range("B2:C" & row).NumberFormat = "0.0%"
    wsSens.Range("D2:D" & row).NumberFormat = "$#,##0"
End Sub
