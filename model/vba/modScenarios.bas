Attribute VB_Name = "modScenarios"
Option Explicit

' Core scenario calculations for the ER Visit Reduction model.

Public Type ScenarioResult
    Name As String
    Reduction As Double
    VisitsAvoided As Double
    MonthlySavings As Double
    AnnualSavings As Double
End Type

Public Function CalcSingle(ByVal name As String, _
                            ByVal baselineVisits As Double, _
                            ByVal costPerVisit As Double, _
                            ByVal reduction As Double) As ScenarioResult
    Dim s As ScenarioResult
    s.Name = name
    s.Reduction = reduction
    s.VisitsAvoided = baselineVisits * reduction
    s.MonthlySavings = s.VisitsAvoided * costPerVisit
    s.AnnualSavings = s.MonthlySavings * 12#
    CalcSingle = s
End Function

Public Function CalcCombined(ByVal baselineVisits As Double, _
                              ByVal costPerVisit As Double, _
                              ByVal rA As Double, _
                              ByVal rB As Double, _
                              ByVal rC As Double) As ScenarioResult
    Dim s As ScenarioResult
    Dim rCombined As Double
    rCombined = 1# - (1# - rA) * (1# - rB) * (1# - rC)
    s.Name = "Combined"
    s.Reduction = rCombined
    s.VisitsAvoided = baselineVisits * rCombined
    s.MonthlySavings = s.VisitsAvoided * costPerVisit
    s.AnnualSavings = s.MonthlySavings * 12#
    CalcCombined = s
End Function

Public Sub WriteScenarioRow(ByVal target As Range, ByRef s As ScenarioResult)
    target.Cells(1, 1).Value = s.Name
    target.Cells(1, 2).Value = s.Reduction
    target.Cells(1, 3).Value = s.VisitsAvoided
    target.Cells(1, 4).Value = s.MonthlySavings
    target.Cells(1, 5).Value = s.AnnualSavings
End Sub
