Attribute VB_Name = "modConfig"
Option Explicit

' Reads numeric and string config values from the Config sheet.
' Layout: column A = key, column B = value.

Public Function GetConfig(ByVal key As String) As Variant
    Dim ws As Worksheet
    Dim r As Range
    Set ws = ThisWorkbook.Worksheets("Config")
    Set r = ws.Range("A:A").Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If r Is Nothing Then
        Err.Raise vbObjectError + 1001, "modConfig.GetConfig", _
            "Config key not found: " & key
    End If
    GetConfig = r.Offset(0, 1).Value
End Function

Public Function GetConfigNum(ByVal key As String) As Double
    GetConfigNum = CDbl(GetConfig(key))
End Function
