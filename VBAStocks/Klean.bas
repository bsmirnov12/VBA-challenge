Attribute VB_Name = "Module2"
Sub Klean()
Attribute Klean.VB_ProcData.VB_Invoke_Func = "k\n14"
'Clears the results of StatsGenerator()
    For Each ws In Worksheets
        ws.Range("I1:Q3200").Clear
    Next
End Sub
