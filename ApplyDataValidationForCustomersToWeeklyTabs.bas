Attribute VB_Name = "Module5"
Sub TilføjKunder()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Integer
    Dim customerNamesRange As String
    
    ' Find the last row with data in column A of "KUNDER" sheet
    With ThisWorkbook.Sheets("KUNDER")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        ' Create a reference to the range with customer names
        customerNamesRange = "'KUNDER'!$A$1:$A$" & lastRow
    End With
    
    ' Apply data validation to each of the 52 weekly tabs
    For i = 1 To 52
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' Set the range for column E from row 1 to row 100
            Set rng = ws.Range("B1:B100")
            
            ' Apply data validation
            On Error Resume Next
            With rng.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=" & customerNamesRange
            End With
            On Error GoTo 0
        End If
        Set ws = Nothing
    Next i
    
    MsgBox "Datavalidering af kunder anvendt til alle ugetabeller!"
End Sub
