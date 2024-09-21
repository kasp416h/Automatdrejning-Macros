Attribute VB_Name = "Module4"
Dim sheetEventHandlers() As New clsSheetEvents

Sub TilføjTegningsnr()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Integer
    Dim lastRow As Long
    Dim tegningsnrRange As Range
    Dim kundenavnRange As Range, tekstRange As Range, tidRange As Range, opstillingRange As Range, stkPrisRange As Range
    Dim sourceSheet As Worksheet
    Dim weeklyTabsCount As Integer
    
    ' Set the sheet where the source data is located (change if necessary)
    Set sourceSheet = ThisWorkbook.Sheets("TEGNINGSNR")
    
    ' Find the last row with data in column A of "TEGNINGSNR" sheet
    With sourceSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        ' Create references to the ranges with the relevant columns
        Set kundenavnRange = .Range("A2:A" & lastRow) ' Kundenavn
        Set tegningsnrRange = .Range("B2:B" & lastRow) ' Tegningsnr
        Set tekstRange = .Range("C2:C" & lastRow) ' Tekst
        Set tidRange = .Range("D2:D" & lastRow) ' Tid
        Set opstillingRange = .Range("E2:E" & lastRow) ' Opstilling
        Set stkPrisRange = .Range("F2:F" & lastRow) ' Stk. pris
    End With
    
    weeklyTabsCount = 52
    ReDim sheetEventHandlers(1 To weeklyTabsCount) ' Resize the array to the number of weekly tabs
    
    ' Apply data validation to each of the 52 weekly tabs
    For i = 1 To weeklyTabsCount
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(i))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' Set the range for column C from row 1 to row 100 (for Tegningsnr)
            Set rng = ws.Range("C1:C100")
            
            ' Apply data validation for Tegningsnr
            On Error Resume Next
            With rng.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="='TEGNINGSNR'!$B$2:$B$" & lastRow
            End With
            On Error GoTo 0
            
             Set sheetEventHandlers(i).ws = ws
        End If
    Next i
    
    MsgBox "Datavalidering anvendt til alle ugetabeller!"
End Sub


