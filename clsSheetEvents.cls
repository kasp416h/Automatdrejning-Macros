VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ws As Worksheet
Attribute ws.VB_VarHelpID = -1

Private Sub ws_Change(ByVal Target As Range)
    Dim selectedTegningsnr As String
    Dim rowToUpdate As Long
    Dim sourceSheet As Worksheet
    Dim lookupRow As Range

    ' Set the source sheet where the data is stored (Tegningsnr, Kundenavn, etc.)
    Set sourceSheet = ThisWorkbook.Sheets("TEGNINGSNR") ' Ensure this sheet name is correct

    ' Check if the change happened in column C (where Tegningsnr is located)
    If Not Intersect(Target, ws.Range("C:C")) Is Nothing Then
        ' Ensure that the change is only in a single cell and in the correct range
        If Target.Cells.Count = 1 And Target.Row >= 1 Then
            selectedTegningsnr = Target.Value
            
            ' Skip if the selectedTegningsnr is empty (don't clear other columns)
            If Len(Trim(selectedTegningsnr)) = 0 Then
                Exit Sub
            End If

            ' Find the corresponding row in the source sheet with the selected Tegningsnr
            On Error Resume Next
            Set lookupRow = sourceSheet.Range("B:B").Find(What:=selectedTegningsnr, LookIn:=xlValues, LookAt:=xlWhole)
            On Error GoTo 0

            ' Check if the Tegningsnr was found
            If Not lookupRow Is Nothing Then
                rowToUpdate = lookupRow.Row

                ' Auto-populate adjacent columns with data from the source sheet
                Target.Offset(0, -1).Value = sourceSheet.Cells(rowToUpdate, 1).Value ' Kundenavn (Column A)
                Target.Offset(0, 1).Value = sourceSheet.Cells(rowToUpdate, 3).Value ' Tekst (Column C)
                Target.Offset(0, 3).Value = (sourceSheet.Cells(rowToUpdate, 4).Value / 3600) ' Tid s / h (Column D)
                Target.Offset(0, 7).Value = sourceSheet.Cells(rowToUpdate, 5).Value ' Opstilling (Column E)
                Target.Offset(0, 8).Value = sourceSheet.Cells(rowToUpdate, 6).Value ' Stk. pris (Column F)
            End If
        End If
    End If
End Sub

