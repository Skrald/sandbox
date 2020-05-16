Attribute VB_Name = "Module1"
Sub SortSheetsTabName()
    Application.ScreenUpdating = False
    Dim iSheets%, i%, j%
    iSheets = Sheets.Count
    For i = 1 To iSheets - 1
        For j = i + 1 To iSheets
            If Sheets(j).Name < Sheets(i).Name Then
                Sheets(j).Move before:=Sheets(i)
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub
