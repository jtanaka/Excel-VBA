Sub 全シートをA1選択状態にする()
updating = Application.ScreenUpdating
Application.ScreenUpdating = False
    If ActiveWorkbook.Sheets.Count = 0 Then
        Exit Sub
    End If

    For Each s In ActiveWorkbook.Sheets
        s.Activate
        s.Cells(1, 1).Select
           
        With ActiveWindow
            .ScrollRow = ActiveCell.row
            .ScrollColumn = ActiveCell.Column
        End With
    Next
    ActiveWorkbook.Sheets(1).Activate

Application.ScreenUpdating = updating
End Sub
