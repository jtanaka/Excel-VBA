Attribute VB_Name = "RangeUtil"

Sub 選択中セル末尾に改行を追加()
updating = Application.ScreenUpdating
Application.ScreenUpdating = False

    For Each c In Selection.Cells
        If TypeOf c Is Range And Right(c.Value, 1) <> vbLf Then
            c.Value = c.Value + vbLf
        End If
    Next

Application.ScreenUpdating = updating
End Sub

Sub 選択中セル末尾の改行を削除()
updating = Application.ScreenUpdating
Application.ScreenUpdating = False

    For Each c In Selection
        If TypeOf c Is Range And Right(c.Text, 1) = vbLf Then
            c.Value = Left(c.Value, Len(c.Value) - 1)
        End If
    Next

Application.ScreenUpdating = updating
End Sub
