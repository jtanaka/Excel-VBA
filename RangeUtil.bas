Attribute VB_Name = "RangeUtil"

Sub �I�𒆃Z�������ɉ��s��ǉ�()
updating = Application.ScreenUpdating
Application.ScreenUpdating = False

    For Each c In Selection.Cells
        If TypeOf c Is Range And Right(c.Value, 1) <> vbLf Then
            c.Value = c.Value + vbLf
        End If
    Next

Application.ScreenUpdating = updating
End Sub

Sub �I�𒆃Z�������̉��s���폜()
updating = Application.ScreenUpdating
Application.ScreenUpdating = False

    For Each c In Selection
        If TypeOf c Is Range And Right(c.Text, 1) = vbLf Then
            c.Value = Left(c.Value, Len(c.Value) - 1)
        End If
    Next

Application.ScreenUpdating = updating
End Sub
