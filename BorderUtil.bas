Sub 罫線外枠のみ()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Sub

Sub 罫線内横線削除()

    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Sub

Sub 罫線内横線点線()

    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub

