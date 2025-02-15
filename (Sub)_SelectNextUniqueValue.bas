Sub SelectNextUniqueValue()
    Dim currentValue As Variant
    Dim currentCell As Range

    ' 현재 셀 값 저장
    Set currentCell = ActiveCell
    currentValue = currentCell.Value

    ' 아래로 반복하며 값 비교
    Do While Not IsEmpty(currentCell.Offset(1, 0).Value)
        Set currentCell = currentCell.Offset(1, 0)
        If currentCell.Value <> currentValue Then
            currentCell.Select
            Exit Sub
        End If
    Loop
End Sub
