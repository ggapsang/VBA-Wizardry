``` 이미지 파일을 특정 셀에 삽입한다
Private Sub InsertImageAtCell(imagePath As String, targetCell As Range)
    Dim ws As Worksheet
    Dim shp As Shape
    Set ws = targetCell.Worksheet

    ' 이미지 삽입
    Set shp = ws.Pictures.Insert(imagePath)

    ' 이미지 위치를 셀과 맞추기
    With shp
        .LockAspectRatio = msoTrue ' 비율 유지
        .Left = targetCell.Left ' 셀의 왼쪽 정렬
        .Top = targetCell.Top ' 셀의 상단 정렬
        .Width = targetCell.Width ' 셀 너비와 맞춤
        .Height = targetCell.Height ' 셀 높이와 맞춤
    End With
End Sub
