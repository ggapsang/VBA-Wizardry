''' 모든 시트 이름을 배열로 반환
Function GetSheetNames() As Variant

    Dim ws As Worksheet
    Dim sheetNames() As String
    Dim i As Long
    
    ReDim sheetNames(1 To Worksheets.Count) ' 시트 개수만큼 배열 크기 설정

    ' 각 시트 이름을 배열에 저장
    i = 1
    For Each ws In Worksheets
        sheetNames(i) = ws.Name
        i = i + 1
    Next ws

    GetSheetNames = sheetNames ' 결과 반환 
End Function
