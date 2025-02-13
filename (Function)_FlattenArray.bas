Function FlattenArray(twoDArray As Variant) As Variant
    Dim r As Long, c As Long, index As Long
    Dim rowCount As Long, colCount As Long
    Dim oneDArray() As Variant
    
    ' 2차원 배열 크기 가져오기
    rowCount = UBound(twoDArray, 1)
    colCount = UBound(twoDArray, 2)
    
    ' 1차원 배열 크기 설정
    ReDim oneDArray(1 To rowCount * colCount)
    
    ' 2차원 배열을 1차원 배열로 변환
    index = 1
    For r = 1 To rowCount
        For c = 1 To colCount
            oneDArray(index) = twoDArray(r, c)
            index = index + 1
        Next c
    Next r
    
    ' 변환된 1차원 배열 반환
    FlattenArray = oneDArray
End Function
