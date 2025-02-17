Function CreateDictionaryFromRange(rng As Range) As Object
''' input : 범위(N x M)에서 첫 번째 열을 key로, 나머지 (M-1)개 열의 값을 배열로 저장하는 딕셔너리 생성 (0부터 시작)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, j As Long
    Dim key As Variant
    Dim values() As Variant
    Dim colCount As Long

    ' 범위 데이터 읽기 (배열로 변환)
    Dim data As Variant
    data = rng.Value ' N x M 배열

    ' 열 개수 계산 (M)
    colCount = UBound(data, 2)

    ' 딕셔너리에 추가
    For i = 1 To UBound(data, 1) ' 첫 번째 행부터 마지막 행까지
        key = data(i, 1) ' 첫 번째 열(Key)

        ' 나머지 열들을 배열로 저장 (0부터 시작)
        ReDim values(0 To colCount - 2)
        For j = 2 To colCount
            values(j - 2) = data(i, j) ' 2열부터 마지막 열까지 저장
        Next j

        ' Key가 없으면 추가
        If Not dict.Exists(key) Then
            dict.Add key, values
        End If
    Next i

    Set CreateDictionaryFromRange = dict
End Function
