''' 배열 안에 특정 문자열이 있으면 True, 그렇지 않으면 False를 반환
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
