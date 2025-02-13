Function ConvertTo2DArray(arr As Variant) As Variant
''' input : 1차원 array
''' return : 2차원 array. 세로로 된 range에 바로 들어갈 수 있음

    Dim i As Long
    Dim tempArray() As Variant
    ReDim tempArray(LBound(arr) To UBound(arr), 1 To 1)
    
    For i = LBound(arr) To UBound(arr)
        tempArray(i, 1) = arr(i)
    Next i
    
    ConvertTo2DArray = tempArray
End Function
