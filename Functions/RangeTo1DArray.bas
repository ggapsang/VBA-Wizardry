``` Range를 1차원 Array로 변환
Function RangeTo1DArray(rng as Range) As Variant
    Dim arr As Variant
    Dim tempArr As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long

    arr = rng.Value

    Redim tempArr(1 To rng.Cells.Count)
    k = 1

    For i = Lbound(arr, 1) To Ubound(arr, 1)
        For j = Lbound(arr, 2) To Ubound(arr, 2)
            tempArr(k) = arr(i, j)
            k = k + 1
        Next j
    Next i

    RangeTo1DArray = tempArr

End Function
