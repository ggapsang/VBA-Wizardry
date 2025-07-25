Sub createPt(resultWs1 As Worksheet, sharedPivotCache As PivotCache, tableNm As String, _
             startCellAddress As String, rowName As String, _
             colName As String, valName As String, _
             Optional pageFieldName_1 As String = "", _
             Optional pageFieldName_2 As String = "", _
             Optional pageFieldName_3 As String = "", _
             Optional summaryFunction As XlPivotFieldCalculation = xlCount)

    Dim pt As PivotTable
    Dim ptRange As Range

    ' 피벗 테이블이 배치될 범위를 정의합니다.
    Set ptRange = resultWs1.Range(startCellAddress)

    Set pt = sharedPivotCache.CreatePivotTable( _
                TableDestination:=ptRange, _
                TableName:=tableNm)

    pt.HasAutoFormat = False

    ' 행 필드 추가
    If rowName <> "" Then
        With pt.PivotFields(rowName)
            .Orientation = xlRowField
            .Position = 1 ' 여러 개의 행 필드가 있을 경우 위치 조정
        End With
    End If

    ' 열 필드 추가 (선택 사항: colName이 제공되었고 비어 있지 않은 경우에만)
    If colName <> "" Then
        With pt.PivotFields(colName)
            .Orientation = xlColumnField
            .Position = 1 ' 여러 개의 열 필드가 있을 경우 위치 조정
        End With
    End If

    ' 값 필드 추가
    If valName <> "" Then
        With pt.PivotFields(valName)
            .Orientation = xlDataField
            .Function = summaryFunction ' 기본값은 갯수(xlCount). 필요에 따라 변경 (예: xlSum, xlAverage)
        End With
    End If
    
    ' 페이지(필터) 값 추가
    If pageFieldName_1 <> "" Then
        With pt.PivotFields(pageFieldName_1)
            .Orientation = xlPageField
            .Position = 1 ' 여러 개의 필터 필드가 있을 경우 위치 조정
        End With
    End If
    
    If pageFieldName_2 <> "" Then
        With pt.PivotFields(pageFieldName_2)
            .Orientation = xlPageField
            .Position = 2 ' 여러 개의 필터 필드가 있을 경우 위치 조정
        End With
    End If
    
    If pageFieldName_3 <> "" Then
        With pt.PivotFields(pageFieldName_3)
            .Orientation = xlPageField
            .Position = 3 ' 여러 개의 필터 필드가 있을 경우 위치 조정
        End With
    End If
    
    Debug.Print "Create Pivot Table : " & tableNm

End Sub
