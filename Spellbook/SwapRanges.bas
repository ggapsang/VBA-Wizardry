Sub SwapRanges()
    Dim range1 As Range
    Dim range2 As Range
    Dim tempValue As Variant
    Dim i As Integer
    Dim inputRange As Range
    Dim inputRangeCell As String

    OptimizeStart
    
    ' 사용자가 선택한 첫 번째 범위
    Set range1 = Selection
    
    ' 두 번째 범위 입력 요청
    Set inputRange = Application.InputBox("두 번째 범위를 입력하거나 선택하세요:", Type:=8)
    
    Set range2 = inputRange
    
    ' 입력받은 두 번째 범위 설정
    
    'If TypeName(inputRange) = "Range" Then
    '    Set range2 = inputRange
    'Else
        'MsgBox "올바른 범위를 선택하지 않았습니다."
        'Exit Sub
    'End If
    
    ' 두 범위의 크기 비교
    If range1.Rows.Count <> range2.Rows.Count Or range1.Columns.Count <> range2.Columns.Count Then
        MsgBox "범위의 크기가 일치하지 않습니다."
        Exit Sub
    End If
    
    ' 값 교환
    For i = 1 To range1.Rows.Count
        tempValue = range1.Cells(i, 1).value
        range1.Cells(i, 1).value = range2.Cells(i, 1).value
        range2.Cells(i, 1).value = tempValue
    Next i

    OptimizeEnd
End Sub




''' 최적화 시작 (화면 업데이트 중지, 자동 계산 비활성화, 경고 메시지 비활성화)
Private Sub OptimizeStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

''' 최적화 종료 (원래 설정 복원)
Private Sub OptimizeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
