Sub SwapRanges()
    Dim range1 As Range, range2 As Range
    Dim tempValues1 As Variant, tempValues2 As Variant
    Dim inputRange As Range

    OptimizeStart
    
    ' 사용자가 선택한 첫 번째 범위
    Set range1 = Selection
    
    ' 두 번째 범위 입력 요청
    Set inputRange = Application.InputBox("두 번째 범위를 입력하거나 선택하세요:", Type:=8)
    Set range2 = inputRange
    
    ' 두 범위의 크기 비교
    If range1.Rows.Count <> range2.Rows.Count Or range1.Columns.Count <> range2.Columns.Count Then
        MsgBox "범위의 크기가 일치하지 않습니다."
        Exit Sub
    End If
    
    ' 배열을 사용하여 값 교환 (메모리 절약 및 성능 향상)
    tempValues1 = range1.Value
    tempValues2 = range2.Value
    
    range1.Value = tempValues2
    range2.Value = tempValues1
    
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
