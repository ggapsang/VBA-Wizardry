''' 프로시저 시작/끝에서 screenupdating 기능 등 vba 속도 구현을 위한 옵션 설정
' 최적화 시작 (화면 업데이트 중지, 자동 계산 비활성화, 경고 메시지 비활성화)
Private Sub OptimizeStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

' 최적화 종료 (원래 설정 복원)
Private Sub OptimizeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
