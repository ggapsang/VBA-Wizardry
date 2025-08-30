Option Explicit

Sub main()

    Dim wb As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    
    Dim filePath As String
    Dim loadWb As Workbook
    Dim loadWs1 As Worksheet
    Dim loadWs2 As Worksheet
    Dim resultWb As Workbook
    Dim resultWs1 As Worksheet
    Dim resultWs2 As Worksheet
    
    Set wb = ThisWorkbook
    Set ws1 = wb.Sheets("sheet name of ws1")
    Set ws2 = wb.Sheets("sheet name of ws2")
    Set ws3 = wb.Sheets("sheet name of ws3")
    Set ws4 = wb.Sheets("sheet name of ws4")
    
    
''' winform으로 작업 대상 파일 가져오기
    
    filePath = GetFilePath()
    If filePath = "" Then
        Exit Sub
    End If
    Set loadWb = Workbooks.Open(filePath)

''' Setup resultWb
OptimizeStart

    Set resultWb = ExportWsToNewWb(loadWs1)
    loadWs2.Copy after:=resultWb.Sheets(1)
    
    Set resultWs1 = resultWb.Sheets(1)
    Set resultWs2 = resultWb.Sheets(2)

    '''''''''''''''''''''''''''''''''''''''''''
    '''''''' 실제 메인 로직은 여기에 작성 '''''''
    '''''''''''''''''''''''''''''''''''''''''''
    
OptimizeEnd

End Sub


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

''' winform을 이용하여 사용자가 특정 파일 선택
Function GetFilePath(Optional AllowMultiSelect As Boolean = False) As String
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "파일 선택"
        .AllowMultiSelect = AllowMultiSelect ' 다중 선택 가능 여부 설정
        .Filters.Clear
        .Filters.Add "모든 파일", "*.*"
        .Filters.Add "Excel 파일", "*.xls; *.xlsx; *.xlsm"
        .Filters.Add "PDF 파일", "*.pdf"

        If .Show = -1 Then ' 사용자가 파일 선택 시
            filePath = .SelectedItems(1) ' 첫 번째 파일의 경로 저장
        Else ' 취소 시 빈 값 반환
            filePath = ""
        End If
    End With
    GetFilePath = filePath ' 경로 반환
End Function

''' 워크시트를 복사해서 새 워크북에 넣음
Function ExportWsToNewWb(ws As Worksheet) As Workbook
    
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ws.Copy after:=wb.Sheets(1)
    wb.Sheets(1).Delete
    
    Application.DisplayAlerts = True
    
    Set ExportWsToNewWb = wb
    
End Function
