'''특정 폴더를 선택하면, 그 안에 있는 엑셀 파일들을 불러와서, 각 파일들의 특정 번호의 시트를 모두 csv 파일로 변환하여 저장하는 매크로

Sub Main
  OptimizeStart
  
  Dim loadWb As Workbook
  Dim loadFolderPath As String, saveFolderPath As String
  Dim fileList As Variant
  Dim i As Long

  loadFolderPath = GetFolderPath("불러올 폴더 선택")
  SaveFolderPath = GetFolderPath("저장할 폴더 선택")
  fileList = GetExcelFilesList(loadFolderPath)

  For i = 1 To Ubound(fileList)

    Dim filePath As String
    Dim ws As Worksheet

    Set loadWb = Workbooks.Open(filePath)
    Set ws = loadWb.Sheets(1) ' 시트 번호를 수정하여 일괄 저장할 시트를 선택하거나 동적으로 입력받도록 매크로를 수정할 수 있음

    SaveWsAsCSV ws, saveFolderPath
    loadWb.Close saveChanges:=False

  Next i

  OptimizeEnd
End Sub


' 최적화 시작 (화면 업데이트 중지, 자동 계산 비활성화, 경고 메시지 비활성화)
Private Sub OptimizeStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

''' winform을 이용하여 사용자가 특정 폴더 선택
Function GetFolderPath(Optional strTitle As String = "") As String
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = strTitle 'title은 적절하게 변경
        If .Show = -1 Then ' 사용자가 경로 선택 시
            folderPath = .SelectedItems(1)
        Else ' 취소 시 빈 값 반환
            folderPath = ""
        End If
    End With
    GetFolderPath = folderPath ' 반환
End Function

''' 선택한 폴더 내의 모든 엑셀 파일 목록을 배열로 반환하는 함수
Function GetExcelFilesList(folderPath) As Variant
    Dim fileName As String
    Dim fileList As Object
    Dim i As Integer
  
    ' 파일 목록을 저장할 Collection 객체 생성
    Set fileList = CreateObject("System.Collections.ArrayList")
    
    ' 폴더 내 엑셀 파일 검색
    fileName = Dir(folderPath & "\*.xls*")
    Do While fileName <> ""
        fileList.Add fileName
        fileName = Dir ' 다음 파일 검색
    Loop
    
    ' 결과를 배열로 변환 후 반환
    If fileList.Count > 0 Then
        GetExcelFilesList = fileList.ToArray()
    Else
        GetExcelFilesList = Array() ' 파일이 없을 경우 빈 배열 반환
    End If
End Function

''' 워크시트를 csv 파일로 저장함
Private Sub SaveWsAsCSV(ws As Worksheet, Optional folderPath As String = "")
    
    Dim fileName As String
    
    If folderPath = "" Then ' 폴더 경로가 지정되지 않은 경우, 사용자가 선택할 수 있도록 탐색기 창 열기
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "CSV 저장 경로 선택"
            If .Show = -1 Then ' 사용자가 경로 선택 시
                folderPath = .SelectedItems(1)
            Else ' 취소 시 종료
                Exit Sub
            End If
        End With
    End If
    
    ' 파일명 설정
    fileName = folderPath & "\" & ws.Name & ".csv"
    
    ' 워크시트를 CSV로 저장
    ws.Copy
    ActiveWorkbook.SaveAs fileName:=fileName, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close False

End Sub


' 최적화 종료 (원래 설정 복원)
Private Sub OptimizeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

