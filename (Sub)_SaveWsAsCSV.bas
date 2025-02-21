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
