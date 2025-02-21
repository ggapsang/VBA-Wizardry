' winform을 이용하여 사용자가 특정 파일 선택
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
