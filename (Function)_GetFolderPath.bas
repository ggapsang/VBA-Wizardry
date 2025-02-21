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
