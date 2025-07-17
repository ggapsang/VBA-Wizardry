Function GetExcelFilesListRecursive(folderPath As String) As Variant
    Dim fso As Object ' FileSystemObject를 사용하기 위한 변수
    Dim folder As Object ' 현재 폴더 객체
    Dim file As Object ' 파일 객체
    Dim subFolder As Object ' 하위 폴더 객체

    Dim fileList As New Collection ' 파일 경로를 저장할 컬렉션

    Set fso = CreateObject("Scripting.FileSystemObject") ' FileSystemObject 생성
    Set folder = fso.GetFolder(folderPath) ' 지정된 폴더 가져오기

    ' 현재 폴더의 엑셀 파일 추가
    For Each file In folder.Files
        ' 임시 파일 (~)로 시작하는 파일) 및 엑셀 파일 확장자 확인
        If Left(file.Name, 1) <> "~" And _
           (LCase(fso.GetExtensionName(file.Name)) = "xls" Or _
            LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or _
            LCase(fso.GetExtensionName(file.Name)) = "xlsm" Or _
            LCase(fso.GetExtensionName(file.Name)) = "xlsb") Then
            fileList.Add file.Path
        End If
    Next file

    ' 하위 폴더 탐색 (재귀 호출)
    For Each subFolder In folder.SubFolders
        Dim subFolderFiles As Variant
        subFolderFiles = GetExcelFilesListRecursive(subFolder.Path) ' 재귀적으로 하위 폴더 탐색

        ' 하위 폴더에서 찾은 파일들을 현재 컬렉션에 추가
        If IsArray(subFolderFiles) Then
            Dim i As Long
            For i = LBound(subFolderFiles) To UBound(subFolderFiles)
                fileList.Add subFolderFiles(i)
            Next i
        End If
    Next subFolder

    ' 컬렉션을 배열로 변환하여 반환
    If fileList.Count > 0 Then
        Dim result() As String
        ReDim result(0 To fileList.Count - 1)
        Dim j As Long
        For j = 0 To fileList.Count - 1
            result(j) = fileList(j + 1) ' 컬렉션은 1부터 시작하므로 j+1
        Next j
        GetExcelFilesListRecursive = result
    Else
        GetExcelFilesListRecursive = Array() ' 파일이 없으면 빈 배열 반환
    End If

    Set fso = Nothing
    Set folder = Nothing
End Function
