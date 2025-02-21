' 선택한 폴더 내의 모든 엑셀 파일 목록을 배열로 반환하는 함수
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
