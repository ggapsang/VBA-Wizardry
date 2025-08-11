Option Explicit

Sub CreateFolderAndFiles()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim baseFolderPath As String
    Dim folderPath As String
    Dim folderLevel_1 As String
    Dim folderLevel_2 As String
    Dim folderLevel_3 As String
    Dim fileName As String
    Dim newWb As Workbook
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        folderLevel_1 = ws.Cells(i, 1).Value
        folderLevel_2 = ws.Cells(i, 2).Value
        folderLevel_3 = ws.Cells(i, 3).Value
        fileName = ws.Cells(i, 4).Value '확장자 포함

        folderPath = ThisWorkbook.Path & "\" & folderLevel_1 & "\" & folderLevel_2 & "\" & folderLevel_3 'ThisWorkbook.Path 또한 변경 가능

        baseFolderPath = ThisWorkbook.Path
        
        folderPath = baseFolderPath & "\" & folderLevel_1
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder folderPath
        End If
        
        folderPath = folderPath & "\" & folderLevel_2
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder folderPath
        End If
        
        folderPath = folderPath & "\" & folderLevel_3
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder folderPath
        End If

        If fileName <> "" Then
            ''' 텍스트 파일인 경우
            'Open folderPath & "\" fileName For Output As #1 '파일 식별을 위한 고유 번호 1 ~ 511
                ' Input: 파일을 읽기 전용으로 엽니다.
                ' Output: 파일을 쓰기 전용으로 엽니다. 파일이 이미 존재하면 내용이 지워지고, 없으면 새로 생성됩니다.
                ' Append: 파일을 추가 모드로 엽니다. 파일이 이미 존재하면 내용의 끝에 데이터가 추가되고, 없으면 새로 생성됩니다.
                ' Random: 임의 접근 모드로 파일을 엽니다.
                ' Binary: 이진 모드로 파일을 엽니다.
                ' Print #1
                ' Close #1
            ''' 엑셀 파일인 경우
            Set newWb = Workbooks.Add
            newWb.SaveAs fileName:=folderPath & "\" & fileName
            newWb.Close SaveChanges:=False
        End If
        
    Next i
    
End Sub
