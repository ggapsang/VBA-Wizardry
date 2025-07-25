Sub ReadTextFile()
  Dim filePath As String
  Dim fileNum As Integer
  Dim lineContent As String

  filePath = "C\PATH\FILENMAE."
  fileNum = FreeFile

  Open filePath For Input As #fileNum
  Do while Not EOF(fileNum)
    Line Input #fileNum, lineContent
    Debug.Print lineContent
  Loop
  Close #fileNum
End Sub

Function ReadTextLines(filePath As String) As Variant
    Dim fileNum As Integer
    Dim lines As Object
    Dim lineContent As String

    Set lines = CreateObject("System.Collections.ArrayList")
    fileNum = FreeFile

    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineContent
        lines.Add lineContent
    Loop
    Close #fileNum
    
    ReadTextLines = lines.ToArray() ' 배열로 변환하여 반환
End Function

Sub TestReadTextLines()
    Dim lines As Variant
    Dim i As Integer
    lines = ReadTextLines("C:\path\to\your\file.txt")
    
    For i = LBound(lines) To UBound(lines)
        Debug.Print lines(i) ' 한 줄씩 출력
    Next i
End Sub
