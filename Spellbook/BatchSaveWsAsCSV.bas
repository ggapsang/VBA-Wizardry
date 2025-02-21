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

private Sub OptimzeStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculation
End Sub
