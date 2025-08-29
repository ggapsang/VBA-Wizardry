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
