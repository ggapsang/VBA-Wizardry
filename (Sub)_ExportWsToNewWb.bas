Private Sub ExportWsToNewWb(ws As Worksheet)
    
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ws.Copy after:=wb.Sheets(1)
    wb.Sheets(1).Delete
    
    wb.Activate
    Application.DisplayAlerts = True
    
End Sub
