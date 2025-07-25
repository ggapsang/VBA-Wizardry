Sub FillDown()

    Dim rngStartCell As Range
    Dim rngEndCell As Range
    Dim rngNewStartCell As Range
    
    Set rngStartCell = ActiveCell
    Set rngEndCell = rngStartCell.End(xlDown).Offset(-1, 0)
    
    rngStartCell.Copy
    
    If rngStartCell.Row < rngEndCell.Row Then
        Range(rngStartCell, rngEndCell).PasteSpecial Paste:=xlPasteAll
    End If
    
    Application.CutCopyMode = False
    rngEndCell.Select
    
    
    Set rngNewStartCell = rngEndCell.End(xlDown)
    
    rngNewStartCell.Select
    
End Sub
