''' 워크시트를 pdf 파일로 저장함(한 페이지 안에 모든 값이 다 들어가도록 인쇄)
Private Sub SaveWsAsPDF(ws As Worksheet, folderPath As String)

    Dim fileName As String
    ' 파일명 설정
    fileName = folderPath & "\payment notice_" & ws.Name & ".pdf"

    ' 페이지 설정: 한 페이지에 맞춤
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    ' PDF로 내보내기
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

End Sub
