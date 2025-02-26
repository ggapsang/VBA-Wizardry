Sub SendEmailWithAttachment(MailTo As String, attachmentFilePath As String, Optional subjectMail As String = "자동 발송된 엑셀 파일", Optional BodyMail As String = "첨부된 엑셀 파일 확인")
    Dim OutApp As Object
    Dim OutMail As Object
    
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Outlook을 열 수 없습니다. Outlook에 먼저 로그인해주세요.", vbExclamation
        Exit Sub
    End If

    ' 이메일 작성 및 전송
    With OutMail
        .To = MailTo
        .Subject = subjectMail
        .Body = subjectBodyMail
        .Attachments.Add attachmentFilePath
        .Send ' .Display로 변경하면 미리보기 가능
    End With

    ' 객체 해제
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    'MsgBox "이메일이 성공적으로 전송되었습니다.", vbInformation
End Sub
