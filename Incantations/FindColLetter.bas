''' 특정 행(해더)에서 특정 문자열(칼럼 이름)을 찾고, 해당 칼럼이 있는 열 위치를 알파벳으로 반환하기
Function FindColLetter(hdr_row As Integer, search_value As Variant, Optional ws As Worksheet = Nothing) As String

    Dim search_rng As Range
    Dim found_cell As Range
    Dim col_letter As String

    ' 워크시트 변수를 설정. 기본값은 ActiveSheet
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If

    Set search_rng = ws.Rows(hdr_row)

    Set found_cell = search_rng.Find(What:=search_value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found_cell Is Nothing Then
        col_letter = Replace(found_cell.Cells.Address(False, False), hdr_row & "", "")
        FindColLetter = col_letter
    Else
        FindColLetter = ""
    End If

End Function
