'실제로 사용했던 고급필터 매크로

Sub View()
    
    Dim wb As Workbook
    Dim wsView As Worksheet
    Dim wsFilter As Worksheet
    Dim wsSrc As Worksheet
    
    ClearResult
    ClearFilter
    
    Set wb = ThisWorkbook
    Set wsView = wb.Sheets("매물검색")
    Set wsFilter = wb.Sheets("필터조건")
    Set wsSrc = wb.Sheets("매물데이터정비리스트")
    
        
    '==============================================================
    ' 1) OR 조건 값 읽기 (분류 / 타입)
    '==============================================================
    Dim colCategory As New Collection
    Dim colType As New Collection
    Dim i As Long
    
    ' 분류 체크박스 (예: B5 = 주거용, B10 = 상업용)
    If wsView.Range("B5").Value = True Then colCategory.Add "주거용"
    If wsView.Range("B10").Value = True Then colCategory.Add "상업용"
    
    ' 타입 체크박스 10개 (E2:E11 체크, 실제 타입명은 B2:B11)
    For i = 2 To 11
        If wsView.Range("E" & i).Value = True Then
            colType.Add wsView.Range("B" & i).Value
        End If
    Next i

     '==============================================================
    ' 2) AND 조건 읽기 → wsView에서 읽음
    '==============================================================

    Dim minBang, maxBang
    minBang = wsView.Range("H2").Value

    Dim parking, pet
    parking = wsView.Range("H3").Value
    pet = wsView.Range("H4").Value

    Dim minSale, maxSale
    Dim minJeonse, maxJeonse
    Dim minWolse, maxWolse
    Dim minDeposit, maxDeposit
    Dim minPyeong, maxPyeong
    Dim minM2, maxM2

    minSale = wsView.Range("H7").Value
    maxSale = wsView.Range("I7").Value
    minJeonse = wsView.Range("H8").Value
    maxJeonse = wsView.Range("I8").Value
    minWolse = wsView.Range("H9").Value
    maxWolse = wsView.Range("I9").Value
    minDeposit = wsView.Range("H10").Value
    maxDeposit = wsView.Range("I10").Value
    minPyeong = wsView.Range("H11").Value
    maxPyeong = wsView.Range("I11").Value
    minM2 = wsView.Range("H12").Value
    maxM2 = wsView.Range("I12").Value

    Dim moveinDate, phoneTail, addr
    moveinDate = wsView.Range("H13").Value
    addr = wsView.Range("D13").Value
    phoneTail = wsView.Range("D14").Value

    '==============================================================
    ' 3) OR 조건 개수 결정
    '==============================================================
    Dim orCount As Long

    orCount = Application.WorksheetFunction.Max( _
                 IIf(colCategory.Count = 0, 1, colCategory.Count), _
                 IIf(colType.Count = 0, 1, colType.Count))
    
    If orCount < 1 Then orCount = 1


    '==============================================================
    ' 4) 조건 테이블 생성
    '==============================================================
    Dim r As Long: r = 2
   
    For i = 1 To orCount
        
        '-------------------------
        ' OR 조건: 분류
        '-------------------------
        If i <= colCategory.Count Then
            wsFilter.Cells(r, 1).Value = colCategory(i)
        Else
            wsFilter.Cells(r, 1).Value = ""
        End If
        
        '-------------------------
        ' OR 조건: 타입
        '-------------------------
        If i <= colType.Count Then
            wsFilter.Cells(r, 2).Value = colType(i)
        Else
            wsFilter.Cells(r, 2).Value = ""
        End If
        
        '-------------------------
        ' AND 조건 (모든 행 동일)
        '-------------------------
        
        ' 방
        If minBang <> "" Then wsFilter.Cells(r, 3).Value = ">=" & minBang

        ' 주차
        'If parking <> "" Then wsFilter.Cells(r, 4).Value = "=" & parking
        If parking = True Then wsFilter.Cells(r, 4).Value = "=" & parking
        
        ' 애완불가
        'If pet <> "" Then wsFilter.Cells(r, 5).Value = "=" & pet
        If pet = True Then wsFilter.Cells(r, 5).Value = "=" & pet
        
        ' 매매
        If minSale <> "" Then wsFilter.Cells(r, 6).Value = ">=" & minSale
        If maxSale <> "" Then wsFilter.Cells(r, 7).Value = "<=" & maxSale
        
        ' 전세
        If minJeonse <> "" Then wsFilter.Cells(r, 8).Value = ">=" & minJeonse
        If maxJeonse <> "" Then wsFilter.Cells(r, 9).Value = "<=" & maxJeonse
        
        ' 월세
        If minWolse <> "" Then wsFilter.Cells(r, 10).Value = ">=" & minWolse
        If maxWolse <> "" Then wsFilter.Cells(r, 11).Value = "<=" & maxWolse
        
        ' 보증금
        If minDeposit <> "" Then wsFilter.Cells(r, 12).Value = ">=" & minDeposit
        If maxDeposit <> "" Then wsFilter.Cells(r, 13).Value = "<=" & maxDeposit
        
        ' 면적(평)
        If minPyeong <> "" Then wsFilter.Cells(r, 14).Value = ">=" & minPyeong
        If maxPyeong <> "" Then wsFilter.Cells(r, 15).Value = "<=" & maxPyeong
        
        ' 면적(m2)
        If minM2 <> "" Then wsFilter.Cells(r, 16).Value = ">=" & minM2
        If maxM2 <> "" Then wsFilter.Cells(r, 17).Value = "<=" & maxM2
        
        ' 연락처(일치)
        If phoneTail <> "" Then wsFilter.Cells(r, 18).Value = "=" & phoneTail
        
        ' 입주일
        If moveinDate <> "" Then wsFilter.Cells(r, 19).Value = ">=" & moveinDate
        
        ' 주소(일치)
        If addr <> "" Then wsFilter.Cells(r, 20).Value = addr
        
        r = r + 1
    Next i

    '==============================================================
    ' 5) 고급필터 실행/정렬
    '==============================================================

    Dim src As Range
    Dim crit As Range
    Dim dest As Range
    
    Set src = wsSrc.Range("A1").CurrentRegion
    Set crit = wsFilter.Range("A1:T" & (orCount + 1))
    Set dest = wsView.Range("A18:AE18")
    
    src.AdvancedFilter Action:=xlFilterCopy, _
                   CriteriaRange:=crit, _
                   CopyToRange:=dest, _
                   Unique:=False
    
    Dim rngResult As Range
    
    Set rngResult = wsView.Range("A18").CurrentRegion
    
    
    rngResult.Sort Key1:=rngResult.Columns(2), _
                   Order1:=xlDescending, _
                   Header:=xlYes

End Sub


Sub clear()

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("매물검색")
    
    ws.Range("B2:B11").ClearContents
    ws.Range("E2:E11").ClearContents
    ws.Range("D13:D14").ClearContents
    ws.Range("H2:H4").ClearContents
    ws.Range("H7:I12").ClearContents
    ws.Range("H13").ClearContents
    
End Sub

Sub ClearResult()

    Sheets("매물검색").Range("A19:AE10000").clear

End Sub

Sub ClearFilter()

    Sheets("필터조건").Range("A2:T1000").clear

End Sub

