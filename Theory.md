## Application Object
<br>

```
Sub Application_Object
    Dim strInfo As String
    strInfo = Application.UserName
    strInfo = Application.Version
    strInfo = Application.Path
    strInfo = Application.StartPath
    strInfo = Application.FullName
End Sub
```

- `.UserName` = 사용자 이름
- `.Verstion` = 현재 설치된 엑셀 버전
- `.Path` = 엑셀이 설치되어 있는 폴더 경로
- `.StartupPath` = 시작 폴더의 위치
- `.FullName` = 통합 문서의 경로, 파일명, 확장자를 모두 표시함 (경로만 표시할 때는 Path 사용)
<br>

```
Sub StatusBar
    Dim dbCount As Double
    Application.DisplayStatusBar = True

    For dbCount = 1 To 1000
        Application.StatusBar = dbCount & " / 1000"
    Next dbCount

    Application.StatusBar = ""
End Sub
```

- `.StatusBar` = 상태 표시줄 진행창 변경
- `.DisplayScrollBars` = 스크롤 바를 숨기거나 보이게 함
<br>

## Workbooks Object

### Object Model 계층 구조
- 엑셀(`Application`) <br>
- 워크북스(`Workbooks`) : Collection Object <br> 
- **워크북(`Workbook`)** <br>
- 워크시트(`Worksheet`)<br>
- 레인지(`Range`)
<br>
### Workbooks 개체의 주요 매서드
- `Workbooks.Add` : 새 워크북(새 엑셀 파일-통합 문서) 생성
- `Workbooks.Open Filename:= "file_name.xlsx"` : 워크북 열기
- `Workbooks("file_name.xlsx").Close` : 워크북 닫기
- `Workbooks.Close` : 모든 워크북 닫기
    ※ `ActiveWorkbook.Close` : (x) 사용 불가능한 방식
    - SaveChanges 인수를 통해 저장 유무를 선택할 수 있음(메세지 상자가 나타나지 않게 함)
        - ex : `ActiveWorkbook.Close savechanges:=True` 또는 `ActiveWorkbook.Close True`
- `Workbooks("file_name.xlsx").Save` : 워크북 저장
- `ActiveWorkbook.Save` 가능함
<br>
<br>
<br>

## Worksheets Object

### Wokrsheet 개체의 주요 매서드
- `Worksheets.Add` : 새 워크시트 삽입(현재 시트의 앞쪽(왼쪽)에 새로운 워크시트가 삽입됨)
    - 정해진 위치에 워크시트를 삽입하기 위해서는
      - Worksheets.Add after:=Worksheets(1), Count:=2 (삽입 위치: 첫 번째 워크시트 뒤, 삽입할 워크시트의 개수 : count)
- `Worksheet.Activate` : 워크시트 활성화
- `Worksheets.Select` : 워크북 내의 모든 시트 선택
- `Worksheets(Array(1, 3, 5)).Select : 서로 인접해 있지 않은 시트를 선택할 때는 배열을 사용함
- `Wokrsheets(2).Name = '시트이름'` : 시트 이름 변경
- `Worksheets("Sheet1").Move after:=Worksheets("Sheet3") : 시트 이동
- `Worksheets("Sheet1").Copy after:=Worksheets(2) : 시트를 복사(2번째 시트 뒤에 놓음)




