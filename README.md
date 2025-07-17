# VBA-Wizardry
VBA code Snippets, patterns, useful functions <br>
5 module, 7 sub processor, 12 functions, 1 snippet (2025.3.24)

- Spellbook : modules
- Runes : sub processors
- Incantations : functions
- Cantrips : sinppets

## Read 
**Arcane_Compendium.md :** Summary of methods and usage of frequently used VBA Objects<br>


## Moduels
**1. BatchSaveWsAsCSV**
  - 특정 폴더 안에 있는 모든 엑셀파일들의 특정 번호의 시트를 csv파일로 저장함
  - 시트 번호를 시트 이름 등으로 수정하여 동적으로 변경되도록 개조 가능

**2. FillDown**
  - 아래로 빈 값 채우기
  - 선택된 셀 부터 다음 채워진 값이 나오기 전까지 모든 비어 있는 셀 값을 선택된 셀의 값으로 채움
  - 단축키 설정하고 사용하면 좋음

**3. MatrixMixer**
  - 두 개의 테이블(시트)의 값을 각 칼럼 이름에 맞게 합치되, 첫 번째 시트의 값을 1순위로, 두번째 시트의 값을 2순위로 하고, 첫 번째 시트의 특정 셀이 비었지만 두 번째 시트의 셀에 값이 있을때는 두 번째 셀의 값을 입력함

**4. SelectNextUniqueValue**
  - 현재 선택된 셀에서 열 기준으로 아래로 내려가면서, 현재 선택된 셀과 값이 다른 셀에서 멈춤
  - 정렬 후 단축키를 넣어 사용할때, 유용한 매크로

**5. SwapRange**
  - 열 기준으로 하여 두 셀의 값을 서로 바꿈

## Sub processors

**1. ExportWsToNewWb**
  - 워크시트를 복사해서 새 워크북에 넣음

**2. InsertImageAtCell**
  - 특정 셀안에 이미지를 넣음

**3. SaveWsAsPDF**
  - 워크시트를 pdf 파일로 저장함(한 페이지 안에 모든 값이 다 들어가도록 인쇄)

**4. SaveWsAsCSV**
  - 워크시트를 csv 파일로 저장함

**5. SelectNextUniqueValue**
  - 현재 선택된 셀에서 아래로 내려가면서, 현재 선택된 셀과 값이 다른 셀에서 멈춤
  - 정렬된 데이터에서 중복된 값 다음으로 이동할 때 유용

**6. SendEmailWithAttachment**
  - 아웃룩에 로그인 된 상태에서 원하는 메일 주소로 파일을 첨부하여 이메일을 전송함 

**7. VBAOptimizer**
   - 프로시저 시작/끝에서 screenupdating 기능 등 vba 속도 구현을 위한 옵션 설정
  
## Functions

**1. ColNoToLetter**
  - 열번호를 엑셀 형식에 맞는 알파벳 칼럼 번호로 바꿔줌 

**2. ConvertTo2DArray**
  - 1차원 배열을 2차원으로 변경. 2차원인 상태에서 세로로 된 Range에 바로 입력될 수 있음

**3. CreateDictionaryFromRange**
  - (N x M) Range에서 첫 번째 열을 key로, 나머지 (M-1)개 열의 값을 배열로 저장하는 딕셔너리 생성

**4. FindColLetter**
   - 특정 행(해더)에서 특정 문자열(칼럼 이름)을 찾고, 해당 칼럼이 있는 열 위치를 알파벳으로 반환하기

**5. FlattenArray**
   - 2차원 배열을 1차원으로 변환

**6. GetExcelFilesList**
  - 선택한 폴더 내의 모든 엑셀 파일 목록을 배열로 반환

**7. GetFilePath**
   - 윈폼(winform) 창을 띄우고, 폴더를 선택하게 한 뒤 해당 파일 경로를 반환

**8. GetFolderPath**
   - 윈폼(winform) 창을 띄우고, 파일을 선택하게 한 뒤 해당 파일 경로를 반환

**9. GetImageList**
  - 워크시트를 받아서 해당 워크시트 안에 있는 이미지 파일들의 배열을 반환

**10. GetSheetsNames**
  - 워크북의 모든 워크시트들의 이름을 배열로 저장

**11. GetExcelFilesListRecursive**
  - 특정 폴더 안에 있는 모든 엑셀 파일(하위 폴더안에 있는 파일 포함)들을 가져옴
  - 임시파일($)은 가져오지 않음

**12. GetUniqueValues**
  - Range를 input으로 받아 Range 안에 고유값만을 배열로 반환

**13. IsInArray**
   - 배열 안에 특정 문자열이 있으면 True, 그렇지 않으면 False를 반환
   - 
## Snippets

**1. ReadFilesByLines**
    - VBA로 텍스트 파일을 읽는 법
    - 파이썬처럼 line by line으로 읽어들이기
