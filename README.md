# VBA-Wizardry
VBA code Snippets, patterns, useful functions
11 Functions, 5 Subprocessor (2025.2.21)


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

**9. GetSheetsNames**
  - 워크북의 모든 워크시트들의 이름을 배열로 저장

**10. GetUniqueValues**
  - Range를 input으로 받아 Range 안에 고유값만을 배열로 반환

**11. IsInArray**
   - 배열 안에 특정 문자열이 있으면 True, 그렇지 않으면 False를 반환


## Sub processors

**1. ExportWsToNewWb**
  - 워크시트를 복사해서 새 워크북에 넣음

**2. SaveWsAsPDF**
  - 워크시트를 pdf 파일로 저장함(한 페이지 안에 모든 값이 다 들어가도록 인쇄)

**3. SaveWsAsCSV**
  - 워크시트를 csv 파일로 저장함

**4. SelectNextUniqueValue**
  - 현재 선택된 셀에서 아래로 내려가면서, 현재 선택된 셀과 값이 다른 셀에서 멈춤
  - 정렬된 데이터에서 중복된 값 다음으로 이동할 때 유용

**5. VBAOptimizer**
   - 프로시저 시작/끝에서 screenupdating 기능 등 vba 속도 구현을 위한 옵션 설정
  
