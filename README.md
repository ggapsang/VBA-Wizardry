# VBA-Wizardry
VBA code Snippets, patterns, useful functions


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

**6. GetFilePath**
   - 윈폼(winform) 창을 띄우고, 폴더를 선택하게 한 뒤 해당 파일 경로를 return함

**7. GetFolderPath**
   - 윈폼(winform) 창을 띄우고, 파일을 선택하게 한 뒤 해당 파일 경로를 return함

**8. GetUniqueValues**
  - Range를 input으로 받아 Range 안에 고유값만을 배열로 return

**9. IsInArray**
   - 배열 안에 특정 문자열이 있으면 True, 그렇지 않으면 False를 반환함


## Sub processors

1. SaveSheetAsPDF
   - 워크시트를 pdf 파일로 저장함(한 페이지 안에 모든 값이 다 들어가도록 인쇄)
  
2. VBAOptimizer
   - 프로시저 시작/끝에서 screenupdating 기능 등 vba 속도 구현을 위한 옵션 설정
  
