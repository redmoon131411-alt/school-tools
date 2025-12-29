function onEdit(e) {
  // 1. 기본 설정
  var sheet = e.source.getActiveSheet(); // 현재 활성화된 시트
  var range = e.range; // 편집된 셀의 범위
  var editedCol = range.getColumn(); // 편집된 열 번호
  var editedRow = range.getRow(); // 편집된 행 번호
  var inputValue = e.value; // 입력된 값 (숫자)

  // 2. 실행 제외 조건 (오류 방지)
  // 입력값이 없거나(삭제 시), 데이터베이스 영역(K~N열)을 수정한 경우는 작동하지 않음
  // K=11, L=12, M=13, N=14
  if (!inputValue || (editedCol >= 11 && editedCol <= 14)) {
    return;
  }

  // 3. 데이터 검색 범위 설정 (K열부터 N열까지의 데이터 가져오기)
  var lastRow = sheet.getLastRow();
  // K열(11번째)부터 N열(14번째)까지의 데이터를 배열로 가져옴
  // 범위: (시작행, 시작열, 행개수, 열개수) -> (1, 11, 마지막행, 4)
  var dataRange = sheet.getRange(1, 11, lastRow, 4).getValues(); 
  
  var foundName = ""; // 찾은 이름을 저장할 변수

  // 4. 데이터 탐색 (배열 순회)
  for (var i = 0; i < dataRange.length; i++) {
    var rowData = dataRange[i];
    var kValue = rowData[0]; // K열 값 (배열 인덱스 0)
    var mValue = rowData[2]; // M열 값 (배열 인덱스 2)

    // 입력값과 K열의 값이 같으면 -> L열(인덱스 1)의 이름 가져오기
    if (kValue == inputValue) {
      foundName = rowData[1];
      break; // 찾았으면 반복 종료
    }
    
    // 입력값과 M열의 값이 같으면 -> N열(인덱스 3)의 이름 가져오기
    else if (mValue == inputValue) {
      foundName = rowData[3];
      break; // 찾았으면 반복 종료
    }
  }

  // 5. 결과 출력
  if (foundName !== "") {
    // 편집한 셀 바로 아래 행(offset(1, 0))에 이름 입력
    range.offset(1, 0).setValue(foundName);
  } else {
    // (선택사항) 찾는 번호가 없을 경우 '없음'이라고 표시하려면 아래 주석 해제
    // range.offset(1, 0).setValue("찾을 수 없음");
  }
}
