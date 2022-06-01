function myFunction() {
  // スプレッドシートのインスタンスを設定する。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let num_sheet = ss.getNumSheets() - 1; // 対象のシートの枚数を取得
  
  // 諸々の変数を設定する。
  const columnOfValue = COLUMN_INDEX_TO_READ_VALUE; // 読み取る列のインデックス
  let rowOfValue = ROW_INDEX_TO_READ_VALUE; // 「他の人に渡す単語」の行インデックス
  let arrOfSubmittedValues = []; // 入力した値
  
  for (let i = 0; i < num_sheet; i++) {
    arrOfSubmittedValues.push(readSubmitValue(ss.getSheets()[i], columnOfValue, rowOfValue));
  }
  arrOfSubmittedValues.push(arrOfSubmittedValues[0]);
  arrOfSubmittedValues.splice(0, 1);
  
  for (let i = 0; i < num_sheet; i++) {
    inputValue(ss.getSheets()[i], columnOfValue, arrOfSubmittedValues, num_sheet);
    
  }
}

function readSubmitValue(sheet, columnOfSubmittedValue, rowOfSubmittedValue) {
  value_submit = sheet.getRange(rowOfSubmittedValue, columnOfSubmittedValue).getValue();
  return value_submit;
}

function inputValue(sheet, columnIndexOfValue, arrOfSubmittedValues, numberOfSheet) {
  let rowIndexOfBaseOfPlayer = ROW_INDEX_OF_BASE_OF_PLAYERS; // 単語を格納する行の最上の行インデックス
  for (let i = 0; i < numberOfSheet; i++) {
    // シート名と同じプレイヤー名の行に書き込まない。
    if (sheet.getSheetName() != sheet.getRange(rowIndexOfBaseOfPlayer + i,columnIndexOfValue - 1).getValue()) {
      sheet.getRange(rowIndexOfBaseOfPlayer + i, columnIndexOfValue).setValue(arrOfSubmittedValues[i]);
      sheet.getRange(rowIndexOfBaseOfPlayer + i, columnIndexOfValue).setBackground(KNOWN_CELL_COLOR);
    }
    else {
      sheet.getRange(rowIndexOfBaseOfPlayer + i, columnIndexOfValue).setValue(UNKNOWN_CELL_VALUE);
      sheet.getRange(rowIndexOfBaseOfPlayer + i, columnIndexOfValue).setBackground(UNKNOWN_CELL_COLOR);
    }
    
  }
}
