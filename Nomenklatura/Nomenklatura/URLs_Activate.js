/** 
 * сделать ссылки активными
 */
function spreadSheet_Sheets_Links_Activate(sheet) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('сводная таблица');
  sheet_Links_Activate(sheet);
  cells_URLs_Format_CLIP();
}

function sheet_Links_Activate(sheet) {

  let range = sheet.getDataRange();
  let a2d = range.getValues();

  let address = '';
  let cell = '';
  let value = '';

  let spread = SpreadsheetApp.getActive();

  for (let row = 0; row < a2d.length; row++) {

    if (row % 100 === 0) {
      spread.toast('Начинаю обрабатывать строки c ' + row + ' по ' + (row + 100) + ' из ' + a2d.length);
    }

    for (let col = 0; col < a2d[0].length; col++) {

      value = a2d[row][col].toString();

      if (value.startsWith('http')) {

        address = columnToLetter(col + 1) + (row + 1);
        cell = sheet.getRange(address);
        cell_Link_Activate(cell);
      }
    }
  }
}


function cell_Link_Activate_Test(range) {
  var cell = SpreadsheetApp.getActive().getSheetByName('Лист32').getRange('B3');
  cell_Link_Activate(cell);
}

/** 
 * Сделать ячейке активную ссылку
 */
function cell_Link_Activate(cell) {

  let value = cell.getValue();

  cell.setRichTextValue(SpreadsheetApp.newRichTextValue()
    .setText(value)
    .setLinkUrl(value)
    .build());
}

function columnToLetter_Test() {
  Logger.log(columnToLetter(27));
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
