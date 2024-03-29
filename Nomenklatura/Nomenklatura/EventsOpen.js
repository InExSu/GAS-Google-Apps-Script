function onEdit(event) {
  //Возникает при изменении ячейки
  // Logger.log('onEdit');

  const sheetName = event.source.getActiveSheet().getName();// лист события

  const col = event.range.getColumn();  //Номер столбца

  if (sheetName == 'сводная таблица') {

    const col_SKU_2 = 2, col_Price_10 = 10;

    if (col == col_SKU_2 || col == col_Price_10) {

      priceFixAdd(event, 'Log изменений листов');
      column_Price_Paint_SKU('Log изменений листов', sheetName);
    }
    Logger.log('Лист ' + sheetName + ', столбец ' + col)
  }
}


/**
 * добавить запись в низ листа
 */
function priceFixAdd(event, sheet_Log_Name) {

  const sheet_Log = SpreadsheetApp.getActive().getSheetByName(sheet_Log_Name);
  const sheet_Source = event.source.getActiveSheet();// лист события
  const row = event.range.getRow();      //Номер строки
  const sku = sheet_Source.getRange(row, 2).getValue();

  // без артикула не фиксирую

  if (sku.match(/^\d{3}-\d{3}-\d{4}$/) != null) {

    const col = event.range.getColumn();  //Номер столбца
    const user = event.user;

    const value_New = event.value;            //Новое значение
    const value_Old = event.oldValue;        //Старое значение

    if (value_Old != value_New) {
      // Дата	Лист_Имя	Строка	Столбец	Было	Стало Пользователь
      // создать массив, 
      const date_Formatted = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss");
      const sheet_Source_Name = sheet_Source.getName();
      // массив-строка
      const a2 = [[date_Formatted, sheet_Source_Name, row, col, value_Old, value_New, user, sku]];
      // ячейку последнюю найти
      const log_Row = sheet_Log.getDataRange().getLastRow() + 1;
      // вставить массив на лист
      sheet_Log.getRange(log_Row, 1, a2.length, a2[0].length).setValues(a2);
    }
  }
}

function sku_parse_Test() {
  var regExp = /^\d{3}-\d{3}-\d{4}$/
  Logger.log("101-011-0003".match(regExp));
  Logger.log("101-011-000".match(regExp));
  Logger.log("нет".match(regExp));
  Logger.log("".match(regExp));
}

function selectionDuplicates() {
  // найти строки различающиеся ростами и если разные цены - сообщить пользователю
  var a2 = SpreadsheetApp.getActiveSpreadsheet().getSelection().getActiveRange().getValues();
  // в одномерный массив
  a2 = a2.flat(Infinity);

  //console.log(a2);

  var duplicates = [];

  /* отсортировать массив, а затем проверить, совпадает ли «следующий элемент» с текущим элементом, и поместить его в массив: */
  var tempArray = [...a2].sort();
  //console.log(tempArray);

  for (let i = 0; i < tempArray.length; i++) {
    if (tempArray[i + 1] === tempArray[i]) {
      duplicates.push(tempArray[i]);
    }
  }

  // массив оставляю уникальные
  duplicates = duplicates.filter(onlyUnique);

  Browser.msgBox(duplicates);
  //console.log(duplicates);

}

function onlyUnique(value, index, self) {
  //проверяет, является ли данное значение первым встречающимся. Если нет, то это дубликат и не будет скопирован.
  return self.indexOf(value) === index;
}

function onOpen() {

  var ui = SpreadsheetApp.getUi();  // Or DocumentApp or FormApp.

  ui.createMenu('Прайсы')

    .addItem('Обрамить', 'formulaCodeFind')

    .addItem('Дубликаты', 'selectionDuplicates')

    .addItem('Создать копию книги', 'spreadsheetCopy')

    .addItem('Ссылки сделать активными на листе сводная таблица', 'spreadSheet_Sheets_Links_Activate')

    // .addItem('Нули формат', 'selectionNullFormatted')

    // .addSeparator()

    // .addSubMenu(ui.createMenu('Sub-menu')

    //   .addItem('Тест', 'sheetActive'))

    .addToUi();

}

function formulaCodeFind() {

  // ячейки выделенные обрамить слева и справа формулой
  const column = columnBySheet();
  if (column === undefined) { return; }

  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  var rowsCount = range.getNumRows();
  var colsCount = range.getNumColumns();

  var cell;
  var cellValue;
  var formula;
  for (var row = 1; row <= rowsCount; row++) {
    for (var col = 1; col <= colsCount; col++) {

      cell = range.getCell(row, col);

      formula = cell.getFormula()
      if (formula != "") {
        return;
      }

      cellValue = cell.getValue();

      if (cellValue == '') {
        return;
      }

      if (!IsNumeric(cellValue)) {

        // нечисла добавить кавычки
        cellValue = '"' + cellValue + '"';
      }

      cellValue = "=IFError(Index('сводная таблица'!" + column + ";MATCH("
        + cellValue + ";'сводная таблица'!$A:$A;0);1);\"код НЕ найден\")";

      cell.setValue(cellValue);

    }
  }
}


function menuItem2() {

  SpreadsheetApp.getUi().alert('You clicked the second menu item!');

  // DocumentApp.getUi().alert('You clicked the second menu item!'); - for DocumentApp

}

function IsNumeric(stringIN) {
  return isFinite(parseFloat(stringIN));
}


function columnBySheet() {
  // столбец в зависимости от имени листа
  const sheetName = SpreadsheetApp.getActiveSheet().getName();
  if (sheetName == "Прайс без НДС") { return "$J:$J"; }
  if (sheetName == "Прайс с НДС") { return "$L:$L"; }
  if (sheetName == "Прайс партнеры без НДС") { return "$M:$M"; }
  if (sheetName == "Прайс партнеры c НДС") { return "$N:$N"; }
  if (sheetName == "Прайс СНГ") { return "$O:$O"; }
  if (sheetName == "Прайс СНГ партнеры") { return "$P:$P"; }
}

function sheetActive() {
  Browser.msgBox(SpreadsheetApp.getActiveSheet().getName())
}

