function column_Price_Paint_SKU(event) {
  // лист лога в массив
  // покрасить столбец цен в общий цвет
  // проходом по массиву столбца артикулов 
  // фильтровать массив лога по артикулу и дате
  // если массив непуст - красить ячейку цены

  const book = SpreadsheetApp.getActive();

  const sheet_Log = book.getSheetByName('Log изменений листов');
  const sheet_SKU = book.getSheetByName('сводная таблица');

  // Столбец цен сбрасываю цвет на по умолчанию
  let range_Column = sheet_SKU.getRange('J:J');
  range_Column.setBackground('#70ad47');

  const a2_SKU = sheet_SKU.getRange('B:B').getValues();
  let a2_Log = sheet_Log.getDataRange().getValues();

  // массив лога - строки с датами ненужными удаляю
  date_Left = date_Create(-7, 0, 0);
  a2_Log = a2_Log.filter(function (row) {
    return row[0] > date_Left;
  })

  // столбец массива 2мерного
  const arrayColumn = (arr, n) => arr.map(x => x[n]);
  const column_SKU_7 = 7;
  const column_Price_10 = 10;

  // Крашу ячейки, если артикул есть в столбце массива
  for (let row = 0; row < a2_SKU.length; row++) {

    const sku = a2_SKU[row][0];

    if (sku.match(/^\d{3}-\d{3}-\d{4}$/) != null) {

      if (arrayColumn(a2_Log, column_SKU_7).includes(sku)) {

        sheet_SKU.getRange(row + 1, column_Price_10).setBackground('#FF9C9C');
      }
    }
  }

  // cell_Price_Paint_If(
  //   array_filter_SKU(
  //     array_filter_Date()));


}


/** 
 * вывести фон ячейки
 */
function cell_Interior_Color() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('сводная таблица');
  var range = sheet.getRange('J1');
  var background = range.getBackground();
  Logger.log(background);
}


function date_Create(days, months, years) {
  var date = new Date();
  date.setDate(date.getDate() + days);
  date.setMonth(date.getMonth() + months);
  date.setFullYear(date.getFullYear() + years);
  return date;
}

function arraySheet_Columns_Remove_Test() {
  let a2 = [
    ['0', '1', '', '3', ''],
    ['2', '2', '', '', ''],
    ['3', '1', '', '', ''],
    ['1', '1', '', '1', ''],
    ['', '', '', '2', '']
  ];

  a1 = [1, 3];
  a2 = arraySheet_Columns_Remove(a2, a1);
  Logger.log(a2);
}

/**
 * в массиве 2мерном удалить столбцы по номерам
 */
function arraySheet_Columns_Remove(a2, a1_Columns) {
  return a2.map(function (a2, ind) {
    return a2.filter(function (a2, ind) { return !a1_Columns.includes(ind); });
  });
}


function test() {
  var data = [
    ['1', '2', '', '3', ''],
    ['2', '2', '', '', ''],
    ['3', '1', '', '', ''],
    ['1', '1', '', '1', ''],
    ['', '', '', '2', '']
  ];
  var countColumns = data[0].length; // Ну или как-то по другому посчитаете
  var removes = []; // Список номеров пустых колонок
  for (var i = 0; i < countColumns; i++) {
    var isRemove = data.every(function (val) { return val[i].length === 0; });
    if (isRemove) {
      removes.push(i); // Добавляем номер колонки
    }
  }
  // Удаляем колонки и получаем новый массив newData
  var newData = data.map(function (val, ind) {
    return val.filter(function (val, ind) { return !removes.includes(ind); });
  });
  console.warn(x);
}