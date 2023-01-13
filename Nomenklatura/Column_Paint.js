/** 
 * запуск процедур раскраски листа по массиву лога
 */
function column_Price_Color_Decor() {
  // вызываю из меню
  // покрасить диапазон в общий цвет
  // столбец в массив столбца
  // проходом по массиву столбца создать массив строк для покраски
  // покрасить несмежный диапазон ячеек
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_Dest = book.getSheetByName('сводная таблица');
  const sheet_Log_ = book.getSheetByName('Log изменений листов');
  let a2_Log = sheet_Log_.getDataRange().getValues();

}

function arraySheet_Columns_Remove(a2, a1_Columns) {
  return data.map(function (a2, ind) {
    return a2.filter(function (val, ind) { return !a1_Columns.includes(ind); });
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