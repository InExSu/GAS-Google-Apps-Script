// Переделка прайса
// из формул с кодом разово сделать копию с артикулами.
// На листе будут три таблицы:
// слева    версия для печати - цены вбиваются пользователями.
// в центре версия с артикулами.
// По этим версиям, скрипт обновления цен в листе "сводная таблица"
// создаст справа отчёт работы в третью таблицу.

function rangePriceColumnUpade() {

  const spread = SpreadsheetApp.getActive();

  const sheet_Sour = spread.getSheetByName('Прайс без НДС');
  const sheet_Dest = spread.getSheetByName('сводная таблица');
  const sheet_Logg = spread.getSheetByName('Log');

  if (headersOk(sheet_Sour, sheet_Dest) === false) {

    let warning = 'Ожидаемые заголовки НЕ совпали. \n Выход!'

    if (sheet_Dest.getName().indexOf('копия') === -1) {
      Browser.msgBox(warning);
    }
    Logger.log(warning);

  } else {

    const a2_Artics = sheet_Sour.getRange('L:Q').getValues();
    const a2_Prices = sheet_Sour.getRange('C:H').getValues();

    const a2_Column_Artics = sheet_Dest.getRange('B:B').getValues();
    const map_Artics = Array2D_2_Map(a2_Column_Artics);

    const range_Column_Prices = sheet_Dest.getRange('J:J');
    const a2_Column_Prices = range_Column_Prices.getValues();

    // копировать массив 2мерный не просто
    const a2_Column_Prices_Old = JSON.parse(JSON.stringify(a2_Column_Prices))

    A2PriceColumnUpdate(a2_Artics, a2_Prices, map_Artics, a2_Column_Prices);

    range_Column_Prices.setValues(a2_Column_Prices);

    rangePriceColumnUpade_Log(sheet_Logg, a2_Column_Artics, a2_Column_Prices_Old, a2_Column_Prices);
  }
}

function A2PriceColumnUpdate(a2_Arti_Range, a2_Price_Range, map_Arti, a2_Price_Colum) {
  // Словарь артикулов - артикул: номер строки

  // Проходом по массиву артикулов
  // 	Если артикул есть в словаре
  // 		Взять цену из массива цен в координатах артикула
  // 			Взять номер строки из словаря
  // 				Вставить в массив цен цену по номеру строки

  for (let row = 0; row < a2_Arti_Range.length; row++) {
    for (let col = 0; col < a2_Arti_Range[0].length; col++) {

      let artic = a2_Arti_Range[row][col];

      if (map_Arti.has(artic)) {

        let row_Price = map_Arti.get(artic);
        let price = a2_Price_Range[row][col];

        a2_Price_Colum[row_Price][0] = price;
      }
    }
  }
}

function rangePriceColumnUpade_Log_Test() {

  const spread = SpreadsheetApp.getActive();
  const sheet_Logg = spread.getSheetByName('Log');

  const a2_Column_Artics = [['a1'], ['a2']];
  const a2_Column_Prices_Old = [[1], [2]];
  const a2_Column_Prices_New = [[11], [22]];

  rangePriceColumnUpade_Log(sheet_Logg, a2_Column_Artics, a2_Column_Prices_Old, a2_Column_Prices_New)
}

function rangePriceColumnUpade_Log(sheet_Logg, a2_Column_Artics, a2_Column_Prices_Old, a2_Column_Prices) {

  sheet_Logg.clear();

  let cell = sheet_Logg.getRange('A1')
  let a2 = [['Лог обновления цен сводная таблица из Прайс без НДС', '', Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm:ss' мск'")]];
  array2d2Range(cell, a2)

  cell = sheet_Logg.getRange('A3')
  a2 = [['Было', 'Стало']]
  array2d2Range(cell, a2)

  a2_Column_Artics
  cell = sheet_Logg.getRange('A4');
  array2d2Range(cell, a2_Column_Artics)

  cell = sheet_Logg.getRange('B4');
  array2d2Range(cell, a2_Column_Prices_Old)

  cell = sheet_Logg.getRange('C4');
  array2d2Range(cell, a2_Column_Prices);
}

function headersOk(sheet_Sour, sheet_Dest) {
  // проверка значений ячеек

  return cell_Value(sheet_Dest.getRange('B1'), 'Артикул') &&
    cell_Value(sheet_Dest.getRange('J1'), 'Цена, руб (Без НДС)') &&
    cell_Value(sheet_Sour.getRange('C7'), 'ШМП-1') &&
    cell_Value(sheet_Sour.getRange('L7'), 'ШМП-1')

}

function cell_Value_Test() {

  const sheet = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС');
  const cell_ = sheet.getRange('C7')

  let value = 'ШМП-1';
  let resul = cell_Value(cell_, value);
  Logger.log(resul);

  value = '';
  resul = cell_Value(cell_, value);
  Logger.log(resul);
}

function cell_Value(cell, value) {
  if (cell.getValue() !== value) {
    const sheet = sheetByRange(cell);
    Logger.log('На листе ' + sheet.getName() + ' в ячейке ' +
      cell.getA1Notation() + ' !== ' + value);
    return false;
  }
  return true;
}


function A2PriceColumnUpdate_Test() {

  let a2_Arti_Range = [
    ['1', ''],
    ['3', '4']];
  let a2_Price_Range = [
    [11, 22],
    [33, 44]];

  let a2_Arti__Colum = ['5', '4', '3', '2', '1'];
  let a2_Price_Colum = [5, 4, 3, 2, 1];

  let map_Arti = Array2D_Column_2_Map(a2_Arti__Colum, 0);

  A2PriceColumnUpdate(a2_Arti_Range, a2_Price_Range, map_Arti, a2_Price_Colum);

  if (a2_Price_Colum[4][0] !== 11) {
    Logger.log('a2_Price_Colum[4][0] !== 11');
  }
}




function sheetByRange(cell) {
  // вернуть лист по диапазону
  // как в Excel range.Parent

  return sheetById(cell.getGridId());
}


function sheetById(id) {
  // вернуть лист по id

  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) {
      return s.getSheetId() === id;
    }
  )[0];
}



function price2VendorCode_Test() {

  cell = SpreadsheetApp.getActive().getSheetByName('Прайс без НДС (копия)').getRange(8, 4).getFormula();

  price2VendorCode(cell);
}

function price2VendorCode(formu) {
  // Разовый этап создания из таблицы с формулами таблицы с артикулами
  // UDF из тексту формулы извлекает код 1С,
  // по коду 1С ищет на листе строку с этим кодом,
  // если в строке есть артикул вернёт его или cell

  // нужны формулы без пробелов
  formu.replaceAll(' ', '');
  let code1 = extractBetween(formu, 'MATCH(', ";'");
  if (code1 == '') {
    code1 = formu;
  }
  else {
    //Logger.log('code1 = ' + code1);
  }
  return code1;
}

function extractBetween_Test() {

  let result = extractBetween('123', '0', '3');
  if (result !== '') {
    Logger.log('extractBetween_Test ошибка: ждал пусто, пришло ' + result);
  }

  result = extractBetween('123', '1', '3');
  if (result !== '2') {
    Logger.log('extractBetween_Test ошибка: ждал 2, пришло ' + result);
  }
  result = extractBetween('12345', '12', '45');
  if (result !== '3') {
    Logger.log('extractBetween_Test ошибка: ждал 3, пришло ' + result);
  }
}

function extractBetween(sMain, sLeft, sRigh) {
  // из строки извлечь строку между подстроками
  // InExSu 

  // добавил 1, чтобы стало возможным условие проверки на 0
  let idxBeg = sMain.indexOf(sLeft) + 1;
  let idxEnd = sMain.indexOf(sRigh) + 1;
  let strOut = '';

  if ((idxBeg * idxEnd) > 0) {
    idxBeg = idxBeg + sLeft.length - 1;
    idxEnd = idxEnd - 1;
    strOut = sMain.slice(idxBeg, idxEnd);
  }
  return strOut;
}

function Array2D_Column_2_Map(array2d, column_key) {
  // из массива 2мерного вернуть словарь - массив ассоциативный: значение столбца и номер строки
  let map_return = new Map();
  let val = '';
  for (var row = 0; row < array2d.length; row++) {
    val = String(array2d[row][column_key]);
    if (val.length > 0) {
      // если ключ повторяется, то обновится значение
      map_return.set(val, row);
    }
  }
  return map_return;
}

function Array2D_2_Map(array2d) {
  // из массива 2мерного вернуть словарь массив ассоциативный

  let map_return = new Map();
  let val = '';

  for (let row = 0; row < array2d.length; row++) {
    for (let col = 0; col < array2d[0].length; col++) {

      val = String(array2d[row][col]);

      if (val.length > 0) {
        // если ключ повторяется, то обновится значение
        map_return.set(val, row);
      }
    }
  }
  return map_return;
}

function sheetByRange(cell) {
  // вернуть лист по диапазону
  // как в Excel range.Parent

  return sheetById(cell.getGridId());
}

function sheetById(id) {
  // вернуть лист по id

  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) {
      return s.getSheetId() === id;
    }
  )[0];
}

function array2d2Range(cell, a2d) {

  // массив 2мерный вставить на лист

  let sheet_id = cell.getGridId();
  let sheet_ob = sheetById(sheet_id);
  let row_numb = cell.getRow();
  let col_numb = cell.getColumn();

  sheet_ob.getRange(row_numb, col_numb, a2d.length, a2d[0].length).setValues(a2d);
}
