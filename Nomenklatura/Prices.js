// Переделка прайса
// из формул с кодом сделать копию с артикулами.
// На листе будут три таблицы:
// слева    версия для печати - цены вбиваются пользователями.
// в центре версия с артикулами.
// По этим версиям, скрипт обновления цен в листе "сводная таблица"
// создаст справа отчёт работы в третью таблицу.

function priceArticoolPivot_RUN() {
  // вызываю по кнопке на листе
  let sheet = SpreadsheetApp.getActive().getActiveSheet();

  // Убеждаюсь, что диапазон артикулов на месте
  if (priceRangeArticoolsCheck(sheet)) {

  } else {
    Browser.msgBox('Диапазон артикулов не похож на ожидаемый');
  }
}

function priceRangeArticoolsCheck(sheet) {
  let value = sheet.getRange("S2").getValues();
  if (value !== 'ШМП-1') { return false };

  value = sheet.getRange("S3").getValues();
  if (value !== 'МАГ-4') { return false };

  return true

}


function priceArticoolPivot() {
  // Скрипт получит на вход:
  // - диапазон прайса с артикулами
  // - диапазон прайса с ценами
  // - столбец артикулов "сводная таблица"
  // - столбец цен            "сводная таблица"

  // Проходом по массиву диапазона прайса с артикулами,
  // ищет значение ячейки в массиве артикулов "сводная таблица".
  // Если находит, то берёт значение из массива цен и
  // вставляет в массив цен "сводная таблица"
  // По окончании вставляет массив цен "сводная таблица" на лист.


}


function rangeCtrlShiftDown_Test() {
  let cell = SpreadsheetApp.getActive().getRange('S1:Y1');
  rangeCtrlShiftDown(cell).activate();
}
function rangeCtrlShiftDown(range) {
  // вернуть прямоугольный диапазон от range по последнюю строку со значеними
  let sheet = sheetByRange(range);
  let row_First = 1;
  let row_Last_ = 1;
  // ToDo: продолжить
  // let row_Last_ = sheetColumnValueRowLastNumber()
  // return sheet.getRange(range);
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


function sheetColumnValueRowLastNumber(range) {
  // принимает  диапазон,
  // возвращает номер последней непустой строки
  // идёт снизу вверх по массиву

  let array2d = range.getValues();
  for (let i = array2d.length - 1; i >= 0; i--) {
    if (array2d[i][0] != null && array2d[i][0] != '') {
      return i + 1;
    };
  };
};


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

