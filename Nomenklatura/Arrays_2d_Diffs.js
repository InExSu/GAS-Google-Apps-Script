// import { Array1D_2_HeadNumbers_LookUp } from './LibraryBigInExSu.js';
// import { Array2D_2_Map } from './LibraryBigInExSu.js';

function arrays2dDiff_Test() {
    // Вернуть массив 2мерный было/стало по ключевому полю
    // в одинаковых столбцах

    let arr1 = [
        ['ст1', 'ст2'],
        ['va1', 'va2']];
    let arr2 = [
        ['ст1', 'ст2'],
        ['va1', 'va2']];

    // массив соответствия номеров заголовков столбцов
    let a11Head = array2dRow2Array1(arr1, 1)
    let a12Head = array2dRow2Array1(arr2, 1)

    let arrCols = Array1D_2_HeadNumbers_LookUp(a11Head, a12Head);
    let arrKeys = Array2D_2_Map(arr2, 1);

    let arr3 = arrays2dDiff(arr1, arr2, arrKeys, arrCols, 1);

}

function arrays2dDiff(arr1, arr2, arrKeys, arr2Cols, colKey) {
    // Вернуть массив 2мерный (по размерам arr2) было/стало по ключевому полю
    // в одинаковых столбцах

    // массив 1 - сводная, массив 2 - Битрикс24
    // проходом по столбцу ключей массива arr1,
    // если ключ есть в словаре,
    // подобрать столбцы дя двух массивов из 
    // если значение 1 и значени 2 различаются
    // добавить в массив 3 значение 1 / значение 2 
    // вернуть массив 3
    let arr3,
        col1,
        col2,
        key_,
        row2,
        val1,
        val2;

    arr3 = JSON.parse(JSON.stringify(arr2));

    for (let row1 = 0; row1 < arr1.length; row1++) {

        key_ = arr1[row1][colKey];

        if (arrKeys.has(key_)) {

            row2 = arrKeys.get(key_);

            for (let indx = 0; indx < arr2Cols.length; indx++) {

                col1 = arr2Cols[indx][0];
                col2 = arr2Cols[indx][1];

                val1 = arr1[row1][col1];
                val2 = arr2[row2][col2];

                if (val1 !== val2) {

                    arr3[row2][col2] = val1 + '/' + val2;
                }
            }
        }
    }
    return arr3;
}

function array2dRow2Array1_Test() {
    let a2 = [[1, 2], [3, 4]];
    let a1 = array2dRow2Array1(a2, 0);
    if (a1[0] == 1) {
        return array2dRow2Array1_Test + " " + true;
    }
    return array2dRow2Array1_Test + " " + false;
}

function array2dRow2Array1(arr2, row) {
    // строку массива 2мерного в массив 1мерный

    let a1 = [];

    for (let col = 0; col < arr2.length; col++) {

        a1.push(arr2[row][col]);
    }
    return a1;
}

// Копия из LibraryBigInExSu
function Array2D_2_Map(array2d, column_key) {
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

// Копия из LibraryBigInExSu
function Array1D_2_HeadNumbers_LookUp(array1d_Old, array1d_New) {

    // из двух 1мерных массивов создать массив 2мерный с соответствия номеров столбцов

    let value;
    let row_new;
    let array2D = [];

    for (var row_old = 0; row_old < array1d_Old.length; row_old++) {

        value = array1d_Old[row_old];

        if (String(value).length > 0) {

            row_new = array1d_New.indexOf(value);

            if (row_new > -1) {
                array2D.push([row_old, row_new]);
            }
        }
    }

    return array2D;
}

// Пусть строки отладки будут внизу
// Array1D_2_HeadNumbers_LookUp();
arrays2dDiff_Test();
