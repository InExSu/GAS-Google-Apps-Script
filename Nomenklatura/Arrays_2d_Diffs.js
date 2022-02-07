function arrays2dDiff_Test() {
    // Вернуть массив 2мерный было/стало по ключевому полю
    // в одинаковых столбцах

    let arr1, arr2, row1Head, row2Head;

    arr1 = [['ст1', 'ст2'], ['val11', 'val12']];
    arr2 = [['ст1', 'ст2'], ['val21', 'val22']];

    // массив соответствия номеров заголовков столбцов
    let a11Head = array2dRow2Array1(arr1, 1)
    let a12Head = array2dRow2Array1(arr2, 1)

    let arrCols = Array1D_2_HeadNumbers_LookUp(a11Head, a12Head);
    let arrKeys = Array2D_2_Map(arr2, colKey);

    let arr3 = arrays2dDiff(arr1, arr2, arrCols, arrKeys);
}

function arrays2dDiff(arr1, arr2, arrKeys, arr1Cols) {
    // Вернуть массив 2мерный (по размерам arr2) было/стало по ключевому полю
    // в одинаковых столбцах

    // массив 1 - сводная, массив 2 - Битрикс24
    // Создать словарь из ключевого столбца массива 2
    // создать копию массива 2
    // создать массив соответствия номеров полей заголовков
    // проходом по ключам массива 1 
    // если ключ есть в словаре
    // если заголовок есть в массиве 2
    // если значение 1 и значени 2 различаются
    // добавить в массив 3 значение 1 / значение 2 
    // вернуть массив 3
    let arr3,
        col1,
        col2,
        row2,
        val1,
        val2;

    arr3 = JSON.parse(JSON.stringify(arr2));

    for (let row1 = 0; row1 < arr1.length; row1++) {
        for (let indx = 0; indx <= arr1Cols.length; col1++) {

            col1 = arr1Cols[indx][0];

            val1 = arr1[row1][col1];

            row2 = arrKeys[val1];

            col2 = arr1Cols[indx][1];

            if (row2 * col2 > 0) {

                val2 = arr2[row2][col2];

                if (val1 !== val2) {
                    arr3[row3][col3] = val1 + '/' + val2;
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

// array2dRow2Array1_Test();
