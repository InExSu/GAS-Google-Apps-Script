function arrays2dDiff(arr1, arr2, row1Head, row2Head) {
    // массив соответствия номеров заголовков столбцов
    let a11Head = array2dRow2Array1(arr1, row1Head)
    let a12Head = array2dRow2Array1(arr2, row2Head)
    let arrCols = Array1D_2_HeadNumbers_LookUp(a11Head, a12Head);
    let arrKeys = Array2D_2_Map(arr2, colKey);
    return arrays2dDiffAction(arr1, arr2, arrCols, arrKeys);
}

function arrays2dDiffAction(arr1, arr2, , arrKeys, arr1Cols) {
    // Вернуть разницу двух массивов по ключевому полю
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

function array2dRow2Array1(arr2, row2Head) {
    // строку массива 2мерного в массив 1мерный

}
