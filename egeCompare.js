const xlsx = require('xlsx');
const path = require('path');

const compareEge = (fis, uws) => {
    const fisbook = xlsx.readFile(fis);
    const uwsbook = xlsx.readFile(uws);

    const sheet1 = fisbook.Sheets[fisbook.SheetNames[0]];
    const sheet2 = uwsbook.Sheets[uwsbook.SheetNames[0]];

    const data1 = xlsx.utils.sheet_to_json(sheet1);
    const data2 = xlsx.utils.sheet_to_json(sheet2);

    const differences = [];

    data1.forEach(row1 => {
        const matchedRow = data2.find(row2 => 
            row2['фамилия'] === row1['фамилия'] &&
            row2['имя'] === row1['имя'] &&
            row2['отчество'] === row1['отчество'] &&
            row2['серия'] === row1['серия'] &&
            row2['номер'] === row1['номер'] &&
            row2['дисциплина'] === row1['дисциплина']
        );

        if (matchedRow) {
            if (row1[' результат'] !== matchedRow[' результат']) {
                differences.push({
                    фамилия: row1['фамилия'],
                    имя: row1['имя'],
                    отчество: row1['отчество'],
                    серия: row1['серия'],
                    номер: row1['номер'],
                    дисциплина: row1['дисциплина'],
                    результатФИСИ: row1[' результат'],
                    результатUWS: matchedRow[' результат']
                });
            }
        } else {
            differences.push({
                фамилия: row1['фамилия'],
                имя: row1['имя'],
                отчество: row1['отчество'],
                серия: row1['серия'],
                номер: row1['номер'],
                дисциплина: row1['дисциплина'],
                результатФИСИ: row1[' результат'],
                результатUWS: 'Результа не наблюдается'
            });
        }
    });

    return differences;
};

const saveDifferencesToExcel = (differences, outputFileName) => {
    const newWorkbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.json_to_sheet(differences);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Differences');
    xlsx.writeFile(newWorkbook, outputFileName);
};

const fisfile = path.resolve(__dirname, 'fis.xlsx');
const uwsfile = path.resolve(__dirname, 'uws.xlsx');
const differences = compareEge(fisfile, uwsfile);

if (differences.length > 0) {
    saveDifferencesToExcel(differences, 'РАЗЛИЧИЯ.xlsx');
    console.log('Различия найдены и сохранены в РАЗЛИЧИЯ.xlsx');
} else {
    console.log('Различий не найдено.');
}


