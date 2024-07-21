const xlsx = require('xlsx');
const path = require('path');

const normalizeString = (str) => {
    return (str ? str.toString().toLowerCase().trim() : '');
};

const convertAllValuesToString = (data) => {
    return data.map(row => {
        const newRow = {};
        for (const key in row) {
            if (Object.hasOwnProperty.call(row, key)) {
                newRow[key] = normalizeString(row[key]);
            }
        }
        return newRow;
    });
};

const compareEge = (fis, uws) => {
    const fisbook = xlsx.readFile(fis);
    const uwsbook = xlsx.readFile(uws);

    const sheet1 = fisbook.Sheets[fisbook.SheetNames[0]];
    const sheet2 = uwsbook.Sheets[uwsbook.SheetNames[0]];

    const data1 = convertAllValuesToString(xlsx.utils.sheet_to_json(sheet1));
    const data2 = convertAllValuesToString(xlsx.utils.sheet_to_json(sheet2));

    const differences = [];

    data1.forEach(row1 => {
        const matchedRow = data2.find(row2 => {
            return row2['фамилия'] === row1['фамилия'] &&
                   row2['имя'] === row1['имя'] &&
                   row2['отчество'] === row1['отчество'] &&
                   row2['серия'] === row1['серия'] &&
                   row2['номер'] === row1['номер'] &&
                   row2['дисциплина'] === row1['дисциплина'];
        });

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
                результатUWS: 'Результат не наблюдается в нашей любимой системе...'
            });
        }
    });

    data2.forEach(row2 => {
        const matchedRow = data1.find(row1 => {
            return row1['фамилия'] === row2['фамилия'] &&
                   row1['имя'] === row2['имя'] &&
                   row1['отчество'] === row2['отчество'] &&
                   row1['серия'] === row2['серия'] &&
                   row1['номер'] === row2['номер'] &&
                   row1['дисциплина'] === row2['дисциплина'];
        });

        if (!matchedRow) {
            differences.push({
                фамилия: row2['фамилия'],
                имя: row2['имя'],
                отчество: row2['отчество'],
                серия: row2['серия'],
                номер: row2['номер'],
                дисциплина: row2['дисциплина'],
                результатФИСИ: 'Результат не наблюдается в Фисе...',
                результатUWS: row2[' результат']
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

const getUniqueFileName = (baseName) => {
    const now = new Date();
    const timestamp = `${now.getFullYear()}_${now.getMonth()}_${now.getDate()}_${now.getMinutes()}_${now.getSeconds()}`; 
    return `${baseName}_${timestamp}.xlsx`;
};

const fisfile = path.resolve(__dirname, 'fis.xlsx');
const uwsfile = path.resolve(__dirname, 'uws.xlsx');
const differences = compareEge(fisfile, uwsfile);

if (differences.length > 0) {
    const outputFileName = getUniqueFileName('РАЗЛИЧИЯ');
    saveDifferencesToExcel(differences, outputFileName);
    console.log(`Различия найдены и сохранены в ${outputFileName}`);
} else {
    console.log('Различий не найдено.');
}
