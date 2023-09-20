const xlsx = require("node-xlsx");
const fs = require('fs');
const path = require('path');

const filePath = path.resolve(__dirname, './list.xlsx')
const COUNT = 5000;

function main() {
    const workSheetsFromFile = xlsx.parse(filePath);

    console.log(workSheetsFromFile);

    // makeFile(workSheetsFromFile[0].data.slice(0, 10))
    const header = workSheetsFromFile[0].data.slice(0, 1);

    loopData(workSheetsFromFile[0].data.slice(1), header)
}

function loopData (dataList, header) {
    let index = 0;
    let fileIndex = 1;

    while (index < dataList.length) {
        makeFile([...header, ...dataList.slice(index, index + COUNT)], fileIndex)
        index += COUNT;
        fileIndex++;
    }
}

function makeFile (data, fileIndex) {
    const bufferData = xlsx.build([{name: 'mySheetName', data: data}]);

    fs.writeFile(`./source/item_${fileIndex}.xlsx`, bufferData, (err) => {
        if (err) {
            console.log(err)
            return;
        }
        console.log('succss')
    } )
}

main();