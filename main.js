const Exceljs = require('exceljs');
const fs = require('fs');

const file = fs.readFileSync('./files/Emmaus-27.txt', 'utf8')
const allItemsArray = file.split("\n").map(function(row){return row.split(",");}) 

let allRelays = []

let panelA = {
    size: 48,
    name: "LCP-A",
    location: "Elec-102",
    relays: [
        {
            number: "1",
            ckt: "hd1-13",
            loadName: "conference 123",
            emergency: false,
        },
        {
            number: "2",
            ckt: "hd1-13",
            loadName: "conference 123",
            emergency: true,
        }
    ]
}


async function copyExcel () {
    let targetWorkbook = new Ex
}


function printPanel (panel) {
    const workbook = new Exceljs.Workbook();
    const worksheet = workbook.addWorksheet(panel.name);

    let sourceWorkbook = new Exceljs.Workbook();
    sourceWorkbook = sourceWorkbook.xlsx.readFile('./files/source.xlsx');
    let sourceWorksheet = workbook.getWorksheet('relays')

    let sourceRange;
    let targetRange;

    switch(panel.size){
        case 8:
            sourceRange = ['A2', 'I9']
            targetRange = ['A1', 'I8']
        break;
        case 16:
            sourceRange = ["A12", 'I23']
            targetRange = ['A1', 'I12']
        break;
        case 32:
            sourceRange = ['A26', 'I45']
            targetRange = ['A1', 'I20']
        break;
        case 48:
            sourceRange = ['A48', "I75"]
            targetRange = ['A1', 'I28']
        break;
    }
    console.log(sourceRange, targetRange)
    worksheet.mergeCells('A1:I2')
    worksheet.getCell('A1').value = 'Hello, World!'
    workbook.xlsx.writeFile('./files/output.xlsx')
}
printPanel(panelA)

function sortAllItems(){
    for (i = 0; i < allItemsArray; i < allItemsArray.length){
        if(allItemsArray[i][0].includes('Relay')){
            allRelays.push(allItemsArray[i])
        }
    }
}
sortAllItems();