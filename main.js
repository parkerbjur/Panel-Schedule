const Exceljs = require('exceljs');
const fs = require('fs');

const file = fs.readFileSync('./files/Emmaus-27.csv', 'utf8') 

console.log(file);