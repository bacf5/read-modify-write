const xlsx = require('xlsx');
const fs = require('fs');

function readFileToJson(filename) {
  const wb = xlsx.readFile(filename, { cellDates: true });
  const firsTabName = wb.SheetNames[0];
}

// link al directorio que queremos leer
const folder = __dirname + '/Septiembre';
// leer archivos de forma sincronica (sync)
const files = fs.readdirSync(folder);
// console.log(__dirname);
console.log(files);
