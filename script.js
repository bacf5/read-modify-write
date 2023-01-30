const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

function readFileToJson(filename) {
  const wb = xlsx.readFile(filename, { cellDates: true });
  const firsTabName = wb.SheetNames[0];
  const ws = wb.Sheets[firsTabName];
  const data = xlsx.utils.sheet_to_json(ws);
  return data;
}

let combinedData = [];

// link al directorio que queremos leer
const sourceFolder = __dirname + '/Septiembre';
// leer archivos de forma sincronica (sync)
const files = fs.readdirSync(sourceFolder);
// console.log(__dirname);
// Loop sobre cada archivo para obtener el nombre.
files.forEach((file) => {
  const fileExtension = path.parse(file).ext;
  if (fileExtension === '.xlsm') {
    // dirección completa del archivo a leer
    const fullFilePath = path.join(sourceFolder, file);
    // console.log(fullFilePath);
    // leer el archivo para sacar la data
    const data = readFileToJson(fullFilePath);
    // combinar la data en un solo array
    combinedData = combinedData.concat(data);
  }
  //   console.log(fileExtension);
});

const newWb = xlsx.utils.book_new();
const newWs = xlsx.utils.json_to_sheet(combinedData);
xlsx.utils.book_append_sheet(newWb, newWs, 'Combined Data');
xlsx.writeFile(newWb, 'NewCombinedData.xlsx');
console.log('done!');
