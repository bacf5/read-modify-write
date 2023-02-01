const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const rangeAwb = { s: { c: 1, r: 23 }, e: { c: 1, r: 38 } };
const rangePas = { s: { c: 1, r: 5 }, e: { c: 1, r: 19 } };

function readFileToJson(filename) {
  // lee el archivo que viene de files
  const wb = xlsx.readFile(filename, { cellDates: true });
  // declaramos la primer sheet
  const firsTabName = wb.SheetNames[0];
  // agarramos la sheet
  const ws = wb.Sheets[firsTabName];
  // leemos el range de las celdas usadas que hay en la sheet
  const range = xlsx.utils.decode_range(ws['!ref']);
  //   console.log(range) ---> { s: { c: 0, r: 0 }, e: { c: 8, r: 43 } }

  // loop de guías entregas
  let numGuias = [];
  for (let rowNum = rangeAwb.s.r; rowNum <= range.e.r; rowNum++) {
    const awb = ws[xlsx.utils.encode_cell({ r: rowNum, c: 1 })];
    let obj = {};
    if (awb === undefined) {
    } else {
      // console.log(awb['v']);
      obj = awb['v'];
      numGuias.push(obj);
      // numGuias.push(awb.v);
    }
    // console.log(numGuias);
  }
  const dataFilter = numGuias.filter((guia) => guia !== undefined);
  // console.log(dataFilter);
  // console.log(numGuias);
  const data = xlsx.utils.json_to_sheet(dataFilter);
  console.log(data);

  // le pasamos la data a la sheet
  // returneamos la data
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
    // console.log(combinedData);
  }
  //   console.log(fileExtension);
});
/*
// Crear nuevo woorkbook
const newWb = xlsx.utils.book_new();
// Crear nuevo worksheet
const newWs = xlsx.utils.json_to_sheet(combinedData);
// Fusiona el worksheet dentro del workbook
xlsx.utils.book_append_sheet(newWb, newWs, 'Combined Data');
// Escribe el archivo y lo genera con el nombre
xlsx.writeFile(newWb, 'NewCombinedData.xlsx');
console.log('done!');
*/
