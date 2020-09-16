'use strict';

const Excel = require('exceljs');

const fs = require('fs');

const oldTxt = 'OLD';
const newTxt = 'NEW';
const inputName = 'original';
const outputName = 'new';

let workbook = new Excel.Workbook();

function readWrite() {
  /* read the file */

  workbook.xlsx
    .readFile(`${outputName}.xlsx`)
    .then(function () {
      let sheetNames = workbook.SheetNames;

      sheetNames.forEach(function (y) {
        let sheet = workbook.Sheets[y];

        /* loop through every cell manually */
        let range = XLSX.utils.decode_range(sheet['!ref']); // get the range

        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            /* find the cell object */
            let cellRef = XLSX.utils.encode_cell({c: C, r: R}); // construct A1 reference for cell
            if (!sheet[cellRef]) continue; // if cell doesn't exist, move on

            let cell = sheet[cellRef];

            /* if the cell is a text cell with the old string, change it */
            if ((cell.t !== 's' && cell.t !== 'str') || !cell.v) continue; // skip if cell is not text

            let v = cell.v;
            let regex = new RegExp(oldTxt, 'g');

            if (v.includes(oldTxt)) cell.v = v.replace(regex, newTxt);
          }
        }
      });

      workbook.xlsx.writeFile(`${outputName}.xlsx`).catch(function (err) {
        console.log(`writeFile Err: ${err}.`);
      });
    })
    .catch(function (err) {
      console.log(`readFile Err: ${err}.`);
    });
}

// new.txt will be created or overwritten by default.
fs.copyFile(`${inputName}.xlsx`, `${outputName}.xlsx`, (err) => {
  if (err) throw err;
  console.log(`${inputName}.xlsx was copied to ${outputName}.xlsx`);
  console.log('read file');
  readWrite();
});
