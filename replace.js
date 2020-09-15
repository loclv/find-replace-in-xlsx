'use strict';

const Excel = require('exceljs');

const fs = require('fs');

const oldTxt = 'OLD';
const newTxt = 'NEW';
const inputName = 'original';
const outputName = 'new';

let workbook = new Excel.Workbook();

function readWrite() {
  workbook.xlsx
    .readFile(`${outputName}.xlsx`)
    .then(function () {
      let sheetNames = workbook.SheetNames;

      sheetNames.forEach(function (y) {
        let sheet = workbook.Sheets[y];

        let range = XLSX.utils.decode_range(sheet['!ref']);

        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            let cellRef = XLSX.utils.encode_cell({c: C, r: R});
            // if cell doesn't exist, move on
            if (!sheet[cellRef]) continue;

            let cell = sheet[cellRef];
            // skip if cell is not text
            if ((cell.t !== 's' && cell.t !== 'str') || !cell.v) continue;

            let v = cell.v;
            let regex = new RegExp(oldTxt, 'g');
            // if the cell is a text cell with the old string, change it
            if (v.includes(oldTxt)) cell.v = v.replace(regex, newTxt);
          }
        }
      });

      workbook.xlsx
        .writeFile(`${outputName}.xlsx`)
        .then(function () {
          console.log('File is written');
        })
        .catch(function (err) {
          console.log(`writeFile Err: ${err}.`);
        });
    })
    .catch(function (err) {
      console.log(`readFile Err: ${err}.`);
    });
}

// new.txt will be created or overwritten by default.
readWrite();
