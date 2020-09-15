'use strict';

const XLSX = require('xlsx');
const fs = require('fs');
const dotenv = require('dotenv');

dotenv.config();

const oldTxt = process.env.OLD_TXT;
const newTxt = process.env.NEW_TXT;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

function findReplace() {
  /* read the file */
  const workbook = XLSX.readFile(`${outputName}.xlsx`); // parse the file
  const sheetNames = workbook.SheetNames;

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

  XLSX.writeFile(workbook, `${outputName}.xlsx`);
}

// new.txt will be created or overwritten by default.
fs.copyFile(`${inputName}.xlsx`, `${outputName}.xlsx`, (err) => {
  if (err) throw err;
  console.log(`${inputName}.xlsx was copied to ${outputName}.xlsx`);
  findReplace();
});
