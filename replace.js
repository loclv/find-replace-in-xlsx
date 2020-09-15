'use strict';

const XLSX = require('xlsx-style');
const fs = require('fs');
const dotenv = require('dotenv');
const yaml = require('js-yaml');

dotenv.config();

const DICT_NAME = process.env.DICT_NAME;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

function readDict() {
  try {
    let fileContents = fs.readFileSync(`./${DICT_NAME}.yaml`, 'utf8');
    const dict = yaml.load(fileContents);

    console.log(dict);

    console.log(`Read ./${DICT_NAME}.yaml`);

    return dict;
  } catch (e) {
    console.log(e);
  }
}

function findReplace(dict) {
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

        for (let key in dict) {
          let oldTxt = key.toString();
          let newTxt = dict[key].toString();
          console.log(oldTxt);
          console.log(newTxt);

          let regex = new RegExp(oldTxt, 'g');
          if (v.includes(oldTxt)) cell.v = v.replace(regex, newTxt);
        }
      }
    }
  });

  XLSX.writeFile(workbook, `${outputName}.xlsx`);
}

// new.txt will be created or overwritten by default.
fs.copyFile(`${inputName}.xlsx`, `${outputName}.xlsx`, (err) => {
  if (err) throw err;
  console.log(`${inputName}.xlsx was copied to ${outputName}.xlsx`);
  const dict = readDict();
  findReplace(dict);
});
