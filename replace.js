'use strict';

const XLSX = require('xlsx-populate');
const fs = require('fs');
const dotenv = require('dotenv');
const getDict = require('./src/getDict');

dotenv.config();

const DICT_NAME = process.env.DICT_NAME;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

function findReplace(dict) {
  const workbook = XLSX.fromFileAsync(`./${outputName}.xlsx`).then(
    (workbook) => {
      for (let key in dict) {
        let oldTxt = key.toString();
        let newTxt = dict[key].toString();

        let regex = new RegExp(oldTxt, 'g');
        workbook.find(regex, (match) => newTxt);
      }

      workbook.toFileAsync(`${outputName}.xlsx`);
    },
  );
}

// new.txt will be created or overwritten by default.
fs.copyFile(`${inputName}.xlsx`, `${outputName}.xlsx`, (err) => {
  if (err) throw err;
  console.log(`${inputName}.xlsx was copied to ${outputName}.xlsx`);
  const dict = getDict(DICT_NAME);
  findReplace(dict);
});
