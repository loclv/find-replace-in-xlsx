'use strict';

const XLSX = require('xlsx-populate');
const fs = require('fs');
const dotenv = require('dotenv');
const getDict = require('./src/getDict');

dotenv.config();

const dictName = process.env.DICT_NAME;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

function findReplace(dict) {
  const workbook = XLSX.fromFileAsync(`./${inputName}.xlsx`).then(
    (workbook) => {
      for (let key in dict) {
        let oldTxt = key.toString();
        let newTxt = dict[key].toString();

        let regex = new RegExp(oldTxt, 'g');
        workbook.find(regex, (match) => newTxt);
      }

      // new.txt will be created or overwritten by default.
      workbook.toFileAsync(`${inputName}-${outputName}.xlsx`);
    },
  );
}

const dict = getDict(dictName);
findReplace(dict);
