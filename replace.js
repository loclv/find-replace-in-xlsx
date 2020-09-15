'use strict';

const XLSX = require('xlsx-populate');
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
  const dict = readDict();
  findReplace(dict);
});
