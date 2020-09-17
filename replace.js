const XLSX = require('xlsx-populate');
const dotenv = require('dotenv');
const getDict = require('./src/getDict');

dotenv.config();

const dictName = process.env.DICT_NAME;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

function findReplace(dict) {
  XLSX.fromFileAsync(`./${inputName}.xlsx`).then((workbook) => {
    const keys = Object.keys(dict);
    const values = Object.values(dict);

    for (let i = 0; i < keys.length; i += 1) {
      const oldTxt = keys[i].toString();
      const newTxt = values[i].toString();

      const regex = new RegExp(oldTxt, 'g');
      workbook.find(regex, () => newTxt);
    }

    // new.txt will be created or overwritten by default.
    workbook.toFileAsync(`${inputName}-${outputName}.xlsx`);
  });
}

const dict = getDict(dictName);
findReplace(dict);
