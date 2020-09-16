'use strict';

const Excel = require('exceljs');

const fs = require('fs');

const dotenv = require('dotenv');
dotenv.config();

const oldTxt = process.env.OLD_TXT;
const newTxt = process.env.NEW_TXT;
const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

let workbook = new Excel.Workbook();

function readWrite() {
  workbook.xlsx.readFile(`${inputName}.xlsx`).then(function () {
    workbook.eachSheet(function (worksheet, sheetId) {
      worksheet.eachRow(function (row, rowNumber) {
        row.eachCell(function (cell, colNumber) {
          let v = cell.value;
          if (!v) return;
          let regex = new RegExp(oldTxt, 'g');
          // if the cell is a text cell with the old string, change it
          if (v.includes(oldTxt)) cell.value = v.replace(regex, newTxt);
        });
        // Commit a completed row to stream
        row.commit();
      });
    });

    workbook.xlsx
      .writeFile(`${outputName}.xlsx`)
      .then(function () {
        console.log('File is written');
      })
      .catch(function (err) {
        console.log(`writeFile Err: ${err}.`);
      });
  });
}

// new.txt will be created or overwritten by default.
readWrite();
