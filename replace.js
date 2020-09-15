var oldTxt = 'OLD';
var newTxt = 'NEW';
var inputName = 'original';
var outputName = 'new';

var XLSX = require('xlsx');

var fs = require('fs');

function findReplace() {
  /* read the file */
  var workbook = XLSX.readFile(`${outputName}.xlsx`); // parse the file
  var sheetNames = workbook.SheetNames;

  sheetNames.forEach(function (y) {
    var sheet = workbook.Sheets[y];

    /* loop through every cell manually */
    var range = XLSX.utils.decode_range(sheet['!ref']); // get the range

    for (var R = range.s.r; R <= range.e.r; ++R) {
      for (var C = range.s.c; C <= range.e.c; ++C) {
        /* find the cell object */
        var cellRef = XLSX.utils.encode_cell({c: C, r: R}); // construct A1 reference for cell
        if (!sheet[cellRef]) continue; // if cell doesn't exist, move on

        var cell = sheet[cellRef];

        /* if the cell is a text cell with the old string, change it */
        if ((cell.t !== 's' && cell.t !== 'str') || !cell.v) continue; // skip if cell is not text

        var v = cell.v;
        var regex = new RegExp(oldTxt, 'g');

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
