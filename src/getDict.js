/* eslint-disable no-console */
const fs = require('fs');
const yaml = require('js-yaml');

function getDict(DICT_NAME) {
  try {
    const fileContents = fs.readFileSync(`./${DICT_NAME}.yaml`, 'utf8');
    const dict = yaml.load(fileContents);

    console.log(dict);

    console.log(`Read ./${DICT_NAME}.yaml`);

    return dict;
  } catch (e) {
    console.log(e);
    return null;
  }
}

module.exports = getDict;
