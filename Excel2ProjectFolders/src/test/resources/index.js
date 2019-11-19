import fs from 'fs';
import { PACT_BECTU_Timesheet } from '@sddevelopment/modular-forms';

const tests = [];

fs.readdirSync(__dirname).forEach(function(file) {
  const currentPath = __dirname + '/' + file;

  // If This is a directory then get index.js from it
  if (fs.lstatSync(currentPath).isDirectory()) {
    let outputJson = {},
      readOnlyFields,
      requiredFields,
      hiddenFields,
      configJson;
    // get model input
    const inputJson = require(currentPath + '/in.json');
    // get expected output
    if (fs.existsSync(currentPath + '/out.json')) {
      outputJson = require(currentPath + '/out.json');
    }
    // get expected output
    if (fs.existsSync(currentPath + '/readOnly.json')) {
      readOnlyFields = require(currentPath + '/readOnly.json');
    }
    // get expected output
    if (fs.existsSync(currentPath + '/required.json')) {
      requiredFields = require(currentPath + '/required.json');
    }
    // get expected output
    if (fs.existsSync(currentPath + '/hidden.json')) {
      hiddenFields = require(currentPath + '/hidden.json');
    }
    // get config
    if (fs.existsSync(currentPath + '/config.json')) {
      configJson = require(currentPath + '/config.json');
    }

    tests.push({
      ...configJson,
      inputModel: inputJson,
      outputModel: outputJson,
      readOnlyFields,
      requiredFields,
      hiddenFields,
    });
  }
});

export default {
  tests,
  template: PACT_BECTU_Timesheet,
};
