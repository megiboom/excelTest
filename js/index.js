const makeExcel = require('./excel.js')
const makeGroup = require('./groupJson.js')
const getData = require('./getData');
const logs = require('./logs');
const config = require('../config/excelConfig.json');

var path = require('path');
var scriptName = path.basename(__filename);

async function main(){
	var msg = "Main Start"
    await logs.writeLogs(msg,scriptName,1);
        
    const oldData = await getData.getData(config.baseUrl+config.getDataUrlOld);
    const oldJsonData = await makeGroup.groupByOld(oldData);
    await logs.writeLogs(JSON.stringify(oldJsonData),scriptName);
    
    const fileName = await makeExcel.getFileName(config.excelDownloadDir);

    await makeExcel.makeExcelOld(oldJsonData,fileName,config.excelDownloadDir)

    const newData = await getData.getData(config.baseUrl+config.getDataUrlNew);
    const newJsonData = await makeGroup.groupByNew(newData);
    await logs.writeLogs(JSON.stringify(newJsonData),scriptName);
    await delay(100);
    await makeExcel.makeExcelNew(newJsonData,fileName,config.excelDownloadDir);
    
    msg = "Main End"
    await logs.writeLogs(msg,scriptName,-1);
}

function delay(ms) {
    return new Promise(resolve => {
      setTimeout(() => {
        resolve()
      }, ms);
    });
  }

main();