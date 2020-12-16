const makeExcel = require('./excel.js')
const makeGroup = require('./groupJson.js')
const getData = require('./getData');
const logs = require('./logs');
const config = require('../config/config.json');

var path = require('path');
var scriptName = path.basename(__filename);

async function main(){
	var msg = "Main Start"
    await logs.writeLogs(msg,scriptName,1);
        
    const oldData = await getData.getData(config.getDataUrlOld);
    const oldJsonData = await makeGroup.groupByOld(oldData);
    
    const fileName = await makeExcel.makeExcelOld(oldJsonData,config.excelDownloadDir)

    const newData = await getData.getData(config.getDataUrlNew);
    const newJsonData = await makeGroup.groupByNew(newData);

    await makeExcel.makeExcelNew(newJsonData,fileName,config.excelDownloadDir);
    msg = "Main End"
    await logs.writeLogs(msg,scriptName,-1);
}

main();