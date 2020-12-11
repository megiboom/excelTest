const makeExcel = require('./excel.js')
const makeGroup = require('./groupJson.js')
const getData = require('./getData');
const logs = require('./logs');

var path = require('path');
var scriptName = path.basename(__filename);

async function main(){
	var msg = "Main Start"
    await logs.writeLogs(msg,scriptName,1);

    //const data = await getData.getData();
    //const jsonData = await makeGroup.groupByOld(data);
    //await makeExcel.makeExcelOld(jsonData);
    
    const oldData = await getData.getData('Y');
    const oldJsonData = await makeGroup.groupByOld(oldData);
    //await logs.writeLogs(JSON.stringify(oldJsonData),scriptName);
    
    await makeExcel.makeExcelOld(oldJsonData)
    const newData = await getData.getData('N');
    const newJsonData = await makeGroup.groupByNew(newData);
    //await logs.writeLogs(JSON.stringify(newJsonData),scriptName);

    await makeExcel.makeExcelNew(newJsonData)
    msg = "Main End"
    await logs.writeLogs(msg,scriptName,-1);
}

main();