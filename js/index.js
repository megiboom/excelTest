const makeExcel = require('./excel.js')
const makeGroup = require('./groupJson.js')
const getData = require('./getData');
const logs = require('./logs');

var path = require('path');
var scriptName = path.basename(__filename);

async function main(){
	var msg = "Main Start"
    await logs.writeLogs(msg,scriptName,1);
    
    const data = await getData.getData();
    const jsonData = await makeGroup.groupBy(data);
    await logs.writeLogs(JSON.stringify(jsonData),scriptName);
    await makeExcel.makeExcel(jsonData)
    msg = "Main End"
    await logs.writeLogs(msg,scriptName,-1);
}

main();