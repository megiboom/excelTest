const axios = require('axios')
const logs = require('./logs');

var path = require('path');
var scriptName = path.basename(__filename);

async function getData(){
	var msg = "getData"
    await logs.writeLogs(msg,scriptName);
    try{
        var rtn = await axios.get('http://localhost:8070/api/power/excel/hourly');
        if(rtn.data==null || rtn.data=="" || rtn.data.resultCd!="0000") {
	        msg = ("getData Fail: ",rtn.data);
            await logs.writeLogs(msg,scriptName);
            return;
        }
        msg = ("getData Success: "+rtn.data.pfe.length);
        await logs.writeLogs(msg,scriptName);

        return rtn.data.pfe;
    }catch(e){
	    msg = ("getData Fail: ",e);
        await logs.writeLogs(msg,scriptName);
        console.error(e);
    }
}

module.exports.getData = getData;