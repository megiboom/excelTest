var path = require('path');
const logs = require('./logs');
var scriptName = path.basename(__filename);

function groupBy(arr){
	var msg = "groupBy Start"
    logs.writeLogs(msg,scriptName);

    const result = arr.reduce((r, { moduleId: moduleId,moduleName:moduleName,groupId:groupId,groupName:groupName, ...object }) => {
        var temp = r.find(o => o.moduleId === moduleId);
        if (!temp) r.push(temp = { moduleId,moduleName,groupId,groupName, data: [] });
        temp.data.push(object);
        return r;
    }, []);
    
    const result1 = result.reduce((r,{groupId:groupId,groupName:groupName, ...object})=>{
        var temp = r.find(o => o.groupId === groupId);
        if (!temp) r.push(temp = { groupId,groupName, moduleList: [] });
        temp.moduleList.push(object);
        return r;
    },[])

	var msg = "groupBy Complete : "+result1.length
    logs.writeLogs(msg,scriptName);
    return result1;
}
module.exports.groupBy=groupBy;