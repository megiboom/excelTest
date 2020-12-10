var path = require('path');
const logs = require('./logs');
var scriptName = path.basename(__filename);

function groupBy(arr){
	var msg = "groupBy Start"
    logs.writeLogs(msg,scriptName);

    const result = arr.reduce((r, { moduleId: moduleId,moduleName:moduleName,groupCd:groupCd,groupName:groupName,loadQty:loadQty,oldYn:oldYn, ...object }) => {
        var temp = r.find(o => o.moduleId === moduleId);
        if (!temp) r.push(temp = { moduleId,moduleName,groupCd,groupName,loadQty,oldYn, data: [] });
        temp.data.push(object);
        return r;
    }, []);
    
    const result1 = result.reduce((r,{groupCd:groupCd,groupName:groupName,oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.groupCd === groupCd);
        if (!temp) r.push(temp = { groupCd,groupName,oldYn, moduleList: [] });
        temp.moduleList.push(object);
        return r;
    },[])

    const result2 = result1.reduce((r,{oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.oldYn === oldYn);
        if (!temp) r.push(temp = { oldYn, groupList: [] });
        temp.groupList.push(object);
        return r;
    },[])

	var msg = "groupBy Complete : "+result1.length
    logs.writeLogs(msg,scriptName);
    return result2;
}

module.exports.groupBy=groupBy;