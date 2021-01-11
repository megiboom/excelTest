var path = require('path');
const logs = require('./logs');
var scriptName = path.basename(__filename);

function groupByOld(arr){
	var msg = "groupBy Start"
    logs.writeLogs(msg,scriptName);

    const result = arr.reduce((r, { moduleId: moduleId,moduleName:moduleName,groupCd:groupCd,groupName:groupName,loadQty:loadQty,oldYn:oldYn, ...object }) => {
        var temp = r.find(o => o.moduleId === moduleId && o.groupCd === groupCd);
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

function groupByNew(arr){
	var msg = "groupBy Start"
    logs.writeLogs(msg,scriptName);

    const groupByModule = arr.reduce((r, { buildCd:buildCd,buildName:buildName,groupCd:groupCd,groupName:groupName,moduleId: moduleId,moduleName:moduleName,loadQty:loadQty,oldYn:oldYn, ...object }) => {
        var temp = r.find(o => o.moduleId === moduleId);
        if (!temp) r.push(temp = { buildCd,buildName,groupCd,groupName,moduleId,moduleName,loadQty,oldYn, data: [] });
        temp.data.push(object);
        return r;
    }, []);
    
    const groupByGroup = groupByModule.reduce((r,{buildCd:buildCd,buildName:buildName,groupCd:groupCd,groupName:groupName,oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.groupCd === groupCd);
        if (!temp) r.push(temp = { buildCd,buildName,groupCd,groupName,oldYn, moduleList: [] });
        temp.moduleList.push(object);
        return r;
    },[])
/*
    const groupBySector = groupByGroup.reduce((r,{floorCd:floorCd,floorName:floorName,sectorCd:sectorCd,sectorName:sectorName,oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.sectorCd === sectorCd);
        if (!temp) r.push(temp = { floorCd,floorName,sectorCd,sectorName,oldYn, groupList: [] });
        temp.groupList.push(object);
        return r;
    },[])

    const groupByFloor = groupBySector.reduce((r,{floorCd:floorCd,floorName:floorName,oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.floorCd === floorCd);
        if (!temp) r.push(temp = { floorCd,floorName,oldYn, sectorList: [] });
        temp.sectorList.push(object);
        return r;
    },[])
*/

    const groupByBuild = groupByGroup.reduce((r,{buildCd:buildCd,buildName:buildName,oldYn:oldYn, ...object})=>{
        var temp = r.find(o => o.buildCd === buildCd);
        if (!temp) r.push(temp = { buildCd,buildName,oldYn, groupList: [] });
        temp.groupList.push(object);
        return r;
    },[])
	var msg = "groupBy Complete : "+groupByBuild.length
    logs.writeLogs(msg,scriptName);
    return groupByBuild;
}


module.exports.groupByOld=groupByOld;
module.exports.groupByNew=groupByNew;