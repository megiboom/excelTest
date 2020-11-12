var json = require('./test.json')

function groupBy(arr){
    const result = arr.result.reduce((r, { moduleId: moduleId,moduleName:moduleName,groupId:groupId,groupName:groupName, ...object }) => {
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

    return result1;
}
//const gArr = groupBy(json);
//console.log(gArr[0].moduleList[0].data);
//console.log(JSON.stringify(gArr[0]));

module.exports.groupBy=groupBy;