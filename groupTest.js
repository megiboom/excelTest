var json = require('./test.json')

function groupBy(arr){
    const result = arr.result.reduce((r, { moduleId: moduleId,groupId:groupId,groupName:groupName, ...object }) => {
        var temp = r.find(o => o.moduleId === moduleId);
        if (!temp) r.push(temp = { moduleId,groupId,groupName, data: [] });
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

const gArr = groupBy(json);
//console.log(gArr.length);
console.log(gArr[0].moduleList.length);
//console.log(JSON.stringify(gArr));

module.exports.groupBy=groupBy;