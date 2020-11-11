var groupBy = require('json-groupby')
var json = require('./test.json')

var groupJ = groupBy(json.result,['groupId'],['deviceId','channelNo'])
//console.log(groupJ)

var obj = {
    "deviceId": 117,
    "groupId": 1,
    "channelNo": 3,
    "current": 100,
    "activePowerQty": 200,
    "regDate": "2020-11-12 10:00:00"
}
//json.result.push(obj)
var result2 = json.result.reduce((r,{groupId: name, ...object})=>{
    //console.log(groupId)
    var temp = r.find(o => {
        o.name === name
    });
    if(!temp) r.push(temp = {name, children: []});
    
    temp.children.push(object);
    return r;
},[])
console.log(result2)

var array = [{ name: "cat", value: 17, group: "animal" }, { name: "dog", value: 6, group: "animal" }, { name: "snak", value: 2, group: "animal" }, { name: "tesla", value: 11, group: "car" }, { name: "bmw", value: 23, group: "car" }],
    result = array.reduce((r, { group: name, ...object }) => {
        var temp = r.find(o => o.name === name);
        if (!temp) r.push(temp = { name, children: [] });
        temp.children.push(object);
        //console.log(r)
        return r;
    }, []);
    
//console.log(result)