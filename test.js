const PlayXlsx=require('./lib/index').PlayXlsx;

var fileName = "test3.xlsx";
var xlsxData = [{
    name:"t1",
    data:[
        ["id","name","age","flag"],
        [1,"zhh",34,true],
        [2,"fz",18,false],
        [3,"ls",30,true],
    ]
},{
    name:"t2",
    data:[
        ["kind","coin"],
        [200,100],
        [204,383],
        [205,"tmp_mark：我去"],
    ]
} ];

PlayXlsx.generateNew(xlsxData).writeFile(fileName);

var xls = PlayXlsx.loadByFile(fileName);

console.log(xls.sheetNum==xlsxData.length);
console.log(JSON.stringify(xls.getSheetByIndex(0).readAll())==JSON.stringify(xlsxData[0].data));
console.log(JSON.stringify(xls.getSheetByIndex(1).readAll())==JSON.stringify(xlsxData[1].data));



