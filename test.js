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
},{
    name:"t3",
    data:[
        "aa",
        "bb",
        235
    ]
} ];

PlayXlsx.generateNew(xlsxData).writeFile(fileName);

var book = PlayXlsx.loadByFile(fileName);
console.log(JSON.stringify(book.sheetNames)==JSON.stringify(xlsxData.map(e=>e.name)));
console.log(JSON.stringify(book.getSheetByIndex(0).read())==JSON.stringify(xlsxData[0].data));
console.log(JSON.stringify(book.getSheetByIndex(1).read())==JSON.stringify(xlsxData[1].data));
console.log(JSON.stringify(book.getSheetByIndex(2).read())==JSON.stringify(xlsxData[2].data));

book.removeSheet(1);
book.writeFile(fileName);
console.log(JSON.stringify(PlayXlsx.loadByFile(fileName).getSheetByIndex(1).read())==JSON.stringify(xlsxData[2].data));
book.copySheet(1,"t2");
book.getSheetByName("t2").writeAll(xlsxData[1].data);
book.writeFile(fileName);
xls = PlayXlsx.loadByFile(fileName);
console.log(JSON.stringify(book.getSheetByIndex(0).read())==JSON.stringify(xlsxData[0].data));
console.log(JSON.stringify(book.getSheetByIndex(1).read())==JSON.stringify(xlsxData[2].data));
console.log(JSON.stringify(book.getSheetByIndex(2).read())==JSON.stringify(xlsxData[1].data));

book.swapSheetByIndex(1,2);
console.log(JSON.stringify(book.sheetNames)==JSON.stringify(xlsxData.map(e=>e.name)));
console.log(JSON.stringify(book.getSheetByIndex(0).read())==JSON.stringify(xlsxData[0].data));
console.log(JSON.stringify(book.getSheetByIndex(1).read())==JSON.stringify(xlsxData[1].data));
console.log(JSON.stringify(book.getSheetByIndex(2).read())==JSON.stringify(xlsxData[2].data));

// PlayXlsx.generateFast(xlsxData[0].data, "t3").writeFile(fileName);
// var xls = PlayXlsx.loadByFile(fileName);
// console.log(JSON.stringify(book.sheetNames)=='["t3"]');
// console.log(JSON.stringify(book.getSheetByIndex(0).readAll())==JSON.stringify(xlsxData[0].data));
