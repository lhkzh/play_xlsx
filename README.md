**lite xlsx in fibjs or nodejs**
**简单的xlsx库**

**读取**
```
const PlayXlsx=require("play_xlsx").PlayXlsx;
var xls = new PlayXlsx();
xls.loadFile("test1.xlsx").then(()=>{
    console.log(xls.sheetsNum, xls.sheetsNames, xls.isSheetVisible(0))
    var sheet = xls.getSheetByIndex(0);
    console.log(sheet.dimension(), JSON.stringify(sheet.readAll()))
});
```
**新建写入**
```
const PlayXlsx=require("play_xlsx").PlayXlsx;
var xls = await PlayXlsx.generateNew([{
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
        [205,399,"tmp_mark"],
    ]
} ]);
xls.writeFile("test2.xlsx");
根据自己环境（非fibjs）需要安装依赖jszip    
参考 https://github.com/lodengo/xlsx