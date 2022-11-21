**lite xlsx in fibjs or nodejs**
**简单的xlsx库**

**新建写入**
```
const PlayXlsx=require("play_xlsx").PlayXlsx;

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

PlayXlsx.generateNew(xlsxData).writeFile("test3.xlsx");
```

**读取**
```
const PlayXlsx=require("play_xlsx").PlayXlsx;
var xls = PlayXlsx.loadByFile("test3.xlsx");
console.log(xls.sheetsNum, xls.sheetsNames, xls.isSheetVisible(0))
var sheet = xls.getSheetByIndex(0);
console.log(sheet.dimension(), JSON.stringify(sheet.readAll()))
```

根据自己环境（非fibjs）    
xlsx改自 https://github.com/lodengo/xlsx   
zip原作者  https://github.com/Teal/TUtils/blob/master/src/zipFile.ts 

**和sheetjs区别**
功能少    
提升require的速度