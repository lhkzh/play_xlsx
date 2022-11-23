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
PlayXlsx.generateFast(xlsxData[0].data).writeFile("test4.xlsx");
```

**读取**
```
const PlayXlsx=require("play_xlsx").PlayXlsx;
var book = PlayXlsx.loadByFile("test3.xlsx");
console.log(book.sheetsNum, book.sheetsNames, book.isSheetVisible(0));
var sheet = book.getSheetByIndex(0);
console.log(sheet.dimension(), JSON.stringify(sheet.readAll()));
console.log(sheet.dimension(), JSON.stringify(sheet.read()));
```

**一些操作**
```
//获取有几个sheet
book.sheetNum
//获取所有sheet的名字
book.sheetNames
//交换sheet顺序
book.swapSheetByIndex 
//根据索引或名字,删除sheet
book.removeSheet 
//修改sheet名字
book.setSheetName
//通过索引获取sheet
book.getSheetByIndex
//通过name获取sheet
book.getSheetByName
//写入文件
book.writeFile

//写入数据
sheet.writeAll()
//标准化读取所有
sheet.readAll()
//根据参数不同随缘返回
sheet.read()
```
       
xlsx改自 https://github.com/lodengo/xlsx   
zip原作者  https://github.com/Teal/TUtils/blob/master/src/zipFile.ts 

**和sheetjs区别**    
功能少    
require效率高
都是同步方法