import * as fs from "fs";
import { parseLtxSax } from "./ltx_sax";
import { Xlsx_base } from "./Xlsx_base";
export class Xlsx_fibjs extends Xlsx_base{
    protected _loadData(data:Buffer) {
        var self = this;
        var zfile = require("zip").open(data, "r");
        zfile.namelist().forEach(f=>{
            if(f.endsWith("/")){
                return;
            }
            var strf = zfile.read(f).toString();
            try{
                self._fe[f] = parseLtxSax(strf);
            }catch(err){
                self._fe[f] = strf;
            }
            if(!self._fe[f]){
                delete self._fe[f];
            }
        });
        // sharedStrings
        var el = this._fe['xl/sharedStrings.xml'];
        if (el) {
            var i = 0;
            el.children.forEach(function(si) {
                var t = si.getChildText('t');
                self._ss['' + i++] = t;
            });
        }
        return this;
    }
    public writeFile(filename:string){
        var data = this.data();
        fs.writeFileSync(filename, data);
    }
    public data() {
        var ms = new (require("io").MemoryStream)();
        var zfile = require("zip").open(ms, "w");
        var self = this;
        Object.keys(self._fe).forEach(function(f) {		
            typeof self._fe[f] == 'string' ? zfile.write(Buffer.from(self._fe[f]), f) : zfile.write(Buffer.from(self._fe[f].root().toString()), f);		
        });
        zfile.close();
        ms.rewind();
        return ms.readAll();
    }
    public static generateNew(sheets: Array<{ name: string, data: any[] }>):Xlsx_base{
        var xls = new Xlsx_fibjs();
        xls._loadData(fs.readFileSync(__dirname+"/../tpl.xlsx"));
        return xls.writeAll(sheets);
    } 
    public static loadByFile(fileName:string):Xlsx_base{
        return this.loadByData(fs.readFileSync(fileName));
    }
    public static loadByData(data:Buffer):Xlsx_base{
        return (new Xlsx_fibjs())._loadData(data);
    } 
}