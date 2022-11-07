import * as fs from "fs";
import * as JSZip from "jszip";
import { parseLtxSax } from "./ltx_sax";
import { Xlsx_base } from "./Xlsx_base";
export class Xlsx_node extends Xlsx_base{
    protected async _loadData(data:Buffer) {
        var self = this;
        var zip = new JSZip();
        await zip.loadAsync(data);
        var alltmps = Object.keys(zip.files).map(async function(f) {
            var e = zip.files[f];
            if (!e.dir) {	
                var strf = await zip.file(f).async("string");
                try{
                    self._fe[f] = parseLtxSax(strf);
                }catch(err){
                    self._fe[f] = strf;
                }
                if(!self._fe[f]){
                    delete self._fe[f];
                }
            }
        });
        await Promise.all(alltmps);
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
    public async writeFile(filename:string){
        var data = await this.data();
        fs.writeFileSync(filename, data);
    }
    public async data() {
        var zip = new JSZip();
        var self = this;
        Object.keys(self._fe).forEach(function(f) {		
            typeof self._fe[f] == 'string' ? zip.file(f, self._fe[f]) : zip.file(f, self._fe[f].root().toString());		
        });
        return zip.generateAsync({
            type : "nodebuffer"
        });
    }
    
    public static async generateNew(sheets: Array<{ name: string, data: any[] }>):Promise<Xlsx_base>{
        var xls = new Xlsx_node();
        await xls._loadData(fs.readFileSync(__dirname+"/../tpl.xlsx"));
        return await xls.writeAll(sheets);
    } 
    public static async loadByFile(fileName:string):Promise<Xlsx_base>{
        return await this.loadByData(fs.readFileSync(fileName));
    }
    public static async loadByData(data:Buffer):Promise<Xlsx_base>{
        return await (new Xlsx_node())._loadData(data);
    } 
}