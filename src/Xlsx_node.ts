import * as fs from "fs";
import { parseLtxSax } from "./ltx_sax";
import { Xlsx_base } from "./Xlsx_base";


export class Xlsx_node extends Xlsx_base {
  protected _loadData(data: Buffer) {
    var self = this;

    require("./ZipFile").ZipFile.fromBuffer(data).entries.forEach(entry=>{
        if(!entry.isDirectory){
          var strf = entry.data.toString();
          var f = entry.fileName;
          try {
            self._fe[f] = parseLtxSax(strf);
          } catch (err) {
            self._fe[f] = strf;
          }
          if (!self._fe[f]) {
            delete self._fe[f];
          }
        }
    });


    // var zip = newJSZip();
    // await zip.loadAsync(data);
    // await Promise.all(Object.keys(zip.files).map(async function (f) {
    //   var e = zip.files[f];
    //   if (!e.dir) {
    //     var strf = await zip.file(f).async("string");
    //     try {
    //       self._fe[f] = parseLtxSax(strf);
    //     } catch (err) {
    //       self._fe[f] = strf;
    //     }
    //     if (!self._fe[f]) {
    //       delete self._fe[f];
    //     }
    //   }
    // }));

    // sharedStrings
    var el = this._fe['xl/sharedStrings.xml'];
    if (el) {
      var i = 0;
      el.children.forEach(function (si) {
        var t = si.getChildText('t');
        self._ss['' + i++] = t;
      });
    }
    return this;
  }
  public async writeFile(filename: string) {
    var data = await this.data();
    fs.writeFileSync(filename, data);
  }
  public data():Buffer {
    var self = this;

    // var zip = newJSZip();
    // Object.keys(self._fe).forEach(function (f) {
    //   typeof self._fe[f] == 'string' ? zip.file(f, self._fe[f]) : zip.file(f, self._fe[f].root().toString());
    // });
    // return zip.generateAsync({
    //   type: "nodebuffer"
    // });

    var zfile = new (require("./ZipFile").ZipFile)();
    Object.keys(self._fe).forEach(f=> {
      var zen = new (require("./ZipFile").ZipEntry)();
      zen.fileName = f;
      if(typeof self._fe[f]=="string"){
        zen.data = Buffer.from(self._fe[f]);
      }else{
        zen.data = Buffer.from(self._fe[f].root().toString());
      }
      zfile.addEntry(zen);
    });
    return zfile.compress();
  }

  public static generateNew(sheets: Array<{ name: string, data: any[] }>): Xlsx_base {
    var xls = new Xlsx_node();
    xls._loadData(fs.readFileSync(__dirname + "/../tpl.xlsx"));
    return xls.writeAll(sheets);
  }
  public static loadByFile(fileName: string): Xlsx_base {
    return this.loadByData(fs.readFileSync(fileName));
  }
  public static loadByData(data: Buffer): Xlsx_base {
    return (new Xlsx_node())._loadData(data);
  }
}