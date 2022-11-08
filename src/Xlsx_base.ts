import LtxElement from "./ltx_Element";

export abstract class Xlsx_base {
  constructor(protected _fe: any = {}, protected _ss: any = {}) {
  }
  protected _loadData(data: Buffer): Xlsx_base {
    return this;
  }
  protected _loadFile(filename: string): Xlsx_base {
    return this;
  }
  public writeFile(filename: string) {
  }
  public data(): Buffer {
    return null;
  }
  public getDownloadHeaders(filename: string) {
    return {
      'Content-Type': 'application/vnd.openxmlformats',
      'Content-Disposition': 'attachment; filename="'
        + encodeURIComponent(filename) + '.xlsx"'
    }
  }
  public readAll(withOutHidden: boolean = true) {
    var result: Array<{ i: number, name: string, data: Array<string | number | boolean> }> = [];
    this._fe['xl/workbook.xml'].getChild('sheets').children.forEach((e, i) => {
      if (e.attrs['state'] != "visible" && withOutHidden) {
        return;
      }
      var sname: string = e.attrs["name"];
      var s = this.getSheetByIndex(i);
      if (s != null) {
        result.push({
          i: i,
          name: sname,
          data: s.readAll()
        });
      }
    });
    return result;
  }

  public get sheetNum(): number {
    return this._fe['xl/workbook.xml'].getChild('sheets').children.length;
  }
  public get sheetNames(): string[] {
    return this._fe['xl/workbook.xml'].getChild('sheets').children.map(e => {
      return e.attrs["name"];
    });
  }
  public getSheetByIndex(i: number) {
    var el = this._fe['xl/worksheets/sheet' + (i + 1) + '.xml'];
    if (!el) return null;
    return new Xlsx_sheet(el, this._ss);
  }
  public getSheetByName(n: string) {
    return this.getSheetByIndex(this.sheetNames.indexOf(n));
  }

  public setSheetName(i: number, name: string) {
    this._fe['xl/workbook.xml'].getChild('sheets').children[i].attrs['name'] = name;
  }
  public setSheetVisible(sheetIndex: number, visible: boolean) {
    var state = visible ? 'visible' : 'hidden';
    this._fe['xl/workbook.xml'].getChild('sheets').children[sheetIndex].attrs['state'] = state;
  }
  public isSheetVisible(sheetIndex: number) {
    var e = this._fe['xl/workbook.xml'].getChild('sheets').children[sheetIndex];
    return e && e.attrs['state'] == "visible";
  }
  public copySheet(i: number, name: string) {
    if (i > 0 && i < this.sheetNum - 1) {
      return false;
    }
    var n = this.sheetNum + 1;
    this._fe['[Content_Types].xml'].c('Override', { PartName: '/xl/worksheets/sheet' + n + '.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' });
    this._fe['xl/workbook.xml'].getChild('sheets').c('sheet', { name: name, sheetId: n, state: 'visible', 'r:id': 'sheetrId' + n });
    this._fe['xl/_rels/workbook.xml.rels'].c('Relationship', { Id: 'sheetrId' + n, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target: 'worksheets/sheet' + n + '.xml' });
    this._fe['xl/worksheets/sheet' + n + '.xml'] = this._fe['xl/worksheets/sheet' + (i + 1) + '.xml'].clone();
    var srcRels = this._fe['xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels'];
    if (srcRels) this._fe['xl/worksheets/_rels/sheet' + n + '.xml.rels'] = srcRels.clone();
    return true;
  }

  protected writeAll(sheets: Array<{ name: string, data: any[] }>) {
    var end = sheets.length - 1;
    for (var i = 0; i < sheets.length; i++) {
      if (i < end) {
        this.copySheet(i, "tmp");
      }
      this.setSheetName(i, sheets[i].name);
      var xss = this.getSheetByIndex(i);
      var xdata = sheets[i].data;
      var max_row = xdata.length;
      var max_col = 1;
      for (var j = 0; j < xdata.length; j++) {
        xss.write("A" + (j + 1), xdata[j]);
        if(Array.isArray(xdata[j])){
          max_col = Math.max(xdata[j].length, max_col);
        }
      }
      var ref = "A1:"+(Xlsx_base.dsum26(max_col))+max_row;
      xss._writeDimension(ref);
    }
    return this;
  }
  private static dsum26(num:number){
    let res = [];
    // 短除法，注意余数为0时，将商减1，对应字母'z'
    while (num > 26) {
      res.push(this.D26[num % 26]);
      num = Math.floor(num / 26);
    }
    // 不要忘了末位
    res.push(this.D26[num]);
    // 倒置
    return res.reverse().join('');
  }
  private static D26 = 'ZABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
  public static generateNew(sheets: Array<{ name: string, data: any[] }>): Xlsx_base {
    return null;
  }
  public static loadByFile(fileName: string): Xlsx_base {
    return null;
  }
  public static loadByData(data: Buffer): Xlsx_base {
    return null;
  }
}


export class Xlsx_sheet {
  constructor(private _el: LtxElement, private _ss: any = {}) {
  }
  public read(ref: string) {
    var splt = ref.split(":");
    if (1 == splt.length) {
      return this._readCell(ref);
    } else {
      return this._readRange(ref);
    }
  }
  public readAll(): Array<string | number | boolean> {
    return <any[]>this.read(this.dimension());
  }
  public dimension(): string {
    var em = this._get_dimension();
    var ref:string;
    if(em!=null){
      var t = em.attr("ref");
      if(t){
        ref = t.toString();
      }
    }
    return ref||":";
  }
  _writeDimension(ref: string) {
    var em = this._get_dimension();
    em && em.attr("ref", ref);
  }
  private _get_dimension():LtxElement{
    for (var i = 0; i <= this._el.children.length; i++) {
      if ('dimension' == this._el.children[i].name) {
        return this._el.children[i];
      }
    }
    return null;
  }
  _readRange(range) {
    range = decode_range(range);

    if (range.s.r == range.e.r) {
      return this._readRow(range.s.r, range.s.c, range.e.c);
    } else if (range.s.c == range.e.c) {
      return this._readCol(range.s.c, range.s.r, range.e.r);
    } else {
      var ret = [];
      for (var r = range.s.r; r <= range.e.r; r++) {
        var row = this._readRow(r, range.s.c, range.e.c);
        ret.push(row);
      }
      return ret;
    }
  }
  _readRow(r, sc, ec) {
    var row = this._el.getChild('sheetData').getChildByAttr('r', '' + r);
    sc = decode_col(sc);
    ec = decode_col(ec);
    var ret = [];
    for (var i = sc; i <= ec; i++) {
      if (row) {
        var cell = row.getChildByAttr('r', encode_col(i) + r);
        ret.push(this._cellv(cell));
      } else {
        ret.push('');
      }
    }
    return ret;
  }
  _readCol(c, sr, er) {
    var ret = [];
    var sd = this._el.getChild('sheetData');
    for (var i = sr; i <= er; i++) {
      var row = sd.getChildByAttr('r', '' + i);
      var cell = row.getChildByAttr('r', c + i);
      ret.push(this._cellv(cell));
    }
    return ret;
  }
  _readCell(cell) {
    var cr = split_cell(cell);
    var r = this._el.getChild('sheetData').getChildByAttr('r', '' + cr[1]);
    var v = '';
    if (r) {
      var c = r.getChildByAttr('r', cell);
      v = this._cellv(c);
    }
    return v;
  }
  _cellv(c) {
    if (!c) return '';
    var v: any = '';
    var t = c.attrs['t'];
    v = c.getChildText('v');
    switch (t) {
      case 'n':
        v = parseFloat(v);
        break;
      case 'b':
        v = v == 1 ? true : false;
        break;
      case 's':
        v = this._ss[v];
        break;
    }
    return v;
  }
  public write(cell, v, append?: boolean) {
    var self = this;
    var cr = split_cell(cell);
    if (Array.isArray(v)) {
      var sr = cr[1]; // 开始行
      var sc = decode_col(cr[0]); // 开始列 int
      if (append) {
        self._writeRow(sr, sc, v);
      } else {
        v.forEach(function (r) {
          if (Array.isArray(r)) {
            var c = sc;
            r.forEach(function (cv) {
              self._writeCell(sr, encode_col(c++), cv);
            });
            sr++;
          } else {
            self._writeCell(sr, encode_col(sc++), r);
          }
        });
      }
    } else {
      self._writeCell(cr[1], cr[0], v);
    }
  }
  _v2cell(v) {
    var cv = {
      v: v,
      s: 0,
      t: 'str'
    };
    var vt = typeof v;
    switch (vt) {
      case 'string': break;
      case 'boolean': cv.t = 'b'; cv.v = v ? 1 : 0; break;
      case 'number': cv.t = 'n'; break;
      default: break;
    }
    return cv;
  }
  _writeCell(ri, c, v) {
    var sd = this._el.getChild('sheetData');
    var r = sd.getChildByAttr('r', '' + ri);
    var cr = '' + c + ri;
    var cv = this._v2cell(v);
    if (r) {
      var c = r.getChildByAttr('r', cr);
      if (c) {
        c.attr('t', cv.t);
        var rcv = c.getChild('v');
        rcv ? rcv.text(cv.v) : c.c('v').t(cv.v);
      } else {
        r.c('c', {
          r: cr,
          s: cv.s,
          t: cv.t
        }).c('v').t(cv.v);
      }
    } else {
      sd.c('row', {
        r: '' + ri
      }).c('c', {
        r: cr,
        s: cv.s,
        t: cv.t
      }).c('v').t(cv.v);
    }
  }
  _writeRow(sr, sc, rows) {
    sr = parseInt(sr);
    var sd = this._el.getChild('sheetData');
    var srow = sd.c('row', { r: '' + sr });
    var self = this;
    rows.forEach(function (r, ri) {
      if (Array.isArray(r)) {
        var wr = ri == 0 ? srow : sd.c('row', { r: '' + (sr + ri) });
        var c = sc;
        r.forEach(function (rcv) {
          var cv = self._v2cell(rcv);
          var cr = '' + encode_col(c++) + (sr + ri);

          wr.c('c', {
            r: cr,
            s: cv.s,
            t: cv.t
          }).c('v').t(cv.v);
        });
      } else {
        var cv = self._v2cell(r);
        var cr = '' + encode_col(sc++) + sr;
        srow.c('c', {
          r: cr,
          s: cv.s,
          t: cv.t
        }).c('v').t(cv.v);
      }
    });
  }
}

function decode_col(colstr) {
  var c = colstr.replace(/^\$([A-Z])/, "$1"), d = 0, i = 0;
  for (; i !== c.length; ++i)
    d = 26 * d + c.charCodeAt(i) - 64;
  return d - 1;
};

function encode_col(col) {
  var s = "";
  for (++col; col; col = Math.floor((col - 1) / 26))
    s = String.fromCharCode(((col - 1) % 26) + 65) + s;
  return s;
};

function split_cell(cstr) {
  return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
};

function decode_range(range) {
  var x = range.split(":").map(function (cell) {
    var splt = split_cell(cell);
    return {
      c: splt[0],
      r: parseInt(splt[1])
    };
  });
  return {
    s: x[0],
    e: x[x.length - 1]
  };
}