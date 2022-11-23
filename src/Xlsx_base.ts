import LtxElement from "./ltx_Element";

export abstract class Xlsx_base {
    constructor(protected _fe: { [index: string]: LtxElement } = {}, protected _ss: { [index: string]: string } = {}) {
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
        var result: Array<{ i: number, name: string, data: Array<Array<Xlsx_Val>> }> = [];
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
        var el = this._fe[this._link_xml(i)];//this._fe['xl/worksheets/sheet' + (i + 1) + '.xml'];
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
        return e && e.attrs['state'] != "hidden";
    }

    public copySheet(i: number, name: string) {
        if (i > 0 && i < this.sheetNum - 1) {
            return false;
        }
        let srcXml = this._link_xml(i);
        if (!srcXml || !this._fe[srcXml]) {
            return false;
        }
        let nTo = this._next_id_n();//this.sheetNum + 1;
        this._fe['[Content_Types].xml'].c('Override', { PartName: '/xl/worksheets/sheet' + nTo + '.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' });
        this._fe['xl/workbook.xml'].getChild('sheets').c('sheet', { name: name, sheetId: nTo, state: 'visible', 'r:id': 'sId' + nTo });
        this._fe['xl/_rels/workbook.xml.rels'].c('Relationship', { Id: 'sId' + nTo, Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target: 'worksheets/sheet' + nTo + '.xml' });
        this._fe['xl/worksheets/sheet' + nTo + '.xml'] = this._fe[srcXml].clone();
        let src_refs = srcXml.replace("worksheets/", "worksheets/_rels/")
        if (this._fe[src_refs]) this._fe['xl/worksheets/_rels/sheet' + nTo + '.xml.rels'] = this._fe[src_refs].clone();
        return true;
    }
    private _next_id_n() {
        let ens: number[] = [];
        this._fe['xl/_rels/workbook.xml.rels'].getChildren("Relationship").forEach(e => {
            var t = e.attr("Target");
            if (t && t.startsWith("worksheets/sheet")) {
                ens.push(parseInt((e.attr("Target") + "").replace(/\D/g, "")));
            }
        });
        if (ens.length < 1) {
            return this.sheetNum + 1;
        }
        ens = ens.sort();
        return ens[ens.length - 1] + 1;
    }
    private _link_xml(i: number) {
        let rid: string = this._fe['xl/workbook.xml'].getChild('sheets').children[i].attr("r:id");
        let rem = this._fe['xl/_rels/workbook.xml.rels'].getChildren("Relationship").find(e => e.attr("Id") == rid);
        return rem ? "xl/" + rem.attr("Target") : null;
    }
    /**
     * 删除sheet
     * @param i 索引 或 名字
     */
    public removeSheet(i: number | string) {
        let tmpNames = this.sheetNames;
        let j = typeof (i) == "string" ? tmpNames.indexOf(i.toString()) : <number>i;
        if (j < 0 || j >= tmpNames.length) {
            return false;
        }
        let Jname = tmpNames[j];
        let Jid = j + 1;
        if (tmpNames.length == 1) {
            this.getSheetByIndex(0).empty();
        } else {
            let CT_em = this._fe['[Content_Types].xml'].getChildren("Override").find(e => {
                var pname: string = e.attr("PartName");
                if (pname.startsWith("/xl/worksheets/sheet") && pname.endsWith(".xml")) {
                    return parseInt(pname.replace(/\D/g, "")) == Jid
                }
                return false;
            });
            if (CT_em) {
                this._fe['[Content_Types].xml'].remove(CT_em);
            }
            let WS_em = this._fe['xl/workbook.xml'].getChild('sheets').children.find(e => {
                if (e.attr("name") == Jname) {
                    return true;
                }
                return false;
            });
            if (WS_em) {
                let WS_em_rid: string = WS_em.attr("r:id");
                this._fe['xl/workbook.xml'].getChild('sheets').remove(WS_em);
                let WS_em_releation = this._fe['xl/_rels/workbook.xml.rels'].getChildren("Relationship").find(e => e.attr("Id") == WS_em_rid);
                if (WS_em_releation) {
                    this._fe['xl/_rels/workbook.xml.rels'].remove(WS_em_releation);
                }
            }
            delete this._fe["xl/worksheets/sheet" + Jid + ".xml"];
            delete this._fe['xl/worksheets/_rels/sheet' + Jid + '.xml.rels'];
        }
    }
    public swapSheetByIndex(index1: number, index2: number) {
        var array = this._fe['xl/workbook.xml'].getChild('sheets').children;
        [array[index1], array[index2]] = [array[index2], array[index1]];
    }
    protected writeAll(sheets: Array<{ name: string, data: any[] }>) {
        var end = sheets.length - 1;
        for (let i = 0; i < sheets.length; i++) {
            if (i < end) {
                this.copySheet(i, "tmp");
            }
            this.setSheetName(i, sheets[i].name);
            this.getSheetByIndex(i).writeAll(sheets[i].data);
        }
        return this;
    }

    public static generateNew(sheets: Array<{ name: string, data: any[] }>): Xlsx_base {
        return null;
    }
    public static generateFast(data: any[], sheetName: string = "Sheet1"): Xlsx_base {
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
    /**
     * 读取ref范围的数据
     * @param ref "A1" or "A1:B3"
     * @returns 
     */
    public read(ref: string = this.dimension()): Xlsx_Val | Array<Xlsx_Val> | Array<Array<Xlsx_Val>> {
        var splt = ref.split(":");
        if (1 == splt.length) {
            return this._readCell(ref);
        } else {
            return this._readRange(ref);
        }
    }
    /**
     * 覆盖式写入数据
     * @param xdata 
     */
    public writeAll(xdata: any[]) {
        this.empty();
        let max_row = xdata.length;
        let max_col = 1;
        for (let j = 0; j < xdata.length; j++) {
            this.write("A" + (j + 1), xdata[j]);
            if (Array.isArray(xdata[j])) {
                max_col = Math.max(xdata[j].length, max_col);
            }
        }
        let ref = "A1:" + (encode_col(max_col - 1)) + max_row;
        this._writeDimension(ref);
    }
    /**
     * 读取本sheet内数据
     */
    public readAll(): Array<Array<Xlsx_Val>> {
        let r: any = this.read();
        if (!Array.isArray(r)) {
            r = [[r]];
        } else if (!Array.isArray(r[0])) {
            r = [r];
        }
        return r;
    }
    /**
     * 数据标记范围
     * @returns A1 OR A1:B3
     */
    public dimension(): string {
        var em = this._get_dimension();
        var ref: string;
        if (em != null) {
            var t = em.attr("ref");
            if (t) {
                ref = t.toString();
            }
        }
        return ref || ":";
    }
    _writeDimension(ref: string) {
        var em = this._get_dimension();
        em && em.attr("ref", ref);
    }
    private _get_dimension(): LtxElement {
        for (var i = 0; i <= this._el.children.length; i++) {
            if ('dimension' == this._el.children[i].name) {
                return this._el.children[i];
            }
        }
        return null;
    }
    _readRange(range: string) {
        const ref = decode_range(range);

        if (ref.s.r == ref.e.r) {
            return this._readRow(ref.s.r, ref.s.c, ref.e.c);
        } else if (ref.s.c == ref.e.c) {
            return this._readCol(ref.s.c, ref.s.r, ref.e.r);
        } else {
            var ret = [];
            for (var r = ref.s.r; r <= ref.e.r; r++) {
                ret.push(this._readRow(r, ref.s.c, ref.e.c));
            }
            return ret;
        }
    }
    _readRow(r: number, sc: string, ec: string) {
        let row = this._el.getChild('sheetData').getChildByAttr('r', '' + r),
            ret: Xlsx_Val[] = [];
        for (var i = decode_col(sc); i <= decode_col(ec); i++) {
            if (row) {
                var cell = row.getChildByAttr('r', encode_col(i) + r);
                ret.push(this._cellv(cell));
            } else {
                ret.push('');
            }
        }
        return ret;
    }
    _readCol(c: string, sr: number, er: number) {
        var ret: Xlsx_Val[] = [],
            sd = this._el.getChild('sheetData');
        for (var i = sr; i <= er; i++) {
            var row = sd.getChildByAttr('r', i.toString());
            var cell = row.getChildByAttr('r', c + i);
            ret.push(this._cellv(cell));
        }
        return ret;
    }
    _readCell(cell: string): Xlsx_Val {
        var cr = split_cell(cell);
        var r: LtxElement = this._el.getChild('sheetData').getChildByAttr('r', '' + cr[1]);
        var v: Xlsx_Val = '';
        if (r) {
            var c = r.getChildByAttr('r', cell);
            v = this._cellv(c);
        }
        return v;
    }
    _cellv(c: LtxElement): Xlsx_Val {
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
    /**
     * 清空本sheet的数据
     */
    public empty() {
        this._writeDimension("A1");
        (this._el.getChild('sheetData') as LtxElement).children = [];
    }
    write(cell: string, v: any, append?: boolean) {
        var self = this;
        var cr = split_cell(cell);
        var cr1n = parseInt(cr[1]);
        if (Array.isArray(v)) {
            var sr = cr1n; // 开始行
            var sc = decode_col(cr[0]); // 开始列 int
            if (append) {
                self._writeRow(cr[1], sc, v);
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
            self._writeCell(cr1n, cr[0], v);
        }
    }
    _v2cell(v: Xlsx_Val) {
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
    _writeCell(ri: number, c: string, v: Xlsx_Val) {
        let sd = this._el.getChild('sheetData');
        let r = sd.getChildByAttr('r', ri.toString());
        let cr = (c + ri).toString();
        let cv = this._v2cell(v);
        if (r) {
            let ce = r.getChildByAttr('r', cr);
            if (ce) {
                ce.attr('t', cv.t);
                let rcv = ce.getChild('v');
                rcv ? rcv.text(cv.v) : ce.c('v').t(cv.v);
            } else {
                r.c('c', {
                    r: cr,
                    s: cv.s,
                    t: cv.t
                }).c('v').t(cv.v);
            }
        } else {
            sd.c('row', {
                r: ri.toString()
            }).c('c', {
                r: cr,
                s: cv.s,
                t: cv.t
            }).c('v').t(cv.v);
        }
    }
    _writeRow(sr_: string, sc: number, rows: any[]) {
        var sr = parseInt(sr_);
        var sd: LtxElement = this._el.getChild('sheetData');
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

export type Xlsx_Val = string | number | boolean;

function decode_col(colstr: string) {
    var c = colstr.replace(/^\$([A-Z])/, "$1"), d = 0, i = 0;
    for (; i !== c.length; ++i)
        d = 26 * d + c.charCodeAt(i) - 64;
    return d - 1;
}

function encode_col(col: number) {
    var s = "";
    for (++col; col; col = Math.floor((col - 1) / 26))
        s = String.fromCharCode(((col - 1) % 26) + 65) + s;
    return s;
}

function split_cell(cstr: string): string[] {
    return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
}

function decode_range(range: string): { s: { c: string, r: number }, e: { c: string, r: number } } {
    let x = range.split(":").map(function (cell) {
        let splt = split_cell(cell);
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