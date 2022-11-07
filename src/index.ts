
import { Xlsx_base, Xlsx_sheet } from "./Xlsx_base";
import { Xlsx_fibjs } from "./Xlsx_fibjs";
import { Xlsx_node } from "./Xlsx_node";

const PlayXlsx: typeof Xlsx_base = global["process"]["versions"]["fibjs"]!==undefined ? Xlsx_fibjs:Xlsx_node;
const PlaySheet: typeof Xlsx_sheet = Xlsx_sheet;

export {
    PlayXlsx,
    PlaySheet
};