import Workbook from './Workbook';
import { select, select1 } from './XMLUtil';

export default class Styles {
    //linear equations for the row height given the font size and type
    private rowScaleParams = {
        'Arial': { a: 1.1776412347, b: 0.88468841, min: 12.8 },
        'default': { a: 1.1776412347, b: 0.88468841, min: 12.8 }
    }
    private allFonts: any[] = [];
    private allXfs: any[] = [];
    constructor(private sheet: Document, public workbook: Workbook) {
        this._readSheetData();
    }

    private _readSheetData = () => {
        this.allFonts = select(`//xl:font`, this.sheet)
            .map(this._readFont);
        this.allXfs = select(`//xl:cellXfs/xl:xf`, this.sheet)
            .map(this._readXf);
        // console.log(this.allFonts, this.allXfs);
    }

    private _readFont = (font: Element) => {
        return {
            sz: select1('number(xl:sz/@val)', font),
            name: select1('string(xl:name/@val)', font),
            family: select1('number(xl:family/@val)', font),
            charset: select1('number(xl:charset/@val)', font),
        }
    }

    private _readXf = (xf: Element) => {
        return {
            fontId: select1('number(@fontId)', xf)
        }
    }

    public getRowSize(styleId: number) {
        let font = this.allFonts[this.allXfs[styleId].fontId];
        let { a, b, min } = this.rowScaleParams[font.name] || this.rowScaleParams.default;
        return Math.max(a * font.sz + b, min);
    }
}
