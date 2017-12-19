import { templateType, matchRangeTemplate } from './Util';
import { select1, nodeAttributes, NodeAttributes } from './XMLUtil';
import * as XLSX from './XLSXUtil';
import WorkSheet from './WorkSheet';

export default class Cell {
    constructor(private attributes: NodeAttributes, private value: string, private form: string, public worksheet: WorkSheet) { };

    // attributes
    public getAttributes = () => this.attributes;
    public ref = () => this.attributes.r;
    public setRef = (ref: XLSX.CellRef) => this.attributes.r = ref;
    public type = () => this.attributes.t;
    public style = () => parseInt(this.attributes.s, 10);
    public getValue = () => this.value;
    public setValue = (value: string) => this.value = value
    public formula = () => this.form;
    public address = () => XLSX.getCellAddress(this.ref());

    // cell operations
    public moveTo = (to: XLSX.CellRef) => this.worksheet._moveCell(this.ref(), to);
    public moveBy = (a: XLSX.CellAddress) => this.moveTo(XLSX.getCellRef(XLSX.addAddress(a, this.address())));
    public cloneTo = (to: XLSX.CellRef) => this.worksheet._cloneCell(this.ref(), to);

    public copy = (attrs) => new Cell({ ...this.attributes, ...attrs }, this.value, this.form, this.worksheet);

    static fromNode = (node: Element, sheet: WorkSheet) => {
        let attributes = nodeAttributes(node);
        let value = select1<string>('string(xl:v)', node);
        let formula = select1<string>('string(xl:f)', node);
        let cellValue: any = value;
        if (attributes.t === 's') {
            cellValue = sheet.workbook.getString(parseInt(value, 10));
            attributes.t = 'str';
        }
        return new Cell(attributes, cellValue, formula, sheet);
    }

    // cell predicates
    static at = (ref: string) => (cell: Cell) => cell.ref() === ref;
    static atOrAfterDim = (ref: string, direction: XLSX.Direction) => (cell: Cell) => {
        return XLSX.cellAtOrAfterDim(cell.address(), XLSX.getCellAddress(ref), direction);
    }
    static atOrAfterRangeDim = (ref: string, direction: XLSX.Direction) => (cell: Cell) => {
        return XLSX.cellAtOrAfterRangeDim(cell.address(), XLSX.getRangeAddress(ref), direction);
    }
    static isCellTemplate = (cell: Cell) => !!templateType(cell.getValue());
    static isRangeTemplate = (cell: Cell) => !!matchRangeTemplate(cell.formula());
    
    //sorter
    static sortDim = (axis: XLSX.Direction, reverse = false) => (cella: Cell, cellb: Cell) => {
        let aa = cella.address();
        let ab = cellb.address();
        if (aa[axis] === ab[axis]) return 0;
        return ((aa[axis] < ab[axis]) !== reverse) ? -1 : 1;
    }

}
