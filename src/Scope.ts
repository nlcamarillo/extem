import * as XLSX from './XLSXUtil';
import { getValue, parseTemplate, JsonataGlobals } from './Util';
import RootScope from './RootScope';

export default class Scope extends RootScope {
    private address: XLSX.RangeAddress;
    constructor(private range: string, private sheet: string, template: string, public type: string) {
        super(template);
        this.address = XLSX.getRangeAddress(range);
    }
    public grow = (a: XLSX.CellAddress) => {
        this.address = [this.address[0], XLSX.addAddress(a, this.address[1])];
        this.updateRange();
        if (this.parentScope instanceof Scope) {
            this.parentScope.grow(a);
        }
    }
    public growDim = (offset: number, direction: XLSX.Direction) => {
        return this.grow(XLSX.relAddress(offset, direction));
    }
    public move = (a: XLSX.CellAddress) => {
        this.address = this.address.map(ra => XLSX.addAddress(a, ra));;
        this.updateRange();
        //move children
        this.getChildren().forEach(c => c.move(a));
    }
    public moveDim = (offset: number, direction: XLSX.Direction) => {
        return this.move(XLSX.relAddress(offset, direction));
    }
    public dim = (direction: XLSX.Direction) => XLSX.getRangeDim(this.range, direction);
    public getRange = () => this.range;
    public getAnchor = () => XLSX.splitRange(this.range)[0];
    public getTemplate = () => this.template;
    public getSheet = () => this.sheet;

    private updateRange = () => {
        this.range = XLSX.getRangeRef(this.address);
    }

    public onSheet = (sheet: string): boolean => this.sheet === sheet;
    public containsCell = (ref: string) => XLSX.inRange(this.range)(ref);
    public containsRange = (range: string) => XLSX.getRangeCells(range).every(this.containsCell);
    public containsScope = (scope: Scope) => {
        return (scope.sheet === this.sheet) && this.containsRange(scope.range);
    }

    static isAfterDim = (ref: string, direction: XLSX.Direction) => (scope: Scope) => {
        return XLSX.rangeAfterCellDim(XLSX.getCellAddress(ref), scope.address, direction);
    }

    public isAtOrBelowRange = (ref: string) => {
        let [{ c: c1, r: r1 }, { c: c2 }] = XLSX.getRangeAddress(ref);
        let a = this.address;
        return (a[0].r >= r1) && (a[0].c <= c2) && (c1 <= a[1].c);
    }

    public isAtOrAfterRange = (ref: string) => {
        let [{ r: r1, c: c1 }, { r: r2 }] = XLSX.getRangeAddress(ref);
        let a = this.address;
        return (a[0].c >= c1) && (a[0].r <= r2) && (r1 <= a[1].r);
    }

    public isAtOrAfterRangeDim = (ref: string, direction: XLSX.Direction) => {
        switch (direction) {
            case XLSX.HORIZONTAL: return this.isAtOrAfterRange(ref);
            case XLSX.VERTICAL: return this.isAtOrBelowRange(ref);
        }
    }

    public evaluate = (context, globals?: JsonataGlobals) => {
        let { path, type } = parseTemplate(this.template);
        return {
            value: getValue(context, path, globals),
            type
        }
    }

    private clone = () => {
        let s = new Scope(this.range, this.sheet, this.template, this.type);
        this.getChildren().forEach(child => s.addChild(child.clone()));
        return s;
    }

    public cloneAsScalar = (index) => {
        let clone = this.clone();
        clone.template = '${$[' + index + ']}';
        return clone;
    }

    public makeScalar = () => {
        this.template = this.template.replace(/^[_|]/g, '$');
    }
}
