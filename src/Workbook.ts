import { TYPE_RANGE, TYPE_CELL, TYPE_SCALAR, TYPE_ROW, TYPE_COLUMN } from './types';
import WorkSheet from './WorkSheet';
import Scope from './Scope';
import RootScope from './RootScope';
import Styles from './Styles';
import Cell from './Cell';
import { readZipFile, readData, writeZipFile, writeStream, contains, matchRangeTemplate } from './Util';
import * as XLSX from './XLSXUtil';
import { select, select1, strToXml, xmlToStr, nodeAttribute } from './XMLUtil';
import * as Util from './Util';

//not really sure where to keep this
const parseScopeCell = (sheet: string) => (cell: Cell): Scope => {
    let ref = cell.ref();
    return new Scope(`${ref}:${ref}`, sheet, cell.getValue(), TYPE_CELL);
}
const parseScopeRange = (sheet: string) => (cell: Cell): Scope => {
    let match = matchRangeTemplate(cell.formula());
    //should always match as we only accept range scope cells here
    if (!match) return new Scope('', '', '', '');
    return new Scope(match[1], sheet, match[2], TYPE_RANGE);
}

export interface EvaluateOptions {
    i18n?: {
        dayNames: string[];
        monthNames: string[];
        timeNames: string[];
    };
    globals?: Util.JsonataGlobals
}

export default class Workbook {
    private zip: any;
    private sheets: Document[];
    private worksheets: WorkSheet[];
    public styles: Styles;
    private workbook: Document;
    private strings: Document;
    private rels: Document;
    private scopes: Scope[];
    private rootScope: RootScope;
    private globals?: Util.JsonataGlobals;
    constructor() { }

    private readZip = async (zip) => {
        this.zip = zip;
        //parse workbook, sheets and strings
        this.rels = await this.readXML('_rels/workbook.xml.rels');
        this.workbook = await this.readXML('workbook.xml');
        this.strings = await this.readXML('sharedStrings.xml');
        this.sheets = await Promise.all(this.sheetPaths().map(this.readXML));
        this.worksheets = this.sheetNames().map((name, index) => {
            return new WorkSheet(this.sheets[index], this, name)
        });
        this.styles = new Styles(await this.readXML('styles.xml'), this);
        //get the ranges and remove them from the sheet
        this.scopes = Array<Scope>().concat(
            this.getRangeScopes(),
            this.getCellScopes()
        );
        // console.log(this.scopes);
        this.rootScope = this.createScopeTree(this.scopes);
        return this;
    }

    //read and write
    public async readFile(filename: string) {
        return this.readZip(await readZipFile(filename));
    }

    public async read(data) {
        return this.readZip(await readData(data));
    }

    private writeZip() {
        //write back workbook, sheets and strings
        this.writeXML('workbook.xml', this.workbook);
        this.writeXML('sharedStrings.xml', this.strings);
        // this.sheetPaths().forEach((path, index) => this.writeXML(path, this.sheets[index]));
        this.sheetPaths().forEach((path, index) => this.writeXML(path, this.worksheets[index]._writeSheetData()));
        return this.zip;
    }
    
    public writeFile(filename): Promise<{}> {
        return writeZipFile(filename, this.writeZip());
    }

    public writeStream(): NodeJS.WriteStream {
        return writeStream(this.writeZip());
    }

    public evaluate = (context: any, options: EvaluateOptions = {}) => {
        if (options.i18n) {
            Util.setI18n(options.i18n);
        }
        this.globals = options.globals;
        this.rootScope.getChildren().forEach(this.interpolate(context));
    }

    //sheet operations
    public renameSheet = (from: string, to: string) => {
        if (!contains(this.sheetNames(), from)) {
            throw new Error('specified sheet is not in the workbook');
        }
        if (contains(this.sheetNames(), to)) {
            throw new Error('sheet name must be unique');
        }
        let node = <Element>select1(`//xl:sheet[@name="${from}"]`, this.workbook);
        node.setAttribute('name', to);
    }
    public sheetNames = () => this.sheetAttrs('name');
    public getSheet = (name: string) => {
        return this.worksheets[this.sheetNames().indexOf(name)]
    };

    //private sheet related
    private sheetAttrs = (attr: string): string[] => <string[]>select<Element>('//xl:sheet', this.workbook).map(nodeAttribute(attr));
    private sheetPaths = () => this.sheetAttrs('r:id').map(this.getRel);
    private getRel = (id: string): string => <string>select1(`string(//r:Relationship[@Id="${id}"]/@Target)`, this.rels);

    //private cell related
    public getString = (index: number) => select1(`string(//xl:si[${index + 1}]/xl:t)`, this.strings);

    //private scope related
    private getRangeScopes = () => {
        return this.sheetNames().reduce((scopes, sheet) => {
            let scopeRanges = this.getSheet(sheet).getScopeRanges();
            return scopes.concat(scopeRanges.map(parseScopeRange(sheet)));
        }, Array<Scope>());
    }
    private getCellScopes = () => {
        return this.sheetNames().reduce((scopes, sheet) => {
            let scopeCells = this.getSheet(sheet).getScopeCells();
            return scopes.concat(scopeCells.map(parseScopeCell(sheet)));
        }, Array<Scope>())
    }
    private createScopeTree = (scopes) => {
        let root = new RootScope('');
        const place = (parent: RootScope) => (scope: Scope) => {
            let placed = parent.getChildren().some(leaf => {
                if (leaf.containsScope(scope)) {
                    //place in leaf
                    place(leaf)(scope);
                    return true;
                }
                return false;
            })
            //could not place in tree, append
            if (!placed) {
                parent.addChild(scope);
            }
        }
        scopes.forEach(place(root));
        return root;
    }

    private cloneCell = (scope: Scope, cell: Cell, to: string, direction: XLSX.Direction) => {
        cell.worksheet.insertCellMoveDim(to, direction);
        cell.cloneTo(to);
        //grow scope
        scope.growDim(1, direction);
        //update all scopes below
        this.scopes
            .filter(Scope.isAfterDim(to, direction))
            .filter(s => s.onSheet(cell.worksheet.getName()))
            .forEach(s => s.moveDim(1, direction));
    }

    private copyRange = (sheet: WorkSheet, fromRange, toRange, direction: XLSX.Direction) => {
        let fromAnchor = XLSX.splitRange(fromRange)[0];
        let toAnchor = XLSX.splitRange(toRange)[0];
        let delta = XLSX.subtractAddress(XLSX.getCellAddress(toAnchor), XLSX.getCellAddress(fromAnchor));
        let size = XLSX.getRangeDim(fromRange, direction);

        //move all cells below the range out of the way
        let after = sheet.getCells()
            .filter(Cell.atOrAfterRangeDim(toRange, direction))
            .sort(Cell.sortDim(direction, true));
        after.forEach(c => {
            // console.log('moving',c.ref(),'by',height);
            c.moveBy(XLSX.relAddress(size, direction));
        });

        //clone all cells in the range
        let inside = sheet.getCells().filter(c => XLSX.inRange(fromRange)(c.ref()));
        // console.log('cloning',inside.length,'cells inside',fromRange)
        inside.forEach(c => {
            let newRef = XLSX.getCellRel(delta, c.ref());
            c.cloneTo(newRef);
        })
    }

    private cloneScope = (scope: Scope, offset: number, direction: XLSX.Direction) => {
        let sheet = this.getSheet(scope.getSheet());
        let oldRange = scope.getRange();
        scope.moveDim(offset, direction);
        let newRange = scope.getRange();

        this.copyRange(sheet, oldRange, newRange, direction);

        //grow parent scope
        if (scope.parentScope instanceof Scope) {
            scope.parentScope.growDim(scope.dim(direction), direction)
        }
        //update all scopes after
        let after = this.scopes
            .filter(s => s.isAtOrAfterRangeDim(scope.getRange(), direction))
            .filter(s => s.onSheet(scope.getSheet()))
            // console.log(after.length, 'scopes need to move down');
            .filter(s => s !== scope);
        after.forEach(s => {
            // console.log('moving scope', s.getRange() ,'down by', scope.height());
            s.moveDim(scope.dim(direction), direction);
        });
        //move all cells after
        // sheet.getCells().filter(c => {
        //     c.isBelowRange(oldRange)
        // }).sort(sortRow(true)).forEach(c => c.moveBy({ c: 0, r: scope.height() }));
    }

    private interpolateScalar = (value, scope: Scope) => {
        // console.log(scope.getRange(), 'interpolating scalar', scope.type, value);
        if (scope.type === TYPE_RANGE) {
            // console.log('interpolate child scopes')
            scope.getChildren().forEach(this.interpolate(value));
        } else if (scope.type === TYPE_CELL) {
            let ref = XLSX.splitRange(scope.getRange())[0];
            // console.log('need to replace cell contents', value, scope.sheet, ref)
            let cell = this.getSheet(scope.getSheet()).getCell(ref);
            // console.log('filling', ref);
            cell.setValue(value);
        }
    }



    private interpolateCell = (scope: Scope, value: any[], direction: XLSX.Direction) => {
        let ref = scope.getAnchor();
        // console.log('repeat cells', value, scope.sheet, ref);
        let sheet = this.getSheet(scope.getSheet());
        let cell = sheet.getCell(ref);
        value.forEach((val, index) => {
            let r = XLSX.getCellOffset(index, direction, ref);
            if (index > 0) {
                // console.log('cloning', direction, val, index, scope.template,scope.type,scope.range);
                this.cloneCell(scope, cell, r, direction);
            }
            sheet.getCell(r).setValue(val);
        });
    }

    private interpolateRange = (scope: Scope, value: any[], direction: XLSX.Direction) => {
        let size = scope.dim(direction);
        if (!(value instanceof Array)) {
            throw new Error('expected value to be an array when evaluating '+scope.getTemplate()+', but it is not: ' + JSON.stringify(value));
        }
        scope.makeScalar();
        let clones = value.map((val, index) => {
            //create a scalar scope for the instance
            let s = scope.cloneAsScalar(index)
            // console.log('cloning scope',scope.getRange(),'to',s.getRange(),'to receive value');
            this.scopes.push(s);
            //setting parent early as cloneScopeDown uses the parent
            s.parentScope = scope;
            return s;
        });
        value.forEach((val, index) => {
            //repeat scope
            if (index > 0) {
                this.cloneScope(clones[index], size * index, direction);
            }
            // console.log('interpolate scope for')
        });
        scope.removeChildren();
        clones.forEach(s => scope.addChild(s));
        // console.log('repeat range scope, interpolating',value)
        clones.forEach(this.interpolate(value))
    }

    private interpolateStack = (value, scope: Scope, direction: XLSX.Direction) => {
        if (!value) { return }
        // console.log(scope.getRange(), 'interpolating column', scope.type);
        //need to repeat scope
        switch (scope.type) {
            case TYPE_RANGE: return this.interpolateRange(scope, value, direction);
            case TYPE_CELL: return this.interpolateCell(scope, value, direction);
        }
    }

    private interpolate = (context) => (scope: Scope) => {
        let globals = this.globals;
        let { value, type } = scope.evaluate(context, globals);
        // console.log(scope.getRange(), ':', type, value);
        switch (type) {
            case TYPE_SCALAR: return this.interpolateScalar(value, scope);
            case TYPE_ROW: return this.interpolateStack(value, scope, XLSX.HORIZONTAL);
            case TYPE_COLUMN: return this.interpolateStack(value, scope, XLSX.VERTICAL);
        }
    }

    //private helpers
    private readXML = path => this.zip.file('xl/' + path).async('text').then(strToXml);
    private writeXML = (path, xml) => this.zip.file('xl/' + path, xmlToStr(xml));

    static readFile(filename: string) {
        return new Workbook().readFile(filename);
    }
    static read(data: Buffer) {
        return new Workbook().read(data);
    }
}
