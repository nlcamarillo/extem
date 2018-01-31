import { keys, not, values } from './Util';
import * as XLSX from './XLSXUtil';
import { select, select1 } from './XMLUtil';
import Cell from './Cell';
import Row from './Row';
import Workbook from './Workbook';
import h, { VNode, VNodeChild } from './hyper';

type Index<T> = { [ref: string]: T };
type CellIndex = Index<Cell>;
type RowIndex = Index<CellIndex>;
type ColIndex = Index<CellIndex>;

export default class WorkSheet {
    private cellIndex: CellIndex = {};
    private rows: Index<Row> = {};
    private ranges: Cell[] = [];
    private rowIndex: RowIndex = {};
    private colIndex: ColIndex = {};
    constructor(private sheet: Document, public workbook: Workbook, private name: string) {
        this._readSheetData();
    };

    //cell storage
    public getCell = (ref: XLSX.CellRef): Cell => this.cellIndex[ref];
    public getCells = (): Cell[] => values(this.cellIndex);
    public getRanges = (): Cell[] => this.ranges;
    public _getRow = (ref: XLSX.CellRef): Cell[] => values(this.rowIndex[ref]).sort(Cell.sortDim(XLSX.HORIZONTAL));
    public _getCol = (ref: XLSX.CellRef): Cell[] => values(this.colIndex[ref]).sort(Cell.sortDim(XLSX.VERTICAL));
    public _occupied = (ref: XLSX.CellRef): boolean => !!this.cellIndex[ref];
    private _addCellToIndex = (cell: Cell) => {
        let ref = cell.ref();
        let a = XLSX.getCellAddress(ref);
        let row = XLSX.encodeRow(a.r);
        let col = XLSX.encodeCol(a.c);
        this.cellIndex[ref] = cell;
        if (!this.rowIndex[row]) { this.rowIndex[row] = {}; }
        if (!this.colIndex[col]) { this.colIndex[col] = {}; }
        this.rowIndex[row][ref] = cell;
        this.colIndex[col][ref] = cell;
    }
    private _addRowToIndex = (row: Row) => {
        let ref = row.ref();
        this.rows[ref] = row;
    }
    private _removeCellFromIndex = (cell: Cell) => {
        let ref = cell.ref();
        let a = XLSX.getCellAddress(ref);
        let row = XLSX.encodeRow(a.r);
        let col = XLSX.encodeCol(a.c);
        delete this.cellIndex[ref];
        delete this.rowIndex[row][ref];
        delete this.colIndex[col][ref];
    }
    public _moveCell = (from: XLSX.CellRef, to: XLSX.CellRef) => {
        if (this._occupied(to)) {
            throw new Error(`target ${to} is occupied, move or clear that first`);
        };
        let cellToMove = this.getCell(from);
        this._removeCellFromIndex(cellToMove);
        cellToMove.setRef(to);
        this._addCellToIndex(cellToMove);
    }
    public _cloneCell = (from: XLSX.CellRef, to: XLSX.CellRef) => {
        if (!this._occupied(from)) {
            throw new Error('cell to clone is not occupied');
        }
        let cell = this.getCell(from);
        let newCell = cell.copy({ r: to });
        this._addCellToIndex(newCell);
    }

    //xml generation
    private _readSheetData = () => {
        let allCells = select<Element>(`//xl:c`, this.sheet)
            .map(cellDOM => Cell.fromNode(cellDOM, this))
        //create ranges
        this.ranges = allCells.filter(Cell.isRangeTemplate);
        //create cells
        let cells = allCells.filter(not(Cell.isRangeTemplate));
        cells.forEach(this._addCellToIndex);
        //create rows from dom
        let rows = select<Element>(`//xl:row`, this.sheet)
            .map(rowDOM => Row.fromNode(rowDOM, this));
        rows.forEach(this._addRowToIndex);
    }

    private _createRowNode = (row: string, children?: VNodeChild[]): VNode => {
        let rowCells = values(this.rowIndex[row]);
        let ht = rowCells.reduce((maxHeight, cell: Cell) => {
            let sizeForCell = this.workbook.styles.getRowSize(cell.style());
            return Math.max(maxHeight, sizeForCell);
        }, 0) || undefined;
        let storedRow = this.rows[row];
        let attributes = storedRow ? storedRow.getAttributes() : { r: row };
        return h('row', { ...attributes, ht }, children);
    }
    private _createCellNode = (cell: Cell): VNode => {
        let value = cell.getValue();
        return h('c', {
            ...cell.getAttributes(),
            t: (typeof value === 'number') ? 'n' : 'str'
        }, [
                !!cell.formula() && h('f', { aca: 'false' }, [cell.formula()]),
                !!value && h('v', {}, [value])
            ]);
    }
    public _writeSheetData = () => {
        let doc = this.sheet;
        let newSheetData = h('sheetData', {}, keys(this.rowIndex).map(row => {
            return this._createRowNode(row, this._getRow(row).map(cell => {
                return this._createCellNode(cell);
            }));
        }));

        //replace the sheetdata
        let sheetData = this.getSheetData();
        if (sheetData.parentNode) {
            sheetData.parentNode.replaceChild(newSheetData(doc), sheetData);
        }

        return doc;
    }


    public getName = () => this.name;

    //sheet operations
    public insertCellMoveDim = (ref: XLSX.CellRef, direction: XLSX.Direction) => {
        return this.getCells()
            .filter(Cell.atOrAfterDim(ref, direction))
            .sort(Cell.sortDim(direction, true))
            .forEach(c => c.moveBy(XLSX.relAddress(1, direction)));
    }

    //scope related
    public getScopeRanges = () => this.getRanges().filter(Cell.isRangeTemplate);
    public getScopeCells = () => this.getCells().filter(Cell.isCellTemplate);

    //private sheet related
    private getSheetData = () => select1<Element>(`//xl:sheetData`, this.sheet);
}
