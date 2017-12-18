import { CellAddress } from './XLSXUtil';

export const TYPE_SCALAR = 'scalar';
export const TYPE_COLUMN = 'column';
export const TYPE_ROW = 'row';
export const TYPE_CELL = 'cell';
export const TYPE_RANGE = 'range';

export type ScopeDef = {
    range: string;
    sheet: string;
    template: string;
    address: CellAddress[];
}
