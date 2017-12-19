export const VERTICAL = 'r';
export const HORIZONTAL = 'c';
export type Direction = 'r' | 'c';

export type CellAddress = {
    c: number;
    r: number;
}
export type RangeAddress = CellAddress[];
export type CellRef = string;
export type RangeRef = string;

//excel address related
export const relAddress = (offset: number, direction: Direction): CellAddress => {
    switch (direction) {
        case HORIZONTAL: return { c: offset, r: 0 };
        case VERTICAL: return { r: offset, c: 0 };
    }
}

export const decodeRow = (str: string): number => parseInt(str, 10) - 1;
export const encodeRow = (row: number): string => (row + 1).toFixed(0);
export const decodeCol = (str: string): number => {
    return str.split('').reduce((n, l) => {
        return 26 * n + l.charCodeAt(0) - 64;
    }, 0) - 1;
}
export const encodeCol = (col: number): string => {
    var s = "";
    for (++col; col; col = Math.floor((col - 1) / 26))
        s = String.fromCharCode(((col - 1) % 26) + 65) + s;
    return s;
}

//cell related
export const getCellRef = ({ c, r }: CellAddress) => encodeCol(c) + encodeRow(r);
export const getCellRel = (a: CellAddress, ref: CellRef) => getCellRef(addAddress(getCellAddress(ref), a));
export const getCellAddress = (ref): CellAddress => {
    let match = ref.match(/([A-Z]*)(\d*)/);
    return { c: decodeCol(match[1]), r: decodeRow(match[2]) }
}
export const getCellOffset = (offset: number, direction: Direction, ref: CellRef) => {
    return getCellRel(relAddress(offset, direction), ref);
}

//range related
export const splitRange = (ref: RangeRef) => ref.split(':').slice(0, 2);
export const getRangeAddress = (ref: RangeRef) => splitRange(ref).map(getCellAddress);
export const getRangeRef = (address: CellAddress[]) => address.map(getCellRef).join(':');
export const getRangeDim = (ref: string, direction: Direction) => {
    let a = getRangeAddress(ref);
    return 1 + a[1][direction] - a[0][direction];
}
export const inRange = (range: RangeRef) => (ref: CellRef) => {
    let a = getCellAddress(ref);
    let r = getRangeAddress(range);
    const within = (s, v, e) => (s <= v) && (v <= e);
    return (
        ((r[0].c < 0) || within(r[0].c, a.c, r[1].c)) &&
        ((r[0].r < 0) || within(r[0].r, a.r, r[1].r))
    )
}
export const getRangeCells = (range: RangeRef) => {
    let c, r, a = getRangeAddress(range);
    let cells: string[] = [];
    for (c = a[0].c; c <= a[1].c; c++) {
        for (r = a[0].r; r <= a[1].r; r++) {
            cells.push(getCellRef({ c, r }));
        }
    }
    return cells;
}

//address arithmatic
export const addAddress = (aa: CellAddress, ab: CellAddress): CellAddress => ({ c: aa.c + ab.c, r: aa.r + ab.r });
export const subtractAddress = (aa: CellAddress, ab: CellAddress): CellAddress => ({ c: aa.c - ab.c, r: aa.r - ab.r });

//range and cell evaluation
//create selectors for the primary and secondary dimension of an address based on the direction
const other = (direction: Direction) => {
    switch (direction) {
        case VERTICAL: return HORIZONTAL;
        case HORIZONTAL: return VERTICAL;
    }
}
const createSelectors = (direction: Direction) => {
    return {
        p: (a: CellAddress) => a[direction],
        s: (a: CellAddress) => a[other(direction)]
    };
};

// const cellAfterRangeDim = (ca: CellAddress, ra: RangeAddress, direction: Direction) => {
//     let {p, s} = createSelectors(direction);
//     return (s(ra[0]) <= s(ca)) && (s(ca) <= s(ra[1])) && (p(ca) > p(ra[1]));
// }
export const cellAtOrAfterDim = (aa: CellAddress, ab: CellAddress, direction: Direction) => {
    let { p, s } = createSelectors(direction);
    let delta = subtractAddress(aa, ab);
    return s(delta) === 0 && p(delta) >= 0;
}
export const rangeAfterCellDim = (ca: CellAddress, ra: RangeAddress, direction: Direction) => {
    let { p, s } = createSelectors(direction);
    return (s(ra[0]) <= s(ca)) && (s(ca) <= s(ra[1])) && (p(ra[0]) > p(ca));
}
export const cellAtOrAfterRangeDim = (ca: CellAddress, ra: RangeAddress, direction: Direction) => {
    let { p, s } = createSelectors(direction);
    return (s(ra[0]) <= s(ca)) && (s(ca) <= s(ra[1])) && (p(ca) >= p(ra[0]));
}
