let fs = require('fs');
let JSZip = require('jszip');
let dateFormat = require('dateformat');
import * as jsonata from 'jsonata';
import { TYPE_COLUMN, TYPE_ROW, TYPE_SCALAR } from './types';

//some functional helpers
export const keys = (obj: Object) => Object.keys(obj);
export const values = (obj: Object) => keys(obj).map(k => obj[k]);
export const not = (predicate: Function) => (value) => !predicate(value);
export const contains = <T>(arr: T[], value:T): boolean => arr.indexOf(value) !== -1;

//zip file access
export function readZipFile(xlsPath: string) {
    return new Promise((resolve, reject) => {
        fs.readFile(xlsPath, 'binary', (err, res) => {
            if (err) { return reject(err); }
            resolve(JSZip.loadAsync(res));
        });
    })
}

export function readData(data) {
    return Promise.resolve(JSZip.loadAsync(data));
}

export function writeZipFile(xlsPath: string, zip) {
    return new Promise((resolve, reject) => {
        zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
            .pipe(fs.createWriteStream(xlsPath))
            .on('finish', resolve)
            .on('error', reject);
    })
}

export function writeStream(zip) {
    return zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true })
}

function formatDate(date: string, format: string = 'yyyy-mm-dd'): string {
    let d = new Date(date);
    if (isNaN(d.getTime())) return '';
    return dateFormat(date, format);
}

export function setI18n(i18n) {
    dateFormat.i18n = i18n;
}

//gets a value from an object using jsonata
let cache = {};
export interface JsonataGlobals {
    [name: string]: (...args: any[]) => any;
}
export const getValue = (obj: any, path: string, globals: JsonataGlobals={}) => {
    if (!cache[path]) {
        // replace excel single and double quotes with normal ones
        let expression = jsonata(path.replace(/”/g, '"').replace(/’/g,'’'));
        // for backwards compat, should be dropped in favor of globals
        expression.assign('formatDate', formatDate);
        // additional globals
        Object.keys(globals).forEach(key => {
            expression.assign(key, globals[key]);
        })
        // expression.assign('values', values);
        cache[path] = expression;
    }
    // let expression = jsonata(path);
    return cache[path].evaluate(obj);
}

//template specifics
export function templateType(str: string) {
    if (!str) return null;
    switch (str.toString().substr(0, 2)) {
        case '${': return TYPE_SCALAR;
        case '|{': return TYPE_COLUMN;
        case '_{': return TYPE_ROW;
        default: return null;
    }
}
export const parseTemplate = (str: string) => {
    let type = templateType(str);
    let match = str.match(/^[\$\|_]\{(.*?)\}$/);
    return { type, path: (match && match[1]) || '' };
}

export const matchRangeTemplate = (str: string) => {
    return str && str.match(/^IFERROR\(N\((.+)?\),\s*"(.*?)"\)$/i);
}
