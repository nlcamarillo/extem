// import * as xpath from 'xpath';
let xpath = require('xpath');
import * as xmldom from 'xmldom';

//namespace aware xpath selector
const xpathSelect = xpath.useNamespaces({
    "xl": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/package/2006/relationships"
});
export const select = <T>(path: string, doc: Node): T[] => <T[]>xpathSelect(path, doc);
export const select1 = <T>(path: string, doc: Node): T => <T>xpathSelect(path, doc, true);

//xml conversion
export const strToXml = (str: string) => new xmldom.DOMParser().parseFromString(str);
export const xmlToStr = (xml: Node) => new xmldom.XMLSerializer().serializeToString(xml);

//xml node helpers
export type NodeAttributes = {[key: string]: string;}
export const nodeAttribute = (name: string) => (node: Element) => node.getAttribute(name);
export const nodeAttributes = (node: Element): NodeAttributes => {
    let attributes = {};
    for (let i = 0; i < node.attributes.length; i += 1) {
        let a = node.attributes.item(i);
        attributes[a.name] = a.value;
    }
    return attributes;
}
