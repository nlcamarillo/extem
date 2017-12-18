/**
 * hyperscript like vnode stuff, they are all functions
 * these functions are composed in a tree and when the root
 * is evaluated against a Dom Document,
 * all elements are created
 */

export type VNode = (doc: Document) => Element;
export type VNodeChild = VNode | string | number | boolean | undefined | null;

export default (nodeName: string, attributes = {}, children: VNodeChild[] = []): VNode => (doc: Document): Element => {
    let el = doc.createElement(nodeName);
    Object.keys(attributes).forEach(key => el.setAttribute(key, attributes[key]));
    children.forEach(vnode => {
        if (typeof vnode === 'function') { el.appendChild((<VNode>vnode)(doc)); }
        if (typeof vnode === 'string') { el.appendChild(doc.createTextNode(vnode)); }
        if (typeof vnode === 'number') { el.appendChild(doc.createTextNode(vnode.toString())); }
    });
    return el;
}
