import WorkSheet from './WorkSheet';
import { nodeAttributes, NodeAttributes } from './XMLUtil';

export default class Row {
    constructor(private attributes: NodeAttributes, public worksheet: WorkSheet) { };

    public getAttributes = () => this.attributes;
    public ref = () => this.attributes.r;

    static fromNode = (node: Element, sheet: WorkSheet) => {
        return new Row(nodeAttributes(node), sheet);
    }
}
