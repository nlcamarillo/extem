import { TYPE_RANGE, TYPE_SCALAR } from './types';
import { templateType } from './Util';
import Scope from './Scope';

export default class RootScope {
    private children: Scope[] = [];
    public parentScope: RootScope | null = null;
    public type: string = TYPE_RANGE;
    constructor(protected template: string) { }


    // public setParent = (parentScope: RootScope) => {
    //     this.parentScope = parentScope;
    // }
    // public getParent = () => this.parentScope;

    public addChild = (childScope: Scope) => {
        childScope.parentScope = this;
        this.children.push(childScope);
    }

    public getChildren = () => this.children;
    public removeChildren = () => this.children = [];

    public templateType = () => templateType(this.template);

    public evaluate = (context): { value: any, type: string | null } => {
        return {
            value: context,
            type: TYPE_SCALAR
        }
    }
}
