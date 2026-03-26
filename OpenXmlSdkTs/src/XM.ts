import { XName, XNamespace } from 'ltxmlts';

export class XM {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/excel/2006/main");

    static readonly f: XName = XM.namespace.getName("f");
    static readonly _ref: XName = XM.namespace.getName("ref");
    static readonly sqref: XName = XM.namespace.getName("sqref");
}
