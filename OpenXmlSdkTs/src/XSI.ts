import { XName, XNamespace } from 'ltxmlts';

export class XSI {
    static readonly namespace: XNamespace = XNamespace.get("http://www.w3.org/2001/XMLSchema-instance");

    static readonly type: XName = XSI.namespace.getName("type");
}
