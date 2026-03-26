import { XName, XNamespace } from 'ltxmlts';

export class SLE {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/drawing/2010/slicer");

    static readonly slicer: XName = SLE.namespace.getName("slicer");
}
