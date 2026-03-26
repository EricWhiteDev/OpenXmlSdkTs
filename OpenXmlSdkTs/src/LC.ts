import { XName, XNamespace } from 'ltxmlts';

export class LC {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas");

    static readonly lockedCanvas: XName = LC.namespace.getName("lockedCanvas");
}
