import { XName, XNamespace } from 'ltxmlts';

export class WP14 {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

    static readonly editId: XName = WP14.namespace.getName("editId");
    static readonly pctHeight: XName = WP14.namespace.getName("pctHeight");
    static readonly pctPosVOffset: XName = WP14.namespace.getName("pctPosVOffset");
    static readonly pctWidth: XName = WP14.namespace.getName("pctWidth");
    static readonly sizeRelH: XName = WP14.namespace.getName("sizeRelH");
    static readonly sizeRelV: XName = WP14.namespace.getName("sizeRelV");
}
