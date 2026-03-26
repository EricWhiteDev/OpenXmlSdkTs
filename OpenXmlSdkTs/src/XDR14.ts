import { XName, XNamespace } from 'ltxmlts';

export class XDR14 {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing");

    static readonly cNvContentPartPr: XName = XDR14.namespace.getName("cNvContentPartPr");
    static readonly cNvPr: XName = XDR14.namespace.getName("cNvPr");
    static readonly nvContentPartPr: XName = XDR14.namespace.getName("nvContentPartPr");
    static readonly nvPr: XName = XDR14.namespace.getName("nvPr");
    static readonly xfrm: XName = XDR14.namespace.getName("xfrm");
}
