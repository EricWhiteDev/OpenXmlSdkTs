import { XName, XNamespace } from 'ltxmlts';

export class XDR {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

    static readonly absoluteAnchor: XName = XDR.namespace.getName("absoluteAnchor");
    static readonly blipFill: XName = XDR.namespace.getName("blipFill");
    static readonly clientData: XName = XDR.namespace.getName("clientData");
    static readonly cNvCxnSpPr: XName = XDR.namespace.getName("cNvCxnSpPr");
    static readonly cNvGraphicFramePr: XName = XDR.namespace.getName("cNvGraphicFramePr");
    static readonly cNvGrpSpPr: XName = XDR.namespace.getName("cNvGrpSpPr");
    static readonly cNvPicPr: XName = XDR.namespace.getName("cNvPicPr");
    static readonly cNvPr: XName = XDR.namespace.getName("cNvPr");
    static readonly cNvSpPr: XName = XDR.namespace.getName("cNvSpPr");
    static readonly col: XName = XDR.namespace.getName("col");
    static readonly colOff: XName = XDR.namespace.getName("colOff");
    static readonly contentPart: XName = XDR.namespace.getName("contentPart");
    static readonly cxnSp: XName = XDR.namespace.getName("cxnSp");
    static readonly ext: XName = XDR.namespace.getName("ext");
    static readonly from: XName = XDR.namespace.getName("from");
    static readonly graphicFrame: XName = XDR.namespace.getName("graphicFrame");
    static readonly grpSp: XName = XDR.namespace.getName("grpSp");
    static readonly grpSpPr: XName = XDR.namespace.getName("grpSpPr");
    static readonly nvCxnSpPr: XName = XDR.namespace.getName("nvCxnSpPr");
    static readonly nvGraphicFramePr: XName = XDR.namespace.getName("nvGraphicFramePr");
    static readonly nvGrpSpPr: XName = XDR.namespace.getName("nvGrpSpPr");
    static readonly nvPicPr: XName = XDR.namespace.getName("nvPicPr");
    static readonly nvSpPr: XName = XDR.namespace.getName("nvSpPr");
    static readonly oneCellAnchor: XName = XDR.namespace.getName("oneCellAnchor");
    static readonly pic: XName = XDR.namespace.getName("pic");
    static readonly pos: XName = XDR.namespace.getName("pos");
    static readonly row: XName = XDR.namespace.getName("row");
    static readonly rowOff: XName = XDR.namespace.getName("rowOff");
    static readonly sp: XName = XDR.namespace.getName("sp");
    static readonly spPr: XName = XDR.namespace.getName("spPr");
    static readonly style: XName = XDR.namespace.getName("style");
    static readonly to: XName = XDR.namespace.getName("to");
    static readonly twoCellAnchor: XName = XDR.namespace.getName("twoCellAnchor");
    static readonly txBody: XName = XDR.namespace.getName("txBody");
    static readonly wsDr: XName = XDR.namespace.getName("wsDr");
    static readonly xfrm: XName = XDR.namespace.getName("xfrm");
}
