import { XName, XNamespace } from 'ltxmlts';

export class CDR {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/chartDrawing");

    static readonly absSizeAnchor: XName = CDR.namespace.getName("absSizeAnchor");
    static readonly blipFill: XName = CDR.namespace.getName("blipFill");
    static readonly cNvCxnSpPr: XName = CDR.namespace.getName("cNvCxnSpPr");
    static readonly cNvGraphicFramePr: XName = CDR.namespace.getName("cNvGraphicFramePr");
    static readonly cNvGrpSpPr: XName = CDR.namespace.getName("cNvGrpSpPr");
    static readonly cNvPicPr: XName = CDR.namespace.getName("cNvPicPr");
    static readonly cNvPr: XName = CDR.namespace.getName("cNvPr");
    static readonly cNvSpPr: XName = CDR.namespace.getName("cNvSpPr");
    static readonly cxnSp: XName = CDR.namespace.getName("cxnSp");
    static readonly ext: XName = CDR.namespace.getName("ext");
    static readonly from: XName = CDR.namespace.getName("from");
    static readonly graphicFrame: XName = CDR.namespace.getName("graphicFrame");
    static readonly grpSp: XName = CDR.namespace.getName("grpSp");
    static readonly grpSpPr: XName = CDR.namespace.getName("grpSpPr");
    static readonly nvCxnSpPr: XName = CDR.namespace.getName("nvCxnSpPr");
    static readonly nvGraphicFramePr: XName = CDR.namespace.getName("nvGraphicFramePr");
    static readonly nvGrpSpPr: XName = CDR.namespace.getName("nvGrpSpPr");
    static readonly nvPicPr: XName = CDR.namespace.getName("nvPicPr");
    static readonly nvSpPr: XName = CDR.namespace.getName("nvSpPr");
    static readonly pic: XName = CDR.namespace.getName("pic");
    static readonly relSizeAnchor: XName = CDR.namespace.getName("relSizeAnchor");
    static readonly sp: XName = CDR.namespace.getName("sp");
    static readonly spPr: XName = CDR.namespace.getName("spPr");
    static readonly style: XName = CDR.namespace.getName("style");
    static readonly to: XName = CDR.namespace.getName("to");
    static readonly txBody: XName = CDR.namespace.getName("txBody");
    static readonly x: XName = CDR.namespace.getName("x");
    static readonly xfrm: XName = CDR.namespace.getName("xfrm");
    static readonly y: XName = CDR.namespace.getName("y");
}
