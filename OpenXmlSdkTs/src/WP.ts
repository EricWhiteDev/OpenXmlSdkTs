import { XName, XNamespace } from 'ltxmlts';

export class WP {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

    static readonly align: XName = WP.namespace.getName("align");
    static readonly anchor: XName = WP.namespace.getName("anchor");
    static readonly cNvGraphicFramePr: XName = WP.namespace.getName("cNvGraphicFramePr");
    static readonly docPr: XName = WP.namespace.getName("docPr");
    static readonly effectExtent: XName = WP.namespace.getName("effectExtent");
    static readonly extent: XName = WP.namespace.getName("extent");
    static readonly inline: XName = WP.namespace.getName("inline");
    static readonly lineTo: XName = WP.namespace.getName("lineTo");
    static readonly positionH: XName = WP.namespace.getName("positionH");
    static readonly positionV: XName = WP.namespace.getName("positionV");
    static readonly posOffset: XName = WP.namespace.getName("posOffset");
    static readonly simplePos: XName = WP.namespace.getName("simplePos");
    static readonly start: XName = WP.namespace.getName("start");
    static readonly wrapNone: XName = WP.namespace.getName("wrapNone");
    static readonly wrapPolygon: XName = WP.namespace.getName("wrapPolygon");
    static readonly wrapSquare: XName = WP.namespace.getName("wrapSquare");
    static readonly wrapThrough: XName = WP.namespace.getName("wrapThrough");
    static readonly wrapTight: XName = WP.namespace.getName("wrapTight");
    static readonly wrapTopAndBottom: XName = WP.namespace.getName("wrapTopAndBottom");
}
