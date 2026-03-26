import { XName, XNamespace } from 'ltxmlts';

export class VML {
    static readonly namespace: XNamespace = XNamespace.get("urn:schemas-microsoft-com:vml");

    static readonly arc: XName = VML.namespace.getName("arc");
    static readonly background: XName = VML.namespace.getName("background");
    static readonly curve: XName = VML.namespace.getName("curve");
    static readonly ext: XName = VML.namespace.getName("ext");
    static readonly f: XName = VML.namespace.getName("f");
    static readonly fill: XName = VML.namespace.getName("fill");
    static readonly formulas: XName = VML.namespace.getName("formulas");
    static readonly group: XName = VML.namespace.getName("group");
    static readonly h: XName = VML.namespace.getName("h");
    static readonly handles: XName = VML.namespace.getName("handles");
    static readonly image: XName = VML.namespace.getName("image");
    static readonly imagedata: XName = VML.namespace.getName("imagedata");
    static readonly line: XName = VML.namespace.getName("line");
    static readonly oval: XName = VML.namespace.getName("oval");
    static readonly path: XName = VML.namespace.getName("path");
    static readonly polyline: XName = VML.namespace.getName("polyline");
    static readonly rect: XName = VML.namespace.getName("rect");
    static readonly roundrect: XName = VML.namespace.getName("roundrect");
    static readonly shadow: XName = VML.namespace.getName("shadow");
    static readonly shape: XName = VML.namespace.getName("shape");
    static readonly shapetype: XName = VML.namespace.getName("shapetype");
    static readonly stroke: XName = VML.namespace.getName("stroke");
    static readonly textbox: XName = VML.namespace.getName("textbox");
    static readonly textpath: XName = VML.namespace.getName("textpath");
}
