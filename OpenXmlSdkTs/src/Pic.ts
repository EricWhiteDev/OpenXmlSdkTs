import { XName, XNamespace } from 'ltxmlts';

export class Pic {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/drawingml/2006/picture");

    static readonly blipFill: XName = Pic.namespace.getName("blipFill");
    static readonly cNvPicPr: XName = Pic.namespace.getName("cNvPicPr");
    static readonly cNvPr: XName = Pic.namespace.getName("cNvPr");
    static readonly nvPicPr: XName = Pic.namespace.getName("nvPicPr");
    static readonly _pic: XName = Pic.namespace.getName("pic");
    static readonly spPr: XName = Pic.namespace.getName("spPr");
}
