import { XName, XNamespace } from 'ltxmlts';

export class DGM14 {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/drawing/2010/diagram");

    static readonly cNvPr: XName = DGM14.namespace.getName("cNvPr");
    static readonly recolorImg: XName = DGM14.namespace.getName("recolorImg");
}
