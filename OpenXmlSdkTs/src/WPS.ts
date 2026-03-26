import { XName, XNamespace } from 'ltxmlts';

export class WPS {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

    static readonly altTxbx: XName = WPS.namespace.getName("altTxbx");
    static readonly bodyPr: XName = WPS.namespace.getName("bodyPr");
    static readonly cNvSpPr: XName = WPS.namespace.getName("cNvSpPr");
    static readonly spPr: XName = WPS.namespace.getName("spPr");
    static readonly style: XName = WPS.namespace.getName("style");
    static readonly textbox: XName = WPS.namespace.getName("textbox");
    static readonly txbx: XName = WPS.namespace.getName("txbx");
    static readonly wsp: XName = WPS.namespace.getName("wsp");
}
