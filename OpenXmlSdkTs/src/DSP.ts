import { XName, XNamespace } from 'ltxmlts';

export class DSP {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/drawing/2008/diagram");

    static readonly dataModelExt: XName = DSP.namespace.getName("dataModelExt");
}
