import { XName, XNamespace } from 'ltxmlts';

export class CUSTPRO {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");

    static readonly Properties: XName = CUSTPRO.namespace.getName("Properties");
    static readonly property: XName = CUSTPRO.namespace.getName("property");
}
