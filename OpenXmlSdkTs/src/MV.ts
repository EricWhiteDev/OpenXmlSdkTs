import { XName, XNamespace } from 'ltxmlts';

export class MV {
    static readonly namespace: XNamespace = XNamespace.get("urn:schemas-microsoft-com:mac:vml");

    static readonly blur: XName = MV.namespace.getName("blur");
    static readonly complextextbox: XName = MV.namespace.getName("complextextbox");
}
