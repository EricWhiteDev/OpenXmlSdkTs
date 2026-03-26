import { XName, XNamespace } from 'ltxmlts';

export class P15 {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office15/powerpoint");

    static readonly extElement: XName = P15.namespace.getName("extElement");
}
