import { XName, XNamespace } from 'ltxmlts';

export class W10 {
    static readonly namespace: XNamespace = XNamespace.get("urn:schemas-microsoft-com:office:word");

    static readonly anchorlock: XName = W10.namespace.getName("anchorlock");
    static readonly borderbottom: XName = W10.namespace.getName("borderbottom");
    static readonly borderleft: XName = W10.namespace.getName("borderleft");
    static readonly borderright: XName = W10.namespace.getName("borderright");
    static readonly bordertop: XName = W10.namespace.getName("bordertop");
    static readonly wrap: XName = W10.namespace.getName("wrap");
}
