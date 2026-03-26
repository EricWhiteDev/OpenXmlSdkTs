import { XName, XNamespace } from 'ltxmlts';

export class R {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

    static readonly blip: XName = R.namespace.getName("blip");
    static readonly cs: XName = R.namespace.getName("cs");
    static readonly dm: XName = R.namespace.getName("dm");
    static readonly embed: XName = R.namespace.getName("embed");
    static readonly href: XName = R.namespace.getName("href");
    static readonly id: XName = R.namespace.getName("id");
    static readonly link: XName = R.namespace.getName("link");
    static readonly lo: XName = R.namespace.getName("lo");
    static readonly pict: XName = R.namespace.getName("pict");
    static readonly qs: XName = R.namespace.getName("qs");
    static readonly verticalDpi: XName = R.namespace.getName("verticalDpi");
}
