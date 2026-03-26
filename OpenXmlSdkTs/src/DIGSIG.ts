import { XName, XNamespace } from 'ltxmlts';

export class DIGSIG {
    static readonly namespace: XNamespace = XNamespace.get("http://schemas.microsoft.com/office/2006/digsig");

    static readonly ApplicationVersion: XName = DIGSIG.namespace.getName("ApplicationVersion");
    static readonly ColorDepth: XName = DIGSIG.namespace.getName("ColorDepth");
    static readonly HorizontalResolution: XName = DIGSIG.namespace.getName("HorizontalResolution");
    static readonly ManifestHashAlgorithm: XName = DIGSIG.namespace.getName("ManifestHashAlgorithm");
    static readonly Monitors: XName = DIGSIG.namespace.getName("Monitors");
    static readonly OfficeVersion: XName = DIGSIG.namespace.getName("OfficeVersion");
    static readonly SetupID: XName = DIGSIG.namespace.getName("SetupID");
    static readonly SignatureComments: XName = DIGSIG.namespace.getName("SignatureComments");
    static readonly SignatureImage: XName = DIGSIG.namespace.getName("SignatureImage");
    static readonly SignatureInfoV1: XName = DIGSIG.namespace.getName("SignatureInfoV1");
    static readonly SignatureProviderDetails: XName = DIGSIG.namespace.getName("SignatureProviderDetails");
    static readonly SignatureProviderId: XName = DIGSIG.namespace.getName("SignatureProviderId");
    static readonly SignatureProviderUrl: XName = DIGSIG.namespace.getName("SignatureProviderUrl");
    static readonly SignatureText: XName = DIGSIG.namespace.getName("SignatureText");
    static readonly SignatureType: XName = DIGSIG.namespace.getName("SignatureType");
    static readonly VerticalResolution: XName = DIGSIG.namespace.getName("VerticalResolution");
    static readonly WindowsVersion: XName = DIGSIG.namespace.getName("WindowsVersion");
}
