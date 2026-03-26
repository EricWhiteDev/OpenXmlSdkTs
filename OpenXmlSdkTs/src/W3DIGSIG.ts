import { XName, XNamespace } from 'ltxmlts';

export class W3DIGSIG {
    static readonly namespace: XNamespace = XNamespace.get("http://www.w3.org/2000/09/xmldsig#");

    static readonly CanonicalizationMethod: XName = W3DIGSIG.namespace.getName("CanonicalizationMethod");
    static readonly DigestMethod: XName = W3DIGSIG.namespace.getName("DigestMethod");
    static readonly DigestValue: XName = W3DIGSIG.namespace.getName("DigestValue");
    static readonly Exponent: XName = W3DIGSIG.namespace.getName("Exponent");
    static readonly KeyInfo: XName = W3DIGSIG.namespace.getName("KeyInfo");
    static readonly KeyValue: XName = W3DIGSIG.namespace.getName("KeyValue");
    static readonly Manifest: XName = W3DIGSIG.namespace.getName("Manifest");
    static readonly Modulus: XName = W3DIGSIG.namespace.getName("Modulus");
    static readonly Object: XName = W3DIGSIG.namespace.getName("Object");
    static readonly Reference: XName = W3DIGSIG.namespace.getName("Reference");
    static readonly RSAKeyValue: XName = W3DIGSIG.namespace.getName("RSAKeyValue");
    static readonly Signature: XName = W3DIGSIG.namespace.getName("Signature");
    static readonly SignatureMethod: XName = W3DIGSIG.namespace.getName("SignatureMethod");
    static readonly SignatureProperties: XName = W3DIGSIG.namespace.getName("SignatureProperties");
    static readonly SignatureProperty: XName = W3DIGSIG.namespace.getName("SignatureProperty");
    static readonly SignatureValue: XName = W3DIGSIG.namespace.getName("SignatureValue");
    static readonly SignedInfo: XName = W3DIGSIG.namespace.getName("SignedInfo");
    static readonly Transform: XName = W3DIGSIG.namespace.getName("Transform");
    static readonly Transforms: XName = W3DIGSIG.namespace.getName("Transforms");
    static readonly X509Certificate: XName = W3DIGSIG.namespace.getName("X509Certificate");
    static readonly X509Data: XName = W3DIGSIG.namespace.getName("X509Data");
    static readonly X509IssuerName: XName = W3DIGSIG.namespace.getName("X509IssuerName");
    static readonly X509SerialNumber: XName = W3DIGSIG.namespace.getName("X509SerialNumber");
}
