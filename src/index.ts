/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

/**
 * OpenXmlSdkTs — Open XML SDK for TypeScript.
 *
 * A library for reading, writing, and manipulating Office Open XML documents
 * (.docx, .xlsx, .pptx) in Node.js and browser environments. Supports three
 * I/O formats: binary Blob, Flat OPC XML, and Base64 string.
 *
 * @packageDocumentation
 */

export { OpenXmlPackage } from "./OpenXmlPackage.js";
export { WmlPackage } from "./WmlPackage.js";
export { SmlPackage } from "./SmlPackage.js";
export { PmlPackage } from "./PmlPackage.js";
export { OpenXmlRelationship } from "./OpenXmlRelationship.js";
export type { Base64String, FlatOpcString, OpcBinary } from "./OpenXmlPackage.js";
export { Utility } from "./Utility.js";
export { OpenXmlPart } from "./OpenXmlPart.js";
export { WmlPart } from "./WmlPart.js";
export { SmlPart } from "./SmlPart.js";
export { PmlPart } from "./PmlPart.js";
export type { PartType } from "./OpenXmlPart.js";
export { ContentType } from "./ContentType.js";
export type { ContentTypeKey, ContentTypeValue } from "./ContentType.js";
export { RelationshipType } from "./RelationshipType.js";
export type { RelationshipTypeKey, RelationshipTypeValue } from "./RelationshipType.js";
export { XDocument, XDeclaration, XElement, XAttribute, XName, XNamespace, XProcessingInstruction, xseq } from "ltxmlts";
export {
  A,
  A14,
  C,
  CDR,
  COM,
  CP,
  CT,
  CUSTPRO,
  DC,
  DCTERMS,
  DGM,
  DGM14,
  DIGSIG,
  DS,
  DSP,
  EP,
  FLATOPC,
  LC,
  M,
  MC,
  MDSSI,
  MP,
  MV,
  NoNamespace,
  O,
  P,
  P14,
  P15,
  Pic,
  PKGREL,
  R,
  S,
  SL,
  SLE,
  VML,
  VT,
  W,
  W10,
  W14,
  W15,
  W16CEX,
  W16CID,
  W3DIGSIG,
  WP,
  WP14,
  WPS,
  X,
  XDR,
  XDR14,
  XM,
  XSI,
} from "./OpenXmlNamespacesAndNames.js";
