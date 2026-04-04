/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { OpenXmlPart } from "./OpenXmlPart";
import { RelationshipType } from "./RelationshipType";

/**
 * Word-specific part with convenience methods for navigating Word document structure.
 *
 * @remarks
 * Extends {@link OpenXmlPart} with typed accessors for headers, footers, styles,
 * comments, numbering, fonts, and other Word-specific parts.
 *
 * @example
 * ```typescript
 * const mainPart = await doc.mainDocumentPart();
 * const headers = await mainPart!.headerParts();
 * const stylesPart = await mainPart!.styleDefinitionsPart();
 * const commentsPart = await mainPart!.wordprocessingCommentsPart();
 * ```
 *
 * @category Class and Type Reference
 */
export class WmlPart extends OpenXmlPart {
  /** Returns all header parts referenced by this part. */
  async headerParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.header) as Promise<WmlPart[]>;
  }

  /** Returns all footer parts referenced by this part. */
  async footerParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.footer) as Promise<WmlPart[]>;
  }

  /** Returns the endnotes part, or `undefined` if not present. */
  async endnotesPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.endnotes) as Promise<WmlPart | undefined>;
  }

  /** Returns the footnotes part, or `undefined` if not present. */
  async footnotesPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.footnotes) as Promise<WmlPart | undefined>;
  }

  /** Returns the comments part, or `undefined` if not present. */
  async wordprocessingCommentsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.wordprocessingComments) as Promise<WmlPart | undefined>;
  }

  /** Returns the font table part, or `undefined` if not present. */
  async fontTablePart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.fontTable) as Promise<WmlPart | undefined>;
  }

  /** Returns the numbering definitions part, or `undefined` if not present. */
  async numberingDefinitionsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.numberingDefinitions) as Promise<WmlPart | undefined>;
  }

  /** Returns the style definitions part, or `undefined` if not present. */
  async styleDefinitionsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.styles) as Promise<WmlPart | undefined>;
  }

  /** Returns the web settings part, or `undefined` if not present. */
  async webSettingsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.webSettings) as Promise<WmlPart | undefined>;
  }

  /** Returns the document settings part, or `undefined` if not present. */
  async documentSettingsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.documentSettings) as Promise<WmlPart | undefined>;
  }

  /** Returns the glossary document part (Quick Parts / AutoText), or `undefined` if not present. */
  async glossaryDocumentPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.glossaryDocument) as Promise<WmlPart | undefined>;
  }

  /** Returns all embedded font parts. */
  async fontParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.font) as Promise<WmlPart[]>;
  }
}
