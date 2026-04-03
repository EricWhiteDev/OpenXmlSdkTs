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

export class WmlPart extends OpenXmlPart {
  async headerParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.header) as Promise<WmlPart[]>;
  }

  async footerParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.footer) as Promise<WmlPart[]>;
  }

  async endnotesPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.endnotes) as Promise<WmlPart | undefined>;
  }

  async footnotesPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.footnotes) as Promise<WmlPart | undefined>;
  }

  async wordprocessingCommentsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.wordprocessingComments) as Promise<WmlPart | undefined>;
  }

  async fontTablePart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.fontTable) as Promise<WmlPart | undefined>;
  }

  async numberingDefinitionsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.numberingDefinitions) as Promise<WmlPart | undefined>;
  }

  async styleDefinitionsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.styles) as Promise<WmlPart | undefined>;
  }

  async webSettingsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.webSettings) as Promise<WmlPart | undefined>;
  }

  async documentSettingsPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.documentSettings) as Promise<WmlPart | undefined>;
  }

  async glossaryDocumentPart(): Promise<WmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.glossaryDocument) as Promise<WmlPart | undefined>;
  }

  async fontParts(): Promise<WmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.font) as Promise<WmlPart[]>;
  }
}
