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

export class SmlPart extends OpenXmlPart {
  async calculationChainPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.calculationChain) as Promise<SmlPart | undefined>;
  }

  async cellMetadataPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.cellMetadata) as Promise<SmlPart | undefined>;
  }

  async sharedStringTablePart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.sharedStringTable) as Promise<SmlPart | undefined>;
  }

  async workbookRevisionHeaderPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.workbookRevisionHeader) as Promise<SmlPart | undefined>;
  }

  async workbookStylesPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.workbookStyles) as Promise<SmlPart | undefined>;
  }

  async worksheetCommentsPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.worksheetComments) as Promise<SmlPart | undefined>;
  }

  async chartsheetParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.chartsheet) as Promise<SmlPart[]>;
  }

  async worksheetParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.worksheet) as Promise<SmlPart[]>;
  }

  async pivotTableParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.pivotTable) as Promise<SmlPart[]>;
  }

  async queryTableParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.queryTable) as Promise<SmlPart[]>;
  }

  async tableDefinitionParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.tableDefinition) as Promise<SmlPart[]>;
  }

  async timeLineParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.timeLine) as Promise<SmlPart[]>;
  }
}
