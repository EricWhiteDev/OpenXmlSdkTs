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
 * Excel-specific part with convenience methods for navigating workbook structure.
 *
 * @remarks
 * Extends {@link OpenXmlPart} with typed accessors for worksheets, shared strings,
 * styles, charts, pivot tables, and other Excel-specific parts.
 *
 * @example
 * ```typescript
 * const workbook = await doc.workbookPart();
 * const worksheets = await workbook!.worksheetParts();
 * for (const ws of worksheets) {
 *   const xDoc = await ws.getXDocument();
 *   const rows = xDoc.root!.element(S.sheetData)?.elements(S.row) ?? [];
 *   console.log(`${ws.getUri()}: ${rows.length} rows`);
 * }
 * ```
 */
export class SmlPart extends OpenXmlPart {
  /** Returns the calculation chain part, or `undefined` if not present. */
  async calculationChainPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.calculationChain) as Promise<SmlPart | undefined>;
  }

  /** Returns the cell metadata part, or `undefined` if not present. */
  async cellMetadataPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.cellMetadata) as Promise<SmlPart | undefined>;
  }

  /** Returns the shared string table part, or `undefined` if not present. */
  async sharedStringTablePart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.sharedStringTable) as Promise<SmlPart | undefined>;
  }

  /** Returns the workbook revision header part, or `undefined` if not present. */
  async workbookRevisionHeaderPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.workbookRevisionHeader) as Promise<SmlPart | undefined>;
  }

  /** Returns the workbook styles part, or `undefined` if not present. */
  async workbookStylesPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.workbookStyles) as Promise<SmlPart | undefined>;
  }

  /** Returns the worksheet comments part, or `undefined` if not present. */
  async worksheetCommentsPart(): Promise<SmlPart | undefined> {
    return this.getPartByRelationshipType(RelationshipType.worksheetComments) as Promise<SmlPart | undefined>;
  }

  /** Returns all chartsheet parts. */
  async chartsheetParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.chartsheet) as Promise<SmlPart[]>;
  }

  /** Returns all worksheet parts. */
  async worksheetParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.worksheet) as Promise<SmlPart[]>;
  }

  /** Returns all pivot table parts. */
  async pivotTableParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.pivotTable) as Promise<SmlPart[]>;
  }

  /** Returns all query table parts. */
  async queryTableParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.queryTable) as Promise<SmlPart[]>;
  }

  /** Returns all table definition parts. */
  async tableDefinitionParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.tableDefinition) as Promise<SmlPart[]>;
  }

  /** Returns all timeline parts. */
  async timeLineParts(): Promise<SmlPart[]> {
    return this.getPartsByRelationshipType(RelationshipType.timeLine) as Promise<SmlPart[]>;
  }
}
