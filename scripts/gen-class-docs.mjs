/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, WidthType, ShadingType, AlignmentType, BorderStyle,
} from 'docx';
import { writeFileSync, mkdirSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DOCS_DIR = join(__dirname, '..', 'docs');

// ─── Palette & constants ─────────────────────────────────────────────────────
const BLUE       = '2E74B5';
const LIGHT_BLUE = 'DEEAF1';
const CODE_BG    = 'F2F2F2';
const DARK       = '2F2F2F';
const MID        = '595959';
const WHITE      = 'FFFFFF';
const BODY       = 'Calibri';
const MONO       = 'Consolas';

// ─── Block helpers ────────────────────────────────────────────────────────────
const gap = () => new Paragraph({ text: '', spacing: { after: 60 } });

function h1(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 40, font: BODY, color: BLUE })],
    spacing: { before: 0, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE } },
  });
}

function h2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, font: BODY, color: BLUE })],
    spacing: { before: 200, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: LIGHT_BLUE } },
  });
}

function body(runs) {
  return new Paragraph({
    children: runs.map(r => new TextRun({
      text: r.text,
      bold: r.bold ?? false,
      italics: r.italic ?? false,
      font: r.code ? MONO : BODY,
      size: r.code ? 18 : 20,
      color: r.code ? '7B3F00' : DARK,
    })),
    spacing: { after: 80 },
  });
}

const plain = text => body([{ text }]);

function label(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 20, font: BODY, color: DARK })],
    spacing: { before: 100, after: 40 },
  });
}

function note(text) {
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: 'FFF9E6' },
    children: [new TextRun({ text: '\u2691  ' + text, size: 18, font: BODY, color: '7B5700', italics: true })],
    indent: { left: 360 },
    spacing: { before: 60, after: 60 },
  });
}

function badge(text, fill, color = WHITE) {
  return new TextRun({ text: ` ${text} `, bold: true, size: 16, font: BODY, color, shading: { type: ShadingType.CLEAR, fill } });
}

function memberHeading(name, kind) {
  const kindColor = {
    property:          '4472C4',
    getter:            '4472C4',
    'getter/setter':   '4472C4',
    method:            '375623',
    'static method':   '7030A0',
    'static property': '7030A0',
    constructor:       'C55A11',
  }[kind] ?? '404040';
  return new Paragraph({
    children: [
      new TextRun({ text: name, bold: true, size: 24, font: MONO, color: DARK }),
      new TextRun({ text: '  ' }),
      badge(kind, kindColor),
    ],
    spacing: { before: 200, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' } },
  });
}

function sigLine(text) {
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: CODE_BG },
    children: [new TextRun({ text, font: MONO, size: 18, color: '1F3864' })],
    indent: { left: 360 },
    spacing: { before: 0, after: 0 },
  });
}

function paramsTable(rows) {
  if (!rows || rows.length === 0) return null;
  const headerRow = new TableRow({
    tableHeader: true,
    children: ['Parameter', 'Type', 'Description'].map((t, i) =>
      new TableCell({
        shading: { type: ShadingType.CLEAR, fill: BLUE },
        width: { size: [22, 22, 56][i], type: WidthType.PERCENTAGE },
        children: [new Paragraph({
          children: [new TextRun({ text: t, bold: true, color: WHITE, size: 18, font: BODY })],
          alignment: AlignmentType.CENTER,
        })],
      })
    ),
  });
  const dataRows = rows.map(([p, t, d], i) =>
    new TableRow({
      children: [p, t, d].map(cell =>
        new TableCell({
          shading: i % 2 === 1 ? { type: ShadingType.CLEAR, fill: LIGHT_BLUE } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: MONO, size: 18, color: DARK })] })],
        })
      ),
    })
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

function twoColTable(headers, rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((t, i) =>
      new TableCell({
        shading: { type: ShadingType.CLEAR, fill: BLUE },
        width: { size: [30, 70][i], type: WidthType.PERCENTAGE },
        children: [new Paragraph({
          children: [new TextRun({ text: t, bold: true, color: WHITE, size: 18, font: BODY })],
          alignment: AlignmentType.CENTER,
        })],
      })
    ),
  });
  const dataRows = rows.map(([a, b], i) =>
    new TableRow({
      children: [a, b].map(cell =>
        new TableCell({
          shading: i % 2 === 1 ? { type: ShadingType.CLEAR, fill: LIGHT_BLUE } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: MONO, size: 16, color: DARK })] })],
        })
      ),
    })
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
}

function codeLines(lines) {
  return lines.map((line, i) => new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: CODE_BG },
    children: [new TextRun({ text: line, font: MONO, size: 18, color: '1F3864' })],
    indent: { left: 360 },
    spacing: { before: 0, after: i === lines.length - 1 ? 80 : 0 },
  }));
}

function inheritanceLine(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: 'Inheritance: ', bold: true, size: 18, font: BODY, color: MID }),
      new TextRun({ text, size: 18, font: MONO, color: DARK }),
    ],
    spacing: { after: 120 },
  });
}

// ─── Member builder ───────────────────────────────────────────────────────────
function member({ name, kind, sigs, params, returns, semantics, example, description }) {
  const blocks = [memberHeading(name, kind)];
  if (Array.isArray(description)) description.forEach(d => blocks.push(plain(d)));
  else blocks.push(plain(description));

  if (sigs && sigs.length > 0) {
    blocks.push(label('Signature'));
    sigs.forEach(s => blocks.push(sigLine(s)));
    blocks.push(gap());
  }

  const pt = paramsTable(params);
  if (pt) { blocks.push(label('Parameters')); blocks.push(pt); blocks.push(gap()); }

  if (returns) {
    blocks.push(label('Returns'));
    blocks.push(body([{ text: returns }]));
  }

  if (semantics) {
    blocks.push(label('Special Semantics'));
    if (Array.isArray(semantics)) semantics.forEach(s => blocks.push(note(s)));
    else blocks.push(note(semantics));
  }

  if (example && example.length > 0) {
    blocks.push(label('Example'));
    blocks.push(...codeLines(example));
  }

  return blocks;
}

// ─── Document factory ─────────────────────────────────────────────────────────
async function writeDoc(filename, classContent) {
  const doc = new Document({ sections: [{ children: classContent }] });
  const buf = await Packer.toBuffer(doc);
  writeFileSync(join(DOCS_DIR, `${filename}.docx`), buf);
  console.log(`  wrote ${filename}.docx`);
}


// ════════════════════════════════════════════════════════════════════════════
// DOCUMENT DEFINITIONS
// ════════════════════════════════════════════════════════════════════════════

// ─── Overview ───────────────────────────────────────────────────────────────
async function genOverview() {
  const content = [
    h1('OpenXmlSdkTs \u2014 Overview'),

    plain('OpenXmlSdkTs is a TypeScript library for reading, writing, and manipulating Office Open XML documents (.docx, .xlsx, .pptx) in Node.js and browser environments. It is the TypeScript equivalent of the .NET Open XML SDK.'),
    plain('The library provides a strongly-typed API for navigating the internal structure of Open XML packages \u2014 parts, relationships, content types \u2014 and for reading and modifying the XML within each part using LINQ-to-XML style operations via the companion library LtXmlTs.'),

    h2('Main Use Cases'),
    plain('\u2022  Office Add-ins \u2014 Manipulate the active document inside Word, Excel, or PowerPoint using the Flat OPC format that Office.js provides.'),
    plain('\u2022  Server-side document generation \u2014 Create or transform documents in Node.js pipelines, APIs, or serverless functions.'),
    plain('\u2022  Document transformation \u2014 Read a document, modify its XML content (styles, text, tables, charts), and save it back in any of three supported formats.'),
    plain('\u2022  Document inspection \u2014 Enumerate parts, relationships, and content types to understand or validate document structure.'),

    h2('Supported I/O Formats'),
    plain('Every package class can open and save documents in three interchangeable formats:'),
    plain('\u2022  Binary (Blob) \u2014 The native .docx / .xlsx / .pptx ZIP format.'),
    plain('\u2022  Flat OPC XML (string) \u2014 A single XML file containing all parts. Required by Office.js Add-ins.'),
    plain('\u2022  Base64 (string) \u2014 The binary ZIP encoded as a Base64 string, useful for APIs and serialization.'),
    plain('The static open() method auto-detects the format from its argument type.'),

    h2('Class Hierarchy'),
    ...codeLines([
      'OpenXmlPackage           (base package)',
      '  \u251C\u2500 WmlPackage           (Word .docx)',
      '  \u251C\u2500 SmlPackage           (Excel .xlsx)',
      '  \u2514\u2500 PmlPackage           (PowerPoint .pptx)',
      '',
      'OpenXmlPart              (base part)',
      '  \u251C\u2500 WmlPart              (Word parts)',
      '  \u251C\u2500 SmlPart              (Excel parts)',
      '  \u2514\u2500 PmlPart              (PowerPoint parts)',
      '',
      'OpenXmlRelationship      (relationship between parts)',
      '',
      'ContentType              (97 content type constants)',
      'RelationshipType         (89 relationship type constants)',
      'Utility                  (static helper methods)',
      '',
      'Namespace classes (48)   (pre-atomized XName/XNamespace)',
      '  W, S, P, A, R, M, NoNamespace, ...',
    ]),

    h2('Quick Start'),
    label('Open a Word document, read paragraphs, add text, and save'),
    ...codeLines([
      'import { WmlPackage, W, XElement } from "openxmlsdkts";',
      'import fs from "fs";',
      '',
      'const buffer = fs.readFileSync("report.docx");',
      'const doc = await WmlPackage.open(new Blob([buffer]));',
      '',
      '// Navigate to the main document part',
      'const mainPart = await doc.mainDocumentPart();',
      'const xDoc = await mainPart!.getXDocument();',
      '',
      '// Read existing paragraphs',
      'const body = xDoc.root!.element(W.body);',
      'const paragraphs = body!.elements(W.p);',
      'console.log(`Document has ${paragraphs.length} paragraphs`);',
      '',
      '// Add a new paragraph',
      'const newPara = new XElement(',
      '  W.p,',
      '  new XElement(W.r, new XElement(W.t, "Hello from OpenXmlSdkTs!"))',
      ');',
      'body!.add(newPara);',
      '',
      '// Write back and save',
      'mainPart!.putXDocument(xDoc);',
      'const blob = await doc.saveToBlobAsync();',
    ]),

    h2('Dependencies'),
    plain('\u2022  ltxmlts \u2014 LINQ to XML for TypeScript. Provides XDocument, XElement, XAttribute, XName, and XNamespace for XML tree manipulation.'),
    plain('\u2022  jszip \u2014 ZIP compression library used internally to read and write the binary Open XML package format.'),

    h2('Pre-Atomized Namespace Classes'),
    plain('OpenXmlSdkTs ships 48 static namespace classes (W, S, P, A, R, NoNamespace, etc.) with pre-initialized XName objects for every element and attribute name in the Open XML specification. Using these atomized names enables O(1) identity-based equality checks instead of string comparison, improving both performance and code readability.'),
    ...codeLines([
      '// Instead of string-based lookup:',
      '// element.element(XName.get("body", "http://schemas.openxml.../wordprocessingml/2006/main"))',
      '',
      '// Use pre-atomized names:',
      'const body = xDoc.root!.element(W.body);',
      'const style = pPr?.element(W.pStyle)?.attribute(NoNamespace.val)?.value;',
    ]),

    h2('Office Add-ins Workflow'),
    plain('Office Add-ins exchange document content with the host application in Flat OPC XML format. OpenXmlSdkTs handles this natively:'),
    ...codeLines([
      '// Receive Flat OPC from Office.js API',
      'const doc = await WmlPackage.open(flatOpcString);',
      '',
      '// Modify the document...',
      'const mainPart = await doc.mainDocumentPart();',
      '// ... make changes ...',
      '',
      '// Return modified Flat OPC to Office.js',
      'const modifiedFlatOpc = await doc.saveToFlatOpcAsync();',
    ]),
  ];
  await writeDoc('OpenXmlSdkTs-Overview', content);
}

// ─── OpenXmlPackage ─────────────────────────────────────────────────────────
async function genOpenXmlPackage() {
  const content = [
    h1('OpenXmlPackage'),
    inheritanceLine('OpenXmlPackage'),
    plain('OpenXmlPackage is the base class for all Office Open XML document packages. It manages the collection of parts and relationships that make up a document, and provides methods to open, navigate, modify, and save packages in three formats: binary (Blob), Flat OPC XML, and Base64.'),
    plain('You typically use one of the format-specific subclasses \u2014 WmlPackage, SmlPackage, or PmlPackage \u2014 rather than OpenXmlPackage directly. However, OpenXmlPackage.open() can open any Open XML document when you do not need format-specific convenience methods.'),

    h2('Type Aliases'),
    ...member({
      name: 'Base64String',
      kind: 'static property',
      sigs: ['type Base64String = string'],
      description: 'A type alias for a string that contains a Base64-encoded Open XML document. Used as an input to open() and as the return type of saveToBase64Async().',
    }),
    ...member({
      name: 'FlatOpcString',
      kind: 'static property',
      sigs: ['type FlatOpcString = string'],
      description: 'A type alias for a string that contains a Flat OPC XML representation of an Open XML document. This is the format used by Office Add-ins to exchange document content with the Office host application.',
    }),
    ...member({
      name: 'DocxBinary',
      kind: 'static property',
      sigs: ['type DocxBinary = Blob'],
      description: 'A type alias for a Blob containing the raw binary (ZIP) content of an Open XML document.',
    }),

    h2('Opening Documents'),
    ...member({
      name: 'open',
      kind: 'static method',
      sigs: ['static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<OpenXmlPackage>'],
      params: [['document', 'Base64String | FlatOpcString | DocxBinary', 'The document to open. The format is auto-detected from the argument type and content.']],
      returns: 'Promise<OpenXmlPackage> \u2014 the opened package.',
      description: 'Opens an Open XML document from any of the three supported formats. When the argument is a string, the method inspects the content to determine whether it is Base64-encoded or Flat OPC XML. When it is a Blob, it is treated as the native binary ZIP format.',
      semantics: 'Use the format-specific subclass (WmlPackage.open, SmlPackage.open, PmlPackage.open) when you need typed part accessors.',
      example: [
        '// Open from binary Blob',
        'const buffer = fs.readFileSync("report.docx");',
        'const pkg = await OpenXmlPackage.open(new Blob([buffer]));',
        '',
        '// Open from Base64',
        'const base64 = fs.readFileSync("report.txt", "utf-8");',
        'const pkg2 = await OpenXmlPackage.open(base64);',
        '',
        '// Open from Flat OPC XML',
        'const flatOpc = fs.readFileSync("report.xml", "utf-8");',
        'const pkg3 = await OpenXmlPackage.open(flatOpc);',
      ],
    }),

    h2('Saving Documents'),
    ...member({
      name: 'saveToBlobAsync',
      kind: 'method',
      sigs: ['async saveToBlobAsync(): Promise<DocxBinary>'],
      returns: 'Promise<DocxBinary> \u2014 a Blob containing the binary ZIP (.docx/.xlsx/.pptx) document.',
      description: 'Serializes the package to its native binary ZIP format and returns it as a Blob. The resulting Blob can be written to a file or sent over the network.',
      example: [
        'const blob = await pkg.saveToBlobAsync();',
        'const buffer = Buffer.from(await blob.arrayBuffer());',
        'fs.writeFileSync("output.docx", buffer);',
      ],
    }),
    ...member({
      name: 'saveToBase64Async',
      kind: 'method',
      sigs: ['async saveToBase64Async(): Promise<Base64String>'],
      returns: 'Promise<Base64String> \u2014 the document encoded as a Base64 string.',
      description: 'Serializes the package to its binary ZIP format, then encodes the result as a Base64 string. Useful for APIs that accept or return documents as strings.',
      example: [
        'const base64 = await pkg.saveToBase64Async();',
        'fs.writeFileSync("output.txt", base64, "utf-8");',
      ],
    }),
    ...member({
      name: 'saveToFlatOpcAsync',
      kind: 'method',
      sigs: ['async saveToFlatOpcAsync(): Promise<FlatOpcString>'],
      returns: 'Promise<FlatOpcString> \u2014 the document as Flat OPC XML.',
      description: 'Serializes the package to Flat OPC XML format, a single XML document that contains all parts. This is the format required by Office.js Add-ins.',
      example: [
        'const flatOpc = await pkg.saveToFlatOpcAsync();',
        'fs.writeFileSync("output.xml", flatOpc, "utf-8");',
      ],
    }),

    h2('Part Management'),
    ...member({
      name: 'getParts',
      kind: 'method',
      sigs: ['getParts(): OpenXmlPart[]'],
      returns: 'OpenXmlPart[] \u2014 all parts in the package, excluding the content types part and relationship parts.',
      description: 'Returns an array of all user-visible parts in the package. Internal parts ([Content_Types].xml and .rels files) are filtered out.',
      example: [
        'const parts = pkg.getParts();',
        'for (const part of parts) {',
        '  console.log(part.getUri(), part.getContentType());',
        '}',
      ],
    }),
    ...member({
      name: 'addPart',
      kind: 'method',
      sigs: ['addPart(uri: string, contentType: string, partType: PartType, data: unknown): OpenXmlPart'],
      params: [
        ['uri', 'string', 'The URI path for the new part (e.g., "/word/newPart.xml").'],
        ['contentType', 'string', 'The MIME content type. Use ContentType constants.'],
        ['partType', 'PartType', 'The data format: "xml", "binary", "base64", or null.'],
        ['data', 'unknown', 'The part data (XDocument for XML parts, string for Base64, etc.).'],
      ],
      returns: 'OpenXmlPart \u2014 the newly added part.',
      description: 'Adds a new part to the package with the specified URI, content type, and data. Throws an error if a part with the same URI already exists.',
      semantics: 'The part is added to both the internal parts map and the [Content_Types].xml. You must also add a relationship to the part for it to be reachable.',
      example: [
        'const xDoc = new XDocument(new XElement("root"));',
        'const newPart = pkg.addPart("/word/custom.xml",',
        '  ContentType.customXmlProperties, "xml", xDoc);',
      ],
    }),
    ...member({
      name: 'deletePart',
      kind: 'method',
      sigs: ['async deletePart(part: OpenXmlPart): Promise<void>'],
      params: [['part', 'OpenXmlPart', 'The part to remove from the package.']],
      description: 'Removes a part from the package. Also removes the part\'s entry from [Content_Types].xml and cleans up any relationships in sibling .rels files that target the deleted part.',
      example: [
        'const commentsPart = await mainPart.wordprocessingCommentsPart();',
        'if (commentsPart) {',
        '  await pkg.deletePart(commentsPart);',
        '}',
      ],
    }),
    ...member({
      name: 'getPartByUri',
      kind: 'method',
      sigs: ['getPartByUri(uri: string): OpenXmlPart | undefined'],
      params: [['uri', 'string', 'The URI path of the part (e.g., "/word/document.xml").']],
      returns: 'OpenXmlPart | undefined \u2014 the part at that URI, or undefined if not found.',
      description: 'Looks up a part by its exact URI path within the package.',
      example: [
        'const part = pkg.getPartByUri("/word/document.xml");',
      ],
    }),
    ...member({
      name: 'getContentType',
      kind: 'method',
      sigs: ['getContentType(uri: string): string'],
      params: [['uri', 'string', 'The URI path of the part.']],
      returns: 'string \u2014 the MIME content type for the part.',
      description: 'Returns the content type for a part URI by looking it up in [Content_Types].xml. First checks Override entries, then falls back to Default entries based on file extension. Throws if the content type cannot be determined.',
      example: [
        'const ct = pkg.getContentType("/word/document.xml");',
        '// "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"',
      ],
    }),

    h2('Relationship Queries'),
    ...member({
      name: 'getRelationships',
      kind: 'method',
      sigs: ['async getRelationships(): Promise<OpenXmlRelationship[]>'],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 all package-level relationships from /_rels/.rels.',
      description: 'Returns all package-level relationships. These are the top-level relationships defined in the package\'s root .rels file, typically pointing to the main document part, core properties, extended properties, and similar top-level parts.',
      example: [
        'const rels = await pkg.getRelationships();',
        'for (const rel of rels) {',
        '  console.log(rel.getId(), rel.getType(), rel.getTarget());',
        '}',
      ],
    }),
    ...member({
      name: 'getRelationshipsByRelationshipType',
      kind: 'method',
      sigs: ['async getRelationshipsByRelationshipType(relationshipType: string): Promise<OpenXmlRelationship[]>'],
      params: [['relationshipType', 'string', 'The relationship type URI. Use RelationshipType constants.']],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 matching package-level relationships.',
      description: 'Returns all package-level relationships that match the given relationship type URI.',
      example: [
        'const rels = await pkg.getRelationshipsByRelationshipType(',
        '  RelationshipType.coreFileProperties);',
      ],
    }),
    ...member({
      name: 'getPartsByRelationshipType',
      kind: 'method',
      sigs: ['async getPartsByRelationshipType(relationshipType: string): Promise<OpenXmlPart[]>'],
      params: [['relationshipType', 'string', 'The relationship type URI.']],
      returns: 'Promise<OpenXmlPart[]> \u2014 the parts targeted by matching relationships.',
      description: 'Returns all parts that are targets of package-level relationships with the given type. This combines getRelationshipsByRelationshipType() with getPartByUri() to resolve the actual parts.',
      example: [
        'const parts = await pkg.getPartsByRelationshipType(',
        '  RelationshipType.extendedFileProperties);',
      ],
    }),
    ...member({
      name: 'getPartByRelationshipType',
      kind: 'method',
      sigs: ['async getPartByRelationshipType(relationshipType: string): Promise<OpenXmlPart | undefined>'],
      params: [['relationshipType', 'string', 'The relationship type URI.']],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the first matching part, or undefined.',
      description: 'Convenience method that returns the first part targeted by a package-level relationship of the given type. Useful when there is expected to be only one relationship of a given type (e.g., the main document part).',
      example: [
        'const corePropsPart = await pkg.getPartByRelationshipType(',
        '  RelationshipType.coreFileProperties);',
      ],
    }),
    ...member({
      name: 'getRelationshipById',
      kind: 'method',
      sigs: ['async getRelationshipById(rId: string): Promise<OpenXmlRelationship | undefined>'],
      params: [['rId', 'string', 'The relationship ID (e.g., "rId1").']],
      returns: 'Promise<OpenXmlRelationship | undefined> \u2014 the matching relationship, or undefined.',
      description: 'Finds a package-level relationship by its unique ID.',
      example: [
        'const rel = await pkg.getRelationshipById("rId1");',
      ],
    }),
    ...member({
      name: 'getPartById',
      kind: 'method',
      sigs: ['async getPartById(rId: string): Promise<OpenXmlPart | undefined>'],
      params: [['rId', 'string', 'The relationship ID.']],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the part targeted by the relationship, or undefined.',
      description: 'Returns the part targeted by the package-level relationship with the given ID.',
      example: [
        'const part = await pkg.getPartById("rId1");',
      ],
    }),
    ...member({
      name: 'getRelationshipsByContentType',
      kind: 'method',
      sigs: ['async getRelationshipsByContentType(contentType: string): Promise<OpenXmlRelationship[]>'],
      params: [['contentType', 'string', 'The MIME content type. Use ContentType constants.']],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 relationships whose targets have the specified content type.',
      description: 'Returns all package-level relationships whose target parts have the specified content type. External relationships are excluded.',
      example: [
        'const rels = await pkg.getRelationshipsByContentType(',
        '  ContentType.mainDocument);',
      ],
    }),
    ...member({
      name: 'getPartsByContentType',
      kind: 'method',
      sigs: ['async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]>'],
      params: [['contentType', 'string', 'The MIME content type.']],
      returns: 'Promise<OpenXmlPart[]> \u2014 parts with the specified content type.',
      description: 'Returns all parts targeted by package-level relationships whose targets have the specified content type.',
      example: [
        'const parts = await pkg.getPartsByContentType(ContentType.theme);',
      ],
    }),

    h2('Properties Parts'),
    ...member({
      name: 'coreFilePropertiesPart',
      kind: 'method',
      sigs: ['async coreFilePropertiesPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the core file properties part (title, author, etc.), or undefined.',
      description: 'Returns the core file properties part, which contains Dublin Core metadata such as title, subject, creator, and dates.',
      example: [
        'const corePart = await pkg.coreFilePropertiesPart();',
        'if (corePart) {',
        '  const xDoc = await corePart.getXDocument();',
        '  console.log(xDoc.toString());',
        '}',
      ],
    }),
    ...member({
      name: 'extendedFilePropertiesPart',
      kind: 'method',
      sigs: ['async extendedFilePropertiesPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the extended file properties part, or undefined.',
      description: 'Returns the extended file properties part, which contains application-specific metadata such as the application name, company, and document statistics (word count, page count, etc.).',
      example: [
        'const extPart = await pkg.extendedFilePropertiesPart();',
      ],
    }),
    ...member({
      name: 'customFilePropertiesPart',
      kind: 'method',
      sigs: ['async customFilePropertiesPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the custom file properties part, or undefined.',
      description: 'Returns the custom file properties part, which contains user-defined name/value metadata pairs.',
      example: [
        'const customPart = await pkg.customFilePropertiesPart();',
      ],
    }),

    h2('Part Relationship Management'),
    ...member({
      name: 'getRelationshipsForPart',
      kind: 'method',
      sigs: ['async getRelationshipsForPart(part: OpenXmlPart): Promise<OpenXmlRelationship[]>'],
      params: [['part', 'OpenXmlPart', 'The part whose relationships to retrieve.']],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 the relationships defined for the given part.',
      description: 'Returns all relationships defined in the .rels file associated with the given part. This is the part-level equivalent of getRelationships().',
      example: [
        'const mainPart = await doc.mainDocumentPart();',
        'const rels = await pkg.getRelationshipsForPart(mainPart!);',
      ],
    }),
    ...member({
      name: 'addRelationship',
      kind: 'method',
      sigs: ['async addRelationship(id: string, type: string, target: string, targetMode?: string): Promise<OpenXmlRelationship>'],
      params: [
        ['id', 'string', 'The relationship ID (e.g., "rId99").'],
        ['type', 'string', 'The relationship type URI. Use RelationshipType constants.'],
        ['target', 'string', 'The target URI (relative or absolute).'],
        ['targetMode', 'string  (optional)', 'The target mode: "Internal" (default) or "External".'],
      ],
      returns: 'Promise<OpenXmlRelationship> \u2014 the newly created relationship.',
      description: 'Adds a new package-level relationship to /_rels/.rels. Creates the .rels part if it does not exist.',
      example: [
        'const rel = await pkg.addRelationship(',
        '  "rId99", RelationshipType.coreFileProperties, "docProps/core.xml");',
      ],
    }),
    ...member({
      name: 'addRelationshipForPart',
      kind: 'method',
      sigs: ['async addRelationshipForPart(part: OpenXmlPart, id: string, type: string, target: string, targetMode?: string): Promise<OpenXmlRelationship>'],
      params: [
        ['part', 'OpenXmlPart', 'The source part that owns the relationship.'],
        ['id', 'string', 'The relationship ID.'],
        ['type', 'string', 'The relationship type URI.'],
        ['target', 'string', 'The target URI.'],
        ['targetMode', 'string  (optional)', '"Internal" (default) or "External".'],
      ],
      returns: 'Promise<OpenXmlRelationship> \u2014 the newly created relationship.',
      description: 'Adds a new relationship in the .rels file associated with the given part. Creates the .rels part if it does not exist.',
      example: [
        'const rel = await pkg.addRelationshipForPart(',
        '  mainPart, "rId10", RelationshipType.image, "media/image1.png");',
      ],
    }),
    ...member({
      name: 'deleteRelationship',
      kind: 'method',
      sigs: ['async deleteRelationship(id: string): Promise<boolean>'],
      params: [['id', 'string', 'The relationship ID to delete.']],
      returns: 'Promise<boolean> \u2014 true if the relationship was deleted.',
      description: 'Deletes a package-level relationship by its ID from /_rels/.rels. Throws if the relationship is not found.',
      example: [
        'await pkg.deleteRelationship("rId99");',
      ],
    }),
    ...member({
      name: 'deleteRelationshipForPart',
      kind: 'method',
      sigs: ['async deleteRelationshipForPart(part: OpenXmlPart, id: string): Promise<boolean>'],
      params: [
        ['part', 'OpenXmlPart', 'The part that owns the relationship.'],
        ['id', 'string', 'The relationship ID to delete.'],
      ],
      returns: 'Promise<boolean> \u2014 true if the relationship was deleted.',
      description: 'Deletes a relationship by its ID from the .rels file associated with the given part. Throws if the relationship is not found.',
      example: [
        'await pkg.deleteRelationshipForPart(mainPart, "rId10");',
      ],
    }),
  ];
  await writeDoc('OpenXmlPackage', content);
}

// ─── OpenXmlPart ────────────────────────────────────────────────────────────
async function genOpenXmlPart() {
  const content = [
    h1('OpenXmlPart'),
    inheritanceLine('OpenXmlPart'),
    plain('OpenXmlPart represents a single part within an Open XML package. A part is a unit of content \u2014 such as the main document XML, a styles definition, an image, or a relationship file \u2014 identified by a URI within the package.'),
    plain('OpenXmlPart provides methods to access the part\'s metadata (URI, content type, data), navigate its relationships to other parts, and read or modify its XML content via XDocument.'),

    h2('Types'),
    ...member({
      name: 'PartType',
      kind: 'static property',
      sigs: ['type PartType = "binary" | "base64" | "xml" | null'],
      description: 'Indicates the data format of a part. "xml" parts contain XML data accessible via getXDocument(). "binary" parts contain raw bytes. "base64" parts contain Base64-encoded binary data. null indicates the type has not been determined yet.',
    }),

    h2('Constructor'),
    ...member({
      name: 'constructor',
      kind: 'constructor',
      sigs: ['constructor(pkg: OpenXmlPackage, uri: string, contentType: string | null, partType: PartType, data: unknown)'],
      params: [
        ['pkg', 'OpenXmlPackage', 'The package that contains this part.'],
        ['uri', 'string', 'The URI path of this part within the package.'],
        ['contentType', 'string | null', 'The MIME content type, or null if not yet determined.'],
        ['partType', 'PartType', 'The data format of the part.'],
        ['data', 'unknown', 'The part data.'],
      ],
      description: 'Creates a new OpenXmlPart. Typically you do not call this directly; parts are created by OpenXmlPackage when opening a document or by addPart().',
    }),

    h2('Accessors'),
    ...member({
      name: 'getUri',
      kind: 'method',
      sigs: ['getUri(): string'],
      returns: 'string \u2014 the URI path of this part (e.g., "/word/document.xml").',
      description: 'Returns the URI path of this part within the package.',
      example: [
        'const uri = part.getUri(); // "/word/document.xml"',
      ],
    }),
    ...member({
      name: 'getPkg',
      kind: 'method',
      sigs: ['getPkg(): OpenXmlPackage'],
      returns: 'OpenXmlPackage \u2014 the package containing this part.',
      description: 'Returns a reference to the OpenXmlPackage that contains this part.',
      example: [
        'const pkg = part.getPkg();',
      ],
    }),
    ...member({
      name: 'getContentType',
      kind: 'method',
      sigs: ['getContentType(): string | null'],
      returns: 'string | null \u2014 the MIME content type, or null.',
      description: 'Returns the MIME content type of this part (e.g., "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml").',
      example: [
        'const ct = part.getContentType();',
      ],
    }),
    ...member({
      name: 'getData',
      kind: 'method',
      sigs: ['getData(): unknown'],
      returns: 'unknown \u2014 the raw data of this part.',
      description: 'Returns the underlying data of this part. For XML parts that have been accessed via getXDocument(), this is an XDocument. For freshly opened binary parts, this is a JSZip file entry. Prefer getXDocument() for XML parts.',
    }),
    ...member({
      name: 'setContentType',
      kind: 'method',
      sigs: ['setContentType(ct: string): void'],
      params: [['ct', 'string', 'The new MIME content type.']],
      description: 'Sets the content type of this part. Typically used internally during package loading.',
    }),
    ...member({
      name: 'getPartType',
      kind: 'method',
      sigs: ['getPartType(): PartType'],
      returns: 'PartType \u2014 "xml", "binary", "base64", or null.',
      description: 'Returns the data format of this part.',
    }),
    ...member({
      name: 'setData',
      kind: 'method',
      sigs: ['setData(data: unknown): void'],
      params: [['data', 'unknown', 'The new data for this part.']],
      description: 'Replaces the data of this part. For XML parts, prefer putXDocument().',
    }),
    ...member({
      name: 'setPartType',
      kind: 'method',
      sigs: ['setPartType(pt: PartType): void'],
      params: [['pt', 'PartType', 'The new part type.']],
      description: 'Sets the data format of this part. Typically used internally during package loading.',
    }),

    h2('XML Access'),
    ...member({
      name: 'getXDocument',
      kind: 'method',
      sigs: ['async getXDocument(): Promise<XDocument>'],
      returns: 'Promise<XDocument> \u2014 the parsed XML document for this part.',
      description: 'Returns the XML content of this part as an XDocument. If the XML has not been parsed yet, it is parsed from the raw data on first access and cached for subsequent calls. Throws if the part is not an XML part.',
      semantics: 'The returned XDocument is the live internal copy. Modifications to it are reflected when the package is saved. Call putXDocument() to replace it entirely.',
      example: [
        'const mainPart = await doc.mainDocumentPart();',
        'const xDoc = await mainPart!.getXDocument();',
        'const body = xDoc.root!.element(W.body);',
        'console.log(body!.elements(W.p).length);',
      ],
    }),
    ...member({
      name: 'putXDocument',
      kind: 'method',
      sigs: ['putXDocument(xDoc: XDocument): void'],
      params: [['xDoc', 'XDocument', 'The XDocument to store as this part\'s content. Must not be null or undefined.']],
      description: 'Replaces this part\'s XML content with the given XDocument. Also sets the part type to "xml". Throws if xDoc is null or undefined.',
      example: [
        'const xDoc = await mainPart!.getXDocument();',
        '// ... modify xDoc ...',
        'mainPart!.putXDocument(xDoc);',
      ],
    }),

    h2('Relationship Navigation'),
    ...member({
      name: 'getRelsPartUri',
      kind: 'method',
      sigs: ['getRelsPartUri(): string'],
      returns: 'string \u2014 the URI of the .rels file for this part.',
      description: 'Returns the URI of the .rels file associated with this part. For example, for "/word/document.xml" this returns "/word/_rels/document.xml.rels".',
      example: [
        'const relsUri = part.getRelsPartUri();',
        '// "/word/_rels/document.xml.rels"',
      ],
    }),
    ...member({
      name: 'getRelsPart',
      kind: 'method',
      sigs: ['getRelsPart(): OpenXmlPart | undefined'],
      returns: 'OpenXmlPart | undefined \u2014 the .rels part for this part, or undefined if none exists.',
      description: 'Returns the .rels part associated with this part, if it exists in the package.',
    }),
    ...member({
      name: 'getRelationships',
      kind: 'method',
      sigs: ['async getRelationships(): Promise<OpenXmlRelationship[]>'],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 all relationships defined for this part.',
      description: 'Returns all relationships from this part\'s .rels file.',
      example: [
        'const rels = await mainPart!.getRelationships();',
        'for (const r of rels) {',
        '  console.log(r.getId(), r.getType());',
        '}',
      ],
    }),
    ...member({
      name: 'getParts',
      kind: 'method',
      sigs: ['async getParts(): Promise<OpenXmlPart[]>'],
      returns: 'Promise<OpenXmlPart[]> \u2014 all parts that this part has relationships to (excluding external targets).',
      description: 'Returns all parts that are targets of this part\'s relationships, excluding external targets. Throws if a relationship target cannot be resolved to a part.',
      example: [
        'const childParts = await mainPart!.getParts();',
      ],
    }),
    ...member({
      name: 'getRelationshipsByRelationshipType',
      kind: 'method',
      sigs: ['async getRelationshipsByRelationshipType(relationshipType: string): Promise<OpenXmlRelationship[]>'],
      params: [['relationshipType', 'string', 'The relationship type URI.']],
      returns: 'Promise<OpenXmlRelationship[]> \u2014 matching relationships.',
      description: 'Returns all relationships of this part that match the given relationship type.',
      example: [
        'const imageRels = await mainPart!.getRelationshipsByRelationshipType(',
        '  RelationshipType.image);',
      ],
    }),
    ...member({
      name: 'getPartsByRelationshipType',
      kind: 'method',
      sigs: ['async getPartsByRelationshipType(relationshipType: string): Promise<OpenXmlPart[]>'],
      params: [['relationshipType', 'string', 'The relationship type URI.']],
      returns: 'Promise<OpenXmlPart[]> \u2014 matching parts.',
      description: 'Returns all parts targeted by relationships of the given type from this part.',
      example: [
        'const imageParts = await mainPart!.getPartsByRelationshipType(',
        '  RelationshipType.image);',
      ],
    }),
    ...member({
      name: 'getPartByRelationshipType',
      kind: 'method',
      sigs: ['async getPartByRelationshipType(relationshipType: string): Promise<OpenXmlPart | undefined>'],
      params: [['relationshipType', 'string', 'The relationship type URI.']],
      returns: 'Promise<OpenXmlPart | undefined> \u2014 the first matching part, or undefined.',
      description: 'Returns the first part targeted by a relationship of the given type from this part.',
      example: [
        'const stylesPart = await mainPart!.getPartByRelationshipType(',
        '  RelationshipType.styles);',
      ],
    }),

    h2('Relationship CRUD'),
    ...member({
      name: 'addRelationship',
      kind: 'method',
      sigs: ['async addRelationship(id: string, type: string, target: string, targetMode?: string): Promise<OpenXmlRelationship>'],
      params: [
        ['id', 'string', 'The relationship ID.'],
        ['type', 'string', 'The relationship type URI.'],
        ['target', 'string', 'The target URI.'],
        ['targetMode', 'string  (optional)', '"Internal" (default) or "External".'],
      ],
      returns: 'Promise<OpenXmlRelationship> \u2014 the newly created relationship.',
      description: 'Adds a new relationship to this part\'s .rels file. Delegates to the package\'s addRelationshipForPart() method.',
      example: [
        'const rel = await mainPart!.addRelationship(',
        '  "rId20", RelationshipType.image, "media/image1.png");',
      ],
    }),
    ...member({
      name: 'deleteRelationship',
      kind: 'method',
      sigs: ['async deleteRelationship(id: string): Promise<boolean>'],
      params: [['id', 'string', 'The relationship ID to delete.']],
      returns: 'Promise<boolean> \u2014 true if deleted.',
      description: 'Deletes a relationship from this part\'s .rels file by ID.',
      example: [
        'await mainPart!.deleteRelationship("rId20");',
      ],
    }),

    h2('Lookup by ID and Content Type'),
    ...member({
      name: 'getRelationshipById',
      kind: 'method',
      sigs: ['async getRelationshipById(rId: string): Promise<OpenXmlRelationship | undefined>'],
      params: [['rId', 'string', 'The relationship ID.']],
      returns: 'Promise<OpenXmlRelationship | undefined>',
      description: 'Finds a relationship of this part by its ID.',
    }),
    ...member({
      name: 'getPartById',
      kind: 'method',
      sigs: ['async getPartById(rId: string): Promise<OpenXmlPart | undefined>'],
      params: [['rId', 'string', 'The relationship ID.']],
      returns: 'Promise<OpenXmlPart | undefined>',
      description: 'Returns the part targeted by the relationship with the given ID.',
    }),
    ...member({
      name: 'getRelationshipsByContentType',
      kind: 'method',
      sigs: ['async getRelationshipsByContentType(contentType: string): Promise<OpenXmlRelationship[]>'],
      params: [['contentType', 'string', 'The MIME content type.']],
      returns: 'Promise<OpenXmlRelationship[]>',
      description: 'Returns all relationships of this part whose target parts have the specified content type. External relationships are excluded.',
    }),
    ...member({
      name: 'getPartsByContentType',
      kind: 'method',
      sigs: ['async getPartsByContentType(contentType: string): Promise<OpenXmlPart[]>'],
      params: [['contentType', 'string', 'The MIME content type.']],
      returns: 'Promise<OpenXmlPart[]>',
      description: 'Returns all parts targeted by relationships of this part whose content type matches.',
    }),

    h2('Convenience Part Accessors'),
    ...member({
      name: 'customXmlPropertiesPart',
      kind: 'method',
      sigs: ['async customXmlPropertiesPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined>',
      description: 'Returns the custom XML properties part related to this part, if one exists.',
    }),
    ...member({
      name: 'themePart',
      kind: 'method',
      sigs: ['async themePart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined>',
      description: 'Returns the theme part related to this part, if one exists.',
    }),
    ...member({
      name: 'thumbnailPart',
      kind: 'method',
      sigs: ['async thumbnailPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined>',
      description: 'Returns the thumbnail part related to this part, if one exists.',
    }),
    ...member({
      name: 'drawingsPart',
      kind: 'method',
      sigs: ['async drawingsPart(): Promise<OpenXmlPart | undefined>'],
      returns: 'Promise<OpenXmlPart | undefined>',
      description: 'Returns the drawings part related to this part, if one exists.',
    }),
    ...member({
      name: 'imageParts',
      kind: 'method',
      sigs: ['async imageParts(): Promise<OpenXmlPart[]>'],
      returns: 'Promise<OpenXmlPart[]> \u2014 all image parts related to this part.',
      description: 'Returns all image parts related to this part.',
    }),
    ...member({
      name: 'customXmlParts',
      kind: 'method',
      sigs: ['async customXmlParts(): Promise<OpenXmlPart[]>'],
      returns: 'Promise<OpenXmlPart[]> \u2014 all custom XML parts related to this part.',
      description: 'Returns all custom XML parts related to this part.',
    }),
  ];
  await writeDoc('OpenXmlPart', content);
}

// ─── OpenXmlRelationship ────────────────────────────────────────────────────
async function genOpenXmlRelationship() {
  const content = [
    h1('OpenXmlRelationship'),
    plain('OpenXmlRelationship represents a single relationship within an Open XML package. Relationships connect packages or parts to target resources (other parts, external URIs, etc.) and are defined in .rels files within the package.'),
    plain('Each relationship has a unique ID, a type URI that identifies the kind of relationship, a target URI, and an optional target mode indicating whether the target is internal or external to the package.'),

    h2('Constructor'),
    ...member({
      name: 'constructor',
      kind: 'constructor',
      sigs: ['constructor(pkg: OpenXmlPackage, part: OpenXmlPart | null, id: string, type: string, target: string, targetMode: string | null)'],
      params: [
        ['pkg', 'OpenXmlPackage', 'The package containing this relationship.'],
        ['part', 'OpenXmlPart | null', 'The source part, or null for package-level relationships.'],
        ['id', 'string', 'The unique relationship ID (e.g., "rId1").'],
        ['type', 'string', 'The relationship type URI.'],
        ['target', 'string', 'The target URI (relative or absolute).'],
        ['targetMode', 'string | null', '"External" for external targets, or null for internal.'],
      ],
      description: 'Creates a new OpenXmlRelationship. Typically created internally when parsing .rels files or when calling addRelationship().',
    }),

    h2('Methods'),
    ...member({
      name: 'getPkg',
      kind: 'method',
      sigs: ['getPkg(): OpenXmlPackage'],
      returns: 'OpenXmlPackage \u2014 the package containing this relationship.',
      description: 'Returns the package that contains this relationship.',
    }),
    ...member({
      name: 'getPart',
      kind: 'method',
      sigs: ['getPart(): OpenXmlPart | null'],
      returns: 'OpenXmlPart | null \u2014 the source part, or null for package-level relationships.',
      description: 'Returns the part that owns this relationship. For package-level relationships (those defined in /_rels/.rels), this returns null.',
      example: [
        'const rel = (await mainPart!.getRelationships())[0];',
        'console.log(rel.getPart()?.getUri()); // "/word/document.xml"',
      ],
    }),
    ...member({
      name: 'getId',
      kind: 'method',
      sigs: ['getId(): string'],
      returns: 'string \u2014 the relationship ID (e.g., "rId1").',
      description: 'Returns the unique identifier of this relationship within its .rels file.',
      example: [
        'console.log(rel.getId()); // "rId1"',
      ],
    }),
    ...member({
      name: 'getType',
      kind: 'method',
      sigs: ['getType(): string'],
      returns: 'string \u2014 the relationship type URI.',
      description: 'Returns the relationship type URI that identifies the kind of relationship. Compare with RelationshipType constants.',
      example: [
        'if (rel.getType() === RelationshipType.styles) {',
        '  console.log("This is a styles relationship");',
        '}',
      ],
    }),
    ...member({
      name: 'getTarget',
      kind: 'method',
      sigs: ['getTarget(): string'],
      returns: 'string \u2014 the target URI as specified in the .rels file.',
      description: 'Returns the raw target URI from the relationship element. This may be a relative path (e.g., "styles.xml") or an absolute path.',
      example: [
        'console.log(rel.getTarget()); // "styles.xml"',
      ],
    }),
    ...member({
      name: 'getTargetMode',
      kind: 'method',
      sigs: ['getTargetMode(): string | null'],
      returns: 'string | null \u2014 "External" for external targets, or null for internal.',
      description: 'Returns the target mode. External relationships point to resources outside the package (such as URLs). Internal relationships (null) point to parts within the package.',
      example: [
        'if (rel.getTargetMode() === "External") {',
        '  console.log("External link:", rel.getTarget());',
        '}',
      ],
    }),
    ...member({
      name: 'getTargetFullName',
      kind: 'method',
      sigs: ['getTargetFullName(): string'],
      returns: 'string \u2014 the fully resolved target URI.',
      description: [
        'Returns the fully resolved target URI. For external relationships, returns the target as-is. For internal relationships, resolves the target relative to the source part\'s directory.',
        'For package-level relationships (no source part), prepends "/" to make it an absolute path. For part-level relationships, prepends the source part\'s directory path.',
      ],
      semantics: 'Use this method to look up the target part via getPartByUri().',
      example: [
        '// Part-level: source is "/word/document.xml", target is "styles.xml"',
        'rel.getTargetFullName(); // "/word/styles.xml"',
        '',
        '// Package-level: target is "word/document.xml"',
        'rel.getTargetFullName(); // "/word/document.xml"',
      ],
    }),
  ];
  await writeDoc('OpenXmlRelationship', content);
}

// ─── WmlPackage ─────────────────────────────────────────────────────────────
async function genWmlPackage() {
  const content = [
    h1('WmlPackage'),
    inheritanceLine('OpenXmlPackage \u2192 WmlPackage'),
    plain('WmlPackage is the package class for Word (.docx) documents. It extends OpenXmlPackage with convenience methods for accessing Word-specific parts such as the main document part and content parts.'),
    plain('All methods inherited from OpenXmlPackage return WmlPart instances (instead of OpenXmlPart) when navigating parts within a Word document.'),

    h2('Opening Documents'),
    ...member({
      name: 'open',
      kind: 'static method',
      sigs: ['static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<WmlPackage>'],
      params: [['document', 'Base64String | FlatOpcString | DocxBinary', 'The Word document to open. Format is auto-detected.']],
      returns: 'Promise<WmlPackage> \u2014 the opened Word document package.',
      description: 'Opens a Word document from any of the three supported formats. Returns a WmlPackage with Word-specific part navigation methods.',
      example: [
        'import { WmlPackage } from "openxmlsdkts";',
        'import fs from "fs";',
        '',
        'const buffer = fs.readFileSync("report.docx");',
        'const doc = await WmlPackage.open(new Blob([buffer]));',
      ],
    }),

    h2('Word-Specific Parts'),
    ...member({
      name: 'mainDocumentPart',
      kind: 'method',
      sigs: ['async mainDocumentPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the main document part (document.xml), or undefined.',
      description: 'Returns the main document part of the Word document. This is the primary XML part containing the document body, and is the starting point for most document manipulation.',
      example: [
        'const mainPart = await doc.mainDocumentPart();',
        'const xDoc = await mainPart!.getXDocument();',
        'const body = xDoc.root!.element(W.body);',
      ],
    }),
    ...member({
      name: 'contentParts',
      kind: 'method',
      sigs: ['async contentParts(): Promise<WmlPart[]>'],
      returns: 'Promise<WmlPart[]> \u2014 the main document part plus all header, footer, endnotes, and footnotes parts.',
      description: 'Returns an array of all content-bearing parts in the Word document: the main document part, followed by all header parts, footer parts, and the endnotes and footnotes parts (if they exist). This is useful for operations that need to process all user-visible content (e.g., search-and-replace across the entire document).',
      semantics: 'Returns an empty array if the document has no main document part.',
      example: [
        'const contentParts = await doc.contentParts();',
        'for (const part of contentParts) {',
        '  const xDoc = await part.getXDocument();',
        '  // Process paragraphs in each content part',
        '}',
      ],
    }),
  ];
  await writeDoc('WmlPackage', content);
}

// ─── WmlPart ────────────────────────────────────────────────────────────────
async function genWmlPart() {
  const content = [
    h1('WmlPart'),
    inheritanceLine('OpenXmlPart \u2192 WmlPart'),
    plain('WmlPart extends OpenXmlPart with convenience methods for navigating Word-specific sub-parts. These methods look up related parts by their standard relationship types, returning strongly-typed WmlPart instances.'),
    plain('All methods inherited from OpenXmlPart are also available on WmlPart.'),

    ...member({
      name: 'headerParts',
      kind: 'method',
      sigs: ['async headerParts(): Promise<WmlPart[]>'],
      returns: 'Promise<WmlPart[]> \u2014 all header parts related to this part.',
      description: 'Returns all header parts (header1.xml, header2.xml, etc.) related to this part. Typically called on the main document part.',
      example: [
        'const mainPart = await doc.mainDocumentPart();',
        'const headers = await mainPart!.headerParts();',
        'for (const hdr of headers) {',
        '  const xDoc = await hdr.getXDocument();',
        '  console.log(hdr.getUri());',
        '}',
      ],
    }),
    ...member({
      name: 'footerParts',
      kind: 'method',
      sigs: ['async footerParts(): Promise<WmlPart[]>'],
      returns: 'Promise<WmlPart[]> \u2014 all footer parts related to this part.',
      description: 'Returns all footer parts (footer1.xml, footer2.xml, etc.) related to this part. Typically called on the main document part.',
      example: [
        'const footers = await mainPart!.footerParts();',
      ],
    }),
    ...member({
      name: 'endnotesPart',
      kind: 'method',
      sigs: ['async endnotesPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the endnotes part, or undefined.',
      description: 'Returns the endnotes part (endnotes.xml) related to this part, if it exists.',
      example: [
        'const endnotes = await mainPart!.endnotesPart();',
      ],
    }),
    ...member({
      name: 'footnotesPart',
      kind: 'method',
      sigs: ['async footnotesPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the footnotes part, or undefined.',
      description: 'Returns the footnotes part (footnotes.xml) related to this part, if it exists.',
      example: [
        'const footnotes = await mainPart!.footnotesPart();',
      ],
    }),
    ...member({
      name: 'wordprocessingCommentsPart',
      kind: 'method',
      sigs: ['async wordprocessingCommentsPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the comments part, or undefined.',
      description: 'Returns the comments part (comments.xml) related to this part, containing all Word comments/annotations.',
      example: [
        'const comments = await mainPart!.wordprocessingCommentsPart();',
        'if (comments) {',
        '  const xDoc = await comments.getXDocument();',
        '  const commentEls = xDoc.root!.elements(W.comment);',
        '  console.log(`${commentEls.length} comments`);',
        '}',
      ],
    }),
    ...member({
      name: 'fontTablePart',
      kind: 'method',
      sigs: ['async fontTablePart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the font table part, or undefined.',
      description: 'Returns the font table part (fontTable.xml) that lists fonts used in the document.',
    }),
    ...member({
      name: 'numberingDefinitionsPart',
      kind: 'method',
      sigs: ['async numberingDefinitionsPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the numbering definitions part, or undefined.',
      description: 'Returns the numbering definitions part (numbering.xml) that defines list numbering styles.',
    }),
    ...member({
      name: 'styleDefinitionsPart',
      kind: 'method',
      sigs: ['async styleDefinitionsPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the styles part, or undefined.',
      description: 'Returns the style definitions part (styles.xml) containing paragraph and character style definitions.',
      example: [
        'const stylesPart = await mainPart!.styleDefinitionsPart();',
        'if (stylesPart) {',
        '  const xDoc = await stylesPart.getXDocument();',
        '  const styles = xDoc.root!.elements(W.style);',
        '  console.log(`${styles.length} styles defined`);',
        '}',
      ],
    }),
    ...member({
      name: 'webSettingsPart',
      kind: 'method',
      sigs: ['async webSettingsPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the web settings part, or undefined.',
      description: 'Returns the web settings part (webSettings.xml) containing settings for when the document is saved as HTML.',
    }),
    ...member({
      name: 'documentSettingsPart',
      kind: 'method',
      sigs: ['async documentSettingsPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the document settings part, or undefined.',
      description: 'Returns the document settings part (settings.xml) containing document-level settings such as compatibility options and revision tracking.',
    }),
    ...member({
      name: 'glossaryDocumentPart',
      kind: 'method',
      sigs: ['async glossaryDocumentPart(): Promise<WmlPart | undefined>'],
      returns: 'Promise<WmlPart | undefined> \u2014 the glossary document part, or undefined.',
      description: 'Returns the glossary document part, which contains building blocks (AutoText, Quick Parts) defined for the document.',
    }),
    ...member({
      name: 'fontParts',
      kind: 'method',
      sigs: ['async fontParts(): Promise<WmlPart[]>'],
      returns: 'Promise<WmlPart[]> \u2014 all embedded font parts.',
      description: 'Returns all embedded font data parts related to this part.',
    }),
  ];
  await writeDoc('WmlPart', content);
}

// ─── SmlPackage ─────────────────────────────────────────────────────────────
async function genSmlPackage() {
  const content = [
    h1('SmlPackage'),
    inheritanceLine('OpenXmlPackage \u2192 SmlPackage'),
    plain('SmlPackage is the package class for Excel (.xlsx) documents. It extends OpenXmlPackage with convenience methods for accessing spreadsheet-specific parts.'),
    plain('All methods inherited from OpenXmlPackage return SmlPart instances when navigating parts within an Excel document.'),

    h2('Opening Documents'),
    ...member({
      name: 'open',
      kind: 'static method',
      sigs: ['static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<SmlPackage>'],
      params: [['document', 'Base64String | FlatOpcString | DocxBinary', 'The Excel document to open. Format is auto-detected.']],
      returns: 'Promise<SmlPackage> \u2014 the opened Excel document package.',
      description: 'Opens an Excel spreadsheet from any of the three supported formats.',
      example: [
        'import { SmlPackage } from "openxmlsdkts";',
        'import fs from "fs";',
        '',
        'const buffer = fs.readFileSync("data.xlsx");',
        'const doc = await SmlPackage.open(new Blob([buffer]));',
      ],
    }),

    h2('Spreadsheet-Specific Parts'),
    ...member({
      name: 'workbookPart',
      kind: 'method',
      sigs: ['async workbookPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the workbook part, or undefined.',
      description: 'Returns the main workbook part (workbook.xml) of the Excel document. This is the starting point for navigating worksheets, shared strings, styles, and other spreadsheet components.',
      example: [
        'const workbook = await doc.workbookPart();',
        'const worksheets = await workbook!.worksheetParts();',
        'console.log(`${worksheets.length} worksheets`);',
      ],
    }),
  ];
  await writeDoc('SmlPackage', content);
}

// ─── SmlPart ────────────────────────────────────────────────────────────────
async function genSmlPart() {
  const content = [
    h1('SmlPart'),
    inheritanceLine('OpenXmlPart \u2192 SmlPart'),
    plain('SmlPart extends OpenXmlPart with convenience methods for navigating Excel-specific sub-parts. These methods look up related parts by their standard relationship types, returning strongly-typed SmlPart instances.'),
    plain('All methods inherited from OpenXmlPart are also available on SmlPart.'),

    ...member({
      name: 'calculationChainPart',
      kind: 'method',
      sigs: ['async calculationChainPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the calculation chain part, or undefined.',
      description: 'Returns the calculation chain part (calcChain.xml) that defines the order in which cells are calculated.',
    }),
    ...member({
      name: 'cellMetadataPart',
      kind: 'method',
      sigs: ['async cellMetadataPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the cell metadata part, or undefined.',
      description: 'Returns the cell metadata part containing additional cell-level metadata.',
    }),
    ...member({
      name: 'sharedStringTablePart',
      kind: 'method',
      sigs: ['async sharedStringTablePart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the shared string table part, or undefined.',
      description: 'Returns the shared string table part (sharedStrings.xml). Excel stores unique string values in this table and references them by index from worksheet cells, which reduces file size when the same text appears in multiple cells.',
      example: [
        'const sst = await workbook!.sharedStringTablePart();',
        'if (sst) {',
        '  const xDoc = await sst.getXDocument();',
        '  const strings = xDoc.root!.elements(S.si);',
        '  console.log(`${strings.length} unique strings`);',
        '}',
      ],
    }),
    ...member({
      name: 'workbookRevisionHeaderPart',
      kind: 'method',
      sigs: ['async workbookRevisionHeaderPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the revision header part, or undefined.',
      description: 'Returns the workbook revision header part, which tracks revision (change tracking) metadata.',
    }),
    ...member({
      name: 'workbookStylesPart',
      kind: 'method',
      sigs: ['async workbookStylesPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the workbook styles part, or undefined.',
      description: 'Returns the workbook styles part (styles.xml) containing cell formatting styles, number formats, fonts, fills, and borders.',
      example: [
        'const stylesPart = await workbook!.workbookStylesPart();',
      ],
    }),
    ...member({
      name: 'worksheetCommentsPart',
      kind: 'method',
      sigs: ['async worksheetCommentsPart(): Promise<SmlPart | undefined>'],
      returns: 'Promise<SmlPart | undefined> \u2014 the worksheet comments part, or undefined.',
      description: 'Returns the worksheet comments part containing cell comments/notes.',
    }),
    ...member({
      name: 'chartsheetParts',
      kind: 'method',
      sigs: ['async chartsheetParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all chartsheet parts.',
      description: 'Returns all chartsheet parts. Chartsheets are sheets that contain only a chart.',
    }),
    ...member({
      name: 'worksheetParts',
      kind: 'method',
      sigs: ['async worksheetParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all worksheet parts.',
      description: 'Returns all worksheet parts (sheet1.xml, sheet2.xml, etc.) in the workbook.',
      example: [
        'const worksheets = await workbook!.worksheetParts();',
        'for (const ws of worksheets) {',
        '  const xDoc = await ws.getXDocument();',
        '  const rows = xDoc.root!.element(S.sheetData)?.elements(S.row) ?? [];',
        '  console.log(`${ws.getUri()}: ${rows.length} rows`);',
        '}',
      ],
    }),
    ...member({
      name: 'pivotTableParts',
      kind: 'method',
      sigs: ['async pivotTableParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all pivot table parts.',
      description: 'Returns all pivot table definition parts.',
    }),
    ...member({
      name: 'queryTableParts',
      kind: 'method',
      sigs: ['async queryTableParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all query table parts.',
      description: 'Returns all query table parts, which define external data queries.',
    }),
    ...member({
      name: 'tableDefinitionParts',
      kind: 'method',
      sigs: ['async tableDefinitionParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all table definition parts.',
      description: 'Returns all table definition parts, which define structured table ranges within worksheets.',
    }),
    ...member({
      name: 'timeLineParts',
      kind: 'method',
      sigs: ['async timeLineParts(): Promise<SmlPart[]>'],
      returns: 'Promise<SmlPart[]> \u2014 all timeline parts.',
      description: 'Returns all timeline parts, which define timeline controls used for filtering pivot table data by date ranges.',
    }),
  ];
  await writeDoc('SmlPart', content);
}

// ─── PmlPackage ─────────────────────────────────────────────────────────────
async function genPmlPackage() {
  const content = [
    h1('PmlPackage'),
    inheritanceLine('OpenXmlPackage \u2192 PmlPackage'),
    plain('PmlPackage is the package class for PowerPoint (.pptx) documents. It extends OpenXmlPackage with convenience methods for accessing presentation-specific parts.'),
    plain('All methods inherited from OpenXmlPackage return PmlPart instances when navigating parts within a PowerPoint document.'),

    h2('Opening Documents'),
    ...member({
      name: 'open',
      kind: 'static method',
      sigs: ['static async open(document: Base64String | FlatOpcString | DocxBinary): Promise<PmlPackage>'],
      params: [['document', 'Base64String | FlatOpcString | DocxBinary', 'The PowerPoint document to open. Format is auto-detected.']],
      returns: 'Promise<PmlPackage> \u2014 the opened PowerPoint document package.',
      description: 'Opens a PowerPoint presentation from any of the three supported formats.',
      example: [
        'import { PmlPackage } from "openxmlsdkts";',
        'import fs from "fs";',
        '',
        'const buffer = fs.readFileSync("deck.pptx");',
        'const doc = await PmlPackage.open(new Blob([buffer]));',
      ],
    }),

    h2('Presentation-Specific Parts'),
    ...member({
      name: 'presentationPart',
      kind: 'method',
      sigs: ['async presentationPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the presentation part, or undefined.',
      description: 'Returns the main presentation part (presentation.xml) of the PowerPoint document. This is the starting point for navigating slides, slide masters, and other presentation components.',
      example: [
        'const presentation = await doc.presentationPart();',
        'const slides = await presentation!.slideParts();',
        'console.log(`${slides.length} slides`);',
      ],
    }),
  ];
  await writeDoc('PmlPackage', content);
}

// ─── PmlPart ────────────────────────────────────────────────────────────────
async function genPmlPart() {
  const content = [
    h1('PmlPart'),
    inheritanceLine('OpenXmlPart \u2192 PmlPart'),
    plain('PmlPart extends OpenXmlPart with convenience methods for navigating PowerPoint-specific sub-parts. These methods look up related parts by their standard relationship types, returning strongly-typed PmlPart instances.'),
    plain('All methods inherited from OpenXmlPart are also available on PmlPart.'),

    ...member({
      name: 'commentAuthorsPart',
      kind: 'method',
      sigs: ['async commentAuthorsPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the comment authors part, or undefined.',
      description: 'Returns the comment authors part, which maps author IDs to author names for presentation comments.',
    }),
    ...member({
      name: 'handoutMasterPart',
      kind: 'method',
      sigs: ['async handoutMasterPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the handout master part, or undefined.',
      description: 'Returns the handout master part, which defines the layout for printed handouts.',
    }),
    ...member({
      name: 'notesMasterPart',
      kind: 'method',
      sigs: ['async notesMasterPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the notes master part, or undefined.',
      description: 'Returns the notes master part, which defines the default layout for speaker notes.',
    }),
    ...member({
      name: 'notesSlidePart',
      kind: 'method',
      sigs: ['async notesSlidePart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the notes slide part, or undefined.',
      description: 'Returns the notes slide part for this slide, containing the speaker notes content.',
    }),
    ...member({
      name: 'presentationPropertiesPart',
      kind: 'method',
      sigs: ['async presentationPropertiesPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the presentation properties part, or undefined.',
      description: 'Returns the presentation properties part (presProps.xml) containing slide show settings, print settings, and color scheme information.',
    }),
    ...member({
      name: 'tableStylesPart',
      kind: 'method',
      sigs: ['async tableStylesPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the table styles part, or undefined.',
      description: 'Returns the table styles part defining available table formatting styles in the presentation.',
    }),
    ...member({
      name: 'userDefinedTagsPart',
      kind: 'method',
      sigs: ['async userDefinedTagsPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the user-defined tags part, or undefined.',
      description: 'Returns the user-defined tags part containing custom name/value pairs attached to the presentation or individual slides.',
    }),
    ...member({
      name: 'viewPropertiesPart',
      kind: 'method',
      sigs: ['async viewPropertiesPart(): Promise<PmlPart | undefined>'],
      returns: 'Promise<PmlPart | undefined> \u2014 the view properties part, or undefined.',
      description: 'Returns the view properties part (viewProps.xml) containing the last-used view settings (Normal, Slide Sorter, etc.).',
    }),
    ...member({
      name: 'slideMasterParts',
      kind: 'method',
      sigs: ['async slideMasterParts(): Promise<PmlPart[]>'],
      returns: 'Promise<PmlPart[]> \u2014 all slide master parts.',
      description: 'Returns all slide master parts. Slide masters define the common layout, formatting, and theme for a group of slides.',
      example: [
        'const masters = await presentation!.slideMasterParts();',
        'console.log(`${masters.length} slide masters`);',
      ],
    }),
    ...member({
      name: 'slideParts',
      kind: 'method',
      sigs: ['async slideParts(): Promise<PmlPart[]>'],
      returns: 'Promise<PmlPart[]> \u2014 all slide parts.',
      description: 'Returns all slide parts (slide1.xml, slide2.xml, etc.) in the presentation.',
      example: [
        'const slides = await presentation!.slideParts();',
        'for (const slide of slides) {',
        '  const xDoc = await slide.getXDocument();',
        '  console.log(slide.getUri());',
        '}',
      ],
    }),
  ];
  await writeDoc('PmlPart', content);
}

// ─── ContentType ────────────────────────────────────────────────────────────
async function genContentType() {
  const contentTypes = [
    ['calculationChain', 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'],
    ['cellMetadata', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'],
    ['chart', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
    ['chartColorStyle', 'application/vnd.ms-office.chartcolorstyle+xml'],
    ['chartDrawing', 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml'],
    ['chartsheet', 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'],
    ['chartStyle', 'application/vnd.ms-office.chartstyle+xml'],
    ['commentAuthors', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'],
    ['connections', 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'],
    ['coreFileProperties', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['customFileProperties', 'application/vnd.openxmlformats-officedocument.custom-properties+xml'],
    ['customization', 'application/vnd.ms-word.keyMapCustomizations+xml'],
    ['customProperty', 'application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty'],
    ['customXmlProperties', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'],
    ['diagramColors', 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml'],
    ['diagramData', 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml'],
    ['diagramLayoutDefinition', 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml'],
    ['diagramPersistLayout', 'application/vnd.ms-office.drawingml.diagramDrawing+xml'],
    ['diagramStyle', 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml'],
    ['dialogsheet', 'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml'],
    ['digitalSignatureOrigin', 'application/vnd.openxmlformats-package.digital-signature-origin'],
    ['documentSettings', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'],
    ['drawings', 'application/vnd.openxmlformats-officedocument.drawing+xml'],
    ['endnotes', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'],
    ['excelAttachedToolbars', 'application/vnd.ms-excel.attachedToolbars'],
    ['extendedFileProperties', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
    ['externalWorkbook', 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'],
    ['fontData', 'application/x-fontdata'],
    ['fontTable', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'],
    ['footer', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'],
    ['footnotes', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'],
    ['gif', 'image/gif'],
    ['glossaryDocument', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml'],
    ['handoutMaster', 'application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml'],
    ['header', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'],
    ['jpeg', 'image/jpeg'],
    ['mainDocument', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'],
    ['notesMaster', 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'],
    ['notesSlide', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'],
    ['numberingDefinitions', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'],
    ['pict', 'image/pict'],
    ['pivotTable', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'],
    ['pivotTableCacheDefinition', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'],
    ['pivotTableCacheRecords', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'],
    ['png', 'image/png'],
    ['presentation', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
    ['presentationProperties', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'],
    ['presentationTemplate', 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'],
    ['queryTable', 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'],
    ['relationships', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['ribbonAndBackstageCustomizations', 'http://schemas.microsoft.com/office/2009/07/customui'],
    ['sharedStringTable', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
    ['singleCellTable', 'application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml'],
    ['slicerCache', 'application/vnd.openxmlformats-officedocument.spreadsheetml.slicerCache+xml'],
    ['slicers', 'application/vnd.openxmlformats-officedocument.spreadsheetml.slicer+xml'],
    ['slide', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'],
    ['slideComments', 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'],
    ['slideLayout', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
    ['slideMaster', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
    ['slideShow', 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml'],
    ['slideSyncData', 'application/vnd.openxmlformats-officedocument.presentationml.slideUpdateInfo+xml'],
    ['styles', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'],
    ['tableDefinition', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'],
    ['tableStyles', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'],
    ['theme', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ['themeOverride', 'application/vnd.openxmlformats-officedocument.themeOverride+xml'],
    ['tiff', 'image/tiff'],
    ['trueTypeFont', 'application/x-font-ttf'],
    ['userDefinedTags', 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml'],
    ['viewProperties', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'],
    ['vmlDrawing', 'application/vnd.openxmlformats-officedocument.vmlDrawing'],
    ['volatileDependencies', 'application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml'],
    ['webSettings', 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'],
    ['wordAttachedToolbars', 'application/vnd.ms-word.attachedToolbars'],
    ['wordprocessingComments', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'],
    ['wordprocessingTemplate', 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'],
    ['workbook', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
    ['workbookRevisionHeader', 'application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml'],
    ['workbookRevisionLog', 'application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml'],
    ['workbookStyles', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    ['workbookTemplate', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'],
    ['workbookUserData', 'application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml'],
    ['worksheet', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],
    ['worksheetComments', 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'],
    ['worksheetSortMap', 'application/vnd.ms-excel.wsSortMap+xml'],
    ['xmlSignature', 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml'],
  ];

  const content = [
    h1('ContentType'),
    plain('ContentType is a static object with 97 readonly properties that map human-readable content type labels to their full MIME type URI strings. Use these constants instead of hard-coding MIME type URIs.'),
    plain('ContentType values are used when adding parts to a package (via addPart()), and when filtering parts or relationships by content type (via getPartsByContentType(), getRelationshipsByContentType(), etc.).'),

    h2('Types'),
    ...member({
      name: 'ContentTypeKey',
      kind: 'static property',
      sigs: ['type ContentTypeKey = keyof typeof ContentType'],
      description: 'A union type of all property names on the ContentType object (e.g., "mainDocument" | "workbook" | "presentation" | ...).',
    }),
    ...member({
      name: 'ContentTypeValue',
      kind: 'static property',
      sigs: ['type ContentTypeValue = (typeof ContentType)[ContentTypeKey]'],
      description: 'A union type of all MIME type URI string values in the ContentType object.',
    }),

    h2('Usage Example'),
    ...codeLines([
      'import { ContentType } from "openxmlsdkts";',
      '',
      '// Use as a constant instead of a raw string',
      'const parts = await pkg.getPartsByContentType(ContentType.mainDocument);',
      '',
      '// Add a part with a typed content type',
      'pkg.addPart("/word/comments.xml",',
      '  ContentType.wordprocessingComments, "xml", xDoc);',
    ]),

    h2('Content Type Reference'),
    plain('The following table lists all 97 content type properties and their MIME type values.'),
    twoColTable(['Property', 'MIME Type'], contentTypes),
  ];
  await writeDoc('ContentType', content);
}

// ─── RelationshipType ───────────────────────────────────────────────────────
async function genRelationshipType() {
  const relTypes = [
    ['alternativeFormatImport', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk'],
    ['calculationChain', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'],
    ['cellMetadata', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'],
    ['chart', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'],
    ['chartColorStyle', 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'],
    ['chartDrawing', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes'],
    ['chartsheet', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet'],
    ['chartStyle', 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'],
    ['commentAuthors', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'],
    ['connections', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'],
    ['coreFileProperties', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'],
    ['customFileProperties', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'],
    ['customization', 'http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations'],
    ['customProperty', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty'],
    ['customXml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'],
    ['customXmlMappings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps'],
    ['customXmlProperties', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'],
    ['diagramColors', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors'],
    ['diagramData', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData'],
    ['diagramLayoutDefinition', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout'],
    ['diagramPersistLayout', 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'],
    ['diagramStyle', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle'],
    ['dialogsheet', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet'],
    ['digitalSignatureOrigin', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'],
    ['documentSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'],
    ['drawings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'],
    ['endnotes', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes'],
    ['excelAttachedToolbars', 'http://schemas.microsoft.com/office/2006/relationships/attachedToolbars'],
    ['extendedFileProperties', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'],
    ['externalWorkbook', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'],
    ['font', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'],
    ['fontTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable'],
    ['footer', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'],
    ['footnotes', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'],
    ['glossaryDocument', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument'],
    ['handoutMaster', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster'],
    ['header', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'],
    ['image', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'],
    ['mainDocument', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'],
    ['notesMaster', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster'],
    ['notesSlide', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'],
    ['numberingDefinitions', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'],
    ['pivotTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'],
    ['pivotTableCacheDefinition', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'],
    ['pivotTableCacheRecords', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'],
    ['presentation', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'],
    ['presentationProperties', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'],
    ['queryTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'],
    ['ribbonAndBackstageCustomizations', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'],
    ['sharedStringTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'],
    ['singleCellTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells'],
    ['slide', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'],
    ['slideComments', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'],
    ['slideLayout', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'],
    ['slideMaster', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'],
    ['slideSyncData', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo'],
    ['styles', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'],
    ['tableDefinition', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'],
    ['tableStyles', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'],
    ['timeLine', 'http://schemas.microsoft.com/office/2011/relationships/timeline'],
    ['theme', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'],
    ['themeOverride', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride'],
    ['thumbnail', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'],
    ['userDefinedTags', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'],
    ['viewProperties', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps'],
    ['vmlDrawing', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'],
    ['volatileDependencies', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies'],
    ['webSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings'],
    ['wordAttachedToolbars', 'http://schemas.microsoft.com/office/2006/relationships/attachedToolbars'],
    ['wordprocessingComments', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'],
    ['workbook', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'],
    ['workbookRevisionHeader', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders'],
    ['workbookRevisionLog', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog'],
    ['workbookStyles', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'],
    ['workbookUserData', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames'],
    ['worksheet', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'],
    ['worksheetComments', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'],
    ['worksheetSortMap', 'http://schemas.microsoft.com/office/2006/relationships/wsSortMap'],
    ['xmlSignature', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature'],
  ];

  const content = [
    h1('RelationshipType'),
    plain('RelationshipType is a static object with 89 readonly properties that map human-readable relationship type labels to their full URI strings. Use these constants instead of hard-coding relationship type URIs.'),
    plain('RelationshipType values are used when querying or adding relationships (via getRelationshipsByRelationshipType(), getPartsByRelationshipType(), addRelationship(), etc.).'),

    h2('Types'),
    ...member({
      name: 'RelationshipTypeKey',
      kind: 'static property',
      sigs: ['type RelationshipTypeKey = keyof typeof RelationshipType'],
      description: 'A union type of all property names on the RelationshipType object.',
    }),
    ...member({
      name: 'RelationshipTypeValue',
      kind: 'static property',
      sigs: ['type RelationshipTypeValue = (typeof RelationshipType)[RelationshipTypeKey]'],
      description: 'A union type of all relationship type URI values.',
    }),

    h2('Usage Example'),
    ...codeLines([
      'import { RelationshipType } from "openxmlsdkts";',
      '',
      '// Query parts by relationship type',
      'const stylePart = await mainPart!.getPartByRelationshipType(',
      '  RelationshipType.styles);',
      '',
      '// Add a relationship using a constant',
      'await mainPart!.addRelationship(',
      '  "rId20", RelationshipType.image, "media/image1.png");',
    ]),

    h2('Relationship Type Reference'),
    plain('The following table lists all 89 relationship type properties and their URI values.'),
    twoColTable(['Property', 'URI'], relTypes),
  ];
  await writeDoc('RelationshipType', content);
}

// ─── Utility ────────────────────────────────────────────────────────────────
async function genUtility() {
  const content = [
    h1('Utility'),
    plain('Utility is a static helper class providing methods for common operations related to Open XML package structure. These methods are used internally by the library and can also be called directly.'),

    ...member({
      name: 'getRelsPartUri',
      kind: 'static method',
      sigs: ['static getRelsPartUri(part: OpenXmlPart): string'],
      params: [['part', 'OpenXmlPart', 'The part whose .rels URI to compute.']],
      returns: 'string \u2014 the URI of the .rels file for the given part.',
      description: 'Computes the URI of the .rels file associated with a part. For example, for a part at "/word/document.xml", the .rels URI is "/word/_rels/document.xml.rels".',
      example: [
        'const relsUri = Utility.getRelsPartUri(mainPart);',
        '// "/word/_rels/document.xml.rels"',
      ],
    }),
    ...member({
      name: 'getRelsPart',
      kind: 'static method',
      sigs: ['static getRelsPart(part: OpenXmlPart): OpenXmlPart | undefined'],
      params: [['part', 'OpenXmlPart', 'The part whose .rels part to look up.']],
      returns: 'OpenXmlPart | undefined \u2014 the .rels part, or undefined if it does not exist.',
      description: 'Returns the .rels part associated with the given part by computing its URI and looking it up in the package.',
      example: [
        'const relsPart = Utility.getRelsPart(mainPart);',
        'if (relsPart) {',
        '  const xDoc = await relsPart.getXDocument();',
        '}',
      ],
    }),
    ...member({
      name: 'isBase64',
      kind: 'static method',
      sigs: ['static isBase64(str: unknown): boolean'],
      params: [['str', 'unknown', 'The value to test.']],
      returns: 'boolean \u2014 true if the value appears to be a Base64-encoded string.',
      description: 'Tests whether a value is a Base64-encoded string by checking the first 500 characters. Returns false if the value is not a string or contains characters outside the Base64 alphabet (A-Z, a-z, 0-9, +, /).',
      semantics: 'This is a heuristic check used internally by OpenXmlPackage.open() to distinguish Base64 strings from Flat OPC XML strings. It does not validate the entire string or check for proper padding.',
      example: [
        'Utility.isBase64("UEsDBBQAAAA...");  // true',
        'Utility.isBase64("<?xml ...");         // false',
        'Utility.isBase64(42);                  // false',
      ],
    }),
  ];
  await writeDoc('Utility', content);
}

// ─── Namespaces ─────────────────────────────────────────────────────────────
async function genNamespaces() {
  const nsClasses = [
    ['A', 'DrawingML main', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
    ['A14', 'DrawingML 2010', 'http://schemas.microsoft.com/office/drawing/2010/main'],
    ['C', 'DrawingML charts', 'http://schemas.openxmlformats.org/drawingml/2006/chart'],
    ['CDR', 'Chart drawing', 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing'],
    ['COM', 'Common Office', 'various'],
    ['CP', 'Core properties', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'],
    ['CT', 'Content types', 'http://schemas.openxmlformats.org/package/2006/content-types'],
    ['CUSTPRO', 'Custom properties', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'],
    ['DC', 'Dublin Core', 'http://purl.org/dc/elements/1.1/'],
    ['DCTERMS', 'Dublin Core terms', 'http://purl.org/dc/terms/'],
    ['DGM', 'Diagrams', 'http://schemas.openxmlformats.org/drawingml/2006/diagram'],
    ['DGM14', 'Diagrams 2010', 'http://schemas.microsoft.com/office/drawing/2010/diagram'],
    ['DIGSIG', 'Digital signatures', 'http://schemas.openxmlformats.org/package/2006/digital-signature'],
    ['DS', 'Data store', 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'],
    ['DSP', 'Data/shape pairs', 'http://schemas.microsoft.com/office/drawing/2008/diagram'],
    ['EP', 'Extended properties', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'],
    ['FLATOPC', 'Flat OPC package', 'http://schemas.microsoft.com/office/2006/xmlPackage'],
    ['LC', 'Literal connections', 'various'],
    ['M', 'Office math', 'http://schemas.openxmlformats.org/officeDocument/2006/math'],
    ['MC', 'Markup compatibility', 'http://schemas.openxmlformats.org/markup-compatibility/2006'],
    ['MDSSI', 'Macro-enabled docs', 'http://schemas.openxmlformats.org/package/2006/digital-signature'],
    ['MP', 'Macro presentation', 'various'],
    ['MV', 'Macro workbook', 'various'],
    ['NoNamespace', 'No namespace attrs', '(empty)'],
    ['O', 'Office VML', 'urn:schemas-microsoft-com:office:office'],
    ['P', 'PresentationML', 'http://schemas.openxmlformats.org/presentationml/2006/main'],
    ['P14', 'PresentationML 2010', 'http://schemas.microsoft.com/office/powerpoint/2010/main'],
    ['P15', 'PresentationML 2012', 'http://schemas.microsoft.com/office/powerpoint/2012/main'],
    ['Pic', 'DrawingML pictures', 'http://schemas.openxmlformats.org/drawingml/2006/picture'],
    ['PKGREL', 'Package rels', 'http://schemas.openxmlformats.org/package/2006/relationships'],
    ['R', 'Relationships', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
    ['S', 'SpreadsheetML', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'],
    ['SL', 'Slicer', 'various'],
    ['SLE', 'Slicer extensions', 'various'],
    ['VML', 'Vector markup', 'urn:schemas-microsoft-com:vml'],
    ['VT', 'VML types', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'],
    ['W', 'WordprocessingML', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
    ['W10', 'Word 2000', 'urn:schemas-microsoft-com:office:word'],
    ['W14', 'Word 2010', 'http://schemas.microsoft.com/office/word/2010/wordml'],
    ['W3DIGSIG', 'W3C digital sigs', 'http://www.w3.org/2000/09/xmldsig#'],
    ['WP', 'WP drawing', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'],
    ['WP14', 'WP drawing 2010', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'],
    ['WPS', 'WP shapes', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'],
    ['X', 'SpreadsheetML ext', 'various'],
    ['XDR', 'Spreadsheet drawing', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'],
    ['XDR14', 'Spreadsheet drw 2010', 'http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing'],
    ['XM', 'XPath markup', 'various'],
    ['XSI', 'XML Schema Instance', 'http://www.w3.org/2001/XMLSchema-instance'],
  ];

  // Build a three-column table
  const headerRow = new TableRow({
    tableHeader: true,
    children: ['Class', 'Description', 'Namespace URI'].map((t, i) =>
      new TableCell({
        shading: { type: ShadingType.CLEAR, fill: BLUE },
        width: { size: [10, 25, 65][i], type: WidthType.PERCENTAGE },
        children: [new Paragraph({
          children: [new TextRun({ text: t, bold: true, color: WHITE, size: 18, font: BODY })],
          alignment: AlignmentType.CENTER,
        })],
      })
    ),
  });
  const dataRows = nsClasses.map(([cls, desc, uri], i) =>
    new TableRow({
      children: [cls, desc, uri].map(cell =>
        new TableCell({
          shading: i % 2 === 1 ? { type: ShadingType.CLEAR, fill: LIGHT_BLUE } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: cell, font: MONO, size: 16, color: DARK })] })],
        })
      ),
    })
  );
  const nsTable = new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });

  const content = [
    h1('Namespace Classes'),
    plain('OpenXmlSdkTs provides 48 static namespace classes that contain pre-atomized XName and XNamespace objects for every element and attribute name in the Open XML specification. These classes are defined in OpenXmlNamespacesAndNames.ts and are re-exported from the main entry point.'),

    h2('Purpose'),
    plain('In Open XML, every element and attribute belongs to a specific XML namespace. Rather than constructing XName objects from strings at runtime, the namespace classes provide pre-initialized static XName properties that enable O(1) identity-based equality checks and improve code readability.'),

    h2('Structure'),
    plain('Each namespace class follows the same pattern:'),
    ...codeLines([
      'class W {',
      '  static readonly namespace: XNamespace = XNamespace.get(',
      '    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
      '  );',
      '  static readonly body: XName = W.namespace.getName("body");',
      '  static readonly p: XName = W.namespace.getName("p");',
      '  static readonly r: XName = W.namespace.getName("r");',
      '  static readonly t: XName = W.namespace.getName("t");',
      '  static readonly pPr: XName = W.namespace.getName("pPr");',
      '  static readonly rPr: XName = W.namespace.getName("rPr");',
      '  // ... hundreds more XName properties',
      '}',
    ]),

    h2('NoNamespace'),
    plain('The NoNamespace class is special: it provides XName objects for attributes that have no namespace prefix (such as val, id, type, name, etc.). These are the most commonly used unqualified attribute names in Open XML.'),
    ...codeLines([
      '// Access an unqualified attribute',
      'const val = element.attribute(NoNamespace.val)?.value;',
      'const id = element.attribute(NoNamespace.id)?.value;',
    ]),

    h2('Usage Examples'),
    label('Navigate Word document XML'),
    ...codeLines([
      'import { W, NoNamespace } from "openxmlsdkts";',
      '',
      '// Find all paragraphs in the document body',
      'const body = xDoc.root!.element(W.body);',
      'const paragraphs = body!.elements(W.p);',
      '',
      '// Get the style of a paragraph',
      'const pPr = paragraphs[0].element(W.pPr);',
      'const styleName = pPr?.element(W.pStyle)?.attribute(NoNamespace.val)?.value;',
      '',
      '// Create a new bold run',
      'const run = new XElement(',
      '  W.r,',
      '  new XElement(W.rPr, new XElement(W.b)),',
      '  new XElement(W.t, "Bold text")',
      ');',
    ]),
    gap(),
    label('Navigate Excel spreadsheet XML'),
    ...codeLines([
      'import { S, NoNamespace } from "openxmlsdkts";',
      '',
      '// Get all rows from a worksheet',
      'const sheetData = xDoc.root!.element(S.sheetData);',
      'const rows = sheetData!.elements(S.row);',
      '',
      '// Read cell values',
      'for (const row of rows) {',
      '  for (const cell of row.elements(S.c)) {',
      '    const ref = cell.attribute(NoNamespace.r)?.value;',
      '    const val = cell.element(S.v)?.value;',
      '    console.log(`${ref}: ${val}`);',
      '  }',
      '}',
    ]),
    gap(),
    label('Navigate PowerPoint presentation XML'),
    ...codeLines([
      'import { P, A } from "openxmlsdkts";',
      '',
      '// Get shapes from a slide',
      'const spTree = xDoc.root!',
      '  .element(P.cSld)',
      '  ?.element(P.spTree);',
      'const shapes = spTree?.elements(P.sp) ?? [];',
    ]),

    h2('Namespace Reference'),
    plain('The following table lists all 48 namespace classes, their descriptions, and namespace URIs.'),
    nsTable,
  ];
  await writeDoc('Namespaces', content);
}


// ════════════════════════════════════════════════════════════════════════════
// MAIN
// ════════════════════════════════════════════════════════════════════════════
async function main() {
  mkdirSync(DOCS_DIR, { recursive: true });
  console.log('Generating OpenXmlSdkTs documentation...');

  await genOverview();
  await genOpenXmlPackage();
  await genOpenXmlPart();
  await genOpenXmlRelationship();
  await genWmlPackage();
  await genWmlPart();
  await genSmlPackage();
  await genSmlPart();
  await genPmlPackage();
  await genPmlPart();
  await genContentType();
  await genRelationshipType();
  await genUtility();
  await genNamespaces();

  console.log('\nDone! Generated 14 DOCX files in docs/');
}

main().catch(console.error);
