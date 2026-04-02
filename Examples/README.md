# OpenXmlSdkTs Examples

This directory contains example Node.js programs that demonstrate basic usage of the OpenXmlSdkTs library for reading, modifying, and writing Office Open XML documents.

## Prerequisites

From the repository root, install all dependencies (including the workspace link to the library):

```bash
npm install
```

## Running Examples

Each example is a standalone TypeScript file that can be run with `node --conditions=import --import tsx`. Run from the **repository root**:

```bash
# Example 1: Load, modify, and save a binary DOCX file
node --conditions=import --import tsx Examples/01-load-modify-save-binary/load-modify-save-binary.ts

# Example 2: Load, modify, and save a Flat OPC file
node --conditions=import --import tsx Examples/02-load-modify-save-flatopc/load-modify-save-flatopc.ts

# Example 3: Round-trip a DOCX through base64 encoding
node --conditions=import --import tsx Examples/03-roundtrip-base64/roundtrip-base64.ts

# Example 4: Round-trip a DOCX through Flat OPC XML
node --conditions=import --import tsx Examples/04-roundtrip-flatopc/roundtrip-flatopc.ts
```

Each example creates an `example-output/` directory inside its folder with the results. You can inspect these output files after running an example.

## Examples Overview

| Example | Description |
|---------|-------------|
| [01-load-modify-save-binary](./01-load-modify-save-binary/) | Opens a binary DOCX, navigates parts (main document, comments), adds a paragraph, and saves |
| [02-load-modify-save-flatopc](./02-load-modify-save-flatopc/) | Converts a DOCX to Flat OPC XML, modifies it, and saves back to both Flat OPC and binary |
| [03-roundtrip-base64](./03-roundtrip-base64/) | Round-trips a DOCX through base64 encoding and verifies document integrity |
| [04-roundtrip-flatopc](./04-roundtrip-flatopc/) | Round-trips a DOCX through Flat OPC XML format and verifies document integrity |

## Pre-Atomized Namespaces

All examples demonstrate the use of **pre-atomized namespaces and names**, a key feature of OpenXmlSdkTs. Instead of working with raw XML namespace strings, you use typed constants like `W.body`, `W.p`, `W.r`, and `W.t`. These are pre-initialized `XName` objects that enable efficient identity-based equality checks (`===`) and type-safe XML navigation. See the inline comments in each example for details.
