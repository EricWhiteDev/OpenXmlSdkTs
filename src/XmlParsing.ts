/*
 * OpenXmlSdkTs (Open XML SDK for TypeScript)
 * Copyright (c) 2026 Eric White
 * eric@ericwhite.com
 * https://www.ericwhite.com
 * linkedin.com/in/ericwhitedev
 * Licensed under the MIT License
 */

import { XDocument } from "ltxmlts";

type PreserveWhitespaceParse = (xml: string, options?: { preserveWhitespace?: boolean }) => XDocument;

const parseWithWhitespace = XDocument.parse as unknown as PreserveWhitespaceParse;

/**
 * Parses OOXML XML while preserving whitespace-only text nodes when supported
 * by the installed ltxmlts version.
 *
 * @remarks
 * The cast keeps OpenXmlSdkTs buildable until the paired ltxmlts release that
 * adds the optional `preserveWhitespace` parse overload is available in the
 * repo lockfile.
 */
export function parseXmlPreservingWhitespace(xml: string): XDocument {
  return parseWithWhitespace(xml, { preserveWhitespace: true });
}
