// A few-kilobyte OpenDocument spreadsheet that materializes millions of cells.
//
// ODS stores a run of identical rows or cells once, with a repeat count on the element, so the
// file's size says nothing about its cell count: the default shape here is one numeric cell
// repeated across 2,000 columns and 2,000 rows -- four million cells from a ~3 KB upload. Every
// guard that runs before SheetJS materializes the grid passes it, which is why the parse runs in
// a dynamic worker of its own. Kept dependency-light (fflate only) so the staging check can build
// the same file outside the test runner.

import { zipSync, strToU8 } from "fflate";

export const ODS_MIME_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

export function repeatedCellsOds(rows = 2_000, columns = 2_000): Uint8Array {
  let content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content` +
    ` xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
    ` xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
    ` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">` +
    `<office:body><office:spreadsheet><table:table table:name="Sheet1">` +
    `<table:table-row table:number-rows-repeated="${rows}">` +
    `<table:table-cell office:value-type="float" office:value="1"` +
    ` table:number-columns-repeated="${columns}"><text:p>1</text:p></table:table-cell>` +
    `</table:table-row></table:table></office:spreadsheet></office:body>` +
    `</office:document-content>`;
  let manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"` +
    ` manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="${ODS_MIME_TYPE}"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;
  // SheetJS refuses a package without one, however little it holds.
  let styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
    ` office:version="1.2"/>`;
  // The `mimetype` entry comes first and uncompressed, as the format requires -- and as the upload
  // path's container check looks for.
  return zipSync({
    mimetype: [strToU8(ODS_MIME_TYPE), { level: 0 }],
    "META-INF/manifest.xml": strToU8(manifest),
    "styles.xml": strToU8(styles),
    "content.xml": strToU8(content),
  });
}
