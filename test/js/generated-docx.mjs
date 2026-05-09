const encoder = new TextEncoder();

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function escapeXml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function writeUint16(output, value) {
  output.push(value & 0xff, (value >>> 8) & 0xff);
}

function writeUint32(output, value) {
  output.push(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff);
}

function writeBytes(output, bytes) {
  for (const byte of bytes) output.push(byte);
}

function createStoredZip(entries) {
  const output = [];
  const centralDirectory = [];
  const metadata = [];

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const dataBytes = typeof entry.data === 'string' ? encoder.encode(entry.data) : entry.data;
    const localHeaderOffset = output.length;
    const entryCrc = crc32(dataBytes);

    writeUint32(output, 0x04034b50);
    writeUint16(output, 20); // version needed
    writeUint16(output, 0); // flags
    writeUint16(output, 0); // compression: store
    writeUint16(output, 0); // modified time
    writeUint16(output, 0); // modified date
    writeUint32(output, entryCrc);
    writeUint32(output, dataBytes.length);
    writeUint32(output, dataBytes.length);
    writeUint16(output, nameBytes.length);
    writeUint16(output, 0); // extra length
    writeBytes(output, nameBytes);
    writeBytes(output, dataBytes);

    metadata.push({ nameBytes, dataBytes, crc: entryCrc, localHeaderOffset });
  }

  const centralDirectoryOffset = output.length;
  for (const entry of metadata) {
    writeUint32(centralDirectory, 0x02014b50);
    writeUint16(centralDirectory, 20); // version made by
    writeUint16(centralDirectory, 20); // version needed
    writeUint16(centralDirectory, 0); // flags
    writeUint16(centralDirectory, 0); // compression: store
    writeUint16(centralDirectory, 0); // modified time
    writeUint16(centralDirectory, 0); // modified date
    writeUint32(centralDirectory, entry.crc);
    writeUint32(centralDirectory, entry.dataBytes.length);
    writeUint32(centralDirectory, entry.dataBytes.length);
    writeUint16(centralDirectory, entry.nameBytes.length);
    writeUint16(centralDirectory, 0); // extra length
    writeUint16(centralDirectory, 0); // comment length
    writeUint16(centralDirectory, 0); // disk number
    writeUint16(centralDirectory, 0); // internal attrs
    writeUint32(centralDirectory, 0); // external attrs
    writeUint32(centralDirectory, entry.localHeaderOffset);
    writeBytes(centralDirectory, entry.nameBytes);
  }

  writeBytes(output, centralDirectory);
  writeUint32(output, 0x06054b50);
  writeUint16(output, 0); // disk
  writeUint16(output, 0); // central dir disk
  writeUint16(output, metadata.length);
  writeUint16(output, metadata.length);
  writeUint32(output, centralDirectory.length);
  writeUint32(output, centralDirectoryOffset);
  writeUint16(output, 0); // zip comment length

  return new Uint8Array(output);
}

export function createGeneratedDocxBytes({ paragraphs = 1000, prefix = 'Generated DOCX SAX paragraph' } = {}) {
  const body = [];
  for (let index = 0; index < paragraphs; index += 1) {
    body.push(`<w:p><w:r><w:t>${escapeXml(prefix)} ${index}</w:t></w:r></w:p>`);
  }

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body.join('\n    ')}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

  return createStoredZip([
    { name: '[Content_Types].xml', data: CONTENT_TYPES },
    { name: '_rels/.rels', data: ROOT_RELS },
    { name: 'word/document.xml', data: documentXml },
  ]);
}
