// Zip layer of the surgical xlsx save (src/renderers/sheet.ts, section XLSX PACKAGE; contract §2.3, §14 P1 WP1):
// parseZip / readEntry / writeZip / crc32 / ZipError. Unchanged entries must be copied byte for byte (order, data
// descriptors, extra fields, non-UTF-8 names, zip comment); packages that cannot be written back exactly throw ZipError.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');

const S = require('../.out/sheet.js');

const sample = (name) => fs.readFileSync(path.join(__dirname, '..', 'samples', name));

/**
 * A hand-made zip (so the tests control every header field). entries: { name, nameBytes?, data, method = 8,
 * utf8 = true, descriptor = false, extra?, comment?, flags?, crc? }; o: { comment?, prefix?, zip64Locator? }.
 */
function makeZip(entries, o = {}) {
  const chunks = [];
  let offset = 0;
  const push = (b) => {
    chunks.push(b);
    offset += b.length;
  };
  if (o.prefix) push(o.prefix);
  const central = [];
  for (const e of entries) {
    const nameBytes = e.nameBytes ?? Buffer.from(e.name, 'utf8');
    const data = Buffer.from(e.data);
    const method = e.method ?? 8;
    const packed = method === 8 ? zlib.deflateRawSync(data) : data;
    const crc = e.crc ?? zlib.crc32(data);
    const extra = e.extra ?? Buffer.alloc(0);
    const comment = Buffer.from(e.comment ?? '', 'utf8');
    const flags = (e.flags ?? 0) | (e.utf8 === false ? 0 : 0x800) | (e.descriptor ? 0x8 : 0);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0x6000, 10);
    local.writeUInt16LE(0x5a21, 12);
    local.writeUInt32LE(e.descriptor ? 0 : crc, 14);
    local.writeUInt32LE(e.descriptor ? 0 : packed.length, 18);
    local.writeUInt32LE(e.descriptor ? 0 : data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(extra.length, 28);
    const lho = offset;
    push(Buffer.concat([local, nameBytes, extra]));
    push(packed);
    if (e.descriptor) {
      const dd = Buffer.alloc(16);
      dd.writeUInt32LE(0x08074b50, 0);
      dd.writeUInt32LE(crc, 4);
      dd.writeUInt32LE(packed.length, 8);
      dd.writeUInt32LE(data.length, 12);
      push(dd);
    }
    const rec = Buffer.alloc(46);
    rec.writeUInt32LE(0x02014b50, 0);
    rec.writeUInt16LE(0x031e, 4);
    rec.writeUInt16LE(20, 6);
    rec.writeUInt16LE(flags, 8);
    rec.writeUInt16LE(method, 10);
    rec.writeUInt16LE(0x6000, 12);
    rec.writeUInt16LE(0x5a21, 14);
    rec.writeUInt32LE(crc, 16);
    rec.writeUInt32LE(e.csize ?? packed.length, 20);
    rec.writeUInt32LE(data.length, 24);
    rec.writeUInt16LE(nameBytes.length, 28);
    rec.writeUInt16LE(extra.length, 30);
    rec.writeUInt16LE(comment.length, 32);
    rec.writeUInt32LE(0x81a40000, 38);
    rec.writeUInt32LE(lho, 42);
    central.push(Buffer.concat([rec, nameBytes, extra, comment]));
  }
  const cdStart = offset;
  for (const rec of central) push(rec);
  const cdSize = offset - cdStart;
  if (o.zip64Locator) {
    const loc = Buffer.alloc(20);
    loc.writeUInt32LE(0x07064b50, 0);
    push(loc);
  }
  const comment = Buffer.from(o.comment ?? '', 'utf8');
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(cdSize, 12);
  end.writeUInt32LE(cdStart, 16);
  end.writeUInt16LE(comment.length, 20);
  push(end);
  push(comment);
  return Buffer.concat(chunks);
}

/** The raw bytes of an entry (local header + data + descriptor). */
const rawSpan = (zip, name) => {
  const e = zip.get(name);
  return zip.buf.subarray(e.lho, e.spanEnd);
};

const ODD_ZIP_ENTRIES = () => [
  { name: 'a.xml', data: '<a>first</a>' },
  { name: 'desc.xml', data: '<d>data descriptor</d>'.repeat(50), descriptor: true },
  { name: 'extra.xml', data: '<e>extra field</e>', extra: Buffer.from([0x55, 0x54, 0x05, 0x00, 0x01, 0x10, 0x20, 0x30, 0x40]) },
  { name: 'latin1-é.xml', nameBytes: Buffer.from('latin1-\xe9.xml', 'latin1'), utf8: false, data: '<n>non-UTF-8 name</n>' },
  { name: 'stored.bin', data: Buffer.from([0, 1, 2, 3, 250, 251]), method: 0, comment: 'entry comment' },
  { name: 'z.xml', data: '<z>last</z>' },
];

// ----- crc32 -----

test('crc32: zlib and table versions agree, also when continued', () => {
  const data = Buffer.from(Array.from({ length: 70000 }, (_, i) => (i * 7919) & 0xff));
  const whole = S.crc32(data);
  assert.equal(whole, zlib.crc32(data) >>> 0);
  assert.equal(S.crc32(data, 0, { table: true }), whole);
  const part = S.crc32(data.subarray(0, 1234), 0, { table: true });
  assert.equal(S.crc32(data.subarray(1234), part, { table: true }), whole);
  assert.equal(S.crc32(Buffer.alloc(0)), 0);
  assert.equal(S.crc32(Buffer.from('123456789'), 0, { table: true }), 0xcbf43926);
});

// ----- parseZip / readEntry -----

test('parseZip + readEntry: every part of styled.xlsx unpacks with the right size and CRC', async () => {
  const zip = S.parseZip(sample('styled.xlsx'));
  assert.ok(zip.entries.length > 10);
  assert.ok(zip.get('XL/WORKBOOK.XML'), 'lookups ignore case');
  for (const entry of zip.entries) {
    const data = await S.readEntry(zip, entry);
    assert.equal(data.length, entry.usize, entry.name);
  }
});

test('parseZip: data descriptors, extra fields, non-UTF-8 names, stored entries and comments are read', async () => {
  const buf = makeZip(ODD_ZIP_ENTRIES(), { comment: 'zip comment', prefix: Buffer.from('PREFIX') });
  const zip = S.parseZip(buf);
  assert.deepEqual(zip.entries.map((e) => e.name), ['a.xml', 'desc.xml', 'extra.xml', 'latin1-é.xml', 'stored.bin', 'z.xml']);
  assert.equal(zip.comment.toString(), 'zip comment');
  assert.equal(zip.prefixEnd, 6);
  assert.equal((await S.readEntry(zip, zip.get('desc.xml'))).toString(), '<d>data descriptor</d>'.repeat(50));
  assert.deepEqual([...(await S.readEntry(zip, zip.get('stored.bin')))], [0, 1, 2, 3, 250, 251]);
});

test('readEntry: a wrong CRC is a ZipError', async () => {
  const zip = S.parseZip(makeZip([{ name: 'a.xml', data: '<a/>', crc: 12345 }]));
  await assert.rejects(S.readEntry(zip, zip.get('a.xml')), S.ZipError);
});

test('readEntry: more or fewer unpacked bytes than the directory says, or damaged deflate data, is a ZipError', async () => {
  const data = '<a>' + 'x'.repeat(200000) + '</a>';
  const zip = S.parseZip(makeZip([{ name: 'a.xml', data }]));
  const entry = zip.get('a.xml');
  assert.equal((await S.readEntry(zip, entry)).toString(), data);
  await assert.rejects(S.readEntry(zip, { ...entry, usize: entry.usize - 1 }), /wrong size/);
  await assert.rejects(S.readEntry(zip, { ...entry, usize: entry.usize + 1 }), /wrong size/);
  const bytes = makeZip([{ name: 'a.xml', data }]);
  const broken = S.parseZip(bytes);
  bytes[broken.get('a.xml').dataStart] = 0xff; // final block with the reserved block type 3
  await assert.rejects(S.readEntry(broken, broken.get('a.xml')), (err) => err instanceof S.ZipError && /cannot be unpacked/.test(err.message));
});

test('parseZip: packages that cannot be written back exactly are ZipErrors', () => {
  const cases = {
    'not a zip': Buffer.from('this is not a zip file at all, just text'),
    'ZIP64 locator': makeZip([{ name: 'a.xml', data: '<a/>' }], { zip64Locator: true }),
    'ZIP64 sizes': makeZip([{ name: 'a.xml', data: '<a/>', csize: 0xffffffff }]),
    encrypted: makeZip([{ name: 'a.xml', data: '<a/>', flags: 0x1 }]),
    'strong encryption': makeZip([{ name: 'a.xml', data: '<a/>', flags: 0x40 }]),
    'method 12 (bzip2)': makeZip([{ name: 'a.xml', data: '<a/>', method: 12 }]),
    'method 14 (lzma)': makeZip([{ name: 'a.xml', data: '<a/>', method: 14 }]),
    'duplicate names': makeZip([{ name: 'xl/a.xml', data: '<a/>' }, { name: 'XL/A.xml', data: '<b/>' }]),
  };
  for (const [label, buf] of Object.entries(cases)) {
    assert.throws(() => S.parseZip(buf), S.ZipError, label);
  }
  const truncated = sample('styled.xlsx').subarray(0, 5000);
  assert.throws(() => S.parseZip(truncated), S.ZipError, 'truncated');
});

test('parseZip: a local header that does not match the directory is a ZipError', () => {
  const buf = makeZip([{ name: 'a.xml', data: '<a/>' }]);
  buf[30] = 0x62; // first byte of the local file name: 'a' -> 'b'
  assert.throws(() => S.parseZip(buf), /does not match/);
});

// ----- writeZip -----

test('writeZip without changes writes the input back byte for byte', async () => {
  for (const buf of [sample('styled.xlsx'), makeZip(ODD_ZIP_ENTRIES(), { comment: 'c', prefix: Buffer.from('PK-prefix') })]) {
    const out = await S.writeZip(S.parseZip(buf), new Map());
    assert.ok(Buffer.from(out.bytes).equals(buf));
    assert.deepEqual(out.changed, []);
  }
});

test('writeZip: unchanged entries are raw copies (order, descriptors, extra fields, names, comments); rewritten ones inflate', async () => {
  const buf = makeZip(ODD_ZIP_ENTRIES(), { comment: 'zip comment' });
  const zip = S.parseZip(buf);
  const out = await S.writeZip(zip, new Map([['a.xml', [Buffer.from('<a>'), Buffer.from('changed</a>')]], ['z.xml', null]]), {
    added: [{ name: 'new/part.xml', parts: [Buffer.from('<new/>')] }],
  });
  assert.deepEqual(out.changed, ['a.xml', 'new/part.xml']);
  assert.deepEqual(out.removed, ['z.xml']);
  assert.equal(out.identical, 4);
  const z2 = S.parseZip(out.bytes);
  assert.deepEqual(z2.entries.map((e) => e.name), ['a.xml', 'desc.xml', 'extra.xml', 'latin1-é.xml', 'stored.bin', 'new/part.xml']);
  for (const name of ['desc.xml', 'extra.xml', 'latin1-é.xml', 'stored.bin']) {
    assert.ok(rawSpan(z2, name).equals(rawSpan(zip, name)), `${name} is byte-identical`);
    const a = zip.get(name).rec;
    const b = z2.get(name).rec;
    assert.ok(a.subarray(0, 42).equals(b.subarray(0, 42)) && a.subarray(46).equals(b.subarray(46)), `${name} directory record kept`);
  }
  assert.ok(z2.get('latin1-é.xml').nameBytes.equals(Buffer.from('latin1-\xe9.xml', 'latin1')));
  assert.equal(z2.comment.toString(), 'zip comment');
  assert.equal((await S.readEntry(z2, z2.get('a.xml'))).toString(), '<a>changed</a>');
  assert.equal((await S.readEntry(z2, z2.get('new/part.xml'))).toString(), '<new/>');
  // A rewritten entry keeps its name bytes, time and date and its directory comment.
  assert.equal(z2.get('a.xml').rec.readUInt32LE(12), zip.get('a.xml').rec.readUInt32LE(12));
  // The output also unpacks with a second, independent reader.
  const JSZip = require('jszip');
  const jz = await JSZip.loadAsync(out.bytes, { checkCRC32: true });
  assert.equal(await jz.file('a.xml').async('string'), '<a>changed</a>');
});

test('writeZip: a rewritten part bigger than 4 MB is deflated (level 1) and reads back', async () => {
  const big = Buffer.alloc(5 << 20, '<row r="1"><c r="A1"><v>1</v></c></row>');
  const zip = S.parseZip(makeZip([{ name: 'a.xml', data: '<a/>' }, { name: 'b.xml', data: '<b/>' }]));
  const out = await S.writeZip(zip, new Map([['b.xml', [big]]]));
  const z2 = S.parseZip(out.bytes);
  assert.ok(z2.get('b.xml').csize < big.length / 10);
  assert.ok((await S.readEntry(z2, z2.get('b.xml'))).equals(big));
});

test('writeZip: changing a part the package does not have, or adding one it has, is refused', async () => {
  const zip = S.parseZip(makeZip([{ name: 'a.xml', data: '<a/>' }]));
  await assert.rejects(S.writeZip(zip, new Map([['missing.xml', [Buffer.from('x')]]])), /no part/);
  await assert.rejects(S.writeZip(zip, new Map(), { added: [{ name: 'A.XML', parts: [Buffer.from('x')] }] }), /exists already/);
});

test('writeZip: the yield hook can cancel', async () => {
  const big = Buffer.alloc(20 << 20, 'x');
  const zip = S.parseZip(makeZip([{ name: 'a.xml', data: '<a/>' }]));
  const cancel = new Error('cancelled');
  await assert.rejects(S.writeZip(zip, new Map([['a.xml', [big]]]), { onYield: () => { throw cancel; } }), (err) => err === cancel);
});
