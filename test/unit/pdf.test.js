// PDF header check and error messages (src/renderers/pdf.ts). Rendering itself happens in the webview with pdf.js.
// Run with `npm test` (it bundles the renderers into test/.out/ first).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { checkPdf, describePdfFailure, PDF_HEADER_SEARCH_BYTES } = require('../.out/pdf.js');

const bytes = (text) => new TextEncoder().encode(text);

test('checkPdf reads the version from the header of sample.pdf', () => {
  const pdf = fs.readFileSync(path.join(__dirname, '..', 'samples', 'sample.pdf'));
  assert.deepEqual(checkPdf(new Uint8Array(pdf.subarray(0, PDF_HEADER_SEARCH_BYTES))), { version: '1.7' });
});

test('checkPdf finds a header after leading junk', () => {
  assert.deepEqual(checkPdf(bytes('junk before the header\n%PDF-2.0\n')), { version: '2.0' });
  assert.deepEqual(checkPdf(bytes('%PDF-')), {}, 'a header without a version is still a PDF');
});

test('checkPdf explains files that are not PDFs', () => {
  assert.match(checkPdf(new Uint8Array(0)).error, /^The file is empty \(0 bytes\)/);
  assert.match(checkPdf(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00])).error, /it is a ZIP package/);
  assert.match(checkPdf(bytes('<!DOCTYPE html><html><body>Sign in</body></html>')).error, /it is a web page \(HTML\)/);
  assert.match(checkPdf(bytes('just some text')).error, /no %PDF header/);
});

test('describePdfFailure: password errors (pdf.js PasswordException)', () => {
  // pdf.js errors can arrive as Error subclasses or as plain structured-cloned objects.
  const needPassword = Object.assign(new Error('No password given'), { name: 'PasswordException', code: 1 });
  assert.equal(describePdfFailure(needPassword), 'This PDF is password-protected. Enter its password to open it.');
  assert.equal(
    describePdfFailure({ name: 'PasswordException', message: 'Incorrect Password', code: 2 }),
    'The password is incorrect. Enter the password of this PDF to open it.',
  );
});

test('describePdfFailure: damaged files, worker failures and unknown errors', () => {
  assert.match(describePdfFailure({ name: 'InvalidPDFException', message: 'Invalid PDF structure.' }), /structure is damaged/);
  assert.match(describePdfFailure(new Error('Setting up fake worker failed: x')), /PDF engine could not start/);
  assert.equal(describePdfFailure('something odd'), 'Could not display the PDF: something odd');
  assert.equal(describePdfFailure(undefined), 'Could not display the PDF: Unknown error.');
});
