import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { buildManifest, hostedBase, ADDIN_ID, VERSION } from '../site/manifest-builder.js';
const template = await readFile(new URL('../site/manifest.template.xml', import.meta.url), 'utf8');

test('GitHub project URL replaces every placeholder and preserves add-in identity', () => {
  const xml = buildManifest(template,'https://example-user.github.io/message-notes/');
  assert.ok(xml.includes(`<Id>${ADDIN_ID}</Id>`));
  assert.ok(xml.includes(`<Version>${VERSION}</Version>`));
  assert.ok(xml.includes('https://example-user.github.io/message-notes/taskpane.html'));
  assert.ok(xml.includes('<AppDomain>https://example-user.github.io</AppDomain>'));
  assert.ok(!/__BASE_URL__|__ORIGIN__|localhost/.test(xml));
});

test('manifest uses one current V1_1 command surface and no FunctionFile', () => {
  assert.equal((template.match(/VersionOverridesV1_1/g) || []).length, 1);
  assert.equal((template.match(/MessageReadCommandSurface/g) || []).length, 1);
  assert.ok(template.includes('<SupportsPinning>true</SupportsPinning>'));
  assert.ok(!template.includes('<FunctionFile'));
  assert.ok(!template.includes('Commands.Url'));
});

test('custom HTTPS hosts and subpaths work', () => {
  assert.ok(buildManifest(template,'https://notes.example.org/department/tools/').includes('https://notes.example.org/department/tools/taskpane.html'));
  assert.ok(buildManifest(template,'https://notes.example.org/').includes('https://notes.example.org/taskpane.html'));
});

test('directory base is origin-safe and normalizes trailing slash', () => {
  assert.deepEqual(hostedBase('https://user.github.io/message-notes/'),{base:'https://user.github.io/message-notes',origin:'https://user.github.io'});
});

test('rejects repository, raw, localhost, insecure, and authenticated URLs', () => {
  for (const url of ['https://github.com/user/repo/','https://raw.githubusercontent.com/user/repo/main/',
    'https://localhost:3000/','https://127.0.0.1/','https://[::1]/','http://example.org/',
    'file:///tmp/index.html','https://name:secret@example.org/','https://example.org/?key=secret','https://example.org/#x']) {
    assert.throws(()=>buildManifest(template,url),Error,url);
  }
});

test('XML escapes allowed path characters', () => {
  const xml=buildManifest(template,'https://example.org/notes&tools/');
  assert.ok(xml.includes('notes&amp;tools/taskpane.html'));
});
