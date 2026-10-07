// No credentials, mailbox access, or external service is used by this generator.
export const ADDIN_ID = '7c6010d9-3b66-45f7-86ca-296e175b2b86';
export const VERSION = '1.1.1.1';
const escapeXml = value => value.replaceAll('&', '&amp;').replaceAll('"', '&quot;')
  .replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll("'", '&apos;');

export function hostedBase(value) {
  const url = new URL(value);
  if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash) {
    throw new Error('Use the published HTTPS website address, without credentials, a query, or a fragment.');
  }
  const hostname = url.hostname.toLowerCase();
  if (hostname === 'localhost' || hostname.endsWith('.localhost') || hostname === '[::1]' || /^127\./.test(hostname)) {
    throw new Error('This package needs a published website, not localhost.');
  }
  if (hostname === 'github.com' || hostname === 'raw.githubusercontent.com') {
    throw new Error('Open the GitHub Pages website, not the repository or a raw-file link.');
  }
  return { base: url.href.replace(/\/+$/, ''), origin: url.origin };
}

export function buildManifest(template, value) {
  const { base, origin } = hostedBase(value);
  if (!template.includes(`<Id>${ADDIN_ID}</Id>`) || !template.includes(`<Version>${VERSION}</Version>`) ||
      !template.includes('__BASE_URL__') || !template.includes('__ORIGIN__')) {
    throw new Error('The manifest template is missing, stale, or not the expected Message Notes template.');
  }
  const xml = template.replaceAll('__BASE_URL__', escapeXml(base)).replaceAll('__ORIGIN__', escapeXml(origin));
  if (/__BASE_URL__|__ORIGIN__|https?:\/\/localhost/i.test(xml)) {
    throw new Error('The manifest still has an unresolved or local URL.');
  }
  return xml;
}
