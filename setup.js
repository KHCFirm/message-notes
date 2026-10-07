import { ADDIN_ID, VERSION, hostedBase, buildManifest } from './manifest-builder.js';
const $ = id => document.getElementById(id);
let manifest = '';
const paths = ['taskpane.html','app.js','core.js','office-bridge.js',
  'demo-bridge.js','styles.css','support.html','assets/icon-16.png','assets/icon-32.png',
  'assets/icon-64.png','assets/icon-80.png','assets/icon-128.png'];

function notice(text, failed = false) {
  $('setupStatus').textContent = text;
  $('setupStatus').className = `notice ${failed ? 'error' : 'info'}`;
}
function downloadableXml(text) {
  const blob = new Blob([text], { type: 'application/xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Message-Notes-GitHub.xml';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
async function prepare() {
  $('downloadManifest').disabled = true;
  $('checkSite').disabled = true;
  manifest = '';
  try {
    const page = new URL(location.href);
    page.search = '';
    page.hash = '';
    const directory = new URL('./', page);
    const { base } = hostedBase(directory.href);
    $('siteAddress').textContent = `${base}/`;
    notice('Checking the published website files...');
    const response = await fetch(`${base}/manifest.template.xml`, { cache: 'no-store' });
    if (!response.ok) throw new Error(`manifest.template.xml returned HTTP ${response.status}. Upload all files from site/.`);
    manifest = buildManifest(await response.text(), base);
    const parsed = new DOMParser().parseFromString(manifest, 'application/xml');
    if (parsed.querySelector('parsererror')) throw new Error('The generated XML could not be parsed.');
    await Promise.all(paths.map(async path => {
      const result = await fetch(`${base}/${path}`, { cache: 'no-store' });
      if (!result.ok) throw new Error(`${path} returned HTTP ${result.status}. Check the upload and Pages deployment.`);
      const type = (result.headers.get('content-type') || '').toLowerCase();
      if (path.endsWith('.js') && !/javascript|ecmascript/.test(type)) throw new Error(`${path} is not being served as JavaScript.`);
      if (path.endsWith('.png') && !type.startsWith('image/')) throw new Error(`${path} is not being served as an image.`);
      if (path === 'taskpane.html' && !(await result.text()).includes('id="note"')) throw new Error('taskpane.html is not the expected application.');
    }));
    $('downloadManifest').disabled = false;
    notice('Website files are available. Download the manifest, then install it in your mailbox. This check does not verify Outlook installation or note saving.');
  } catch (error) {
    manifest = '';
    notice(`${error.message} Open this setup page at your published HTTPS site, not by double-clicking the local file.`, true);
  } finally { $('checkSite').disabled = false; }
}
$('appId').textContent = ADDIN_ID;
$('appVersion').textContent = VERSION;
$('downloadManifest').addEventListener('click', () => { if (manifest) downloadableXml(manifest); });
$('checkSite').addEventListener('click', prepare);
prepare();
