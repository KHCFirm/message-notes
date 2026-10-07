import { NotesController, MAX_TEXT } from './core.js';
import { OfficeBridge } from './office-bridge.js';

const $ = id => document.getElementById(id);
const demo = new URLSearchParams(location.search).get('demo') === '1';
let controller;
let itemId = null;
function show(id, text) { $(id).textContent = text; $(id).hidden = !text; }
function render(s) {
  if (itemId !== s.item?.id) { $('confirmReload').hidden = true; itemId = s.item?.id; }
  $('subject').textContent = s.item?.subject || 'Select an email';
  $('sender').textContent = s.item?.sender || 'Open a received or sent email to see its note.';
  if ($('note').value !== s.text) $('note').value = s.text;
  $('note').disabled = false;
  $('note').readOnly = s.phase !== 'ready';
  $('note').setAttribute('aria-invalid', String(!s.validation.ok));
  $('counter').textContent = `${s.text.length.toLocaleString()} / ${MAX_TEXT.toLocaleString()}`;
  $('counter').classList.toggle('over', !s.validation.ok);
  $('save').disabled = s.phase !== 'ready' || !s.dirty || !s.validation.ok || s.conflict;
  $('save').textContent = s.phase === 'saving' ? 'Saving...' : 'Save note';
  $('clear').disabled = s.phase !== 'ready' || !s.text;
  $('copy').disabled = !s.text;
  $('reload').disabled = !s.item || s.phase === 'saving' || s.phase === 'loading';
  const status = s.phase === 'loading' ? 'Loading...' : s.phase === 'saving' ? 'Saving...' : s.error ? 'Needs attention' : s.dirty ? 'Unsaved' : s.record ? 'Saved' : 'No note';
  $('status').textContent = status;
  $('status').className = `status ${s.error ? 'failed' : s.dirty ? 'unsaved' : s.record ? 'saved' : ''}`;
  $('lastSaved').textContent = s.record ? `Last saved ${new Date(s.record.updatedAt).toLocaleString(undefined,{month:'short',day:'numeric',hour:'numeric',minute:'2-digit'})}` : 'No saved note';
  show('error', s.error || s.validation.message);
  show('notice', s.message);
  const count = s.otherDrafts.length;
  show('draftNotice', count ? `${count} other email${count === 1 ? ' has' : 's have'} unsaved changes: ${s.otherDrafts.slice(0,3).join('; ')}${count > 3 ? '; ...' : ''}. Return to those emails and save before closing this pane.` : '');
}

$('note').addEventListener('input', e => controller?.edit(e.target.value));
$('save').addEventListener('click', () => controller?.save());
$('clear').addEventListener('click', () => { controller?.edit(''); $('note').focus(); });
$('reload').addEventListener('click', () => {
  if (controller.state.dirty) $('confirmReload').hidden = false;
  else controller.reloadDiscardingDraft();
});
$('discard').addEventListener('click', () => { $('confirmReload').hidden = true; controller.reloadDiscardingDraft(); });
$('cancelDiscard').addEventListener('click', () => { $('confirmReload').hidden = true; $('note').focus(); });
$('copy').addEventListener('click', async () => {
  try { await navigator.clipboard.writeText($('note').value); show('notice','Note copied to your clipboard.'); }
  catch { $('note').focus(); $('note').select(); show('notice','Text selected. Press Ctrl+C to copy it.'); }
});
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); controller?.save(); }
});
window.addEventListener('beforeunload', e => {
  if (controller?.hasUnsaved()) { e.preventDefault(); e.returnValue = ''; }
});

function loadOffice() {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Outlook could not initialize. Open Message Notes from the Apps menu on an email. Check that the app host and Microsoft Office.js are reachable.')), 20000);
    const script = document.createElement('script');
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.onerror = () => { clearTimeout(timer); reject(new Error('Microsoft Office.js could not load. Check your internet connection or web filtering policy.')); };
    script.onload = () => {
      Office.onReady(info => {
        clearTimeout(timer);
        if (info.host !== Office.HostType.Outlook) reject(new Error('Open this app from an email in Outlook. This browser page is not connected to a mailbox. Use taskpane.html?demo=1 for the sample-only preview.'));
        else resolve();
      });
    };
    document.head.appendChild(script);
  });
}

async function start() {
  let adapter;
  if (demo) {
    const { DemoBridge } = await import('./demo-bridge.js');
    adapter = new DemoBridge();
    $('demoTools').hidden = false;
    $('storageLabel').textContent = 'Demo only. No connection to Outlook or a mailbox.';
    $('demoItem').addEventListener('change', e => adapter.change(Number(e.target.value)));
  } else {
    await loadOffice();
    if (!Office.context.requirements.isSetSupported('Mailbox','1.5')) throw new Error('This Outlook client does not support the required Mailbox 1.5 APIs. Use a supported Microsoft 365 account in new Outlook or Outlook on the web.');
    adapter = new OfficeBridge();
  }
  controller = new NotesController(adapter, render);
  // Fail closed if ItemChanged registration fails: stale message context is unsafe.
  await adapter.subscribe(() => controller.select());
  await controller.select();
}
start().catch(error => {
  $('subject').textContent = 'Open Message Notes in Outlook';
  $('sender').textContent = 'The mailbox connection is not active.';
  $('status').textContent = 'Not connected';
  show('fatal',error.message);
});
