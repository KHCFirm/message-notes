import { NotesController } from './core.js';
import { OfficeBridge } from './office-bridge.js';

const $ = id => document.getElementById(id);
const demo = new URLSearchParams(location.search).get('demo') === '1';
const AUTOSAVE_DELAY = 650;
const FONT_SIZES = new Set(['12px', '14px', '16px', '18px', '20px', '24px']);
const ALLOWED_TAGS = new Set(['BR', 'DIV', 'P', 'SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'FONT']);

let controller;
let activeItemId = null;
let lastStateHtml = null;
let autosaveTimer = null;
let savedRange = null;

function show(id, text) {
  $(id).textContent = text;
  $(id).hidden = !text;
}

function allowedColor(value) {
  const text = String(value || '').trim();
  return /^(#[0-9a-f]{3,8}|rgba?\([\d\s.,%]+\))$/i.test(text) ? text : '';
}

function normalizedFontSize(value) {
  const text = String(value || '').trim().toLowerCase();
  return FONT_SIZES.has(text) ? text : '';
}

function legacyFontSize(value) {
  const map = { '1': '12px', '2': '12px', '3': '14px', '4': '18px', '5': '24px', '6': '24px', '7': '24px' };
  return map[String(value || '')] || '';
}

function appendSanitized(node, target) {
  if (node.nodeType === Node.TEXT_NODE) {
    target.append(document.createTextNode(node.nodeValue || ''));
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;

  if (!ALLOWED_TAGS.has(node.tagName)) {
    for (const child of [...node.childNodes]) appendSanitized(child, target);
    return;
  }

  if (node.tagName === 'BR') {
    target.append(document.createElement('br'));
    return;
  }

  const mapped = {
    B: 'strong', STRONG: 'strong', I: 'em', EM: 'em', U: 'u',
    DIV: 'div', P: 'div', SPAN: 'span', FONT: 'span'
  }[node.tagName];
  const clean = document.createElement(mapped);

  if (mapped === 'span') {
    const styles = [];
    const weight = String(node.style?.fontWeight || '').toLowerCase();
    if (weight === 'bold' || Number.parseInt(weight, 10) >= 600) styles.push('font-weight:bold');
    if (String(node.style?.fontStyle || '').toLowerCase() === 'italic') styles.push('font-style:italic');
    const decoration = `${node.style?.textDecoration || ''} ${node.style?.textDecorationLine || ''}`.toLowerCase();
    if (decoration.includes('underline')) styles.push('text-decoration:underline');

    const color = allowedColor(node.style?.color || (node.tagName === 'FONT' ? node.getAttribute('color') : ''));
    if (color) styles.push(`color:${color}`);

    const size = normalizedFontSize(node.style?.fontSize) || (node.tagName === 'FONT' ? legacyFontSize(node.getAttribute('size')) : '');
    if (size) styles.push(`font-size:${size}`);

    if (styles.length) clean.setAttribute('style', styles.join(';'));
  }

  for (const child of [...node.childNodes]) appendSanitized(child, clean);
  target.append(clean);
}

function sanitizeHtml(raw) {
  const input = document.createElement('template');
  input.innerHTML = String(raw || '');
  const output = document.createElement('div');
  for (const node of [...input.content.childNodes]) appendSanitized(node, output);

  // Empty formatting shells and blank lines are not a saved note.
  const text = output.textContent.replace(/\u00a0/g, ' ').trim();
  if (!text) return '';
  return output.innerHTML;
}

function editorHtml() {
  return sanitizeHtml($('note').innerHTML);
}

function setEditorHtml(html) {
  $('note').innerHTML = html || '';
}

function selectionIsInEditor(range) {
  const editor = $('note');
  return !!range && (editor === range.commonAncestorContainer || editor.contains(range.commonAncestorContainer));
}

function rememberSelection() {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (selectionIsInEditor(range)) savedRange = range.cloneRange();
}

function restoreSelection() {
  if (!savedRange || !selectionIsInEditor(savedRange)) return false;
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(savedRange);
  return true;
}

function editorChanged() {
  if (!controller) return;
  const html = editorHtml();
  controller.edit(html);
  scheduleAutosave();
  rememberSelection();
}

function applyCommand(command, value = null) {
  const editor = $('note');
  if (editor.getAttribute('contenteditable') !== 'true') return;
  editor.focus();
  restoreSelection();
  try { document.execCommand('styleWithCSS', false, true); } catch {}
  document.execCommand(command, false, value);
  rememberSelection();
  editorChanged();
}

// Outlook's embedded browser can drop the live text selection when the native
// color picker opens. Apply color to the Range we saved before the picker took
// focus instead of relying on execCommand('foreColor') to find that selection.
function applyTextColor(value) {
  const color = allowedColor(value);
  const editor = $('note');
  if (!color || editor.getAttribute('contenteditable') !== 'true') return;

  editor.focus({ preventScroll: true });
  if (!restoreSelection()) return;

  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!selectionIsInEditor(range)) return;

  if (range.collapsed) {
    // For a caret-only selection, use the browser's typing-state command so
    // subsequently typed text adopts the chosen color.
    try { document.execCommand('styleWithCSS', false, true); } catch {}
    document.execCommand('foreColor', false, color);
  } else {
    // Range.extractContents works even when the selection crosses nested bold,
    // italic, underline, or size spans. Re-select the result so another color
    // can be chosen without selecting the text again.
    const span = document.createElement('span');
    span.style.color = color;
    span.append(range.extractContents());
    range.insertNode(span);

    const coloredRange = document.createRange();
    coloredRange.selectNodeContents(span);
    selection.removeAllRanges();
    selection.addRange(coloredRange);
    savedRange = coloredRange.cloneRange();
  }

  editorChanged();
}

function applyFontSize(px) {
  const editor = $('note');
  if (editor.getAttribute('contenteditable') !== 'true') return;
  editor.focus();
  restoreSelection();
  try { document.execCommand('styleWithCSS', false, false); } catch {}
  document.execCommand('fontSize', false, '7');
  for (const font of [...editor.querySelectorAll('font[size="7"]')]) {
    const span = document.createElement('span');
    span.style.fontSize = `${px}px`;
    while (font.firstChild) span.append(font.firstChild);
    font.replaceWith(span);
  }
  try { document.execCommand('styleWithCSS', false, true); } catch {}
  rememberSelection();
  editorChanged();
}

function scheduleAutosave(delay = AUTOSAVE_DELAY) {
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(async () => {
    autosaveTimer = null;
    if (!controller) return;
    const s = controller.state;
    if (s.phase === 'ready' && s.dirty && !s.error && !s.conflict) await controller.save();
  }, delay);
}

async function saveNow() {
  clearTimeout(autosaveTimer);
  autosaveTimer = null;
  await controller?.save();
}

function render(s) {
  const editor = $('note');
  const itemChanged = activeItemId !== s.item?.id;
  if (itemChanged) {
    activeItemId = s.item?.id ?? null;
    savedRange = null;
  }

  const canEdit = !!s.item && (s.phase === 'ready' || s.phase === 'saving');
  editor.setAttribute('contenteditable', canEdit ? 'true' : 'false');
  editor.setAttribute('aria-disabled', String(!canEdit));

  // Do not rewrite contenteditable while the user is typing; that would move the caret.
  const stateHtmlChanged = s.html !== lastStateHtml;
  if (itemChanged || (stateHtmlChanged && document.activeElement !== editor)) setEditorHtml(s.html);
  lastStateHtml = s.html;

  const blocked = !s.item || s.phase === 'loading' || s.phase === 'error' || s.phase === 'saving' ||
    !s.dirty || !s.validation.ok || s.conflict;
  $('save').disabled = blocked;
  $('save').textContent = s.phase === 'saving' ? 'Saving…' : 'Save';

  let status = 'Ready';
  if (!s.item) status = 'Open an email';
  else if (s.phase === 'loading') status = 'Loading…';
  else if (s.phase === 'saving') status = 'Saving…';
  else if (s.error) status = 'Save failed';
  else if (s.dirty) status = 'Autosaving…';
  else if (s.record) status = 'Saved';
  $('status').textContent = status;
  $('status').className = `save-status${s.error ? ' failed' : (!s.dirty && s.record ? ' saved' : '')}`;

  show('error', s.error || s.validation.message);
  $('recover').hidden = !s.conflict;

  // If the user typed again while an autosave was in flight, queue the new snapshot now.
  if (s.phase === 'ready' && s.dirty && !s.error && !s.conflict && !autosaveTimer) scheduleAutosave();
}

$('note').addEventListener('input', editorChanged);
$('note').addEventListener('keyup', rememberSelection);
$('note').addEventListener('mouseup', rememberSelection);
$('note').addEventListener('paste', event => {
  event.preventDefault();
  const text = event.clipboardData?.getData('text/plain') || '';
  document.execCommand('insertText', false, text);
  queueMicrotask(editorChanged);
});
$('note').addEventListener('focusout', event => {
  const next = event.relatedTarget;
  if (next?.closest?.('.format-toolbar')) return;
  if (next?.id === 'save') return;
  if (controller?.state.phase === 'ready' && controller.state.dirty && !controller.state.error) saveNow();
});

document.addEventListener('selectionchange', rememberSelection);

for (const [id, command] of [['bold', 'bold'], ['italic', 'italic'], ['underline', 'underline']]) {
  $(id).addEventListener('mousedown', event => event.preventDefault());
  $(id).addEventListener('click', () => applyCommand(command));
}

$('fontSize').addEventListener('pointerdown', rememberSelection);
$('fontSize').addEventListener('change', event => applyFontSize(event.target.value));
let colorApplyTimer = null;
function queueTextColor(value, immediate = false) {
  document.querySelector('.color-control')?.style.setProperty('--chosen-color', value);
  clearTimeout(colorApplyTimer);
  colorApplyTimer = setTimeout(() => {
    colorApplyTimer = null;
    applyTextColor(value);
  }, immediate ? 0 : 80);
}

$('textColor').addEventListener('pointerdown', rememberSelection);
// Chromium/WebView variants differ on whether a native color picker emits
// `input`, `change`, or both. Listen to both and debounce duplicate events.
$('textColor').addEventListener('input', event => queueTextColor(event.target.value));
$('textColor').addEventListener('change', event => queueTextColor(event.target.value, true));
$('save').addEventListener('click', saveNow);
$('recover').addEventListener('click', () => controller?.reloadDiscardingDraft());

document.addEventListener('keydown', event => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
    event.preventDefault();
    saveNow();
  }
});

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden' && controller?.state.phase === 'ready' && controller.state.dirty && !controller.state.error) {
    saveNow();
  }
});

function loadOffice() {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Outlook could not initialize. Open Message Notes from the Apps menu on an email.')), 20000);
    const script = document.createElement('script');
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.onerror = () => {
      clearTimeout(timer);
      reject(new Error('Microsoft Office.js could not load. Check your internet connection or web filtering policy.'));
    };
    script.onload = () => {
      Office.onReady(info => {
        clearTimeout(timer);
        if (info.host !== Office.HostType.Outlook) reject(new Error('Open Message Notes from an email in Outlook.'));
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
    $('demoItem').addEventListener('change', event => adapter.change(Number(event.target.value)));
  } else {
    await loadOffice();
    if (!Office.context.requirements.isSetSupported('Mailbox', '1.5')) {
      throw new Error('This Outlook client does not support the required Mailbox 1.5 APIs.');
    }
    adapter = new OfficeBridge();
  }

  controller = new NotesController(adapter, render);
  await adapter.subscribe(() => controller.select());
  await controller.select();
}

start().catch(error => {
  $('status').textContent = 'Not connected';
  show('fatal', error.message);
  $('note').setAttribute('contenteditable', 'false');
});
