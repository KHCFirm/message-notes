/** Pure note model; shared by the real Outlook adapter and automated tests. */
export const PROPERTY = 'messageNote';
export const MAX_TEXT = 2200;
export const MAX_STORAGE = 2500;
const PREVIEW_DATE = '2099-12-31T23:59:59.999Z';

export function recordFor(text, updatedAt = new Date().toISOString()) {
  return text === '' ? null : { v: 1, text, updatedAt };
}

export function parseRecord(value) {
  if (value == null) return null;
  if (typeof value !== 'object' || value.v !== 1 || typeof value.text !== 'string' ||
      typeof value.updatedAt !== 'string' || Number.isNaN(Date.parse(value.updatedAt))) {
    throw new Error('This note has an unrecognized format. It has not been changed. Contact your administrator.');
  }
  return { v: 1, text: value.text, updatedAt: value.updatedAt };
}

export const fingerprint = value => JSON.stringify(value ?? null);

export function validateText(text) {
  const size = JSON.stringify(text === '' ? {} : { [PROPERTY]: recordFor(text, PREVIEW_DATE) }).length;
  if (text.length > MAX_TEXT) return { ok: false, size, message: `This note is too long. Use ${MAX_TEXT.toLocaleString()} characters or fewer. Your text has not been truncated.` };
  if (size > MAX_STORAGE) return { ok: false, size, message: 'This note exceeds Outlook\'s storage limit. Shorten it; quotes, new lines, and some special characters use extra space.' };
  return { ok: true, size, message: '' };
}

export class SelectionChangedError extends Error {
  constructor() { super('The selected email changed. Return to the original email and save again.'); this.name = 'SelectionChangedError'; }
}
export class ConflictError extends Error {
  constructor() { super('The mailbox note changed in another window. Copy your text, then use Reload to review the saved version before making changes.'); this.name = 'ConflictError'; }
}

/**
 * Adapter contract: current() -> item|null; read(item) -> record|null;
 * write(item, record, expectedFingerprint) -> Promise<record|null>.
 * Drafts deliberately stay in memory, never browser storage.
 */
export class NotesController {
  constructor(adapter, onChange = () => {}) {
    this.adapter = adapter;
    this.onChange = onChange;
    this.generation = 0;
    this.drafts = new Map();
    this.pending = new Map();
    this.state = this.blank();
  }
  blank() { return { item: null, phase: 'empty', text: '', record: null, dirty: false, error: '', conflict: false, message: '' }; }
  emit() {
    this.onChange({ ...this.state, validation: validateText(this.state.text),
      otherDrafts: [...this.drafts].filter(([id]) => id !== this.state.item?.id).map(([, d]) => d.subject) });
  }
  captureDraft() {
    const s = this.state;
    if (!s.item || (s.phase !== 'ready' && s.phase !== 'saving')) return;
    if (s.dirty) this.drafts.set(s.item.id, { text: s.text, base: s.conflict && this.drafts.has(s.item.id) ? this.drafts.get(s.item.id).base : fingerprint(s.record), subject: s.item.subject });
    else this.drafts.delete(s.item.id);
  }
  async select() {
    this.captureDraft();
    const gen = ++this.generation;
    const item = this.adapter.current();
    this.state = { ...this.blank(), item, phase: item ? 'loading' : 'empty' };
    this.emit();
    if (!item) return;
    // Returning to an item with a save already in flight must wait for that save.
    if (this.pending.has(item.id)) await this.pending.get(item.id).catch(() => {});
    if (gen !== this.generation) return;
    try {
      const record = await this.adapter.read(item);
      if (gen !== this.generation) return;
      const draft = this.drafts.get(item.id);
      const changed = !!draft && draft.base !== fingerprint(record);
      const text = draft ? draft.text : record?.text ?? '';
      this.state = { item, phase: 'ready', record, text,
        dirty: text !== (record?.text ?? ''), conflict: changed,
        error: changed ? new ConflictError().message : '',
        message: draft ? 'Unsaved draft restored. Save before closing this pane.' : '' };
      if (!this.state.dirty) { this.drafts.delete(item.id); this.state.conflict = false; this.state.error = ''; }
    } catch (error) {
      if (gen !== this.generation) return;
      const draft = this.drafts.get(item.id);
      this.state = { ...this.state, phase: 'error', text: draft?.text ?? '', error: error.message, dirty: !!draft };
    }
    this.emit();
  }
  edit(text) {
    if (this.state.phase !== 'ready') return;
    this.state.text = text;
    this.state.dirty = text !== (this.state.record?.text ?? '');
    if (!this.state.conflict) this.state.error = '';
    this.state.message = '';
    this.captureDraft();
    this.emit();
  }
  async save() {
    const s = this.state;
    if (s.phase !== 'ready' || !s.dirty || s.conflict || !validateText(s.text).ok) return false;
    if (this.adapter.current()?.id !== s.item.id) {
      s.error = new SelectionChangedError().message; this.emit(); return false;
    }
    const gen = this.generation;
    const item = s.item;
    const text = s.text;
    const record = recordFor(text);
    const expected = fingerprint(s.record);
    this.captureDraft();
    s.phase = 'saving'; s.error = ''; s.message = ''; this.emit();
    const operation = Promise.resolve().then(() => this.adapter.write(item, record, expected));
    this.pending.set(item.id, operation);
    try {
      const saved = await operation;
      if (this.drafts.get(item.id)?.text === text) this.drafts.delete(item.id);
      if (gen === this.generation) this.state = { ...s, phase: 'ready', record: saved, dirty: false,
        conflict: false, error: '', message: saved ? 'Saved to this email.' : 'Note removed from this email.' };
      return true;
    } catch (error) {
      if (gen === this.generation) this.state = { ...s, phase: 'ready', error: error.message, conflict: error instanceof ConflictError };
      return false;
    } finally {
      if (this.pending.get(item.id) === operation) this.pending.delete(item.id);
      this.emit();
    }
  }
  async reloadDiscardingDraft() {
    if (this.state.phase === 'saving') return;
    if (this.state.item) this.drafts.delete(this.state.item.id);
    this.state.dirty = false;
    await this.select();
  }
  hasUnsaved() { return this.state.dirty || this.drafts.size > 0 || this.pending.size > 0; }
}
