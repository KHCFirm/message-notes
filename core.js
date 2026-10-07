/** Pure note model; shared by the real Outlook adapter and automated tests. */
export const PROPERTY = 'messageNote';
export const MAX_STORAGE = 2500;
const PREVIEW_DATE = '2099-12-31T23:59:59.999Z';

function escapeHtml(text) {
  return String(text)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
    .replace(/\r\n?|\n/g, '<br>');
}

export function recordFor(html, updatedAt = new Date().toISOString()) {
  return html === '' ? null : { v: 2, html, updatedAt };
}

export function parseRecord(value) {
  if (value == null) return null;

  // Backward compatibility with the original plain-text format.
  if (typeof value === 'object' && value.v === 1 && typeof value.text === 'string' &&
      typeof value.updatedAt === 'string' && !Number.isNaN(Date.parse(value.updatedAt))) {
    return { v: 2, html: escapeHtml(value.text), updatedAt: value.updatedAt };
  }

  if (typeof value !== 'object' || value.v !== 2 || typeof value.html !== 'string' ||
      typeof value.updatedAt !== 'string' || Number.isNaN(Date.parse(value.updatedAt))) {
    throw new Error('This note has an unrecognized format. It has not been changed. Contact your administrator.');
  }
  return { v: 2, html: value.html, updatedAt: value.updatedAt };
}

export const fingerprint = value => JSON.stringify(value ?? null);

export function validateHtml(html) {
  const size = JSON.stringify(html === '' ? {} : { [PROPERTY]: recordFor(html, PREVIEW_DATE) }).length;
  if (size > MAX_STORAGE) {
    return {
      ok: false,
      size,
      message: 'This formatted note is too large for Outlook\'s per-message note storage. Shorten it or remove some formatting, then save again.'
    };
  }
  return { ok: true, size, message: '' };
}

export class SelectionChangedError extends Error {
  constructor() { super('The selected email changed before the note finished saving. Return to that email and save again.'); this.name = 'SelectionChangedError'; }
}
export class ConflictError extends Error {
  constructor() { super('This note changed in another window. Reload the saved copy before making more changes.'); this.name = 'ConflictError'; }
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

  blank() {
    return { item: null, phase: 'empty', html: '', record: null, dirty: false, error: '', conflict: false, message: '' };
  }

  emit() {
    this.onChange({
      ...this.state,
      validation: validateHtml(this.state.html),
      otherDrafts: [...this.drafts]
        .filter(([id]) => id !== this.state.item?.id)
        .map(([, draft]) => draft.subject)
    });
  }

  captureDraft() {
    const s = this.state;
    if (!s.item || (s.phase !== 'ready' && s.phase !== 'saving')) return;
    if (s.dirty) {
      const previous = this.drafts.get(s.item.id);
      this.drafts.set(s.item.id, {
        html: s.html,
        base: s.conflict && previous ? previous.base : fingerprint(s.record),
        subject: s.item.subject
      });
    } else {
      this.drafts.delete(s.item.id);
    }
  }

  async select() {
    this.captureDraft();
    const gen = ++this.generation;
    const item = this.adapter.current();
    this.state = { ...this.blank(), item, phase: item ? 'loading' : 'empty' };
    this.emit();
    if (!item) return;

    // Returning to an item with an autosave already in flight waits for that save.
    if (this.pending.has(item.id)) await this.pending.get(item.id).catch(() => {});
    if (gen !== this.generation) return;

    try {
      const record = await this.adapter.read(item);
      if (gen !== this.generation) return;
      const draft = this.drafts.get(item.id);
      const changed = !!draft && draft.base !== fingerprint(record);
      const html = draft ? draft.html : record?.html ?? '';
      this.state = {
        item,
        phase: 'ready',
        record,
        html,
        dirty: html !== (record?.html ?? ''),
        conflict: changed,
        error: changed ? new ConflictError().message : '',
        message: ''
      };
      if (!this.state.dirty) {
        this.drafts.delete(item.id);
        this.state.conflict = false;
        this.state.error = '';
      }
    } catch (error) {
      if (gen !== this.generation) return;
      const draft = this.drafts.get(item.id);
      this.state = {
        ...this.state,
        phase: 'error',
        html: draft?.html ?? '',
        error: error.message,
        dirty: !!draft
      };
    }
    this.emit();
  }

  edit(html) {
    if (this.state.phase !== 'ready' && this.state.phase !== 'saving') return;
    this.state.html = html;
    this.state.dirty = html !== (this.state.record?.html ?? '');
    if (!this.state.conflict) this.state.error = '';
    this.state.message = '';
    this.captureDraft();
    this.emit();
  }

  async save() {
    const s = this.state;
    if (s.phase === 'saving') return false;
    if (s.phase !== 'ready' || !s.dirty || s.conflict || !validateHtml(s.html).ok) return false;
    if (this.adapter.current()?.id !== s.item.id) {
      s.error = new SelectionChangedError().message;
      this.emit();
      return false;
    }

    const gen = this.generation;
    const item = s.item;
    const html = s.html;
    const record = recordFor(html);
    const expected = fingerprint(s.record);
    this.captureDraft();
    s.phase = 'saving';
    s.error = '';
    s.message = '';
    this.emit();

    const operation = Promise.resolve().then(() => this.adapter.write(item, record, expected));
    this.pending.set(item.id, operation);

    try {
      const saved = await operation;
      const draft = this.drafts.get(item.id);
      if (draft?.html === html) {
        this.drafts.delete(item.id);
      } else if (draft) {
        // The user kept typing while the previous snapshot was saving. Rebase that
        // newer draft onto the just-saved mailbox record so the next autosave is safe.
        draft.base = fingerprint(saved);
        this.drafts.set(item.id, draft);
      }

      if (gen === this.generation && this.state.item?.id === item.id) {
        const currentHtml = this.state.html;
        const dirty = currentHtml !== (saved?.html ?? '');
        this.state = {
          ...this.state,
          phase: 'ready',
          record: saved,
          dirty,
          conflict: false,
          error: '',
          message: ''
        };
        if (dirty) {
          this.drafts.set(item.id, {
            html: currentHtml,
            base: fingerprint(saved),
            subject: item.subject
          });
        } else {
          this.drafts.delete(item.id);
        }
      }
      return true;
    } catch (error) {
      if (gen === this.generation && this.state.item?.id === item.id) {
        this.state = {
          ...this.state,
          phase: 'ready',
          error: error.message,
          conflict: error instanceof ConflictError,
          dirty: true
        };
        this.captureDraft();
      }
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

  hasUnsaved() {
    return this.state.dirty || this.drafts.size > 0 || this.pending.size > 0;
  }
}
