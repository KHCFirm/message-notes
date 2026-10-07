import { fingerprint, SelectionChangedError, ConflictError } from './core.js';
/** Explicit demo only. Fake email data; memory only; never calls Outlook. */
export class DemoBridge {
  constructor() {
    this.items = [
      { id: 'sample-1', subject: 'Updated documents for review', sender: 'Alex Morgan' },
      { id: 'sample-2', subject: 'Friday meeting follow-up', sender: 'Jordan Lee' },
      { id: 'sample-3', subject: 'Updated documents for review', sender: 'Sam Rivera' }
    ];
    this.index = 0;
    this.records = new Map();
  }
  current() { return this.items[this.index] ?? null; }
  async read(item) { await new Promise(r => setTimeout(r, 100)); if (this.current()?.id !== item.id) throw new SelectionChangedError(); return this.records.get(item.id) ?? null; }
  async write(item, record, expected) {
    if (this.current()?.id !== item.id) throw new SelectionChangedError();
    if (fingerprint(this.records.get(item.id)) !== expected) throw new ConflictError();
    await new Promise(r => setTimeout(r, 180));
    if (record) this.records.set(item.id, structuredClone(record)); else this.records.delete(item.id);
    return record;
  }
  async subscribe(handler) { this.handler = handler; }
  change(index) { this.index = index; this.handler?.(); }
}
