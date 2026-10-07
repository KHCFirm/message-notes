import { PROPERTY, parseRecord, fingerprint, validateText, SelectionChangedError, ConflictError } from './core.js';

function asPromise(run, timeout = 30000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Outlook did not respond. Your draft is still in this pane. Reload the email to check the mailbox before retrying.')), timeout);
    try {
      run(result => {
        clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error?.message || 'Outlook could not complete the operation. Check your connection and mailbox permissions.'));
      });
    } catch (error) { clearTimeout(timer); reject(error); }
  });
}

export class OfficeBridge {
  current() {
    const item = Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || typeof item.subject !== 'string' || !item.itemId) return null;
    return { id: item.itemId, subject: item.subject || '(No subject)',
      sender: item.from?.displayName || item.from?.emailAddress || item.sender?.displayName || '', handle: item };
  }
  assertSelected(item) { if (this.current()?.id !== item.id) throw new SelectionChangedError(); }
  async load(item) {
    this.assertSelected(item);
    const properties = await asPromise(done => item.handle.loadCustomPropertiesAsync(done));
    this.assertSelected(item);
    return properties;
  }
  async read(item) {
    const properties = await this.load(item);
    return parseRecord(properties.get(PROPERTY));
  }
  async write(item, record, expected) {
    const validation = validateText(record?.text ?? '');
    if (!validation.ok) throw new Error(validation.message);
    // Refresh before writing: detect most concurrent edits and avoid a stale property bag.
    // Office.js provides no atomic compare-and-swap; see the documented limitation.
    const properties = await this.load(item);
    if (fingerprint(parseRecord(properties.get(PROPERTY))) !== expected) throw new ConflictError();
    if (record) properties.set(PROPERTY, record);
    else properties.remove(PROPERTY);
    // No await between the final selection check and starting the mailbox write.
    this.assertSelected(item);
    await asPromise(done => properties.saveAsync(done));
    return record;
  }
  async subscribe(handler) {
    await asPromise(done => Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler, done));
  }
}
