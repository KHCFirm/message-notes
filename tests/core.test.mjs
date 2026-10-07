import test from 'node:test';
import assert from 'node:assert/strict';
import { NotesController, recordFor, parseRecord, validateText, fingerprint, ConflictError, SelectionChangedError } from '../site/core.js';
const deferred = () => { let resolve,reject;const promise=new Promise((a,b)=>{resolve=a;reject=b;});return {promise,resolve,reject}; };
const tick = () => new Promise(resolve=>setImmediate(resolve));
class FakeAdapter {
  constructor() { this.item={id:'a',subject:'Same subject',sender:'A'};this.records=new Map();this.writes=[]; }
  current(){return this.item;}
  async read(item){return this.records.get(item.id)??null;}
  async write(item,record,expected){
    if(this.item?.id!==item.id)throw new SelectionChangedError();
    if(fingerprint(this.records.get(item.id))!==expected)throw new ConflictError();
    this.writes.push(item.id);this.records.set(item.id,record);return record;
  }
  choose(id){this.item=id?{id,subject:'Same subject',sender:id}:null;}
}
async function setup(){const adapter=new FakeAdapter();const c=new NotesController(adapter);await c.select();return {adapter,c};}

test('record validation and deletion format',()=>{
  assert.equal(recordFor(''),null);assert.equal(parseRecord(undefined),null);
  const r=recordFor('Call tomorrow');assert.deepEqual(parseRecord(r),r);
  assert.throws(()=>parseRecord({v:2,text:'unknown',updatedAt:'2026-10-06'}));
  assert.throws(()=>parseRecord({v:1,text:'x',updatedAt:'not a date'}));
});
test('2,200-character limit does not truncate text',()=>{
  const text='x'.repeat(2201);assert.equal(validateText(text).ok,false);assert.equal(text.length,2201);
  assert.equal(validateText('x'.repeat(2200)).ok,true);
});
test('serialized storage counts escapes, Unicode, and newlines',()=>{
  assert.equal(validateText('"'.repeat(2000)).ok,false);
  assert.equal(validateText('\n'.repeat(2000)).ok,false);
  const text='Call: \u201cOK\u201d\nNext step: \ud83d\udcdd';assert.equal(parseRecord(recordFor(text)).text,text);assert.equal(validateText(text).ok,true);
});
test('same subject on two emails still stores two independent notes',async()=>{
  const {adapter,c}=await setup();c.edit('Alpha');assert.equal(await c.save(),true);
  adapter.choose('b');await c.select();assert.equal(c.state.text,'');c.edit('Beta');await c.save();
  adapter.choose('a');await c.select();assert.equal(c.state.text,'Alpha');
  assert.equal(adapter.records.get('b').text,'Beta');
});
test('draft is restored after switching and is not silently saved',async()=>{
  const {adapter,c}=await setup();c.edit('Unsaved A');adapter.choose('b');await c.select();assert.equal(adapter.records.size,0);
  assert.equal(c.drafts.size,1);adapter.choose('a');await c.select();assert.equal(c.state.text,'Unsaved A');assert.equal(c.state.dirty,true);
});
test('clear text only removes the saved note after Save',async()=>{
  const {adapter,c}=await setup();c.edit('Keep until saved');await c.save();c.edit('');assert.equal(adapter.records.get('a').text,'Keep until saved');
  await c.save();assert.equal(adapter.records.get('a'),null);assert.equal(c.hasUnsaved(),false);
});
test('read errors disable editing and do not overwrite the mailbox',async()=>{
  const {adapter,c}=await setup();adapter.read=async()=>{throw new Error('Load failed');};await c.select();c.edit('bad overwrite');
  assert.equal(c.state.phase,'error');assert.equal(await c.save(),false);assert.equal(adapter.records.size,0);
});
test('save errors retain the draft and never claim success',async()=>{
  const {adapter,c}=await setup();adapter.write=async()=>{throw new Error('Offline');};c.edit('Pending');assert.equal(await c.save(),false);
  assert.equal(c.state.dirty,true);assert.equal(c.state.text,'Pending');assert.match(c.state.error,/Offline/);assert.equal(c.drafts.size,1);
});
test('late load callbacks do not replace a newly selected email',async()=>{
  const {adapter,c}=await setup();const delayed=deferred();adapter.read=async item=>item.id==='a'?delayed.promise:recordFor('B');
  const old=c.select();adapter.choose('b');await c.select();delayed.resolve(recordFor('A'));await old;
  assert.equal(c.state.item.id,'b');assert.equal(c.state.text,'B');
});
test('selection changes before ItemChanged arrives block saving',async()=>{
  const {adapter,c}=await setup();c.edit('This belongs to A');adapter.choose('b');assert.equal(await c.save(),false);assert.deepEqual(adapter.writes,[]);
});
test('finishing a save for A never replaces the UI or note for B',async()=>{
  const {adapter,c}=await setup();const delayed=deferred();adapter.write=async(item,record)=>{await delayed.promise;adapter.records.set(item.id,record);return record;};
  c.edit('A');const save=c.save();await tick();adapter.choose('b');await c.select();c.edit('B draft');delayed.resolve();await save;
  assert.equal(c.state.item.id,'b');assert.equal(c.state.text,'B draft');assert.equal(adapter.records.get('a').text,'A');assert.equal(adapter.records.has('b'),false);
});
test('returning to A waits for its pending save before loading it',async()=>{
  const {adapter,c}=await setup();const delayed=deferred();adapter.write=async(item,record)=>{await delayed.promise;adapter.records.set(item.id,record);return record;};
  c.edit('A');const save=c.save();await tick();adapter.choose('b');await c.select();adapter.choose('a');const back=c.select();
  assert.equal(c.state.phase,'loading');delayed.resolve();await Promise.all([save,back]);assert.equal(c.state.text,'A');assert.equal(c.state.dirty,false);
});
test('duplicate Save clicks start only one write',async()=>{
  const {adapter,c}=await setup();c.edit('One');await Promise.all([c.save(),c.save()]);assert.equal(adapter.writes.length,1);
});
test('concurrent edits are detected without overwriting the newer note',async()=>{
  const {adapter,c}=await setup();c.edit('My draft');adapter.records.set('a',recordFor('Other window'));
  assert.equal(await c.save(),false);assert.equal(c.state.conflict,true);assert.equal(adapter.records.get('a').text,'Other window');
  await c.reloadDiscardingDraft();assert.equal(c.state.text,'Other window');assert.equal(c.state.dirty,false);
});
test('a conflict remains blocked across repeated message switches',async()=>{
  const {adapter,c}=await setup();c.edit('Old draft');adapter.choose('b');await c.select();adapter.records.set('a',recordFor('Remote update'));
  adapter.choose('a');await c.select();assert.equal(c.state.conflict,true);
  adapter.choose('b');await c.select();adapter.choose('a');await c.select();assert.equal(c.state.conflict,true);assert.equal(await c.save(),false);
});
test('null selection clears the editor but keeps unsaved drafts in memory',async()=>{
  const {adapter,c}=await setup();c.edit('Keep me');adapter.choose(null);await c.select();assert.equal(c.state.phase,'empty');assert.equal(c.state.text,'');assert.equal(c.hasUnsaved(),true);
  adapter.choose('a');await c.select();assert.equal(c.state.text,'Keep me');
});
test('reloading with discard removes only the current email draft',async()=>{
  const {adapter,c}=await setup();c.edit('A draft');adapter.choose('b');await c.select();c.edit('B draft');await c.reloadDiscardingDraft();
  assert.equal(c.state.text,'');assert.equal(c.drafts.has('a'),true);assert.equal(c.drafts.has('b'),false);
});
test('a failed reload retains the unsaved text for copying',async()=>{
  const {adapter,c}=await setup();c.edit('Important draft');adapter.read=async()=>{throw new Error('Disconnected');};await c.select();
  assert.equal(c.state.text,'Important draft');assert.equal(c.state.phase,'error');assert.equal(c.hasUnsaved(),true);
});
