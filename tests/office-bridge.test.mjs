import test from 'node:test';
import assert from 'node:assert/strict';
import { OfficeBridge } from '../site/office-bridge.js';
import { PROPERTY,recordFor,fingerprint,ConflictError,SelectionChangedError } from '../site/core.js';
function host() {
  const db=new Map();const calls=[];
  function makeItem(id){return {itemId:id,subject:'Test email',from:{displayName:'Test sender'},itemType:'message',loadCustomPropertiesAsync(callback){
    const snapshot=structuredClone(db.get(id)||{});
    calls.push(['load',id]);
    queueMicrotask(()=>callback({status:'succeeded',value:{get:key=>snapshot[key],set:(key,value)=>{snapshot[key]=value;},remove:key=>delete snapshot[key],saveAsync(done){calls.push(['save',id]);db.set(id,structuredClone(snapshot));queueMicrotask(()=>done({status:'succeeded'}));}}}));
  }};}
  globalThis.Office={AsyncResultStatus:{Succeeded:'succeeded'},MailboxEnums:{ItemType:{Message:'message'}},EventType:{ItemChanged:'changed'},context:{mailbox:{item:makeItem('a'),addHandlerAsync(type,handler,cb){calls.push(['subscribe',type]);cb({status:'succeeded'});}}}};
  return {db,calls,makeItem,bridge:new OfficeBridge()};
}
test('Office adapter writes and reads item custom properties, not the email body',async()=>{
  const {bridge,db,calls}=host();const item=bridge.current();const record=recordFor('Mailbox note');
  await bridge.write(item,record,fingerprint(null));assert.equal(db.get('a')[PROPERTY].text,'Mailbox note');assert.deepEqual(await bridge.read(item),record);assert.equal(calls.some(c=>c[0]==='save'),true);
});
test('Office adapter deletes only its own property',async()=>{
  const {bridge,db}=host();const r=recordFor('Remove');db.set('a',{[PROPERTY]:r,otherSetting:'keep'});
  await bridge.write(bridge.current(),null,fingerprint(r));assert.deepEqual(db.get('a'),{otherSetting:'keep'});
});
test('Office adapter preflight blocks concurrent changes',async()=>{
  const {bridge,db,calls}=host();db.set('a',{[PROPERTY]:recordFor('Remote')});
  await assert.rejects(bridge.write(bridge.current(),recordFor('Local'),fingerprint(null)),ConflictError);
  assert.equal(calls.some(c=>c[0]==='save'),false);
});
test('Office adapter rejects an item that is no longer selected',async()=>{
  const {bridge,makeItem,calls}=host();const a=bridge.current();Office.context.mailbox.item=makeItem('b');
  await assert.rejects(bridge.write(a,recordFor('Never to B'),fingerprint(null)),SelectionChangedError);assert.equal(calls.length,0);
});
test('Office adapter rejects a switch during asynchronous property loading',async()=>{
  const {bridge,makeItem,calls}=host();const item=bridge.current();const write=bridge.write(item,recordFor('A'),fingerprint(null));Office.context.mailbox.item=makeItem('b');
  await assert.rejects(write,SelectionChangedError);assert.equal(calls.some(c=>c[0]==='save'),false);
});
test('Office adapter handles no selected message and registers ItemChanged',async()=>{
  const {bridge,calls}=host();await bridge.subscribe(()=>{});Office.context.mailbox.item=null;assert.equal(bridge.current(),null);assert.deepEqual(calls,[['subscribe','changed']]);
});
