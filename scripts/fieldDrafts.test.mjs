import assert from 'node:assert/strict';
import {createFieldDrafts,createTokenWriteQueue} from '../fieldDrafts.js';
const drafts=createFieldDrafts(),old=drafts.set('one','attribute:Strength','12'),newer=drafts.set('one','attribute:Strength','16'),other=drafts.set('two','attribute:Strength','7');
drafts.finish('one','attribute:Strength',old);assert.equal(drafts.get('one','attribute:Strength').value,'16');drafts.finish('one','attribute:Strength',newer);assert.equal(drafts.get('one','attribute:Strength'),undefined);assert.equal(drafts.get('two','attribute:Strength'),other);drafts.clear('two');assert.equal(drafts.get('two','attribute:Strength'),undefined);
const enqueue=createTokenWriteQueue(),events=[];let release;const gate=new Promise(resolve=>release=resolve);
const a=enqueue('one',async()=>{events.push('one start');await gate;events.push('one end');}),b=enqueue('one',async()=>events.push('one second')),c=enqueue('two',async()=>events.push('two independent'));
await c;assert.deepEqual(events,['one start','two independent']);release();await Promise.all([a,b]);assert.deepEqual(events,['one start','two independent','one end','one second']);
await assert.rejects(enqueue('one',async()=>{throw new Error('save failed');}));assert.equal(await enqueue('one',async()=>42),42);
console.log('Per-token field drafts, newer-edit protection, write order and failure recovery passed.');
