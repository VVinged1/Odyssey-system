import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const root=path.resolve(fileURLToPath(new URL('..',import.meta.url)));
const mainSource=await fs.readFile(process.env.ODYSSEY_BASELINE_MAIN?path.resolve(root,process.env.ODYSSEY_BASELINE_MAIN):path.join(root,'main.js'),'utf8');
const bundle=await build({stdin:{contents:mainSource,resolveDir:root,sourcefile:'main.js'},bundle:true,format:'esm',platform:'browser',write:false,alias:{'@owlbear-rodeo/sdk':path.join(root,'scripts/fixtures/mockSdk.js')}});
const gmBundle=await build({entryPoints:[path.join(root,'gm-extension/main.js')],bundle:true,format:'esm',platform:'browser',write:false,alias:{'@owlbear-rodeo/sdk':path.join(root,'scripts/fixtures/mockSdk.js')}});
const server=http.createServer(async(req,res)=>{try {const url=new URL(req.url,'http://localhost');if(url.pathname==='/gm-extension/assets/main.js'){res.setHeader('Content-Type','text/javascript');res.end(gmBundle.outputFiles[0].text);return;}if(url.pathname==='/assets/main.js'){res.setHeader('Content-Type','text/javascript');res.end(bundle.outputFiles[0].text);return;}const file=path.resolve(root,'.'+decodeURIComponent(url.pathname==='/'?'/index.html':url.pathname));if(!file.startsWith(root+path.sep)){res.writeHead(403);res.end();return;}res.setHeader('Content-Type',file.endsWith('.css')?'text/css':file.endsWith('.html')?'text/html':'application/octet-stream');res.end(await fs.readFile(file));}catch{res.writeHead(404);res.end();}});
await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
const browser=await chromium.launch(process.env.PLAYWRIGHT_CHANNEL?{channel:process.env.PLAYWRIGHT_CHANNEL}:{});

const pages=[],errors=[];let releaseHeld=[];
try{
 const gm1=await browser.newPage({viewport:{width:1000,height:1100}}),gm2=await browser.newPage({viewport:{width:1000,height:1100}});pages.push(gm1,gm2);
 for(const page of pages){page.setDefaultTimeout(10000);page.on('pageerror',e=>errors.push(e.message));await page.goto(`http://127.0.0.1:${server.address().port}/`);await page.locator('[data-action="set-odyssey-attribute"]').first().waitFor();}
 await gm2.evaluate(()=>window.odysseyTest.select('enemy'));await gm2.waitForFunction(()=>document.querySelector('#selectedTokenPanel').dataset.tokenId==='enemy');
 const authoritative=new Map((await gm1.evaluate(()=>window.odysseyTest.state.items)).map(item=>[item.id,item]));let holdHero=false;
 for(const [index,page]of pages.entries()){
  await page.exposeFunction('publishOdysseyItems',async rows=>{
   if(holdHero&&rows.some(row=>row.id==='hero'))await new Promise(resolve=>releaseHeld.push(resolve));
   for(const row of rows)authoritative.set(row.id,row);const snapshot=[...authoritative.values()];
   await Promise.all(pages.map(view=>view.evaluate(items=>{window.odysseyTest.state.items=structuredClone(items);window.odysseyTest.refresh();},snapshot)));
  });
  await page.evaluate(index=>{window.odysseyTest.state.name='GM '+(index+1);window.odysseyTest.state.afterItemWrite=async rows=>{const hero=rows.some(row=>row.id==='hero');if(hero)window.odysseyTest.state.heroWritesInFlight=(window.odysseyTest.state.heroWritesInFlight||0)+1;try{await window.publishOdysseyItems(rows);}finally{if(hero)window.odysseyTest.state.heroWritesInFlight--;}};},index);
 }
 const open=async(page,text)=>{const summary=page.locator('summary').filter({hasText:text}).first();if(!await summary.locator('..').evaluate(e=>e.open))await summary.click();};
 await open(gm1,'Body Parts');await gm1.locator('[data-action="open-body-part-add"]').click();await gm1.locator('[data-body-field="new-name"]').fill('Extra arm draft');await gm1.locator('[data-body-field="new-max"]').fill('8');await gm1.locator('[data-body-field="new-armor"]').focus();
 await gm2.locator('[data-action="set-odyssey-attribute"][data-attribute="Strength"]').fill('7');await gm2.locator('[data-action="set-odyssey-attribute"][data-attribute="Agility"]').focus();await gm2.waitForFunction(()=>window.odysseyTest.data('enemy').odyssey.attributes.Strength===7);
 await gm1.waitForFunction(()=>window.odysseyTest.data('enemy').odyssey.attributes.Strength===7);assert.equal(await gm1.locator('[data-body-field="new-name"]').inputValue(),'Extra arm draft');assert.equal(await gm1.locator('[data-body-field="new-max"]').inputValue(),'8');
 // Hold a real save acknowledgement while the other GM updates a different token.
 holdHero=true;await gm1.evaluate(()=>{const input=document.querySelector('[data-action="set-odyssey-attribute"][data-attribute="Strength"]');input.value='18';input.dispatchEvent(new Event('input',{bubbles:true}));});await gm1.locator('[data-action="set-odyssey-attribute"][data-attribute="Agility"]').focus();await gm1.waitForFunction(()=>window.odysseyTest.state.heroWritesInFlight>0);
 await gm2.locator('[data-action="set-odyssey-attribute"][data-attribute="Endurance"]').fill('9');await gm2.locator('[data-action="set-odyssey-attribute"][data-attribute="Agility"]').focus();await gm1.waitForFunction(()=>window.odysseyTest.data('enemy').odyssey.attributes.Endurance===9);
 assert.equal(await gm1.locator('[data-action="set-odyssey-attribute"][data-attribute="Strength"]').inputValue(),'18');assert.equal(await gm1.locator('[data-body-field="new-name"]').inputValue(),'Extra arm draft');
 holdHero=false;for(const release of releaseHeld.splice(0))release();await gm1.waitForFunction(()=>!window.odysseyTest.state.heroWritesInFlight);await gm2.waitForFunction(()=>window.odysseyTest.data('hero').odyssey.attributes.Strength===18);
 assert.equal(await gm2.locator('[data-action="set-odyssey-attribute"][data-attribute="Strength"]').inputValue(),'7');assert.equal(authoritative.get('hero').metadata['com.codex.body-hp/data'].odyssey.attributes.Strength,18);assert.equal(authoritative.get('enemy').metadata['com.codex.body-hp/data'].odyssey.attributes.Strength,7);
 // Switching tokens must not carry another character's pending value into the new form.
 await gm1.evaluate(()=>window.odysseyTest.select('enemy'));await gm1.waitForFunction(()=>document.querySelector('#selectedTokenPanel').dataset.tokenId==='enemy');assert.equal(await gm1.locator('[data-action="set-odyssey-attribute"][data-attribute="Strength"]').inputValue(),'7');assert.deepEqual(errors,[]);
 console.log('Two GM pages passed: different characters, unfocused form draft, delayed save, remote updates and token isolation.');
}catch(error){for(const page of pages)console.error(await page.evaluate(()=>({status:document.querySelector('#statusBox')?.textContent,token:document.querySelector('#selectedTokenPanel')?.dataset.tokenId})).catch(()=>({})));throw error;}
finally{for(const release of releaseHeld)release();for(const page of pages)await page.evaluate(()=>{if(window.odysseyTest)window.odysseyTest.state.afterItemWrite=null;}).catch(()=>{});await browser.close();server.closeAllConnections();await new Promise(resolve=>server.close(resolve));}
