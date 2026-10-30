import fs from 'node:fs';import vm from 'node:vm';import assert from 'node:assert/strict';
import {initiativeIconUrls,installInitiativeMenus} from '../initiative/contextMenu.js';
// Exercise the installed SDK's actual icon normalization and context-menu transport.
const context=vm.createContext({window:{location:{origin:'https://vvinged1.github.io'}},Promise});
vm.runInContext(fs.readFileSync(new URL('../node_modules/@owlbear-rodeo/sdk/lib/common/normalize.js',import.meta.url),'utf8').replace(/^export /gm,''),context);
vm.runInContext(fs.readFileSync(new URL('../node_modules/@owlbear-rodeo/sdk/lib/api/ContextMenuApi.js',import.meta.url),'utf8').replace(/^import .*;\r?$/gm,'').replace('export default ContextMenuApi;','globalThis.ContextMenuApi=ContextMenuApi;'),context);
const malformed=context.normalizeUrl('data:image/svg+xml,<svg/>');assert.throws(()=>new URL(malformed),'Regression: the old inline icon becomes an invalid SDK URL.');
const icons=initiativeIconUrls('https://vvinged1.github.io/Odyssey-system/initiative/background.html?v=1.3.11');assert.deepEqual(icons,{add:'https://vvinged1.github.io/Odyssey-system/add.svg',remove:'https://vvinged1.github.io/Odyssey-system/remove.svg'});
const messages=[],listeners=new Map();const transport={on:(name,callback)=>listeners.set(name,callback),sendAsync:async(name,data)=>{messages.push({name,data});return {};}};
const menu=new context.ContextMenuApi(transport),sdk={contextMenu:menu};await installInitiativeMenus(sdk,icons);
const registrations=messages.filter(m=>m.name==='OBR_CONTEXT_MENU_CREATE');assert.equal(registrations.length,2);for(const registration of registrations){const icon=registration.data.icons[0];assert.equal(new URL(icon.icon).protocol,'https:');assert(icon.label.includes('Odyssey Initiative'));assert.equal(typeof menu.contextMenus[registration.data.id].onClick,'function');}
assert.equal(registrations[0].data.icons[0].icon,icons.add);assert.equal(registrations[1].data.icons[0].icon,icons.remove);
assert.equal(initiativeIconUrls('http://127.0.0.1:1234/initiative/background.html').add,'http://127.0.0.1:1234/add.svg');
console.log('Real SDK transport passed: old data-URL failure reproduced; both menus now send valid absolute HTTP(S) icon URLs.');
