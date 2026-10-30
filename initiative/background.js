import OBR from '@owlbear-rodeo/sdk';
import {installInitiativeMenus,initiativeIconUrls} from './contextMenu.js';
OBR.onReady(()=>{void installInitiativeMenus(OBR,initiativeIconUrls(window.location.href)).catch(error=>{console.error('[Odyssey Initiative] Context menu registration failed',error);void OBR.notification.show('Odyssey Initiative · menu registration failed: '+error.message,'ERROR');});});
