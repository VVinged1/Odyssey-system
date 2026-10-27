import {distanceMeters} from './rangeRules.js';
// Uses Owlbear's own grid measurement, including hex grids and diagonal rules.
export function distanceLabel(cells,scale) {
  if(!Number.isFinite(cells)||cells<0)throw new Error('Distance is unavailable.');
  const digits=Math.max(0,Math.min(5,Number(scale?.parsed?.digits)||0)),multiplier=Number(scale?.parsed?.multiplier)||1;
  const format=(n,d,min=0)=>Number(n.toFixed(d)).toLocaleString('en-US',{minimumFractionDigits:min,maximumFractionDigits:d,useGrouping:false});
  const unit=String(scale?.parsed?.unit||'');
  return format(cells,2)+' cells'+(unit?' · '+format(cells*multiplier,digits,digits)+unit:'');
}
export function createTargetDistance({OBR,buildRuler,onRead,context='main'}) {
  const key='com.codex.body-hp/target-distance',id=crypto.randomUUID();
  let request=0,generation=0,timer=null,latest=null,added=false,chain=Promise.resolve(),scalePromise=null;
  const cache=new Map();
  const initialize=OBR.scene.local.getItems().then(items=>{const ids=items.filter(i=>i.metadata?.[key]===context).map(i=>i.id);if(ids.length)return OBR.scene.local.deleteItems(ids);}).catch(()=>{});
  const center=async target=>target.id?(await OBR.scene.items.getItemBounds([target.id])).center:target;
  async function read(source,target) {
    const cacheKey=JSON.stringify([source.id,source.position,source.scale,source.rotation,source.grid,source.image?.width,source.image?.height,target.id,target.position||target,target.scale,target.rotation,target.grid,target.image?.width,target.image?.height]);
    if(!cache.has(cacheKey)) {
      if(cache.size>=32)cache.clear();
      const promise=(async()=>{
        const [from,to,scale]=await Promise.all([center(source),center(target),scalePromise??=(OBR.scene.grid.getScale())]);
        const cells=await OBR.scene.grid.getDistance(from,to);return {sourceId:source.id,targetId:target.id||'',from,to,cells,meters:distanceMeters(cells,scale),label:distanceLabel(cells,scale)};
      })();cache.set(cacheKey,promise);promise.catch(()=>{if(cache.get(cacheKey)===promise)cache.delete(cacheKey);});
    }
    return cache.get(cacheKey);
  }
  const publish=async(source,target,preview,epoch)=>{
    const current=++request;
    try {
      const reading=await read(source,target);if(epoch!==generation||current!==request)return preview?null:reading;
      onRead(reading);
      if(preview) {
        chain=chain.catch(()=>{}).then(async()=>{
          await initialize;if(epoch!==generation||current!==request)return;
          if(!added) {
            const ruler=buildRuler().id(id).name('Odyssey target distance').startPosition(reading.from).endPosition(reading.to).measurement(reading.label).variant('DASHED').layer('RULER').locked(true).disableHit(true).metadata({[key]:context}).build();
            await OBR.scene.local.addItems([ruler]);added=true;
          }else await OBR.scene.local.updateItems([id],items=>{if(epoch===generation&&current===request&&items[0]){items[0].startPosition=reading.from;items[0].endPosition=reading.to;items[0].measurement=reading.label;}},true);
        });await chain;
      }
      return reading;
    }catch{if(epoch===generation&&current===request)onRead(null);return null;}
  };
  return {
    preview(source,target){latest={source,target};if(timer)return;const epoch=generation;timer=setTimeout(()=>{timer=null;const next=latest;if(next)void publish(next.source,next.target,true,epoch);},80);},
    measure(source,target){return publish(source,target,false,generation);},
    invalidate(){generation++;request++;cache.clear();scalePromise=null;if(timer){clearTimeout(timer);timer=null;}if(latest)this.preview(latest.source,latest.target);},
    async clear(){generation++;request++;latest=null;if(timer){clearTimeout(timer);timer=null;}chain=chain.catch(()=>{}).then(async()=>{await initialize;if(added){await OBR.scene.local.deleteItems([id]);added=false;}});await chain.catch(()=>{});},
  };
}
