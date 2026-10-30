import {useEffect,useRef,useState} from 'react';
import OBR from '@owlbear-rodeo/sdk';
const key='odyssey-initiative/window-size',defaults={width:760,height:116};
export function useInitiativeSize(){
 const limits=useRef({width:1200,height:500}),start=useRef<{x:number;y:number;width:number;height:number}|null>(null),pending=useRef<number|null>(null);
 const clamp=(value:{width:number;height:number})=>({width:Math.round(Math.min(limits.current.width,Math.max(180,Number(value.width)||defaults.width))),height:Math.round(Math.min(limits.current.height,Math.max(110,Number(value.height)||defaults.height)))});
 const [size,setSize]=useState(()=>{try{return clamp(JSON.parse(localStorage.getItem(key)||'null')||defaults);}catch{return defaults;}}),current=useRef(size);
 const apply=(value:typeof size)=>{current.current=clamp(value);setSize(current.current);void Promise.all([OBR.action.setWidth(current.current.width),OBR.action.setHeight(current.current.height)]).catch(console.error);};
 const save=()=>{try{localStorage.setItem(key,JSON.stringify(current.current));}catch{}};
 useEffect(()=>{let alive=true;void Promise.all([OBR.viewport.getWidth(),OBR.viewport.getHeight()]).then(([width,height])=>{if(!alive)return;limits.current={width:Math.max(180,Math.min(1200,width-24)),height:Math.max(110,Math.min(500,height-24))};apply(current.current);});return()=>{alive=false;if(pending.current!==null)cancelAnimationFrame(pending.current);};},[]);
 const handle=<button type="button" className="initiative-resize" aria-label="Resize initiative" title="Drag to resize" onPointerDown={event=>{event.currentTarget.setPointerCapture(event.pointerId);start.current={x:event.screenX,y:event.screenY,...current.current};event.preventDefault();}}
  onPointerMove={event=>{if(!start.current)return;const value=clamp({width:start.current.width+event.screenX-start.current.x,height:start.current.height+event.screenY-start.current.y});current.current=value;if(pending.current===null)pending.current=requestAnimationFrame(()=>{pending.current=null;apply(current.current);});}}
  onPointerUp={event=>{if(!start.current)return;start.current=null;if(event.currentTarget.hasPointerCapture(event.pointerId))event.currentTarget.releasePointerCapture(event.pointerId);apply(current.current);save();}}
  onPointerCancel={()=>{start.current=null;save();}}
  onKeyDown={event=>{const change=({ArrowLeft:[-16,0],ArrowRight:[16,0],ArrowUp:[0,-12],ArrowDown:[0,12]} as Record<string,number[]>)[event.key];if(change){event.preventDefault();apply({width:current.current.width+change[0],height:current.current.height+change[1]});save();}}}>◢</button>;
 return {size,handle};
}
