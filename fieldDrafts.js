// Per-token drafts survive unrelated scene updates until their own save completes.
export function createFieldDrafts() {
 const tokens=new Map();let revision=0;
 return {
  set(tokenId,key,value){if(!tokens.has(tokenId))tokens.set(tokenId,new Map());const draft={value,revision:++revision};tokens.get(tokenId).set(key,draft);return draft;},
  get(tokenId,key){return tokens.get(tokenId)?.get(key);},
  finish(tokenId,key,draft){const rows=tokens.get(tokenId);if(rows?.get(key)!==draft)return;rows.delete(key);if(!rows.size)tokens.delete(tokenId);},
  clear(tokenId){tokens.delete(tokenId);}
 };
}
export function createTokenWriteQueue() {
 const pending=new Map();return (tokenId,write)=>{
  const task=(pending.get(tokenId)||Promise.resolve()).catch(()=>{}).then(write);pending.set(tokenId,task);
  const cleanup=()=>{if(pending.get(tokenId)===task)pending.delete(tokenId);};task.then(cleanup,cleanup);return task;
 };
}
