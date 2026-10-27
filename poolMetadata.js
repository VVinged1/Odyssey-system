export function createPoolMetadata({getMetadata,onMetadataChange,onChange}) {
  let revision=0,request=0;
  onMetadataChange(metadata=>{revision++;onChange(metadata||{});});
  return async()=>{const current=++request,before=revision;const metadata=await getMetadata();if(current===request&&before===revision)onChange(metadata||{});return metadata;};
}
