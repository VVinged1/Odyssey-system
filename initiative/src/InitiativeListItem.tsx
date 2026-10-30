import ListItem from '@mui/material/ListItem';import Input from '@mui/material/Input';import Avatar from '@mui/material/Avatar';import Tooltip from '@mui/material/Tooltip';
import VisibilityOffRounded from '@mui/icons-material/VisibilityOffRounded';
import {InitiativeItem} from './InitiativeItem';import {focusInitiativeItem} from './focusInitiativeItem';
type Props={initiative:InitiativeItem;onCountChange:(count:string)=>void;onRemove:()=>void;showHidden:boolean;editable:boolean};
export function InitiativeListItem({initiative,onCountChange,onRemove,showHidden,editable}:Props){
 if(!initiative.visible&&!showHidden)return null;
 return <ListItem className="initiative-card" data-initiative-id={initiative.id} data-active={initiative.active} selected={initiative.active}
  sx={{width:'calc(var(--initiative-avatar-size) + 20px)',flex:'0 0 auto',height:'calc(var(--initiative-avatar-size) + 38px)',p:.5,display:'flex',flexDirection:'column',alignItems:'center',border:'1px solid',borderColor:initiative.active?'primary.main':'divider',borderRadius:'8px',backgroundColor:'background.paper'}}>
  {editable&&<Tooltip disableInteractive title={'Remove '+initiative.name+' from initiative'} arrow><button type="button" className="initiative-remove" data-initiative-remove={initiative.id} aria-label={'Remove '+initiative.name+' from initiative'} onClick={event=>{event.stopPropagation();onRemove();}}>×</button></Tooltip>}
  <Tooltip disableInteractive title={initiative.name} arrow enterDelay={150}><button type="button" className="initiative-token" aria-label={'Focus '+initiative.name} onClick={()=>void focusInitiativeItem(initiative.id)}>
   <Avatar src={initiative.imageUrl} alt="" variant="rounded" sx={{width:'var(--initiative-avatar-size)',height:'var(--initiative-avatar-size)',fontSize:16,color:'#e5f0ff',bgcolor:'action.hover','& img':{objectFit:'contain'}}}>{initiative.name.slice(0,2)}</Avatar>
   {!initiative.visible&&<VisibilityOffRounded sx={{position:'absolute',right:1,bottom:1,fontSize:14,bgcolor:'background.paper',borderRadius:1}}/>}
  </button></Tooltip>
  <Input disableUnderline disabled={!editable} value={initiative.count} inputProps={{'aria-label':'Initiative for '+initiative.name,inputMode:'decimal',sx:{textAlign:'center',fontSize:19,fontWeight:700,p:0}}}
    sx={{width:'100%',mt:'auto','& .MuiInputBase-input.Mui-disabled':{WebkitTextFillColor:'#e5f0ff',opacity:1}}} onChange={event=>onCountChange(event.target.value)}/>
 </ListItem>;
}
