import React from 'react';import Box from '@mui/material/Box';
export function InitiativeHeader({subtitle,action}:{subtitle?:string;action?:React.ReactNode}){return <Box component="nav" aria-label={subtitle||'Initiative controls'} sx={{flexShrink:0,p:.5}}>{action}</Box>;}
