import {createTheme,ThemeProvider} from '@mui/material/styles';import React from 'react';
const theme=createTheme({palette:{mode:'dark',primary:{main:'#76b8ff'},background:{default:'#081322',paper:'#111f33'},text:{primary:'#e5f0ff',secondary:'#9fb8d5'},divider:'#2c4567'},shape:{borderRadius:8},typography:{fontFamily:'Segoe UI, system-ui, sans-serif'},components:{MuiButtonBase:{defaultProps:{disableRipple:true}}}});
export function PluginThemeProvider({children}:{children?:React.ReactNode}){return <ThemeProvider theme={theme}>{children}</ThemeProvider>;}
