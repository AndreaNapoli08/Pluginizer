import * as React from 'react';
import { useState } from 'react';
import IconButton from '@mui/material/IconButton';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';

export const FirstStyles = () => {
    const ReferenceStyle = async () => {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load();
            await context.sync();

            if(selection.styleBuiltIn == "Hyperlink"){
                selection.styleBuiltIn = "Normal";
            }else{
                selection.styleBuiltIn = "Hyperlink" // mette il testo selezionato nello stile HyperLink
            }
        });
    }

    const DefinitionStyle = async () => {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load();
            await context.sync();
            
            if(selection.styleBuiltIn == "Heading6"){
                selection.styleBuiltIn = "Normal";
            }else{
                selection.styleBuiltIn = "Heading6" 
            }

        });
    }

    const FootnoteStyle = async () => {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load();
            await context.sync();
            if(selection.styleBuiltIn == "IntenseEmphasis"){
                selection.styleBuiltIn = "Normal";
            }else{
                selection.styleBuiltIn = "IntenseEmphasis"
            }
        });
    }

  return (
    <div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={ReferenceStyle}>
                <span style={{fontSize: "18px"}}>Reference</span>
                <LinkIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={DefinitionStyle}>
                <span style={{fontSize: "18px"}}>Definition</span>
                <LiveHelpIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={FootnoteStyle}>
                <span style={{fontSize: "18px"}}>Footnote</span>
                <NoteAltIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
    </div>
  )
}