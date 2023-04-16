import * as React from 'react';
import { useState } from 'react';
import IconButton from '@mui/material/IconButton';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';

export const FirstStyles = ({onFontStyle, onFirst}) => {

    const updateStyle = async (style) => {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load();
            await context.sync();

            onFirst(selection.styleBuiltIn);
            if (selection.styleBuiltIn === style) {
                selection.styleBuiltIn = "Normal";
            } else {
                selection.styleBuiltIn = style;
            }

            onFontStyle(style);
            onFontStyle("");
        });
    }

  return (
    <div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={() => updateStyle('Hyperlink')}>
                <span style={{fontSize: "18px"}}>Reference</span>
                <LinkIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={() => updateStyle('Heading6')}>
                <span style={{fontSize: "18px"}}>Definition</span>
                <LiveHelpIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
        <div style={{marginBottom:"15px"}}>
            <IconButton color="inherit" style={{borderRadius: '10px'}} onClick={() => updateStyle('IntenseEmphasis')}>
                <span style={{fontSize: "18px"}}>Footnote</span>
                <NoteAltIcon style={{marginLeft: "10px"}} />
            </IconButton>
        </div>
    </div>
  )
}