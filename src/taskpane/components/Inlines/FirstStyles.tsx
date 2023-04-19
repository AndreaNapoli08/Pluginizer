import * as React from 'react';
import { useState } from 'react';
import IconButton from '@mui/material/IconButton';
import LinkIcon from '@mui/icons-material/Link';
import LiveHelpIcon from '@mui/icons-material/LiveHelp';
import NoteAltIcon from '@mui/icons-material/NoteAlt';

export const FirstStyles = ({onFontStyle, onFirst, expandedText}) => {
    const isLetterOrNumber = (char) => {
        if (typeof char === "undefined") {
          return false;
        }else{
          return /^[a-zA-Z0-9]+$/.test(char);
        }
    }

    const updateStyle = async (style) => {
        await Word.run(async (context) => {
            let selection = context.document.getSelection();
            selection.load();
            await context.sync();

            if(expandedText != selection.text){
                const startIndex = expandedText.indexOf(selection.text);
                const charBefore = expandedText[startIndex - 1];
                let text = selection.text;
                let spaceCount = text.split(" ").length;
               
                //selezione in avanti fino ad uno di quei caratteri
                let rngNextSent = selection.getNextTextRangeOrNullObject([".", " ", ",","!"]);
                selection = selection.expandToOrNullObject(rngNextSent.getRange("Start"));
                await context.sync();
                
                // selezione all'indietro   
                if(isLetterOrNumber(charBefore)){
                  let paragraph = selection.paragraphs.getFirst();
                  paragraph.load("text");
                  await context.sync();
        
                  let rangeToSelect = paragraph.getRange("Start").expandTo(selection);
                  let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"], false);
                  textBeforeSelection.load("items");
                  await context.sync();
                  let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
                  let rangeToExpand = lastItem.getRange("Start");
                  selection = selection.expandToOrNullObject(rangeToExpand);
                  await context.sync();
                }
                
                selection.select();
                selection.load();
                await context.sync();
            }

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