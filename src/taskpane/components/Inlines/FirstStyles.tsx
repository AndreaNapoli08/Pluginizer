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
            selection.load("paragraphs, text, styleBuiltIn");
            await context.sync();
            let paragraphCount = selection.paragraphs.items.length; 

            if(expandedText != selection.text){
                const startIndex = expandedText.indexOf(selection.text);
                const charBefore = expandedText[startIndex - 1];
                let text = selection.text;
                let spaceCount = text.split(" ").length;
                //selezione in avanti fino ad uno di quei caratteri
                const nextCharRanges = selection.getTextRanges([" ", ".", ",", ";", "!", "?", ":", "\n", "\r"], true);
                nextCharRanges.load("items");
                await context.sync();
                
                if (nextCharRanges.items.length > 0) {
                    if(paragraphCount>1){ // se più paragraphi sono compresi, andare a capo lo prende come una parola e quindi spaceCount va incrementato con il numero di paragrafi -1
                        spaceCount = spaceCount + paragraphCount - 1;
                    }
                    for(let i = 0; i < spaceCount; i++){
                        selection = selection.expandTo(nextCharRanges.items[i]);
                    }
                }

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
                selection.load("styleBuiltIn");
                await context.sync();
            }

            onFirst(selection.styleBuiltIn); // da cambiare per tutte le occorrenze
            
            if (selection.styleBuiltIn === style || selection.styleBuiltIn == "Other") {
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