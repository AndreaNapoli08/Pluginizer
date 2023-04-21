import * as React from 'react';
import { useState, useEffect } from 'react';
import IconButton from '@mui/material/IconButton';

export const TipografiaButton = ({setDis, onFirstOccurence, onButtonStyle, expandedText}) => {
 
  const isLetterOrNumber = (char) => {
    if (typeof char === "undefined") {
      return false;
    }else{
      return /^[a-zA-Z0-9]+$/.test(char);
    }
  }

  const toggleFontStyle = async (style) => {
    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      
      selection.load("text, font, paragraphs");
      await context.sync();
      if (selection.isNullObject) {
        return;
      }
      
      let paragraphCount = selection.paragraphs.items.length; // conta il numero di paragrafi all'interno della selezione
    
      // Expand to end of sentence
      if(expandedText != selection.text){
        const startIndex = expandedText.indexOf(selection.text);
        const charBefore = expandedText[startIndex - 1];
        let text = selection.text;
        let spaceCount = text.split(" ").length; // conta il numero di parole nella selezione
  
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
          let textBeforeSelection = rangeToSelect.getTextRanges([" ", ".", ",", ";"]);
          textBeforeSelection.load("items");
          await context.sync();
          let lastItem = textBeforeSelection.items[textBeforeSelection.items.length - spaceCount];
          let rangeToExpand = lastItem.getRange("Start");
          selection = selection.expandToOrNullObject(rangeToExpand);
          await context.sync();
        }

        selection.select();
        selection.load("text, font");
        await context.sync();
      }

      switch (style) {
        case 'bold':
          onFirstOccurence(selection.font.bold);
          selection.font.bold = !selection.font.bold;
          onButtonStyle("bold");
          onButtonStyle(""); // lo setto a "" così quando ci sarà una nuova selezione non rimane salvato bold nella variabile
          break;
        case 'italic':
          onFirstOccurence(selection.font.italic);
          selection.font.italic = !selection.font.italic;
          onButtonStyle("italic");
          onButtonStyle("");
          break;
        case 'underline':
          if (selection.font.underline === "Mixed" || selection.font.underline === "None") {
            onFirstOccurence("None");
            selection.font.underline = "Single";
          } else {  
            onFirstOccurence("Single");
            selection.font.underline = "None";
          }
          onButtonStyle("underline");
          onButtonStyle("");
          break;
        default:
          break;
      }

      await context.sync();
    });
  } 

  return (
    <div style={{marginTop: '20px'}}>
      <div style={{display: "flex", justifyContent: "center", alignItems: "center"}}>
        <IconButton 
          disabled={setDis}
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('bold')}>
            <b>G</b>
        </IconButton>
        <IconButton 
          disabled={setDis} 
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('italic')}>
            <i>I</i>
        </IconButton>
        <IconButton 
          disabled={setDis} 
          color="inherit" 
          style = {{
            marginRight: "10px",
            border: "1px solid black",
            borderRadius: "10px",
            width: "75px",
            height: "40px"
          }}
          onClick={() => toggleFontStyle('underline')}>
            <u>S</u>
        </IconButton>
      </div>
    </div>
  );
}
